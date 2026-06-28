# Column Reference

This document is the working contract for sheet columns used by the MAIN and SUB flows. It explains which sheets provide data, which sheets receive data, and which columns drive routing, enrichment, validation, and optional flags.

Use this as a debugging and revision index, not as a second configuration layer. The canonical runtime contract still lives in `00_Config.gs`, `03_SheetsAndValidation.gs`, and the writer modules listed below.

## Contract Ownership

| Area | Owner file | Purpose |
| --- | --- | --- |
| Source headers, aliases, sheet names, routing policy, optional sheet policy, type registry | `00_Config.gs` | Defines the canonical names and business rules that the pipeline expects. |
| Destination sheet templates and Raw Data header assurance | `03_SheetsAndValidation.gs` | Defines the exact destination layouts that the workbook should contain. |
| Runtime mapping from Raw rows to operational sheets | `05b_Pipeline_RoutingOperational.gs` | Maps source values into destination columns. |
| Optional sheet writers | `05c_Pipeline_OptionalSheets.gs` | Builds B2B, Special Case, EV-Bike, and related optional outputs. |
| SUB ingestion and operational updates | `06a_EntryPoints.gs`, `06c_PostProcessAndUtils.gs` | Reads Raw OLD / Raw NEW and updates Submission plus operational sheets. |

When a column is renamed, added, or removed, update `00_Config.gs`, `03_SheetsAndValidation.gs`, the affected writer, and this file together.

## Flow Map

| Flow | Source sheet | Main destination sheets | Notes |
| --- | --- | --- | --- |
| MAIN | `Raw Data` | Operational sheets, optional sheets, logs/details, overview outputs | Uses canonical raw headers from `CONFIG.headers` and destination templates from `SV03_TEMPLATES`. |
| SUB | `Raw OLD`, `Raw NEW` | `Submission`, `Ask Detail`, `OR - OLD`, `SC - Farhan`, `SC - Meilani`, `SC - Meindar`, `Start`, `Finish`, `PO`, `Exclusion`, `B2B`, `EV-Bike`, `Special Case` | Uses `SUB_FLOW_SPEC` and raw row builders in `06a_EntryPoints.gs`. |

## How To Read Column Sources

Use these categories when debugging a blank or wrong value:

| Source type | Meaning | Debug first |
| --- | --- | --- |
| Raw-driven | Copied or normalized from `Raw Data`, `Raw OLD`, or `Raw NEW`. | Check the raw header name, alias support, and `CONFIG.headers`. |
| Derived | Calculated by workflow functions from one or more raw fields. | Check the function named in this document and the input columns it needs. |
| Manual/restored | Entered by the user in an operational or optional sheet, then preserved across clear/rebuild runs. | Check the pre-clear snapshot, Raw backup, hidden backup sheets, and restore audit logs. |
| Template/validation | Exists because the sheet layout, dropdown, checkbox, or formatting requires it. | Check `SV03_TEMPLATES`, dropdown sync, and template row copy behavior. |

If a column is not present in Raw Data or Raw NEW, do not assume it is missing data. Some columns intentionally come from sheet state, runtime derivation, hidden backup sheets, or template validation. Spreadsheet workflows, naturally, enjoy making "source of truth" a group project.

## Manual And Restored Columns

Operational manual fields are user-managed state. They are not ordinary raw-source columns.

| Column | Source type | Source / restore path | Notes |
| --- | --- | --- | --- |
| `Update Status` | Manual/restored | User input in operational sheets; backed up from ops to Raw when enabled; restored by claim after route. | Rich text is preserved where possible. |
| `Timestamp` | Manual/restored | User or workflow timestamp in operational sheets; backed up/restored by claim. | Number format is preserved by rich snapshot/temp backup paths. |
| `Status` | Manual/restored / template | User dropdown in operational sheets and optional sheets. | Writers intentionally avoid writing this column directly in some flush paths to avoid dropdown validation traps. |
| `Remarks` | Manual/restored | User notes in operational sheets; persisted into `Raw Data.Remarks`; restored by claim after route. | Rich text and wrap are restored where possible. |
| `Update Status Asso`, `Timestamp Asso`, `Update Status Admin`, `Timestamp Admin` | Manual/restored | Raw Data tail columns and admin/associate workflow state. | Written back as derived/raw tail columns when present. |
| `Start Date`, `End Date`, `Details` | Manual/restored or derived by sheet | Operational/B2B tracking fields; Special Case can fill policy dates/details when present. | EV-Bike removes these legacy columns if they still exist. |
| `OR` | Manual/restored / template | Checkbox/manual field in PO and Special Case layouts; can also be derived from raw own-risk markers. | Do not treat it as a universal raw field. |

Manual restore pipeline:

1. MAIN ensures Raw headers and tail columns exist, then builds `rawValues`.
2. `backupRemarksOpsToRawInMemory06b_` copies non-empty operational `Remarks` into in-memory Raw values.
3. If `PIPELINE_FLAGS.ENABLE_BACKUP_FROM_OPS` is enabled, `backupOpsToRawFull_` backs up operational manual fields into Raw with richer formatting support.
4. Before clearing operational sheets, `snapshotOpsManualColumnsRich06c_` captures `Update Status`, `Timestamp`, `Status`, and `Remarks` by `Claim Number`.
5. MAIN also writes `_OPS_MAIN_SUB_TEMP` through `persistOpsManualTempForSub06c_`. SUB uses this one-shot backup to restore by `Claim Number + Service Center`.
6. `persistOpsManualBackupSheet06c_` writes `_OPS_MANUAL_BACKUP` as a hidden fallback backup by `PIC + Claim Number`.
7. `clearOperationalSheets_` snapshots the same four manual columns per target sheet, clears operational data, then route writers rebuild rows.
8. During flush, `routeRawToOperationalSheetsInMemory_` restores blank manual cells from the pre-clear sheet snapshot by `Claim Number`.
9. After route, `restoreOpsFieldsFromRawBackup_`, `applyUpdateStatusRichTextToOperational_`, and `applyRemarksRichTextToOperational_` restore value and rich-text formatting from Raw backup paths.
10. `restoreOpsManualColumnsRich06c_` reapplies rich text, wrap, timestamp formats, and dropdown-sensitive values from the pre-clear snapshot.
11. `auditOpsManualRestore06c_` checks for restore gaps. If gaps remain, `restoreOpsManualFromBackupSheet06c_` attempts fallback restore from `_OPS_MANUAL_BACKUP`.

SUB restore behavior:

| Backup sheet | Written by | Restored by | Match key | Purpose |
| --- | --- | --- | --- | --- |
| `_OPS_MAIN_SUB_TEMP` | `persistOpsManualTempForSub06c_` during MAIN | `restoreOpsManualFromMainTempForSub06c_` during SUB | `Claim Number + Service Center` | Preserves manual fields when SUB relocates claims between operational sheets. |
| `_OPS_MANUAL_BACKUP` | `persistOpsManualBackupSheet06c_` | `restoreOpsManualFromBackupSheet06c_` | `PIC + Claim Number` | Fallback when normal rich snapshot restore misses manual fields. |

Important restore rule: restore only fills blank destination manual cells. If the rebuilt row already has a value, the existing destination value wins.

## Header Normalization And Aliases

The pipeline normalizes headers before lookup. Prefer canonical headers, but these legacy aliases are still recognized in several flows.

| Legacy / display header | Canonical meaning |
| --- | --- |
| `LSA` | `last_status_aging` / `Last Status Aging` |
| `ALA` | `activity_log_aging` / `Activity Log Aging` |
| `sum_insured` | `sum_insured_amount` |
| `repair_replace_amount` | `claim_amount` |
| `or_amount` | `claim_own_risk_amount` |
| `service_center`, `Service Center Name` | `sc_name` / `Service Center` |
| `db_link`, `DB Link` | `dashboard_link` |

## MAIN Raw Data 2026 Rename Map

Runtime keeps backward-compatible aliases, but MAIN should prefer these Raw Data source headers.

| Legacy source | Current source |
| --- | --- |
| `policy_start_date` | `policy_start_datetime` |
| `policy_end_date` | `policy_end_datetime` |
| `activation_datetime` | `policy_issued_datetime` |
| `claim_submission_date` | `claim_submitted_datetime` |
| `claim_submission_months` | `claim_submitted_month` |
| `sc_name` | `repairer_location_store_name` |
| `sc_city` | `repairer_location_city_name` |
| `last_status` | `claim_last_status_name` |
| `last_status_aging` | `days_aging_from_last_activity` |
| `repair_replace_done_date` | `repair_done_datetime` / `replace_done_datetime` |
| `device_checkout_date` | `device_checkout_datetime` |
| `last_update` | `last_update_datetime` |
| `last_activity_log_date` | `last_activity_log_datetime` |
| `last_activity_log` | `last_activity_log_name` |
| `qoala_pic_name` | `pic_name` |

## Source Sheets

### Raw Data

`Raw Data` is the MAIN source sheet. It should contain the canonical raw headers in `CONFIG.headers`, plus managed custom tail columns.

Core identity and policy columns:

| Column | Used for |
| --- | --- |
| `claim_number` | Primary claim key, duplicate checks, DB classification, Flex/B2B token detection, destination `Claim Number`. |
| `business_partner_name` | Destination partner name, B2B partner detection, EV-Bike partner detection. |
| `insurance_partner_name` | Destination `Insurance`. |
| `insurance_partner_code`, `insurance_code` | Insurance fallback and SUB submission fields. |
| `product_name` | Destination `Product`, Flex product detection, Special Case context. |
| `source_system_name` | Duplicate source comparison and DB/source classification. |
| `dashboard_link` | Destination `DB Link`; when blank, runtime derives the internal claim URL from `claim_number`. |
| `DB` | Destination DB value when already available. |

Claim date and aging columns:

| Column | Used for |
| --- | --- |
| `claim_submitted_datetime` | Primary `Submission Date`, B2B/Special Case/EV-Bike output date, first-month and policy-age comparisons. Legacy `claim_submission_date` remains an alias. |
| `claim_submitted_month` | Destination `Submission by Month` when present; otherwise derived from `claim_submitted_datetime`. |
| `claim_last_updated_datetime`, `last_update_datetime` | Destination `Last Status Date` / status recency fields. Legacy `last_update` remains an alias. |
| `last_activity_log_datetime`, `activity_log_datetime` | Activity log recency and destination datetime fields. |
| `month_policy_aging` | Strict Second-Year (Market Value) flag. The claim is flagged only when this Raw Data value is greater than `12`. |
| `days_to_end_policies` | Policy remaining calculation and Special Case flagging. |
| `days_aging_from_submission` | Destination `TAT`, Exclusion TAT fallback, aging output. |
| `days_aging_from_last_activity` / `LSA` | Destination `Last Status Aging`; sort and operational monitoring. Legacy `last_status_aging` remains an alias. |
| `activity_log_aging` / `ALA` | Destination `Activity Log Aging`. |
| `Q-L (Months)` | Special Case and optional output context. |
| `M-L (Months)` | Managed Raw Data tail metric. |
| `M-Q (Months)` | First-month policy and Special Case calculations. |

Status and routing columns:

| Column | Used for |
| --- | --- |
| `claim_last_status_name` | Operational routing, `Status Type`, `Position`, `Submission` trigger, destination `Last Status`. Legacy `last_status` remains an alias. |
| `activity_log`, `last_activity_log_name` | Destination `Activity Log`. |
| `repairer_location_store_name` | Destination `Service Center`, SC sheet routing, Branch derivation, Service Center flagging. Legacy `sc_name` remains an alias. |
| `repairer_location_city_name` | Service-center city reference/debug field. Legacy `sc_city` remains an alias. |
| `id_business_partner_category_name` | Destination `Buss. Category`. |
| `pm_name` | Destination `PM Name`. |
| `apm_name` | Destination `APM Name`. |
| `device_checkin_option_name` | Destination `Service Type` on `Start`. |
| `device_checkout_option_name` | Destination `Service Type` on `Finish`. |
| `Associate` | Admin/operator assignment when the destination profile supports it. |
| `Update Status`, `Timestamp`, `Status`, `Remarks` | Manual or workflow status tracking. Preserved where possible. |

Device and customer columns:

| Column | Used for |
| --- | --- |
| `device_type` | Destination `Device Type`; forbidden in Admin destination templates. |
| `device_brand` | Destination `Device Brand`. |
| `imei_number`, `device_imei` | Destination `IMEI/SN`. |
| `customer_name` | EV-Bike `Owner Name`. |
| `qoala_policy_number` | Duplicate detection and EV-Bike `Policy Number`. |
| `outlet_name` variants | Destination `Store Name` when available. |
| `pa_name` variants | Destination `PA Name` when available. |
| `spa_name` variants | Destination `SPA Name` when available. |

Financial columns:

| Column | Used for |
| --- | --- |
| `sum_insured_amount` | Destination `Sum Insured` / `Sum Insured Amount`, EV-Bike `Sum Insured`, Special Case context. |
| `claim_amount` | Destination `Claim Amount` / `Repair/Replace Amount`. |
| `claim_own_risk_amount` | Destination `Claim Own Risk Amount` / `OR Amount`. |
| `nett_claim_amount` | Destination `Nett Claim Amount`. |
| `OR` | Destination checkbox/manual own-risk marker, only expected in PO and Special Case layouts. |

Managed Raw Data tail columns:

| Column | Purpose |
| --- | --- |
| `Update Status` | Manual workflow status. |
| `Timestamp` | Manual or automated update timestamp. |
| `Status` | Manual status value. |
| `Remarks` | Manual note. |
| `Q-L (Months)` | Claim-to-last-status metric. |
| `M-L (Months)` | Policy-month-to-last-status metric. |
| `M-Q (Months)` | Policy-month-to-claim metric. |
| `Update Status Asso` | Associate workflow status. |
| `Timestamp Asso` | Associate workflow timestamp. |
| `Update Status Admin` | Admin workflow status. |
| `Timestamp Admin` | Admin workflow timestamp. |

### Raw OLD And Raw NEW

`Raw OLD` and `Raw NEW` are SUB source sheets. They feed the `Submission` sheet and update operational sheets.

Required SUB raw columns:

| Column | Used for |
| --- | --- |
| `claim_number` | Primary key and destination `Claim Number`. |
| `claim_submitted_datetime` | Submission date fallback. |
| `dashboard_link` | Destination `DB Link`. |
| `partner_name` | Destination `Partner Name`. |
| `insurance_partner_code`, `insurance_code` | Destination `Insurance`. |
| `device_type` | Destination `Device Type`. |
| `device_imei` | Destination `IMEI/SN`. |
| `last_status` | Submission trigger and operational routing. |
| `last_status_aging` / `LSA` | Destination `Last Status Aging`; sort key. |
| `activity_log_aging` / `ALA` | Destination `Activity Log Aging`. |
| `sc_name`, `service_center` | Destination `Service Center`; SC routing. |
| `days_aging_from_submission` | Destination `TAT`. |
| `claim_last_updated_datetime` | Destination `Last Status Date`. |
| `activity_log`, `last_activity_log` | Destination `Activity Log`. |
| `activity_log_datetime` | Activity log datetime context. |
| outlet / PA / SPA variants | Destination `Store Name`, `PA Name`, `SPA Name` when available. |

SUB submission triggers:

| Source | Trigger | DB value |
| --- | --- | --- |
| `Raw OLD` | `last_status = SUBMITTED` | `OLD` |
| `Raw NEW` | `last_status = CLAIM_INITIATE` | `NEW` |

## Destination Sheets

### Submission

`Submission` is built from `Raw OLD` and `Raw NEW` in SUB, using `SUB_FLOW_SPEC.OP_HEADERS`.

| Submission column | Source |
| --- | --- |
| `Claim Number` | `claim_number` |
| `Last Status Aging` | `last_status_aging` / `LSA` |
| `Activity Log Aging` | `activity_log_aging` / `ALA` |
| `Last Status` | `last_status` |
| `Service Center` | `sc_name` / `service_center` |
| `Submission Date` | `claim_submitted_datetime` or `claim_submission_date` fallback |
| `DB` | `OLD` or `NEW` from source raw sheet and trigger rule |
| `DB Link` | `dashboard_link` / `db_link` |
| `Partner Name` | `partner_name` |
| `Insurance` | `insurance_partner_code` / `insurance_code` |
| `Device Type` | `device_type` |
| `IMEI/SN` | `device_imei` |
| `TAT` | `days_aging_from_submission` |

Default SUB sort keys are `Last Status Aging` descending, `Last Status` ascending, then `DB` ascending.

### Operational Sheets

Operational sheets include `Ask Detail`, `OR - OLD`, `SC - Farhan`, `SC - Meilani`, `SC - Meindar`, `Start`, `Finish`, `PO`, and profile-specific worklists.

Common PIC destination columns:

| Destination column | Source type | Source / rule |
| --- | --- | --- |
| `Submission Date` | Raw-driven | `claim_submitted_datetime`; legacy fallback `claim_submission_date`. |
| `Submission Datetime` | Raw-driven | `claim_submitted_datetime`; legacy fallback `claim_submission_date`. |
| `Submission by Month` | Derived/raw-driven | `claim_submitted_month` when present; otherwise derived from `Submission Date`. |
| `Claim Number` | Raw-driven | `claim_number`. |
| `DB Link` | Raw-driven/derived | `dashboard_link`; when blank, built from `Claim Number`. |
| `DB` | Derived/raw-driven | Claim-token classification first; raw `DB` / `source_system_name` fallback. |
| `Partner Name`, `Partner` | Raw-driven | `business_partner_name`. |
| `Insurance` | Raw-driven/normalized | `insurance_partner_name` or insurance code fallback, shortened by insurance mapper. |
| `Device Type` | Raw-driven | `device_type`. |
| `Product` | Raw-driven | `product_name`. |
| `Device Brand` | Raw-driven | `device_brand`. |
| `IMEI/SN` | Raw-driven | `imei_number`, `device_imei`, or serial number variants. |
| `Last Status` | Raw-driven | `claim_last_status_name`; legacy fallback `last_status`. |
| `Last Status Date` / `Last Status Datetime` | Raw-driven | `claim_last_updated_datetime`, `last_update_datetime`, or last-activity datetime fallback. |
| `Service Center` / `Service Center Name` | Raw-driven | `repairer_location_store_name`; legacy fallback `sc_name` / service center variants. |
| `Last Status Aging` / `LSA` | Raw-driven | `days_aging_from_last_activity`; legacy fallback `last_status_aging` / `LSA`. |
| `Activity Log` | Raw-driven | `activity_log` or `last_activity_log_name`. |
| `Activity Log Datetime` | Raw-driven | `activity_log_datetime` or `last_activity_log_datetime`. |
| `Activity Log Aging` / `ALA` | Raw-driven | `activity_log_aging` / `ALA`. |
| `TAT` | Raw-driven | `days_aging_from_submission`. |
| `Start Date`, `End Date`, `Details` | Manual/restored | Manual operational tracking fields, unless a specific optional writer fills them. |
| `Sum Insured Amount`, `Sum Insured` | Raw-driven | `sum_insured_amount` and aliases. |
| `Claim Amount` / `Repair/Replace Amount` | Raw-driven | `claim_amount`; may fall back to `nett_claim_amount` in some writers. |
| `Claim Own Risk Amount` / `OR Amount` | Raw-driven | `claim_own_risk_amount`, `or_amount`, or own-risk aliases. |
| `Nett Claim Amount` | Raw-driven | `nett_claim_amount`. |
| `Selisih` | Derived | `Sum Insured Amount - Claim Amount`. |
| `% Approval` | Derived | `Claim Amount / Sum Insured Amount` when both numeric and sum insured is not zero. |
| `Update Status`, `Timestamp`, `Status`, `Remarks` | Manual/restored | Preserved through the manual restore pipeline above. |
| `Status Type` | Derived | Derived from `Last Status` via status type maps. |
| `Buss. Category` | Raw-driven | `id_business_partner_category_name`. |
| `PM Name` | Raw-driven | `pm_name`. |
| `APM Name` | Raw-driven | `apm_name`. |
| `Aging Post.` | Raw-driven by sheet | `Aging Start` on `Start`, `Aging SC Receive` on SC sheets, `Aging Ins Approve` on `PO`, `Aging Finish` on `Finish`. |
| `Service Type` | Raw-driven by sheet | `device_checkin_option_name` on `Start`; `device_checkout_option_name` on `Finish`. |

SC destination sheets also include:

| Destination column | Source type | Source / rule |
| --- | --- | --- |
| `Type` | Derived/template | Derived from `Last Status` through `OPS_ROUTING_POLICY.TYPE_BY_LAST_STATUS`; written only when the column exists. |
| `Branch` | Derived | Derived from `Service Center` using branch keyword mapping in `06c_PostProcessAndUtils.gs`. |

PO destination sheets also include:

| Destination column | Source type | Source / rule |
| --- | --- | --- |
| `OR` | Manual/restored / template | Own-risk checkbox/manual field. |

Admin templates intentionally exclude operational PIC-only columns such as `Device Type`, `Last Status Aging`, `Activity Log Aging`, `TAT`, `OR`, and `OR Amount`.

Operational routing and reset rules:

| Function | Role |
| --- | --- |
| `clearOperationalSheets_` | Snapshots manual fields, clears operational sheets, preserves template row formatting/data validation, and clears bad `Submission Date` validations. |
| `routeRawToOperationalSheetsInMemory_` | Routes each Raw row by `Last Status`, service-center keyword split, and fallback `SC - Unmapped`. |
| `buildSheetWriters_` | Converts one Raw row into each destination sheet row using only columns that exist in that sheet. |
| `applyOperationalClaimHighlightsByRaw_` | Applies claim-cell highlights and notes from Raw-driven flags after template formatting. |
| `enrichOperationalSheetsFromRaw06_` | Fills post-route enrichment such as `Activity Log`, `Last Status Date`, `Status Type`, `LSA`, `ALA`, and `TAT`. |
| `applyStrictSubmissionDateAndMonth06b_` | Re-syncs `Submission Date` and `Submission by Month` after optional processors to avoid stale values. |

### B2B

`B2B` is generated by `processB2B_`.

Detection rule:

| Signal | Meaning |
| --- | --- |
| Partner name matches configured B2B partner patterns | Row is a B2B candidate. |
| `Claim Number` contains `SMR` | Row is a B2B claim candidate even if partner matching is incomplete. |

Output columns:

| B2B column | Source type | Source / rule |
| --- | --- | --- |
| `Submission Date` | Raw-driven | `claim_submitted_datetime`; legacy fallback `claim_submission_date`. |
| `Claim Number` | Raw-driven | `claim_number`. |
| `DB Link` | Raw-driven/derived | `dashboard_link`; fallback built from `Claim Number`. |
| `DB` | Derived/raw-driven | `source_system_name` or DB classification. |
| `Partner Name` | Raw-driven | `business_partner_name`. |
| `Insurance` | Raw-driven/normalized | `insurance_partner_name` / insurance code fallback. |
| `Device Type` | Raw-driven | `device_type`. |
| `Product` | Raw-driven | `product_name`. |
| `Device Brand` | Raw-driven | `device_brand`. |
| `IMEI/SN` | Raw-driven | `imei_number` / `device_imei`. |
| `Last Status` | Raw-driven | `claim_last_status_name`; legacy fallback `last_status`. |
| `Service Center` | Raw-driven | `repairer_location_store_name`; legacy fallback `sc_name`. |
| `Last Status Aging` / `LSA` | Raw-driven | `days_aging_from_last_activity`; legacy fallback `last_status_aging` / `LSA`. |
| `Activity Log Aging` / `ALA` | Raw-driven | `activity_log_aging` / `ALA`. |
| `TAT` | Raw-driven | `days_aging_from_submission`. |
| `Start Date`, `End Date`, `Details` | Manual/restored | Manual tracking fields; schema can be auto-healed if missing. |
| `Sum Insured Amount` | Raw-driven | `sum_insured_amount`. |
| `Claim Amount` | Raw-driven | Claim amount / nett amount fallback depending on writer path. |
| `Claim Own Risk Amount` | Raw-driven | `claim_own_risk_amount`. |
| `Nett Claim Amount` | Raw-driven | `nett_claim_amount`. |
| `% Approval` | Derived | Approval ratio when numeric source values support it. |
| `Status Type` | Derived | Added only when the B2B schema contains `Last Status`. |

B2B can also include fallback candidates from `Submission` when a B2B claim is not present in the current Raw window. That fallback exists so short Raw extracts do not accidentally drop active B2B work.

### Special Case

`Special Case` is generated by `processSpecialCase_`. The sheet is treated as a fixed, user-managed schema; the script should not auto-add arbitrary columns.

Output columns:

| Special Case column | Source type | Source / rule |
| --- | --- | --- |
| `Submission Date` | Raw-driven | `claim_submitted_datetime`; Special Case uses this as the source of truth. |
| `Claim Number` | Raw-driven | `claim_number`. |
| `DB Link` | Raw-driven/derived | `dashboard_link`; fallback built from `Claim Number`. |
| `DB` | Derived/raw-driven | DB classification from claim token, then source value fallback. |
| `Partner Name` | Raw-driven | `business_partner_name`. |
| `Insurance` | Raw-driven/normalized | `insurance_partner_name` / insurance code fallback. |
| `Device Type` | Raw-driven | `device_type`. |
| `Last Status` | Raw-driven | `claim_last_status_name`; legacy fallback `last_status`. |
| `Service Center` | Raw-driven | `repairer_location_store_name`; legacy fallback `sc_name`. |
| `Last Status Aging` / `LSA` | Raw-driven | `days_aging_from_last_activity`; legacy fallback `last_status_aging` / `LSA`. |
| `Activity Log` | Raw-driven | `activity_log` / `last_activity_log_name`. |
| `Activity Log Aging` / `ALA` | Raw-driven | `activity_log_aging` / `ALA`. |
| `TAT` | Raw-driven | `days_aging_from_submission`. |
| `Last Status Date` | Raw-driven | `claim_last_updated_datetime`. |
| `Q-L (Months)` | Raw-driven/derived tail | Raw tail metric when present. |
| `Product` | Raw-driven | `product_name`. |
| `Sum Insured Amount`, `Sum Insured` | Raw-driven | `sum_insured_amount`. |
| `Claim Amount` / `Repair/Replace Amount` | Raw-driven | Uses nett claim amount as claim amount output in current writer path. |
| `OR` | Template/manual/derived | Own-risk checkbox output when present. |
| `Claim Own Risk Amount` / `OR Amount` | Raw-driven | `claim_own_risk_amount`. |
| `Nett Claim Amount` | Raw-driven | `nett_claim_amount`. |
| `Selisih` | Derived | `Sum Insured Amount - Claim Amount`. |
| `Reason` | Derived | Joined flags: `Flex`, `Second-Year (Market Value)`, `First-Month Policy`, `Policy Remaining <= 1 Month`. |
| `Start Date`, `End Date` | Raw-driven when present | Policy start/end dates used for context and details. |
| `Details` | Derived when present | Human-readable reason detail; written as a column value and/or claim note depending on layout. |
| `Status Type` | Derived | Added only when the schema contains `Last Status`. |

Special Case flags:

| Flag | Function / policy | Based on |
| --- | --- | --- |
| `Flex` | `processSpecialCase_`, `buildOperationalClaimHighlightSetsFromRaw_` | Partner matches configured special/Flex partner patterns, `Claim Number` contains `SFX`, or `product_name` contains `flex`. |
| `Second-Year (Market Value)` | `SECOND_YEAR_MARKET_VALUE_POLICY`, `processSpecialCase_`, `buildOperationalClaimHighlightSetsFromRaw_` | Strictly `Raw Data.month_policy_aging > 12`. Other month metrics must not be used for this flag. |
| `First-Month Policy` | `processSpecialCase_`, `buildOperationalClaimHighlightSetsFromRaw_` | `claim_submitted_datetime - policy_start_datetime` is non-negative and less than `30` days. |
| `Policy Remaining <= 1 Month` | `processSpecialCase_`, `buildOperationalClaimHighlightSetsFromRaw_` | `policy_end_datetime - claim_submitted_datetime` is non-negative and less than `30` days. |

Special Case write policy:

| Behavior | Rule |
| --- | --- |
| Fixed schema | The sheet is user-managed; missing arbitrary columns are not auto-added. |
| UPSERT mode | Existing claims are updated instead of duplicating rows. |
| Excluded status pruning | Claims that become done/closed/excluded can be removed when configured. |
| Notes/highlights | Flex can color/note Special Case claim cells; policy-age notes/highlights are primarily applied to operational sheets. |

### EV-Bike

`EV-Bike` is generated by `processEVBike_`.

Detection rule:

| Signal | Meaning |
| --- | --- |
| Partner name matches configured EV-Bike partner patterns | Row is an EV-Bike candidate. |
| Policy number is in the configured exclusion list | Row is skipped. |

Output columns:

| EV-Bike column | Source type | Source / rule |
| --- | --- | --- |
| `Submission Date` | Raw-driven or Submission fallback | `claim_submitted_datetime`; fallback from `Submission.Submission Date` when Raw does not contain the EV-Bike claim yet. |
| `Claim Number` | Raw-driven or Submission fallback | `claim_number` / `Submission.Claim Number`. |
| `DB Link` | Raw-driven/derived | `dashboard_link`; fallback from `Submission.DB Link` or built from claim number. |
| `Owner Name` | Raw-driven or Submission fallback | `customer_name` or `Submission.Owner Name` / `Customer Name`. |
| `Policy Number` | Raw-driven or Submission fallback | `qoala_policy_number` or Submission policy-number variants. |
| `Partner Name` | Raw-driven or Submission fallback | `business_partner_name` or `Submission.Partner Name`. |
| `Insurance` | Raw-driven/normalized | `insurance_partner_name` / insurance code fallback. |
| `Sum Insured` | Raw-driven or Submission fallback | `sum_insured_amount` or `Submission.Sum Insured`. |
| `TAT` | Derived fallback | Derived from `Submission Date` to current date when needed. |
| `Last Status` | Raw-driven or Submission fallback | Current raw last status or `Submission.Last Status`. |
| `Status` | Manual/restored | Manual EV-Bike status field; preserved and not overwritten when possible. |

Note: the runtime writer currently removes deprecated `Start Date`, `End Date`, and `Details` columns from EV-Bike if they are still present.

EV-Bike has an overlay behavior: Raw Data rows are the main source, but `Submission` can add or refresh EV-Bike rows that are not yet present in Raw. This is intentional because EV-Bike operational visibility should not depend on the current Raw extract window.

### Exclusion

`Exclusion` uses operational claim data and computes TAT-like timing from submission and last-status dates.

| Column / value | Source / rule |
| --- | --- |
| `Claim Number` | Operational row claim key. |
| `Submission Date` | Operational submission date. |
| `Last Status Date` | Operational last-status date. |
| `TAT` | `Last Status Date - Submission Date`, clamped at zero when needed. |

## Derived Classification Rules

### DB Classification

| DB | Based on |
| --- | --- |
| `OLD` | Claim number contains `SFP`, `SFX`, or `SMR`, or source policy marks it as old service. |
| `NEW` | Claim number contains `VVMAR` or `GADLD`, or source policy marks it as new service. |

### B2B Claim

A claim is treated as B2B when either the partner name matches configured B2B partner patterns or the claim number contains `SMR`.

### Special Case And Policy-Age Flags

These flags are used in two places:

- `processSpecialCase_` writes candidate rows into `Special Case`.
- `buildOperationalClaimHighlightSetsFromRaw_` and `applyOperationalClaimHighlightsByRaw_` add operational-sheet claim-cell highlights and notes.

| Flag | Required source columns | Exact rule | Output impact |
| --- | --- | --- | --- |
| `Flex` | `business_partner_name`, `claim_number`, `product_name` | Partner matches configured special/Flex partner pattern, or claim contains `SFX`, or product contains `flex`. | Special Case `Reason`; claim-cell color/note. |
| `Second-Year (Market Value)` | `month_policy_aging` | Strictly `month_policy_aging > 12`. The policy explicitly forbids substituting `Q-L (Months)`, `M-L (Months)`, or other month fields for this flag. | Special Case `Reason`; operational claim-cell note/detail includes submission date, policy start/end, and month policy aging. |
| `First-Month Policy` | `claim_submitted_datetime`, `policy_start_datetime` | Submission date minus policy start date is `>= 0` and `< 30` days. | Special Case `Reason`; operational claim-cell note/detail includes policy age in days. |
| `Policy Remaining <= 1 Month` | `claim_submitted_datetime`, `policy_end_datetime` | Policy end date minus submission date is `>= 0` and `< 30` days. | Special Case `Reason`; operational claim-cell note/detail. |

Important: `Second-Year (Market Value)` is not "12 months or more". Runtime uses `> 12`, so exactly `12` is not flagged. This is intentional per `SECOND_YEAR_MARKET_VALUE_POLICY`.

### Service Center Routing

Service Center routing uses `repairer_location_store_name`; legacy `sc_name` / `Service Center` remains a fallback.

| Result | Based on |
| --- | --- |
| `SC - Farhan` | Service Center name matches Farhan keyword rules. |
| `SC - Meilani` | Service Center name matches Meilani keyword rules. |
| `SC - Meindar` | Service Center name matches Meindar/Ivan keyword rules. |
| `SC - Unmapped` | No Service Center keyword matches. |

`Branch` is also derived from Service Center text using `BRANCH_KEYWORDS`.

### Status Type And Position

`Status Type` and `Position` are derived from `claim_last_status_name` / legacy `last_status` using configured status maps. These values drive monitoring, presentation, and operational grouping.

### Duplicate Claim

Duplicate detection uses:

| Column | Purpose |
| --- | --- |
| `qoala_policy_number` | Main policy key for duplicate comparison. |
| `source_system_name` | Helps separate old/new source behavior. |
| `claim_number` | Claim identity and token classification. |
| `claim_submitted_datetime` | Submission date comparison; legacy `claim_submission_date` remains an alias. |
| `claim_last_status_name` | Status context; legacy `last_status` remains an alias. |

Configured duplicate comparison allows a maximum submission-date difference of `62` days.

### Claim Highlighting

Claim highlighting and notes use configured marker policies:

| Marker | Based on | Priority / note behavior |
| --- | --- | --- |
| `EXPIRED` | `days_to_end_policies <= 0`. | Existing expired-policy marker. |
| `FLEX` | Flex rule above. | Can appear in Special Case and operational notes. |
| `B2B` | B2B partner or `SMR` claim token. | Uses B2B claim note. |
| `DUPLICATE` | Duplicate detection result. | Uses duplicate comparison across policy/source/status/submission date. |
| `SECOND_YEAR` | `month_policy_aging > 12`. | Operational claim-cell note/detail; Special Case reason. |
| `FIRST_MONTH_POLICY` | Submission date within first 30 days of policy start. | Operational claim-cell note/detail; Special Case reason. |
| `REMAINING_1_MONTH` | Submission date within last 30 days before policy end. | Operational claim-cell note/detail; Special Case reason. |

When multiple markers apply, operational highlighting uses the configured priority in `getOperationalClaimHighlightPolicy_`; do not infer priority from table order here.

## Debugging Source Questions

Use this quick path when a value looks wrong:

| Symptom | Check |
| --- | --- |
| Destination column is blank | Confirm the destination header exists; writers only set columns that exist in that sheet. |
| Raw-driven value is blank | Check canonical raw header, alias support, and whether `applyRawHeaderAliases_` covers the workbook header. |
| Manual field disappeared after MAIN/SUB | Check `RESTORE_AUDIT`, `_OPS_MANUAL_BACKUP`, `_OPS_MAIN_SUB_TEMP`, and whether the claim/service-center key changed. |
| `Status` dropdown formatting broke | Check template row copy and whether a writer attempted to write the `Status` column directly. |
| Second-Year not flagged | Check `month_policy_aging`; exactly `12` does not qualify, only values greater than `12`. |
| First-Month / Policy Remaining not flagged | Check parsed `claim_submitted_datetime` plus policy start/end datetime. Invalid dates fail closed. |
| EV-Bike row missing | Check Raw partner pattern, policy exclusion list, then `Submission` overlay fallback. |
| B2B row missing | Check partner pattern, `SMR` claim token, excluded last-status filtering, and `Submission` fallback. |

## Change Checklist

Use this checklist for every column-structure change:

1. Add or update the canonical source header in `00_Config.gs`.
2. Add any required alias in the header normalization helpers if old workbooks still use the old name.
3. Update `COLUMN_TYPES` when the column needs date, datetime, integer, money, checkbox, or percentage formatting.
4. Update `SV03_TEMPLATES` in `03_SheetsAndValidation.gs` when a destination sheet layout changes.
5. Update the runtime mapper or optional writer that reads or writes the column.
6. Update this document.
7. Run the static smoke check and diff validation before shipping.
