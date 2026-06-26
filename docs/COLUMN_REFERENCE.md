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
| `dashboard_link` | Destination `DB Link`. |
| `DB` | Destination DB value when already available. |

Claim date and aging columns:

| Column | Used for |
| --- | --- |
| `claim_submission_date` | Primary `Submission Date`, B2B/Special Case/EV-Bike output date, first-month and policy-age comparisons. |
| `claim_submitted_datetime` | SUB submission date fallback and datetime output. |
| `claim_last_updated_datetime`, `last_update` | Destination `Last Status Date` / status recency fields. |
| `last_activity_log_date`, `activity_log_datetime`, `last_activity_log_datetime` | Activity log recency and destination datetime fields. |
| `month_policy_aging` | Second-Year (Market Value) flag. Threshold is configured as `12`. |
| `days_to_end_policies` | Policy remaining calculation and Special Case flagging. |
| `days_aging_from_submission` | Destination `TAT`, Exclusion TAT fallback, aging output. |
| `last_status_aging` / `LSA` | Destination `Last Status Aging`; sort and operational monitoring. |
| `activity_log_aging` / `ALA` | Destination `Activity Log Aging`. |
| `Q-L (Months)` | Special Case and optional output context. |
| `M-L (Months)` | Managed Raw Data tail metric. |
| `M-Q (Months)` | First-month policy and Special Case calculations. |

Status and routing columns:

| Column | Used for |
| --- | --- |
| `last_status` | Operational routing, `Status Type`, `Position`, `Submission` trigger, destination `Last Status`. |
| `activity_log`, `last_activity_log` | Destination `Activity Log`. |
| `sc_name` | Destination `Service Center`, SC sheet routing, Branch derivation, Service Center flagging. |
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

| Destination column | Source / rule |
| --- | --- |
| `Submission Date` | `claim_submission_date` |
| `Claim Number` | `claim_number` |
| `DB Link` | `dashboard_link` |
| `DB` | `DB` or classification from claim/source tokens |
| `Partner Name` | `business_partner_name` |
| `Insurance` | `insurance_partner_name` |
| `Device Type` | `device_type` |
| `Product` | `product_name` |
| `Device Brand` | `device_brand` |
| `IMEI/SN` | `imei_number` / `device_imei` |
| `Last Status` | `last_status` |
| `Last Status Date` | `claim_last_updated_datetime` / `last_update` |
| `Service Center` | `sc_name` |
| `Last Status Aging` | `last_status_aging` / `LSA` |
| `Activity Log` | `activity_log` / `last_activity_log` |
| `Activity Log Aging` | `activity_log_aging` / `ALA` |
| `TAT` | `days_aging_from_submission` |
| `Start Date`, `End Date`, `Details` | Manual operational tracking fields. |
| `Sum Insured Amount` | `sum_insured_amount` |
| `Claim Amount` | `claim_amount` |
| `Claim Own Risk Amount` | `claim_own_risk_amount` |
| `Nett Claim Amount` | `nett_claim_amount` |
| `% Approval` | Derived approval percentage when source values support it. |
| `Update Status`, `Timestamp`, `Status` | Manual workflow fields, preserved when rows are rewritten. |
| `Status Type` | Derived from `last_status`. |

SC destination sheets also include:

| Destination column | Source / rule |
| --- | --- |
| `Type` | Derived from `last_status` through service-center type policy. |
| `Branch` | Derived from `Service Center` using branch keyword mapping. |

PO destination sheets also include:

| Destination column | Source / rule |
| --- | --- |
| `OR` | Own-risk checkbox/manual field. |

Admin templates intentionally exclude operational PIC-only columns such as `Device Type`, `Last Status Aging`, `Activity Log Aging`, `TAT`, `OR`, and `OR Amount`.

### B2B

`B2B` is generated by `processB2B_`.

Detection rule:

| Signal | Meaning |
| --- | --- |
| Partner name matches configured B2B partner patterns | Row is a B2B candidate. |
| `Claim Number` contains `SMR` | Row is a B2B claim candidate even if partner matching is incomplete. |

Output columns:

| B2B column | Source / rule |
| --- | --- |
| `Submission Date` | `claim_submission_date` |
| `Claim Number` | `claim_number` |
| `DB Link` | `dashboard_link` |
| `DB` | `source_system_name` or DB classification |
| `Partner Name` | `business_partner_name` |
| `Insurance` | `insurance_partner_name` |
| `Device Type` | `device_type` |
| `Product` | `product_name` |
| `Device Brand` | `device_brand` |
| `IMEI/SN` | `imei_number` / `device_imei` |
| `Last Status` | `last_status` |
| `Service Center` | `sc_name` |
| `Last Status Aging` | `last_status_aging` / `LSA` |
| `Activity Log Aging` | `activity_log_aging` / `ALA` |
| `TAT` | `days_aging_from_submission` |
| `Start Date`, `End Date`, `Details` | Manual tracking fields. |
| `Sum Insured Amount` | `sum_insured_amount` |
| `Claim Amount` | `claim_amount` |
| `Claim Own Risk Amount` | `claim_own_risk_amount` |
| `Nett Claim Amount` | `nett_claim_amount` |
| `% Approval` | Derived approval percentage when possible. |

### Special Case

`Special Case` is generated by `processSpecialCase_`. The sheet is treated as a fixed, user-managed schema; the script should not auto-add arbitrary columns.

Output columns:

| Special Case column | Source / rule |
| --- | --- |
| `Submission Date` | `claim_submission_date` |
| `Claim Number` | `claim_number` |
| `DB Link` | `dashboard_link` |
| `DB` | DB classification or source value |
| `Partner Name` | `business_partner_name` |
| `Insurance` | `insurance_partner_name` |
| `Device Type` | `device_type` |
| `Last Status` | `last_status` |
| `Service Center` | `sc_name` |
| `Last Status Aging` | `last_status_aging` / `LSA` |
| `Activity Log Aging` | `activity_log_aging` / `ALA` |
| `TAT` | `days_aging_from_submission` |
| `Last Status Date` | `claim_last_updated_datetime` / `last_update` |
| `Q-L (Months)` | `Q-L (Months)` |
| `Product` | `product_name` |
| `Sum Insured Amount` | `sum_insured_amount` |
| `Claim Amount` | `claim_amount` |
| `OR` | Own-risk/manual checkbox field. |
| `Claim Own Risk Amount` | `claim_own_risk_amount` |
| `Nett Claim Amount` | `nett_claim_amount` |
| `Selisih` | Difference between relevant claim/insured/nett values. |
| `Reason` | Derived reason such as Flex, Second-Year, First-Month Policy, or Policy Remaining <= 1 Month. |

Special Case flags:

| Flag | Based on |
| --- | --- |
| `Flex` | `Claim Number` contains `SFX`, or `product_name` contains a configured Flex token. |
| `Second-Year (Market Value)` | `month_policy_aging` is greater than or equal to configured threshold `12`. |
| `First-Month Policy` | `M-Q (Months)` / policy-to-claim age indicates a first-month claim. |
| `Policy Remaining <= 1 Month` | Policy end date / remaining-day metric shows the policy is near expiry. |

### EV-Bike

`EV-Bike` is generated by `processEVBike_`.

Detection rule:

| Signal | Meaning |
| --- | --- |
| Partner name matches configured EV-Bike partner patterns | Row is an EV-Bike candidate. |
| Policy number is in the configured exclusion list | Row is skipped. |

Output columns:

| EV-Bike column | Source / rule |
| --- | --- |
| `Submission Date` | `claim_submission_date` |
| `Claim Number` | `claim_number` |
| `DB Link` | `dashboard_link` |
| `Owner Name` | `customer_name` |
| `Policy Number` | `qoala_policy_number` |
| `Partner Name` | `business_partner_name` |
| `Insurance` | `insurance_partner_name` |
| `Sum Insured` | `sum_insured_amount` |
| `Status` | Manual status field; preserved and not overwritten when possible. |

Note: the runtime writer currently removes deprecated `Start Date`, `End Date`, and `Details` columns from EV-Bike if they are still present.

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

### Service Center Routing

Service Center routing uses `sc_name` / `Service Center`.

| Result | Based on |
| --- | --- |
| `SC - Farhan` | Service Center name matches Farhan keyword rules. |
| `SC - Meilani` | Service Center name matches Meilani keyword rules. |
| `SC - Meindar` | Service Center name matches Meindar/Ivan keyword rules. |
| `SC - Unmapped` | No Service Center keyword matches. |

`Branch` is also derived from Service Center text using `BRANCH_KEYWORDS`.

### Status Type And Position

`Status Type` and `Position` are derived from `last_status` using configured status maps. These values drive monitoring, presentation, and operational grouping.

### Duplicate Claim

Duplicate detection uses:

| Column | Purpose |
| --- | --- |
| `qoala_policy_number` | Main policy key for duplicate comparison. |
| `source_system_name` | Helps separate old/new source behavior. |
| `claim_number` | Claim identity and token classification. |
| `claim_submission_date` | Submission date comparison. |
| `last_status` | Status context. |

Configured duplicate comparison allows a maximum submission-date difference of `62` days.

### Claim Highlighting

Claim highlighting and notes use configured marker policies:

| Marker | Based on |
| --- | --- |
| `EXPIRED` | Expired-policy detection and configured highlight policy. |
| `FLEX` | Flex claim/product tokens. |
| `B2B` | B2B partner or `SMR` claim token. |
| `DUPLICATE` | Duplicate detection result. |

## Change Checklist

Use this checklist for every column-structure change:

1. Add or update the canonical source header in `00_Config.gs`.
2. Add any required alias in the header normalization helpers if old workbooks still use the old name.
3. Update `COLUMN_TYPES` when the column needs date, datetime, integer, money, checkbox, or percentage formatting.
4. Update `SV03_TEMPLATES` in `03_SheetsAndValidation.gs` when a destination sheet layout changes.
5. Update the runtime mapper or optional writer that reads or writes the column.
6. Update this document.
7. Run the static smoke check and diff validation before shipping.
