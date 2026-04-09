/***************************************************************
 * CLAIM PIPELINE ETL + ROUTING + MIRRORING (ENTERPRISE EDITION)
 * Module: 00_Config.gs
 ***************************************************************/
'use strict';

/** =========================
 * App namespace + registry (introspection)
 * =========================
 * Apps Script has a single global namespace. Keep the public surface small by
 * registering flows/modules here for easy system introspection.
 */
var App = App || {};
App.Registry = App.Registry || (function () {
  const _flows = Object.create(null);
  const _modules = Object.create(null);

  function registerFlow(key, def) {
    if (!key) return;
    _flows[String(key)] = Object.freeze(Object.assign({ key: String(key) }, def || {}));
  }

  function registerModule(key, def) {
    if (!key) return;
    _modules[String(key)] = Object.freeze(Object.assign({ key: String(key) }, def || {}));
  }

  function snapshot() {
    return {
      flows: Object.assign({}, _flows),
      modules: Object.assign({}, _modules)
    };
  }

  return Object.freeze({
    registerFlow: registerFlow,
    registerModule: registerModule,
    snapshot: snapshot
  });
})();

/** VERSION */
App.APP_VERSION = App.APP_VERSION || '2026.03.26-enterprise-refactor-r3-hardening';

/** Schema version for managed sheets */
const SCHEMA_VERSION = 5;

/**
 * DRY_RUN:
 * - true  => no writes to spreadsheets, no file trashing
 * - false => normal behavior
 */
const DRY_RUN = false;

/**
 * High-level config index.
 * This is documentation-as-data for maintainers and introspection tools,
 * not a runtime switchboard.
 */
const CONFIG_SECTION_INDEX = Object.freeze({
  foundation: Object.freeze(['App.Registry', 'APP_VERSION', 'SCHEMA_VERSION', 'DRY_RUN']),
  ingestion: Object.freeze(['EMAIL_INGEST_POLICY', 'SUB_EMAIL_INGEST_POLICY', 'SUB_FLOW_SPEC', 'FORM_INGEST_POLICY']),
  workbook: Object.freeze(['MASTER_SPREADSHEET_ID', 'MASTER_RAW_SHEET_NAME', 'WEBAPP_MOVEMENT_POLICY', 'WORKBOOK_PROFILES']),
  routing: Object.freeze(['OPS_ROUTING_POLICY', 'STATUS_TYPE_BY_LAST_STATUS', 'POSITION_BY_LAST_STATUS', 'FINISH_STATUSES']),
  validationAndPresentation: Object.freeze(['VALIDATION_POLICY', 'VALIDATION_FALLBACKS', 'CHECKBOX_POLICY', 'LINK_POLICY', 'COLUMN_TYPES', 'COLUMN_ALIGNMENT']),
  optionalSheets: Object.freeze(['SPECIAL_CASE_WRITER_POLICY', 'EVBIKE_POLICY', 'EXCLUSION_TAT_POLICY', 'OPTIONAL_SHEETS_FLAGS']),
  observability: Object.freeze(['LOG_POLICY', 'MAPPING_ERROR_LOG_POLICY', 'DETAILS_LOG_POLICY', 'UI_FLAGS'])
});

/** =========================
 * Status groups (routing helpers)
 * =========================
 * NOTE: keep these in sync with routing overrides (05a/05b) + SC post-processing (06c).
 */
const FINISH_STATUSES = Object.freeze([
  'DONE_REPAIR',
  'WAITING_WALKIN_FINISH',
  'COURIER_PICKED_UP',
  'WAITING_COURIER_FINISH',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH',
  'SERVICE_CENTER_CLAIM_DONE',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH',
  'COURIER_CLAIM_PICKUP_FINISH',
  'COURIER_CLAIM_PICKUP_FINISH_DONE'
]);

const SC_SHEET_NAMES = Object.freeze([
  'SC - Farhan',
  'SC - Meilani',
  'SC - Meindar'
]);

const SC_TYPE_DROPDOWN_OPTIONS = Object.freeze([
  'SC - Rcvd',
  'SC - Est',
  'Insurance',
  'OR',
  'Finish',
  'SC - Wait Rep',
  'SC - On Rep'
]);

function getScTypeDropdownOptions_() {
  return (typeof SC_TYPE_DROPDOWN_OPTIONS !== 'undefined' && Array.isArray(SC_TYPE_DROPDOWN_OPTIONS))
    ? SC_TYPE_DROPDOWN_OPTIONS.slice()
    : ['SC - Rcvd','SC - Est','Insurance','OR','Finish','SC - Wait Rep','SC - On Rep'];
}


/**
 * Branch mapping based on Service Center string (SC sheets).
 * Matching is case-insensitive substring.
 */
const BRANCH_KEYWORDS = Object.freeze({
  'Mitracare': ['mitracare'],
  'Sitcomtara': ['sitcomtara'],
  'iBox': ['ibox'],
  'GSI': ['gsi'],
  'Andalas': ['andalas'],
  'Klikcare': ['klikcare'],
  'J-Bros': ['j-bros', 'jbros'],
  'Makmur Era Abadi': ['makmur era abadi'],
  'Manado Mitra Bersama': ['manado mitra bersama'],
  'CV Kayu Awet Sejahtera': ['cv kayu awet sejahtera', 'kayu awet sejahtera'],
  'GH Store': ['gh store'],
  'Unicom': ['unicom'],
  'Samsung Authorized Service Centre by Unicom': ['samsung authorized service centre by unicom', 'authorized service centre by unicom', 'samsung authorized service center by unicom', 'authorized service center by unicom'],
  'Xiaomi Authorized': ['xiaomi authorized', 'xiaomi'],
  'Samsung Exclusive': ['samsung exclusive', 'samsung'],
  'Carlcare': ['carlcare'],
  'B-Store': ['b-store', 'bstore'],
  'EzCare': ['ezcare', 'ez care'],
  'Deltasindo': ['deltasindo'],
  'MDP': ['mdp']
});


/** =========================
 * Master workbook (single source of truth)
 * =========================
 * All ingestion (email + optional form upload) writes into this workbook.
 */
const MASTER_SPREADSHEET_ID = getPropString_(
  'MASTER_SPREADSHEET_ID',
  '1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc'
);
const MASTER_RAW_SHEET_NAME = getPropString_('MASTER_RAW_SHEET_NAME', 'Raw Data');

/** =========================
 * WebApp Project (movement tracking)
 * =========================
 * Snapshot baseline MUST live in the WebApp Project spreadsheet (NOT Raw OLD/NEW).
 */
const WEBAPP_PROJECT_SPREADSHEET_ID = getPropString_(
  'WEBAPP_PROJECT_SPREADSHEET_ID',
  '1anPGHYa8Ej19jZJMC3bKyReZ-O6Qki2WBvtRl6rNTTk'
);

const WEBAPP_MOVEMENT_POLICY = Object.freeze({
  ENABLE: getPropBool_('WEBAPP_MOVEMENT_ENABLE', true),
  SPREADSHEET_ID: WEBAPP_PROJECT_SPREADSHEET_ID,

  /**
   * Snapshot sheets live in the WebApp Project spreadsheet.
   * Flow contract:
   * - Before SUB starts: copy Overview Claim -> Raw OLD/NEW into SNAPSHOT_PREV_OLD/NEW
   * - After  SUB ends:   copy Overview Claim -> Raw OLD/NEW into SNAPSHOT_CURR_OLD/NEW
   */
  SHEETS: Object.freeze({
    DAILY: 'Daily',
    PAST: 'Past',

    SNAPSHOT_PREV_OLD: 'SNAPSHOT_PREV_OLD',
    SNAPSHOT_CURR_OLD: 'SNAPSHOT_CURR_OLD',

    SNAPSHOT_PREV_NEW: 'SNAPSHOT_PREV_NEW',
    SNAPSHOT_CURR_NEW: 'SNAPSHOT_CURR_NEW'
  }),

  // Column names MUST match snapshot headers (WebApp Project)
  SNAPSHOT_COLUMNS: Object.freeze([
    'Claim Number',
    'Last Status',
    'Last Update Datetime',
    'Activity Log',
    'Activity Log Datetime',
    'Service Center Name',
    'Branch',
    'Position',
    'Status Type'
  ]),

  // Column names MUST match Daily/Past headers (WebApp Project)
  DAILY_COLUMNS: Object.freeze([
    'Timestamp',
    'DB',
    'Claim Number',
    'Change Type',

    'Last Status (Before)',
    'Last Update Datetime (Before)',
    'Last Status (After)',
    'Last Update Datetime (After)',
    'Gap Time Status (Minutes)',
    'Gap Time Status',

    'Activity Log',
    'Activity Log Datetime',

    'Service Center Name',
    'Branch',
    'Position',
    'Status Type',

    'Event ID'
  ]),

  // Existing Event-ID lookup for Past uses recent rows only (performance guard).
  PAST_EVENT_SCAN_MAX_ROWS: getPropInt_('WEBAPP_PAST_EVENT_SCAN_MAX_ROWS', 5000)
});


/**
 * Second-Year (Market Value) detection policy.
 * Requirement: STRICTLY use month_policy_aging from Raw Data (>12 months).
 */
const SECOND_YEAR_MARKET_VALUE_POLICY = Object.freeze({
  RAW_HEADER: 'month_policy_aging',
  THRESHOLD_MONTHS: 12
});

/** =========================
 * Email ingestion (QUEUE-based, deterministic)
 * =========================
 * Spec:
 * - Gmail filters route matching emails into QUEUED_* labels.
 * - Script only reads from queued label + unread + has:attachment.
 * - Success cleanup: markRead + removeLabel + moveToTrash.
 * - Error: leave unread + keep label (automatic retry).
 */

const GMAIL_QUEUE_LABELS = Object.freeze({
  MAIN: getPropString_('GMAIL_QUEUE_LABEL_MAIN', 'QUEUED_MAIN'),
  SUB:  getPropString_('GMAIL_QUEUE_LABEL_SUB',  'QUEUED_SUB')
});

/**
 * MAIN flow (daily ~08:00)
 * - MUST be sourced only from label:QUEUED_MAIN
 */
const EMAIL_INGEST_POLICY = Object.freeze({
  ENABLE: getPropBool_('EMAIL_INGEST_ENABLE', true),

  FLOW: 'MAIN',
  QUEUED_LABEL: GMAIL_QUEUE_LABELS.MAIN,
  MAX_EMAILS_PER_RUN: 1,

  FROM: getPropString_('EMAIL_INGEST_FROM', 'data-reporting@qoala.id'),
  SUBJECT: getPropString_(
    'EMAIL_INGEST_SUBJECT',
    'Daily Claim Pending Monitoring'
  ),

  // Attachment name starts with this prefix (case-sensitive match is safer).
  ATTACHMENT_NAME_PREFIX: getPropString_(
    'EMAIL_INGEST_ATTACHMENT_NAME_PREFIX',
    '[QGP][ID] Claim Daily Monitoring'
  ),

  // LEGACY: kept for backward-compat, but queue-mode uses trash cleanup.
  PROCESSED_LABEL: getPropString_('EMAIL_INGEST_PROCESSED_LABEL', 'PROCESSED_MAIN'),

  // Queue-based query: always deterministic, never "sweeps" Gmail.
  SEARCH_QUERY: getPropString_(
    'EMAIL_INGEST_SEARCH_QUERY',
    'label:' + GMAIL_QUEUE_LABELS.MAIN + ' has:attachment is:unread -in:trash -in:spam'
  )
});

/**
 * SUB flow (hourly)
 * - MUST be sourced only from label:QUEUED_SUB
 * - NEW vs OLD attachment detection:
 *   - NEW: filename contains "(Standardization)"
 *   - OLD: filename contains "List of Claims with Aging" and not standardization
 */
const SUB_EMAIL_INGEST_POLICY = Object.freeze({
  ENABLE: getPropBool_('SUB_EMAIL_INGEST_ENABLE', true),

  FLOW: 'SUB',
  QUEUED_LABEL: GMAIL_QUEUE_LABELS.SUB,
  MAX_EMAILS_PER_RUN: 1,

  FROM: getPropString_('SUB_EMAIL_INGEST_FROM', 'data-reporting@qoala.id'),
  SUBJECT_EXACT: getPropString_('SUB_EMAIL_INGEST_SUBJECT', 'Claim Monitoring Operational Dashboard'),

  ATTACHMENT_MARKERS: Object.freeze({
    NEW_CONTAINS: getPropString_('SUB_ATTACH_NEW_CONTAINS', '(Standardization)'),
    OLD_CONTAINS: getPropString_('SUB_ATTACH_OLD_CONTAINS', 'List of Claims with Aging')
  }),

  // Queue-based query: only queued label + strict From/Subject.
  SEARCH_QUERY: getPropString_(
    'SUB_EMAIL_INGEST_SEARCH_QUERY',
    'label:' + GMAIL_QUEUE_LABELS.SUB +
      ' data-reporting@qoala.id subject:"Claim Monitoring Operational Dashboard" has:attachment is:unread -in:trash -in:spam'
  )
});


/**
 * SUB (QUEUED_SUB) operational update spec.
 *
 * Purpose:
 * - Convert 2 XLSX attachments (OLD then NEW) to temporary spreadsheets.
 * - Copy full datasets into "Raw OLD" and "Raw NEW".
 * - Update operational sheets (lightweight) by Claim Number:
 *   - Last Status Aging
 *   - Activity Log Aging
 *   - Last Status
 *   - Service Center
 * - Append to "Submission" if:
 *   - OLD: last_status == SUBMITTED and Claim Number not found in Submission
 *   - NEW: last_status == CLAIM_INITIATE and Claim Number not found in Submission
 * - After SUB flow completes: sort operational sheets by:
 *   Last Status Aging (Z→A), Last Status (A→Z), DB (A→Z)
 */
const SUB_FLOW_SPEC = Object.freeze({
  // Raw sheets where the full attachments are copied to.
  RAW_OLD_SHEET_NAME: getPropString_('SUB_RAW_OLD_SHEET_NAME', 'Raw OLD'),
  RAW_NEW_SHEET_NAME: getPropString_('SUB_RAW_NEW_SHEET_NAME', 'Raw NEW'),

  // Target operational sheets for lightweight updates (allow-list).
  OPERATIONAL_SHEETS: Object.freeze([
    'Submission',
    'Ask Detail',
    'OR - OLD',
    'SC - Farhan',
    'SC - Meilani',
    'SC - Meindar',
    'Start',
    'Finish',
    'PO',
    'Exclusion',
    'B2B',
    'EV-Bike',
    'Special Case'
  ]),

  // Standard operational headers (destination)
  OP_HEADERS: Object.freeze({
    CLAIM_NUMBER: 'Claim Number',
    LAST_STATUS_AGING: 'Last Status Aging',
    ACTIVITY_LOG_AGING: 'Activity Log Aging',
    LAST_STATUS: 'Last Status',
    SERVICE_CENTER: 'Service Center',
    SUBMISSION_DATE: 'Submission Date',
    DB: 'DB',
    DB_LINK: 'DB Link',
    PARTNER_NAME: 'Partner Name',
    INSURANCE: 'Insurance',
    DEVICE_TYPE: 'Device Type',
    IMEI_SN: 'IMEI/SN'
  }),

  // Standard raw headers (source) used by SUB flow
  RAW_HEADERS: Object.freeze({
    CLAIM_NUMBER: 'claim_number',
    SUBMITTED_DATETIME: 'claim_submitted_datetime',
    DASHBOARD_LINK: 'dashboard_link',
    PARTNER_NAME: 'partner_name',
    INSURANCE_CODE: 'insurance_partner_code', // OLD commonly uses this
    INSURANCE_CODE_ALT: 'insurance_code',     // NEW commonly uses this
    DEVICE_TYPE: 'device_type',
    DEVICE_IMEI: 'device_imei',
    LAST_STATUS: 'last_status',
    LAST_STATUS_AGING: 'last_status_aging',
    ACTIVITY_LOG_AGING: 'activity_log_aging',
    SERVICE_CENTER: 'sc_name'
  }),

  // Submission append rules
  SUBMISSION_RULES: Object.freeze({
    OLD: Object.freeze({
      TRIGGER_RAW_LAST_STATUS: 'SUBMITTED',
      DB_VALUE: 'OLD'
    }),
    NEW: Object.freeze({
      TRIGGER_RAW_LAST_STATUS: 'CLAIM_INITIATE',
      DB_VALUE: 'NEW'
    })
  }),

  // Sorting spec for operational sheets after SUB finishes
  SORT_SPECS: Object.freeze([
    Object.freeze({ header: 'Last Status Aging', ascending: false }),
    Object.freeze({ header: 'Last Status', ascending: true }),
    Object.freeze({ header: 'DB', ascending: true })
  ])
});


/** =========================
 * Secondary flow: Form upload ingest (manual/optional)
 * ========================= */
const FORM_INGEST_POLICY = Object.freeze({
  ENABLE: getPropBool_('FORM_INGEST_ENABLE', true),

  /**
   * Form fields (case-insensitive header match).
   * Recommended:
   * - Flow field contains: MAIN / SUB
   * - File upload field may allow multiple files (for SUB: OLD+NEW in one field).
   */
  FLOW_FIELD_NAME: getPropString_(
    'FORM_FLOW_FIELD_NAME',
    'Flow'
  ),

  // Primary file upload question (can accept multiple files).
  FILE_UPLOAD_FIELD_NAME: getPropString_(
    'FORM_FILE_UPLOAD_FIELD_NAME',
    'Metabase - Upload Claim Data'
  ),

  // Optional: split SUB uploads into 2 separate file-upload questions.
  // If empty, SUB will be detected/picked from FILE_UPLOAD_FIELD_NAME uploads.
  SUB_OLD_FILE_UPLOAD_FIELD_NAME: getPropString_(
    'FORM_SUB_OLD_FILE_UPLOAD_FIELD_NAME',
    ''
  ),
  SUB_NEW_FILE_UPLOAD_FIELD_NAME: getPropString_(
    'FORM_SUB_NEW_FILE_UPLOAD_FIELD_NAME',
    ''
  )
});
/** Lock */
const LOCK_TIMEOUT_MS = 30000;

/**
 * Legacy rule: suppress some logs/details for submissions < this year
 * (Operational routing MUST NOT rely on this; this is log hygiene only.)
 */
const LEGACY_SKIP_YEAR_BEFORE = 2025;

/** Display formats (applied via setNumberFormat, not stringification) */
const FORMATS = Object.freeze({
  DATE: 'd mmm yy',            // 6 Nov 25
  DATETIME: 'd mmm yy, HH:mm', // 6 Dec 25, 14:46
  DATETIME_LONG: 'MMMM d, yyyy, h:mm AM/PM', // February 16, 2026, 5:49 PM
  TIMESTAMP: 'd mmm, HH:mm',   // 15 Dec, 15:17
  INT: '0',
  MONEY0: '#,##0',
  PERCENT0: '0%'
});

/** Pipeline flags */
const PIPELINE_FLAGS = Object.freeze({
  CLEAR_LOG_BEFORE_RUN: true,
  ENABLE_BACKUP_FROM_OPS: true,
  FAIL_IF_MAIN_NOT_FOUND: true,

  /**
   * Safety gate:
   * - true  => pipeline may trash uploaded files (only after successful route)
   * - false => never trash
   * Actual enforcement happens in 06 (after "DONE").
   */
  TRASH_UPLOADED_FILES: true
});

/**
 * Logging policy:
 * User spec:
 * - "Log" should show life-sign early (BOOT/START)
 * - rest can be bulk/segmented
 *
 * Enforcement lives in 02/06 logging funcs (flush policy).
 */
const LOG_POLICY = Object.freeze({
  BOOT_IMMEDIATE_LINE: true,     // write a line at run start
  BOOT_IMMEDIATE_FLUSH: true,    // flush right after boot line
  SEGMENT_LINES_BULK: true,      // segment logs can be batched
  SEGMENT_FLUSH_EVERY: 0         // 0 = no forced flush per segment, >0 flush every N segments
});

/**
 * Log sheet requirement (mapping errors):
 * - For unmapped partner or unmapped last_status, log a short entry:
 *   <Partner/Last Status> | <Claim Number> | <Relevant Date (submission or last status date)>
 */
const MAPPING_ERROR_LOG_POLICY = Object.freeze({
  ENABLE: true,
  SHEET_NAME_FALLBACK: 'Log',

  EVENT: Object.freeze({
    UNMAPPED_PARTNER: 'UNMAPPED_PARTNER',
    UNMAPPED_LAST_STATUS: 'UNMAPPED_LAST_STATUS'
  }),

  // Minimal fields expected in the log line / row
  FIELDS: Object.freeze({
    KEY: 'Claim Number',
    VALUE: 'Partner/Last Status',
    DATE: 'Relevant Date'
  })
});


/** Details logging policy (sheet: "Details") */
const DETAILS_LOG_POLICY = Object.freeze({
  /**
   * Spec:
   * - Do not report "Partner is not mapped to any Associate (Associate left blank)."
   *   when Submission Date is before June 2025.
   */
  UNMAPPED_PARTNER_MIN_SUBMISSION_DATE: getPropString_(
    'DETAILS_UNMAPPED_PARTNER_MIN_SUBMISSION_DATE',
    '2025-06-01' // ISO date (inclusive)
  )
});

/** UI flags */
const UI_FLAGS = Object.freeze({
  REALTIME_PROGRESS: false,
  REALTIME_LOG_LINES: false,
  REALTIME_DETAILS_APPEND: false,
  FLUSH_AFTER_LAYOUT_CLEAR: false
});

/**
 * Workbook profiles:
 * - PIC: full operational + optional sheets
 * - ADMIN: core sheets only (NO B2B/SpecialCase/EV-Bike)
 */
const WORKBOOK_PROFILES = Object.freeze({
  PIC: 'PIC',
  ADMIN: 'ADMIN'
});

/** Special Case flags */
const SPECIAL_CASE_FLAGS = Object.freeze({
  MODE: 'UPSERT', // 'APPEND_NEW_ONLY' | 'UPSERT' | 'REBUILD'

  // Rules (final rule switches live in SPECIAL_CASE_RULES below, via Script Properties overrides)
  ENABLE_FLEX_RULE: true,
  ENABLE_Q_L_OVER_12_RULE: true,
  ENABLE_POLICY_REMAINING_LE_1_RULE: true,
  ENABLE_Q_L_0_1_RULE: true, // include claims where Q-L (Months) is 0 or 1

  // Hard gates
  MIN_SUBMISSION_YEAR: 2025, // Special Case only for claim_submission_date year >= 2025

  // Spec: Special Case must only contain ongoing claims → skip done/closed statuses
  SKIP_EXCLUDED_LAST_STATUSES: true,

  // UX
  COLORIZE_CLAIM_CELL: true
});

/**
 * Special Case writer policy (used by 05c).
 * - UPSERT by Claim Number (no REBUILD) to preserve manual inputs.
 * - If a claim becomes excluded (done/closed), the row should be deleted (descending delete).
 */
const SPECIAL_CASE_WRITER_POLICY = Object.freeze({
  SHEET_NAME: 'Special Case',
  KEY_HEADER: 'Claim Number',

  // Columns managed by the user (must never be overwritten by script)
  MANUAL_HEADERS: Object.freeze([
    'Claim Amount',
    'Repair/Replace Amount', // legacy

    'Top Up',
    'Selisih',
    'Update Status',
    'Timestamp',
    'Status'
  ]),

  // New columns sourced from Raw Data
  START_DATE_HEADER: 'Start Date',
  END_DATE_HEADER: 'End Date',
  DETAILS_HEADER: 'Details',

  PRUNE_WHEN_EXCLUDED: true,
  SKIP_IF_KEY_BLANK: true
});


/** Optional sheet skip rules */
const OPTIONAL_SHEETS_FLAGS = Object.freeze({
  B2B_SKIP_EXCLUDED_LAST_STATUSES: true,
  EVBIKE_SKIP_EXCLUDED_LAST_STATUSES: true,
  SPECIAL_CASE_SKIP_EXCLUDED_LAST_STATUSES: true // aligned with Special Case spec
});

/** -------------------------
 * Script Properties helpers
 * ------------------------- */
function getPropString_(key, fallback) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    return (v != null && String(v).trim() !== '') ? String(v).trim() : fallback;
  } catch (e) {
    return fallback;
  }
}
function getPropBool_(key, fallbackBool) {
  const raw = getPropString_(key, '');
  if (!raw) return !!fallbackBool;
  const s = String(raw).trim().toLowerCase();
  return ['true', 'yes', '1', 'y', 'checked', 'on'].indexOf(s) > -1;
}
function getPropInt_(key, fallbackInt) {
  const raw = getPropString_(key, '');
  const n = parseInt(raw, 10);
  return Number.isFinite(n) ? n : fallbackInt;
}
function getTz_() { return Session.getScriptTimeZone(); }
function equalsIgnoreCase_(a, b) { return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase(); }
function isWildcard_(v) {
  const s = String(v == null ? '' : v).trim();
  return s === '*' || s.toLowerCase() === 'all' || s.toLowerCase() === 'any';
}
function isPicAllowed_(pic, rule) {
  // Empty rule or wildcard means allowed for all PICs.
  if (rule == null) return true;
  const r = String(rule).trim();
  if (!r) return true;
  if (isWildcard_(r)) return true;
  return equalsIgnoreCase_(pic, r);
}

/**
 * Deep-freeze helper for nested config objects/arrays.
 * Use only for plain objects/arrays (do not pass SpreadsheetApp objects).
 */
function deepFreeze_(obj) {
  if (!obj || typeof obj !== 'object' || Object.isFrozen(obj)) return obj;
  Object.freeze(obj);
  const props = Object.getOwnPropertyNames(obj);
  for (let i = 0; i < props.length; i++) {
    const v = obj[props[i]];
    if (v && typeof v === 'object' && !Object.isFrozen(v)) deepFreeze_(v);
  }
  return obj;
}


/** Rule overrides via Script Properties */
const SPECIAL_CASE_RULES = Object.freeze({
  ENABLE_FLEX_RULE: getPropBool_('SC_ENABLE_FLEX_RULE', SPECIAL_CASE_FLAGS.ENABLE_FLEX_RULE),
  ENABLE_Q_L_OVER_12_RULE: getPropBool_('SC_ENABLE_Q_L_OVER_12_RULE', SPECIAL_CASE_FLAGS.ENABLE_Q_L_OVER_12_RULE),
  ENABLE_POLICY_REMAINING_LE_1_RULE: getPropBool_('SC_ENABLE_POLICY_REMAINING_LE_1_RULE', SPECIAL_CASE_FLAGS.ENABLE_POLICY_REMAINING_LE_1_RULE),
  ENABLE_Q_L_0_1_RULE: getPropBool_('SC_ENABLE_Q_L_0_1_RULE', SPECIAL_CASE_FLAGS.ENABLE_Q_L_0_1_RULE),

  MIN_SUBMISSION_YEAR: getPropInt_('SC_MIN_SUBMISSION_YEAR', SPECIAL_CASE_FLAGS.MIN_SUBMISSION_YEAR),

  // Keep as switch even if spec says true, for emergency ops toggling
  SKIP_EXCLUDED_LAST_STATUSES: getPropBool_('SC_SKIP_EXCLUDED_LAST_STATUSES', SPECIAL_CASE_FLAGS.SKIP_EXCLUDED_LAST_STATUSES)
});

const OPTIONAL_SHEETS_RULES = Object.freeze({
  // Wildcard '*' means "all" (no PIC gating). Modules must honor this.
  EVBIKE_ONLY_FOR_PIC: getPropString_('EVBIKE_ONLY_FOR_PIC', '*'),
  B2B_ONLY_FOR_PIC: getPropString_('B2B_ONLY_FOR_PIC', '*')
});

/** -------------------------
 * EV-Bike routing policy (latest spec)
 * -------------------------
 * Requirement:
 * - Every EV-Bike claim that appears in 'Submission' must be upserted into sheet 'EV-Bike'.
 * - If the Claim Number already exists in 'EV-Bike', rewrite that row (no duplicates).
 * - If the Submission Date is older than N days, the EV-Bike claim may be removed from 'Submission'.
 */
/**
 * EV-Bike policy (latest spec):
 * - Only processed when PIC = Farhan (else skip create/update/delete).
 * - Include all EV-Bike claims (sources include sheet 'Submission').
 * - Upsert by Claim Number; overwrite only managed columns; never touch manual Status.
 * - If EV-Bike sheet is filtered, do not reset/unfilter while writing.
 */
const EVBIKE_POLICY = Object.freeze({
  ENABLE: true,

  // Gate: only for this PIC (default: Farhan)
  ONLY_FOR_PIC: OPTIONAL_SHEETS_RULES.EVBIKE_ONLY_FOR_PIC,

  // Sheet identities
  EVBIKE_SHEET_NAME: 'EV-Bike',
  SUBMISSION_SHEET_NAME: 'Submission',
  CLAIM_NUMBER_HEADER: 'Claim Number',

  // Skip policy numbers even if they match EV-Bike criteria
  EXCLUDED_POLICY_NUMBERS: Object.freeze([
    'GODA-20250729-4SHFZ',
    'GODA-20250729-X97UC',
    'GODA-20250729-KESB4'
  ]),

  // Upsert behavior: overwrite only these columns for existing Claim Number
  MANAGED_HEADERS_ON_UPSERT: Object.freeze([
    'Submission Date',
    'Owner Name',
    'Policy Number',
    'Partner Name',
    'Insurance',
    'Sum Insured',
    'DB Link'
  ]),

  // Manual columns (must never be overwritten)
  MANUAL_HEADERS: Object.freeze(['Status']),

  // Runtime behavior
  UPSERT_ON_EVERY_RUN: true,
  DO_NOT_RESET_FILTER: true
});

function getEvBikeExcludedPolicyNumberSet_() {
  if (!RUNTIME.evBikeExcludedPolicySet) {
    RUNTIME.evBikeExcludedPolicySet = new Set(EVBIKE_POLICY.EXCLUDED_POLICY_NUMBERS || []);
  }
  return RUNTIME.evBikeExcludedPolicySet;
}





/** -------------------------
 * Exclusion TAT policy (latest spec)
 * -------------------------
 * Requirement:
 * - Exclusion.TAT (days) = Submission Date - Last Status Date.
 */
const EXCLUSION_TAT_POLICY = Object.freeze({
  SHEET_NAME: 'Exclusion',
  SUBMISSION_DATE_HEADER: 'Submission Date',
  LAST_STATUS_DATE_HEADER: 'Last Status Date',
  OUTPUT_HEADER: 'TAT',

  // Days since submission until the last status date (calendar-day diff).
  DIRECTION: 'LAST_STATUS_MINUS_SUBMISSION', // 'LAST_STATUS_MINUS_SUBMISSION' | 'SUBMISSION_MINUS_LAST_STATUS'

  ROUNDING: 'FLOOR', // floor of day difference (integer days)
  CLAMP_MIN_ZERO: true
});



/** Form response source */
const RESPONSES_SPREADSHEET_ID = getPropString_(
  'RESPONSES_SPREADSHEET_ID',
  '1TC9YjDo6qxWq17zPYEBqIryhaYbUMqGtaSH0F-G8IwE'
);
const RESPONSES_SHEET_NAME = getPropString_('RESPONSES_SHEET_NAME', 'Form Responses 1');

/** -------------------------
 * Dropdown validation: Auto-heal policy
 * -------------------------
 * IMPORTANT:
 * - We DO NOT treat any static list here as source-of-truth.
 * - Auto-heal will read existing dropdown options from:
 *   (a) Raw Data columns (Status, Associate)
 *   (b) Target sheet columns (Status, Associate)
 * Then union + write back rule/options to both sides.
 *
 * This prevents: "violates data validation rules ... Please enter one of ..."
 */
const VALIDATION_POLICY = Object.freeze({
  ENABLE_AUTO_HEAL: true,

  /**
   * Where to pull canonical dropdown options from:
   * - "RAW" uses Raw Data column rule/options as primary
   * - "TARGET" uses current target sheet column rule/options as primary
   * - "BIDIR" unions both (recommended for your spec)
   */
  MODE: 'BIDIR', // 'RAW' | 'TARGET' | 'BIDIR'

  /**
   * Preserve visuals:
   * - Some Google Sheets "dropdown chip colors" can be lost if we rebuild the rule.
   * - Strategy in 03/05:
   *   - Prefer copying the existing rule object from source range when possible.
   *   - Only rebuild when needed.
   */
  PRESERVE_RULE_BY_COPY: true,

  /**
   * Blank-safe:
   * - Blank values are allowed and must not be treated as error.
   */
  ALLOW_BLANK: true,

  /**
   * Hard fallback list (ONLY if we fail to read any rule/options).
   * This is NOT a routing/validation authority.
   */
  USE_FALLBACK_WHEN_EMPTY: true
});

/**
 * Fallback dropdown lists (NOT source-of-truth).
 * Kept to prevent total failure if both Raw+Target have no rule/options.
 */
const VALIDATION_FALLBACKS = Object.freeze({
  ASSOCIATE: Object.freeze(['Meilani', 'Farhan', 'Suci', 'Adi']),
  STATUS: Object.freeze([
    'Pending Admin',
    'Pending SC',
    'Pending Partner',
    'DONE',
    'Pending Insurance',
    'Pending TO',
    'Pending Finance',
    'Pending Meilani',
    'Pending Cust'
  ])
});

/**
 * Backward-compat alias:
 * - Some existing modules may still reference VALIDATION_LISTS.
 * - We keep it, but treat it as fallback only.
 */
const VALIDATION_LISTS = Object.freeze({
  PIC: VALIDATION_FALLBACKS.ASSOCIATE,
  UPDATE_STATUS: VALIDATION_FALLBACKS.STATUS
});


/**
 * Insurance short mapping rules.
 * - Case-insensitive substring match against Raw: insurance_partner_name
 * - Output blank when no match (avoid misleading "Other")
 * - Keep this in 00 for easy maintenance (config/data only).
 */
const INSURANCE_SHORT_RULES = Object.freeze([
  Object.freeze({ needle: 'great eastern general insurance', short: 'GEGI' }),
  Object.freeze({ needle: 'tokio marine', short: 'TMI' }),
  Object.freeze({ needle: 'msig indonesia', short: 'MSIG' }),
  Object.freeze({ needle: 'seainsure', short: 'MIGI' }),
  Object.freeze({ needle: 'sompo insurance', short: 'Sompo' }),
  Object.freeze({ needle: 'axa mandiri insurance', short: 'AXA' }),
  Object.freeze({ needle: 'simas insurtech', short: 'Simas' }),
  // Zurich often appears with variant spelling.
  Object.freeze({ needle: 'zurich', short: 'Zurich' }),
  Object.freeze({ needle: 'zcurich', short: 'Zurich' })
]);

/** -------------------------
 * Checkbox policy (latest spec)
 * -------------------------
 * Requirement:
 * - Column "OR" must be real checkbox validation (not TRUE/FALSE text).
 * - Mirroring must preserve checkbox validation in targets.
 */
const CHECKBOX_POLICY = Object.freeze({
  ENABLE_ENFORCE: true,
  HEADER_NAMES: Object.freeze(['OR']), // sheet header display name
  /**
   * Strategy:
   * - If cell range has no checkbox validation, we "insertCheckboxes" / setValidation
   * - Values are written as boolean (true/false), not string.
   */
  ENFORCE_WHEN_MISSING: true
});

/** -------------------------
 * Link rendering policy (latest spec)
 * -------------------------
 * Fix for Admin "DB Link" showing #ERROR! instead of LINK text:
 * - Avoid HYPERLINK() formulas (locale , vs ; issues / empty url issues).
 * - Render using RichTextValue with link URL.
 */
const LINK_POLICY = Object.freeze({
  MODE: 'RICHTEXT',    // 'RICHTEXT' | 'FORMULA'
  DISPLAY_TEXT: 'LINK' // what user sees in the cell
});

/** -------------------------
 * Associate column policy (latest spec)
 * -------------------------
 * Requirement:
 * - Do NOT auto-add "Associate" on Raw Data or operational sheets.
 * - If it exists (legacy), preserve it but do not depend on it for routing or validation.
 */
const ASSOCIATE_COLUMN_POLICY = Object.freeze({
  /**
   * Spec update (2026-01-15):
   * - Do NOT auto-add "Associate" on Raw Data or operational sheets.
   * - If the column already exists (legacy), preserve it, but do not depend on it.
   */
  PIC_DESTINATION_ASSOCIATE: 'DISABLED',
  ADMIN_DESTINATION_ASSOCIATE: 'DISABLED',

  // Explicit allowlist for modules that want a simple guard (kept for backward-compat)
  OPS_ENABLED_PICS: Object.freeze([]),

  // Backward-compat keys (older modules may still read these)
  ASSOCIATE_DV_SOURCE_SHEET: 'Raw Data',
  ADMIN_ASSOCIATE_DV_SOURCE_SHEET: 'Raw Data'
});


/** -------------------------
 * Date coercion policy (latest spec)
 * -------------------------
 * Requirement:
 * - Admin "Last Status Date" must not be forced to datetime if raw data is date-only.
 */
const DATE_COERCION_POLICY = Object.freeze({
  ADMIN_LAST_STATUS_DATE_MODE: 'AUTO_DATE_OR_DATETIME' // modules decide based on presence of time component
});

/** -------------------------
 * Special Case reason labels (user-facing)
 * ------------------------- */
const SPECIAL_CASE_REASON_LABELS = Object.freeze({
  // Canonical labels (user-facing) — allow multiple reasons per claim
  FLEX: 'Flex',
  QL_GT_12: 'Second-Year (Market Value)',
  QL_LE_1: 'First-Month Policy',
  MQ_LE_1: 'Policy Remaining ≤ 1 Month',

  // Backward-compat aliases (older code paths)
  POLICY_REMAINING_LE_1: 'Policy Remaining ≤ 1 Month',
  QL_0_1: 'First-Month Policy'
});


/** -------------------------
 * Special Case coloring policy (user spec)
 * - Coloring applies ONLY in "Special Case" sheet and ONLY on the claim cell.
 * - Priority: FLEX > QL>12 > QL<=1 > MQ<=1
 * ------------------------- */
const SPECIAL_CASE_COLORING = Object.freeze({
  ENABLE: !!SPECIAL_CASE_FLAGS.COLORIZE_CLAIM_CELL,
  CLAIM_BG_BY_TRIGGER: Object.freeze({
    FLEX: '#f4c7c3',       // pink (highest priority)
    QL_OVER_12: '#fff2cc', // yellow
    QL_LE_1: '#cfe2f3',    // light blue
    MQ_LE_1: '#e1d5f7'     // purple
  }),
  PRIORITY: Object.freeze(['FLEX', 'QL_OVER_12', 'QL_LE_1', 'MQ_LE_1'])
});

/** -------------------------
 * Status Type mapping (mandatory operational column)
 * -------------------------
 * Requirement: derive Status Type from Last Status using user-provided mapping.
 * Any unknown status will be classified as 'UNKNOWN' to aid auditing.
 */
const STATUS_TYPE_BY_LAST_STATUS = Object.freeze({
  'CLAIM_INITIATE': 'OPEN',
  'QOALA_ASK_DETAIL': 'FOLLOW UP',
  'CUSTOMER_RESUBMIT_DOCUMENT': 'OPEN',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_REQ_QOALA': 'OPEN',
  'CLAIM_EXPIRE': 'CLOSE',
  'QOALA_CLAIM_REOPEN': 'OPEN',
  'QOALA_CLAIM_APPROVE_WALKIN': 'FOLLOW UP',
  'WAITING_WALKIN_START': 'FOLLOW UP',
  'CLAIM_EXPIRE_WALKIN': 'CLOSE',
  'QOALA_CLAIM_REOPEN_WALKIN': 'FOLLOW UP',
  'QOALA_CLAIM_APPROVE_PICKUP': 'FOLLOW UP',
  'WAITING_PICKUP_START': 'FOLLOW UP',
  'COURIER_PICKUP_START': 'FOLLOW UP',
  'COURIER_PICKUP_START_DONE': 'FOLLOW UP',
  'SUBMITTED': 'OPEN',
  'QOALA_CLAIM_ASK_DETAIL': 'FOLLOW UP',
  'QOALA_CLAIM_RESUBMIT_DOC': 'FOLLOW UP',
  'WAITING_PAYMENT': 'FOLLOW UP',
  'DONE_EXPIRED': 'CLOSE',
  'WAITING_COURIER_START': 'FOLLOW UP',
  'CLAIM_ADDED_SC': 'FOLLOW UP',
  'RECEIVED_SC': 'FOLLOW UP',
  'ESTIMATE_COST': 'OPEN',
  'ON_PROGRESS': 'FOLLOW UP',
  'INSURANCE_ASK_DETAIL': 'FOLLOW UP',
  'CX_UPLOAD_DOC': 'FOLLOW UP',
  'APPROVED': 'FOLLOW UP',
  'INSURANCE_APPROVED': 'FOLLOW UP',
  'REPLACED': 'OPEN',
  'DONE_REPAIR': 'DONE - FOLLOW UP',
  'WAITING_WALKIN_FINISH': 'DONE - FOLLOW UP',
  'COURIER_PICKED_UP': 'FOLLOW UP',
  'WAITING_COURIER_FINISH': 'DONE - FOLLOW UP',
  'SERVICE_CENTER_CLAIM_RECEIVE': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_ESTIMATE': 'OPEN',
  'QOALA_CLAIM_RESUBMIT_ESTIMATE': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_RESUBMIT_ESTIMATE': 'OPEN',
  'SERVICE_CENTER_CLAIM_WAITING_REPAIR': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_ON_PROGRESS': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_CHANGE_IMEI': 'FOLLOW UP',
  'QOALA_CLAIM_APPROVE_REPAIR': 'FOLLOW UP',
  'QOALA_CLAIM_APPROVE_REPLACE': 'FOLLOW UP',
  'INSURANCE_CLAIM_REVIEW': 'FOLLOW UP',
  'INSURANCE_CLAIM_APPROVE_REPAIR': 'FOLLOW UP',
  'INSURANCE_CLAIM_ASK_DETAIL_ADDITIONAL': 'FOLLOW UP',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_ADDITIONAL': 'FOLLOW UP',
  'CLAIM_EXPIRE_INSURANCE': 'FOLLOW UP',
  'QOALA_CLAIM_REOPEN_INSURANCE_CASHLESS': 'OPEN',
  'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'FOLLOW UP',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'FOLLOW UP',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR_EXPIRED': 'FOLLOW UP',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH': 'DONE - FOLLOW UP',
  'SERVICE_CENTER_CLAIM_DONE': 'DONE',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH': 'DONE - FOLLOW UP',
  'COURIER_CLAIM_PICKUP_FINISH': 'FOLLOW UP',
  'COURIER_CLAIM_PICKUP_FINISH_DONE': 'DONE',
  'INSURANCE_APPROVED_REPLACED': 'FOLLOW UP',
  'DONE_REJECTED': 'CLOSE',
  'DONE': 'DONE',
  'DONE_REPLACED': 'DONE',
  'DONE_REJECT': 'CLOSE',
  'QOALA_REQUEST_SALVAGE': 'FOLLOW UP',
  'PAYMENT_NOT_COMPLETE': 'CLOSE',
  'DONE_CANCEL': 'CLOSE',
  'INSURANCE_REJECTED': 'FOLLOW UP',
  'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'FOLLOW UP',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'FOLLOW UP',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE_EXPIRED': 'FOLLOW UP',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_RREPLACE': 'FOLLOW UP',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'FOLLOW UP',
  'INSURANCE_CLAIM_APPROVE_REPLACE': 'FOLLOW UP',
  'SERVICE_CENTER_REPAIR_CANCELLED_FOR_REPLACE': 'OPEN',
  'QOALA_PROCESS_REPLACE': 'OPEN',
  'QOALA_PROCESS_REPLACE_WALKIN': 'OPEN',
  'CUSTOMER_WAITING_EXCESS_REPLACE_WALKIN': 'FOLLOW UP',
  'CUSTOMER_PAID_EXCESS_REPLACE_WALKIN': 'FOLLOW UP',
  'QOALA_WAITING_CUSTOMER_REPLACE': 'FOLLOW UP',
  'CUSTOMER_RECEIVE_REPLACE': 'DONE',
  'QOALA_PROCESS_REPLACE_PICKUP': 'OPEN',
  'CUSTOMER_WAITING_EXCESS_REPLACE_PICKUP': 'FOLLOW UP',
  'CUSTOMER_PAID_EXCESS_REPLACE_PICKUP': 'FOLLOW UP',
  'COURIER_WAITING_REPLACE_PICKUP': 'FOLLOW UP',
  'COURIER_REPLACE_PICKUP': 'FOLLOW UP',
  'COURIER_REPLACE_PICKUP_DONE': 'DONE',
  'QOALA_CLAIM_REJECT': 'CLOSE',
  'QOALA_CLAIM_REJECT_PICKUP': 'CLOSE - FOLLOW UP',
  'QOALA_CLAIM_REJECT_WALKIN': 'CLOSE - FOLLOW UP',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_WALKIN': 'FOLLOW UP',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_PICKUP': 'FOLLOW UP',
  'INSURANCE_CLAIM_REJECT_WALKIN': 'FOLLOW UP',
  'INSURANCE_CLAIM_REJECT_PICKUP': 'FOLLOW UP',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT': 'CLOSE - FOLLOW UP',
  'SERVICE_CENTER_CLAIM_DONE_REJECT': 'CLOSE - REJECT',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_REJECT': 'CLOSE - FOLLOW UP',
  'COURIER_CLAIM_PICKUP_REJECT': 'FOLLOW UP',
  'COURIER_CLAIM_PICKUP_REJECT_DONE': 'CLOSE - REJECT',
  'INSURANCE_CLAIM_WAITING_PAID_REPAIR': 'DONE - WAIT INSURANCE',
  'INSURANCE_CLAIM_PAID_REPAIR': 'DONE - INSURANCE',
  'INSURANCE_CLAIM_WAITING_PAID_REPLACE': 'DONE - WAIT INSURANCE',
  'INSURANCE_CLAIM_PAID_REPLACE': 'DONE - INSURANCE'
});

/** -------------------------
 * Position mapping (Front/Middle/Back)
 * -------------------------
 * Requirement: Position MUST be derived from Last Status and stored as:
 * - 'Front' | 'Middle' | 'Back'
 *
 * Notes:
 * - 'DB' (OLD/NEW) is recorded on Daily/Past; Position mapping is shared across DBs
 *   because status codes are globally unique.
 * - Unknown statuses must map to 'Unknown' to surface schema drift.
 */
const POSITION_BY_LAST_STATUS = Object.freeze({
  // ===== OLD =====
  'SUBMITTED': 'Front',
  'QOALA_CLAIM_ASK_DETAIL': 'Front',
  'QOALA_CLAIM_RESUBMIT_DOC': 'Front',
  'WAITING_PAYMENT': 'Front',
  'DONE_EXPIRED': 'Front',
  'WAITING_WALKIN_START': 'Front',
  'WAITING_COURIER_START': 'Front',
  'COURIER_PICKED_UP': 'Front',
  'PAYMENT_NOT_COMPLETE': 'Front',

  'CLAIM_ADDED_SC': 'Middle',
  'RECEIVED_SC': 'Middle',
  'ESTIMATE_COST': 'Middle',
  'ON_PROGRESS': 'Middle',
  'INSURANCE_ASK_DETAIL': 'Middle',
  'CX_UPLOAD_DOC': 'Middle',
  'APPROVED': 'Middle',
  'INSURANCE_APPROVED': 'Middle',
  'REPLACED': 'Middle',
  'DONE_REPAIR': 'Middle',
  'WAITING_WALKIN_FINISH': 'Middle',
  'WAITING_COURIER_FINISH': 'Middle',
  'DONE_REJECTED': 'Middle',
  'DONE_REJECT': 'Middle',
  'QOALA_REQUEST_SALVAGE': 'Middle',
  'DONE_CANCEL': 'Middle',
  'INSURANCE_REJECTED': 'Middle',

  'INSURANCE_APPROVED_REPLACED': 'Back',
  'DONE': 'Back',
  'DONE_REPLACED': 'Back',

  // ===== NEW =====
  'CLAIM_INITIATE': 'Front',
  'QOALA_ASK_DETAIL': 'Front',
  'CUSTOMER_RESUBMIT_DOCUMENT': 'Front',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_REQ_QOALA': 'Front',
  'CLAIM_EXPIRE': 'Front',
  'QOALA_CLAIM_REOPEN': 'Front',
  'QOALA_CLAIM_APPROVE_WALKIN': 'Front',
  'CLAIM_EXPIRE_WALKIN': 'Front',
  'QOALA_CLAIM_REOPEN_WALKIN': 'Front',
  'QOALA_CLAIM_APPROVE_PICKUP': 'Front',
  'WAITING_PICKUP_START': 'Front',
  'COURIER_PICKUP_START': 'Front',
  'COURIER_PICKUP_START_DONE': 'Front',
  'QOALA_CLAIM_REJECT': 'Front',

  'SERVICE_CENTER_CLAIM_RECEIVE': 'Middle',
  'SERVICE_CENTER_CLAIM_ESTIMATE': 'Middle',
  'QOALA_CLAIM_RESUBMIT_ESTIMATE': 'Middle',
  'SERVICE_CENTER_CLAIM_RESUBMIT_ESTIMATE': 'Middle',
  'SERVICE_CENTER_CLAIM_WAITING_REPAIR': 'Middle',
  'SERVICE_CENTER_CLAIM_ON_PROGRESS': 'Middle',
  'SERVICE_CENTER_CLAIM_CHANGE_IMEI': 'Middle',
  'QOALA_CLAIM_APPROVE_REPAIR': 'Middle',
  'QOALA_CLAIM_APPROVE_REPLACE': 'Middle',
  'INSURANCE_CLAIM_REVIEW': 'Middle',
  'INSURANCE_CLAIM_APPROVE_REPAIR': 'Middle',
  'INSURANCE_CLAIM_ASK_DETAIL_ADDITIONAL': 'Middle',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_ADDITIONAL': 'Middle',
  'CLAIM_EXPIRE_INSURANCE': 'Middle',
  'QOALA_CLAIM_REOPEN_INSURANCE_CASHLESS': 'Middle',
  'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'Middle',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'Middle',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR_EXPIRED': 'Middle',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPAIR': 'Middle',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN': 'Middle',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH': 'Middle',
  'SERVICE_CENTER_CLAIM_DONE': 'Middle',
  'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP': 'Middle',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH': 'Middle',
  'COURIER_CLAIM_PICKUP_FINISH': 'Middle',
  'COURIER_CLAIM_PICKUP_FINISH_DONE': 'Middle',

  'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'Back',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'Back',
  'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE_EXPIRED': 'Back',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_RREPLACE': 'Back',
  'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPLACE': 'Back',
  'INSURANCE_CLAIM_APPROVE_REPLACE': 'Back',
  'SERVICE_CENTER_REPAIR_CANCELLED_FOR_REPLACE': 'Back',
  'QOALA_PROCESS_REPLACE': 'Back',
  'QOALA_PROCESS_REPLACE_WALKIN': 'Back',
  'CUSTOMER_WAITING_EXCESS_REPLACE_WALKIN': 'Back',
  'CUSTOMER_PAID_EXCESS_REPLACE_WALKIN': 'Back',
  'QOALA_WAITING_CUSTOMER_REPLACE': 'Back',
  'CUSTOMER_RECEIVE_REPLACE': 'Back',
  'QOALA_PROCESS_REPLACE_PICKUP': 'Back',
  'CUSTOMER_WAITING_EXCESS_REPLACE_PICKUP': 'Back',
  'CUSTOMER_PAID_EXCESS_REPLACE_PICKUP': 'Back',
  'COURIER_WAITING_REPLACE_PICKUP': 'Back',
  'COURIER_REPLACE_PICKUP': 'Back',
  'COURIER_REPLACE_PICKUP_DONE': 'Back',

  'QOALA_CLAIM_REJECT_PICKUP': 'Middle',
  'QOALA_CLAIM_REJECT_WALKIN': 'Middle',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_WALKIN': 'Middle',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_PICKUP': 'Middle',
  'INSURANCE_CLAIM_REJECT_WALKIN': 'Middle',
  'INSURANCE_CLAIM_REJECT_PICKUP': 'Middle',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT': 'Middle',
  'SERVICE_CENTER_CLAIM_DONE_REJECT': 'Middle',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_REJECT': 'Middle',
  'COURIER_CLAIM_PICKUP_REJECT': 'Middle',
  'COURIER_CLAIM_PICKUP_REJECT_DONE': 'Middle',

  'INSURANCE_CLAIM_WAITING_PAID_REPAIR': 'Back',
  'INSURANCE_CLAIM_PAID_REPAIR': 'Back',
  'INSURANCE_CLAIM_WAITING_PAID_REPLACE': 'Back',
  'INSURANCE_CLAIM_PAID_REPLACE': 'Back'
});

/**
 * Helper: safe Position derivation.
 * Always returns a string (defaults to 'Unknown').
 */
function getPositionFromLastStatus_(lastStatus) {
  const key = String(lastStatus || '').trim();
  return POSITION_BY_LAST_STATUS[key] || 'Unknown';
}

/** -------------------------
 * Highlight + note policy (operational sheets only; driven by Raw Data sets)
 * ------------------------- */
const CLAIM_HIGHLIGHT_POLICY = Object.freeze({
  ENABLE: true,

  /**
   * Recommended implementation strategy (enforced in 05b):
   * - FOLLOW_NOTE: use the note text on the Claim Number cell to decide fill color (safest).
   * - FROM_RAW: derive flags from Raw Data (days_to_end_policies / claim_number tokens).
   */
  MODE: getPropString_('CLAIM_HIGHLIGHT_MODE', 'FOLLOW_NOTE'), // 'FOLLOW_NOTE' | 'FROM_RAW'

  TARGET_SCOPE: 'OPERATIONAL_ONLY',

  // Apply to both Admin and PIC workbooks
  APPLY_TO_WORKBOOK_PROFILES: Object.freeze(['ADMIN', 'PIC']),

  // Column identification
  CLAIM_NUMBER_HEADER: 'Claim Number',
  CLAIM_NUMBER_HEADER_ALIASES: Object.freeze(['Claim Number', 'Claim No', 'Claim No.', 'Claim #', 'Claim#']),

  // Marker colors
  COLORS: Object.freeze({
    EXPIRED: '#fff2cc', // light yellow
    FLEX: '#f4c7c3',    // light red/pink
    B2B: '#c9daf8',     // light blue
    DUPLICATE: '#dd7e6b' // duplicate claim
  }),

  // Canonical notes (must be written by the pipeline)
  NOTES_CANONICAL: Object.freeze({
    EXPIRED: 'Policy already expired.',
    FLEX: 'Flex claim.',
    B2B: 'B2B claim.',
    DUPLICATE_PREFIX: 'Duplicate Claim - Refer to Claim Number'
  }),

  // Accept legacy note strings as matchers so colorization still works.
  NOTE_MATCHERS: Object.freeze({
    EXPIRED: Object.freeze(['Policy already expired.']),
    FLEX: Object.freeze(['Flex claim.', 'Flex.', 'FLEX claim.']),
    B2B: Object.freeze(['B2B claim.']),
    DUPLICATE: Object.freeze(['Duplicate Claim - Refer to Claim Number'])
  }),

  /**
   * Cleanup behavior:
   * - When a row is not flagged, the pipeline may clear ONLY these marker backgrounds,
   *   so the system can recover from any previous accidental mass-fill without breaking other styling.
   */
  CLEAR_MARKER_BACKGROUNDS_WHEN_NOT_FLAGGED: true,
  MARKER_COLORS: Object.freeze(['#fff2cc', '#f4c7c3', '#c9daf8', '#dd7e6b'])
});


/** -------------------------
 * Fixed schema policy (no auto-add columns)
 * - Special Case & Exclusion must respect whatever columns already exist (manual schema).
 * ------------------------- */
const FIXED_SCHEMA_POLICY = Object.freeze({
  DISALLOW_AUTO_ADD_COLUMNS_IN_SHEETS: Object.freeze(['Special Case', 'Exclusion'])
});

/** -------------------------
 * Claim detection tokens (avoid scattered hardcodes)
 * ------------------------- */
const CLAIM_DETECTION_TOKENS = Object.freeze({
  // Existing markers
  FLEX_CLAIM_NUMBER_SUBSTRING: 'SFX',
  FLEX_PRODUCT_SUBSTRING: 'flex',
  B2B_CLAIM_NUMBER_SUBSTRING: 'SMR',

  // Duplicate detection (by qoala_policy_number, gated by source_system_name)
  DUPLICATE_POLICY_SUBSTRINGS: Object.freeze(['SFP', 'SFX', 'SMR']),
  DUPLICATE_SOURCE_NEW: 'NEW SERVICE',
  DUPLICATE_SOURCE_OLD: 'OLD SERVICE'
});

/**
 * DB classification (latest spec):
 * - OLD if Claim Number contains any of: SFP, SFX, SMR
 * - NEW if Claim Number contains any of: VVMAR, GADLD
 */
const DB_CLASSIFICATION_TOKENS = Object.freeze({
  OLD_SUBSTRINGS: Object.freeze(['SFP', 'SFX', 'SMR']),
  NEW_SUBSTRINGS: Object.freeze(['VVMAR', 'GADLD'])
});


/**
 * Duplicate claim detection policy:
 * - Only NEW SERVICE rows with qoala_policy_number containing SFP/SFX/SMR are eligible.
 * - Match against OLD SERVICE with same qoala_policy_number.
 * - Apply only if submission date delta is within ~2 months (day-based guard).
 */
const DUPLICATE_DETECTION_POLICY = Object.freeze({
  ENABLE: true,

  // Date delta guard (use day-based comparison to avoid month-length pitfalls)
  MAX_MONTH_DIFF: 2,
  MAX_DAY_DIFF: 62,

  RAW_HEADERS: Object.freeze({
    POLICY_NUMBER: 'qoala_policy_number',
    SOURCE_SYSTEM: 'source_system_name',
    CLAIM_NUMBER: 'claim_number',
    SUBMISSION_DATE: 'claim_submission_date',
    LAST_STATUS: 'last_status'
  }),

  // UX
  COLOR: '#dd7e6b',
  NOTE_PREFIX: 'Duplicate Claim - Refer to Claim Number'
});


/**
 * Excluded "done/closed" last_status values.
 * Used by: Exclusion sheet, optional sheets, and (per spec) Special Case.
 */
const EXCLUDED_LAST_STATUSES_BASE = Object.freeze([
  'DONE_REJECTED',
  'DONE',
  'DONE_REPLACED',
  'DONE_REJECT',
  'QOALA_REQUEST_SALVAGE',
  'QOALA_CLAIM_REJECT',
  'SERVICE_CENTER_CLAIM_DONE_REJECT',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT',
  'INSURANCE_CLAIM_WAITING_PAID_REPAIR',
  'INSURANCE_CLAIM_PAID_REPAIR',
  'INSURANCE_CLAIM_WAITING_PAID_REPLACE',
  'INSURANCE_CLAIM_PAID_REPLACE',
  'CUSTOMER_RECEIVE_REPLACE'
]);

/**
 * Build excluded status Set for current run.
 * Optional override: Script Property `EXCLUDED_LAST_STATUSES_CSV` (comma-separated).
 */
function buildExcludedLastStatusesSet_() {
  const overrideCsv = getPropString_('EXCLUDED_LAST_STATUSES_CSV', '');
  if (overrideCsv && overrideCsv.trim()) {
    const arr = overrideCsv.split(',').map(s => s.trim()).filter(Boolean);
    return new Set(arr);
  }
  return new Set(EXCLUDED_LAST_STATUSES_BASE);
}


/**
 * Cached excluded status Set for this run (reset via resetRuntime_()).
 * Use this instead of rebuilding the Set repeatedly.
 */
function getExcludedLastStatusesSet_() {
  if (!RUNTIME.excludedLastStatusesSet) {
    RUNTIME.excludedLastStatusesSet = buildExcludedLastStatusesSet_();
  }
  return RUNTIME.excludedLastStatusesSet;
}

/** Core config */

/** =========================
 * Raw Data: custom tail columns (must be preserved; never overwritten)
 * ========================= */
const RAW_DATA_CUSTOM_TAIL_HEADERS = Object.freeze([
  'Update Status',
  'Timestamp',
  'Status',
  'Remarks',
  'Q-L (Months)',
  'M-L (Months)',
  'M-Q (Months)',
  'Update Status Asso',
  'Timestamp Asso',
  'Update Status Admin',
  'Timestamp Admin'
]);

/** =========================
 * Operational routing (single master workbook)
 * =========================
 * Notes:
 * - SC - Farhan, SC - Meilani, and SC - Meindar share the same last_status universe, but are split by sc_name keywords.
 * - "Type" (dropdown) on both SC sheets is filled from last_status categories.
 */
const OPS_ROUTING_POLICY = Object.freeze({
  SHEETS: Object.freeze({
    SUBMISSION: 'Submission',
    ASK_DETAIL: 'Ask Detail',
    OR_OLD: 'OR - OLD',
    START: 'Start',
    FINISH: 'Finish',
    SC_FARHAN: 'SC - Farhan',
    SC_MEILANI: 'SC - Meilani',
    SC_IVAN: 'SC - Meindar',
    PO: 'PO',
    EXCLUSION: 'Exclusion'
  }),

  LAST_STATUS_BY_SHEET: Object.freeze({
    'Submission': Object.freeze(['SUBMITTED', 'CLAIM_INITIATE']),

    'Ask Detail': Object.freeze([
      'QOALA_CLAIM_ASK_DETAIL',
      'QOALA_CLAIM_RESUBMIT_DOC',
      'QOALA_ASK_DETAIL',
      'CUSTOMER_RESUBMIT_DOCUMENT',
      'QOALA_CLAIM_RESUBMIT_DOCUMENT_REQ_QOALA',
      'CLAIM_EXPIRE',
      'QOALA_CLAIM_REOPEN'
    ]),

    // Legacy queue; keep exactly as requested.
    'OR - OLD': Object.freeze(['WAITING_PAYMENT']),

    'Start': Object.freeze([
      'DONE_EXPIRED',
      'WAITING_WALKIN_START',
      'WAITING_COURIER_START',
      'QOALA_CLAIM_APPROVE_WALKIN',
      'CLAIM_EXPIRE_WALKIN',
      'QOALA_CLAIM_REOPEN_WALKIN',
      'QOALA_CLAIM_APPROVE_PICKUP',
      'WAITING_PICKUP_START',
      'COURIER_PICKUP_START',
      'COURIER_PICKUP_START_DONE'
    ]),

    'Finish': Object.freeze([
      'DONE_REPAIR',
      'WAITING_WALKIN_FINISH',
      'COURIER_PICKED_UP',
      'WAITING_COURIER_FINISH',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN',
      'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH',
      'SERVICE_CENTER_CLAIM_DONE',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP',
      'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH_DONE'
    ]),

    // SC universe (shared by Farhan/Meilani; split via sc_name keyword)
    '__SC_SHARED__': Object.freeze([
      'CLAIM_ADDED_SC',
      'RECEIVED_SC',
      'ESTIMATE_COST',
      'ON_PROGRESS',
      'DONE_REPAIR',
      'WAITING_WALKIN_FINISH',
      'COURIER_PICKED_UP',
      'WAITING_COURIER_FINISH',
      'INSURANCE_ASK_DETAIL',
      'CX_UPLOAD_DOC',
      'APPROVED',
      'INSURANCE_APPROVED',
      'REPLACED',
      'SERVICE_CENTER_CLAIM_RECEIVE',
      'SERVICE_CENTER_CLAIM_ESTIMATE',
      'QOALA_CLAIM_RESUBMIT_ESTIMATE',
      'SERVICE_CENTER_CLAIM_RESUBMIT_ESTIMATE',
      'SERVICE_CENTER_CLAIM_WAITING_REPAIR',
      'SERVICE_CENTER_CLAIM_ON_PROGRESS',
      'SERVICE_CENTER_CLAIM_CHANGE_IMEI',
      'QOALA_CLAIM_APPROVE_REPAIR',
      'QOALA_CLAIM_APPROVE_REPLACE',
      'INSURANCE_CLAIM_REVIEW',
      'INSURANCE_CLAIM_APPROVE_REPAIR',
      'INSURANCE_CLAIM_ASK_DETAIL_ADDITIONAL',
      'QOALA_CLAIM_RESUBMIT_DOCUMENT_ADDITIONAL',
      'CLAIM_EXPIRE_INSURANCE',
      'QOALA_CLAIM_REOPEN_INSURANCE_CASHLESS',
      'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPAIR',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR_EXPIRED',
      'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPAIR',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN',
      'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH',
      'SERVICE_CENTER_CLAIM_DONE',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP',
      'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH_DONE'
    ]),

    'PO': Object.freeze([
      'INSURANCE_APPROVED_REPLACED',
      'INSURANCE_CLAIM_APPROVE_REPLACE',
      'SERVICE_CENTER_REPAIR_CANCELLED_FOR_REPLACE',
      'QOALA_PROCESS_REPLACE',
      'QOALA_PROCESS_REPLACE_WALKIN',
      'CUSTOMER_WAITING_EXCESS_REPLACE_WALKIN',
      'CUSTOMER_PAID_EXCESS_REPLACE_WALKIN',
      'QOALA_WAITING_CUSTOMER_REPLACE',
      'CUSTOMER_RECEIVE_REPLACE',
      'QOALA_PROCESS_REPLACE_PICKUP',
      'CUSTOMER_WAITING_EXCESS_REPLACE_PICKUP',
      'CUSTOMER_PAID_EXCESS_REPLACE_PICKUP',
      'COURIER_WAITING_REPLACE_PICKUP',
      'COURIER_REPLACE_PICKUP',
      'COURIER_REPLACE_PICKUP_DONE',
      'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPLACE',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPLACE_EXPIRED',
      'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_RREPLACE',
      'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPLACE'
    ]),

    'Exclusion': Object.freeze([
      'DONE_REJECTED',
      'DONE',
      'DONE_REPLACED',
      'DONE_REJECT',
      'QOALA_REQUEST_SALVAGE',
      'INSURANCE_REJECTED',
      'QOALA_CLAIM_REJECT',
      'QOALA_CLAIM_REJECT_PICKUP',
      'QOALA_CLAIM_REJECT_WALKIN',
      'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_WALKIN',
      'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_PICKUP',
      'INSURANCE_CLAIM_REJECT_WALKIN',
      'INSURANCE_CLAIM_REJECT_PICKUP',
      'SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT',
      'SERVICE_CENTER_CLAIM_DONE_REJECT',
      'SERVICE_CENTER_CLAIM_WAITING_PICKUP_REJECT',
      'COURIER_CLAIM_PICKUP_REJECT',
      'COURIER_CLAIM_PICKUP_REJECT_DONE',
      'INSURANCE_CLAIM_WAITING_PAID_REPAIR',
      'INSURANCE_CLAIM_PAID_REPAIR',
      'INSURANCE_CLAIM_WAITING_PAID_REPLACE',
      'INSURANCE_CLAIM_PAID_REPLACE'
    ])
  }),

  // Split SC sheets by sc_name keyword match (case-insensitive substring).
  SC_NAME_KEYWORDS: Object.freeze({
    'SC - Farhan': Object.freeze([
      'Mitracare',
      'Sitcomtara',
      'iBox',
      'GSI'
    ]),
    'SC - Meindar': Object.freeze([
      'Klikcare',
      'J-Bros',
      'Makmur Era Abadi',
      'Manado Mitra Bersama',
      'CV Kayu Awet Sejahtera',
      'Kayu Awet Sejahtera',
      'MDP',
      'Deltasindo',
      'EzCare',
      'Ez Care',
      'B-Store',
      'Multikom',
      'GH Store'
    ]),
    'SC - Meilani': Object.freeze([
      'Andalas',
      'Unicom',
      'Authorized Service Centre by Unicom',
      'Samsung Authorized Service Centre by Unicom',
      'Authorized Service Center by Unicom',
      'Samsung Authorized Service Center by Unicom',
      'Xiaomi Authorized',
      'Samsung Exclusive',
      'Carlcare'
    ])
  }),

  // If sc_name does not match any keyword list, route into this sheet (and log it).
  SC_FALLBACK_SHEET: 'SC - Unmapped',

  // "Type" dropdown fill rules for SC sheets.
  TYPE_BY_LAST_STATUS: Object.freeze({
    'SC - Rcvd': Object.freeze([
      'SERVICE_CENTER_CLAIM_RECEIVE',
      'CLAIM_ADDED_SC',
      'RECEIVED_SC'
    ]),
    'SC - Est': Object.freeze([
      'SERVICE_CENTER_CLAIM_ESTIMATE',
      'QOALA_CLAIM_RESUBMIT_ESTIMATE',
      'SERVICE_CENTER_CLAIM_RESUBMIT_ESTIMATE',
      'ESTIMATE_COST'
    ]),
    'Insurance': Object.freeze([
      'INSURANCE_ASK_DETAIL',
      'CX_UPLOAD_DOC',
      'APPROVED',
      'INSURANCE_APPROVED',
      'REPLACED',
      'QOALA_CLAIM_APPROVE_REPAIR',
      'QOALA_CLAIM_APPROVE_REPLACE',
      'INSURANCE_CLAIM_REVIEW',
      'INSURANCE_CLAIM_APPROVE_REPAIR',
      'INSURANCE_CLAIM_ASK_DETAIL_ADDITIONAL',
      'QOALA_CLAIM_RESUBMIT_DOCUMENT_ADDITIONAL',
      'CLAIM_EXPIRE_INSURANCE',
      'QOALA_CLAIM_REOPEN_INSURANCE_CASHLESS'
    ]),
    'OR': Object.freeze([
      'CUSTOMER_WAITING_PAYMENT_DEDUCTIBLE_EXCESS_FEE_REPAIR',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR',
      'CUSTOMER_APPROVE_DEDUCTIBLE_EXCESS_FEE_REPAIR_EXPIRED',
      'CUSTOMER_PAID_DEDUCTIBLE_EXCESS_FEE_REPAIR'
    ]),
    'Finish': Object.freeze([
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN',
      'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH',
      'SERVICE_CENTER_CLAIM_DONE',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP',
      'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH_DONE',
      'DONE_REPAIR',
      'WAITING_WALKIN_FINISH',
      'COURIER_PICKED_UP',
      'WAITING_COURIER_FINISH'
    ]),
    'SC - Wait Rep': Object.freeze([
      'INSURANCE_CLAIM_APPROVE_REPAIR',
      'SERVICE_CENTER_CLAIM_WAITING_REPAIR',
      'INSURANCE_APPROVED'
    ]),
    'SC - On Rep': Object.freeze([
      'SERVICE_CENTER_CLAIM_ON_PROGRESS',
      'SERVICE_CENTER_CLAIM_CHANGE_IMEI',
      'ON_PROGRESS'
    ])
  })
});

/** -------------------------
 * Mapping Team Claim policy (PIC mapping source)
 * -------------------------
 * Sheet: "[UPDATED] Mapping Team Claim"
 * Notes:
 * - Partner lists start at row 4.
 * - For new PIC "Adi", partner list starts at column G (cell G4 downward).
 * - Other PICs keep their existing mapping columns (handled in their respective modules).
 */
const MAPPING_TEAM_CLAIM_POLICY = Object.freeze({
  ENABLE: false, // deprecated (PIC mapping removed; process all data into master)
  PARTNER_START_ROW: 4,
  PARTNER_START_COLUMN_BY_PIC: Object.freeze({
    Adi: 'G'
  })
});



/** -------------------------
 * Column placement rules (your spec)
 * -------------------------
 * You said:
 * - "Last Status Date", "Last Status Aging", "OR Amount" ONLY in "Special Case"
 * - BUT "OR" ALSO required in "PO"
 * - For Admin, do NOT add Device Type / LSA / ALA / TAT / OR / OR Amount to destinations
 * Latest spec update:
 * - Admin may have "Last Status Date" and must not be forced to datetime if raw is date-only.
 */
const COLUMN_PLACEMENT_RULES = Object.freeze({
  // Latest spec: aging + amount columns are allowed broadly (PIC operational + optional sheets).
  // Keep this object for backward-compat (03/05/06 may consult it), but do not restrict LSA/ALA/amount columns to Special Case.
  ONLY_IN_SHEETS: Object.freeze({}),

  ALSO_ALLOWED_IN_SHEETS: Object.freeze({
    'OR': Object.freeze(['Special Case', 'PO'])
  }),

  // Admin workbook is allowed to omit aging/OR/amount helper columns; prevent auto-ensure from adding them.
  // (Legacy names are kept to avoid accidental regressions.)
  ADMIN_FORBIDDEN_COLUMNS: Object.freeze([
    'Device Type',
    'LSA', 'ALA', 'TAT',
    'Last Status Aging', 'Activity Log Aging',
    'OR', 'OR Amount', 'Claim Own Risk Amount'
    // NOTE: 'Last Status Date' is NOT forbidden (latest spec)
  ])
});

/** -------------------------
 * Typed column registry (prevents date/number miswrites)
 * -------------------------
 * New dynamic types added:
 * - DATE_AUTO: choose DATE vs DATETIME based on value having time component
 */
const COLUMN_TYPES = Object.freeze({
  RAW: Object.freeze({
    'policy_start_date': 'DATE',
    'policy_end_date': 'DATE',
    'claim_submission_date': 'DATE',
    'claim_submitted_datetime': 'DATETIME',
    'last_activity_log_date': 'DATETIME',
    'last_update': 'DATETIME',
    'claim_last_updated_datetime': 'DATETIME',
    'activity_log_datetime': 'DATETIME',
    'last_activity_log_datetime': 'DATETIME',

    'days_to_end_policies': 'INT',
    'days_aging_from_submission': 'INT',
    'last_status_aging': 'INT',
    'activity_log_aging': 'INT',
    'Q-L (Months)': 'INT',
    'month_policy_aging': 'INT',

    'sum_insured_amount': 'MONEY0',
    'claim_own_risk_amount': 'MONEY0',
    'nett_claim_amount': 'MONEY0'
  }),

  /**
   * Legacy:
   * - Keep this for backward-compat only.
   * - 03/05 should migrate to DEST_BY_SHEET and enforce COLUMN_PLACEMENT_RULES.
   */
  DEST_COMMON: Object.freeze({
    'Submission Date': 'DATE',
    'Submitted Datetime': 'DATETIME',
    'Timestamp': 'TIMESTAMP',

    'Last Status Aging': 'INT',
    'Activity Log Aging': 'INT',
    'TAT': 'INT',
    'Sum Insured Amount': 'MONEY0',
    'Sum Insured': 'MONEY0', // legacy
    'Claim Amount': 'MONEY0',
    'Claim Own Risk Amount': 'MONEY0',
    'Selisih': 'MONEY0',
    'Nett Claim Amount': 'MONEY0',
    'Q-L (Months)': 'INT'
  }),

  /**
   * New: Per-sheet destination typing (enables "Special Case only" columns)
   */
  DEST_BY_SHEET: Object.freeze({
    'Special Case': Object.freeze({
      'Submission Date': 'DATE',
      'Start Date': 'DATE',
      'End Date': 'DATE',
      'Submitted Datetime': 'DATETIME',
      'Timestamp': 'TIMESTAMP',
      'Last Status Date': 'DATE_AUTO',
      'Last Status Aging': 'INT',
      'Activity Log Aging': 'INT',
      'LSA': 'INT',
      'ALA': 'INT',
      'TAT': 'INT',

      'Sum Insured Amount': 'MONEY0',
      'Sum Insured': 'MONEY0', // legacy
      'Claim Amount': 'MONEY0',
      'Repair/Replace Amount': 'MONEY0', // legacy
      'Claim Own Risk Amount': 'MONEY0',
      'OR Amount': 'MONEY0', // legacy
      'Selisih': 'MONEY0',

      'OR': 'CHECKBOX',
      'Nett Claim Amount': 'MONEY0',
      'Q-L (Months)': 'INT'
    }),
    'PO': Object.freeze({
      'Submission Date': 'DATE',
      'Submitted Datetime': 'DATETIME',
      'Timestamp': 'TIMESTAMP',
      'Sum Insured': 'MONEY0',
      'OR': 'CHECKBOX', // allowed exception
      'Nett Claim Amount': 'MONEY0',
      'Q-L (Months)': 'INT'
    }),

    // Admin operational sheets (support Last Status Date with AUTO handling)
    'Ask Detail': Object.freeze({
      'Last Status Date': 'DATE_AUTO'
    }),
    'Start': Object.freeze({
      'Last Status Date': 'DATE_AUTO'
    }),
    'Finish': Object.freeze({
      'Last Status Date': 'DATE_AUTO'
    }),
    'OR': Object.freeze({
      'Last Status Date': 'DATE_AUTO',
      'OR': 'CHECKBOX'
    }),

    // Default minimal typing
    'Exclusion': Object.freeze({
      'Submission Date': 'DATE',
      'Submitted Datetime': 'DATETIME',
      'Timestamp': 'TIMESTAMP',
      'Last Status Date': 'DATE_AUTO',
      'TAT': 'INT',
      'Sum Insured': 'MONEY0',
      'Nett Claim Amount': 'MONEY0',
      'Q-L (Months)': 'INT'
    }),

    '_DEFAULT_': Object.freeze({
      'Submission Date': 'DATE',
      'Submitted Datetime': 'DATETIME',
      'Timestamp': 'TIMESTAMP',
      'Last Status Date': 'DATE_AUTO',
      'LSA': 'INT',
      'ALA': 'INT',
      'TAT': 'INT',
      'Sum Insured': 'MONEY0',
      'Nett Claim Amount': 'MONEY0',
      'Sum Insured Amount': 'MONEY0',
      'Claim Amount': 'MONEY0',
      'Claim Own Risk Amount': 'MONEY0',
      '% Approval': 'NUMBER',
      'Q-L (Months)': 'INT'
    })
  })
});

/** Alignment registry (applied by 03/05 in one batch per sheet) */
const COLUMN_ALIGNMENT = Object.freeze({
  CENTER: Object.freeze([
    'Submission Date',
    'Submitted Datetime',
    'Timestamp',
    'Last Status Date',
    'OR'
  ]),
  RIGHT: Object.freeze([
    // Aging
    'Last Status Aging',
    'Activity Log Aging',
    'TAT',
    'LSA', 'ALA', // legacy

    // Amounts
    'Sum Insured Amount',
    'Sum Insured', // legacy
    'Claim Amount',
    'Repair/Replace Amount', // legacy
    'Claim Own Risk Amount',
    'OR Amount', // legacy
    'Nett Claim Amount',
    'Selisih',

    // Other numeric
    'Q-L (Months)',
    '% Approval'
  ]),
  LEFT: Object.freeze([]) // default fallback for everything else
});


/** -------------------------
 * Raw Data column ordering (end-of-run normalization)
 * -------------------------
 * Requirement: reorder these headers to the front (keep all other columns after, preserving relative order).
 */
const RAW_DATA_REORDER_POLICY = Object.freeze({
  // Disabled by default to preserve custom tail columns on the far-right of Raw Data.
  ENABLE: getPropBool_('RAW_DATA_REORDER_ENABLE', false),
  SHEET_NAME: 'Raw Data',
  PRIORITY_HEADERS: Object.freeze([
    
    'qoala_policy_number',
    'source_system_name',
    'claim_number',
    'claim_submission_date',
    'policy_start_date',
    'policy_end_date',
    'last_status',
    'last_update',
    'claim_last_updated_datetime',
    'month_policy_aging',
    'last_status_aging',
    'last_activity_log_date',
    'days_to_end_policies',
    'business_partner_name',
    'insurance_partner_name',
    'insurance_partner_code',
    'dashboard_link',
    'sc_name',
    'device_type',
    'DB',
    'device_brand',
    'imei_number',
    'sum_insured_amount',
    'claim_amount',
    'claim_own_risk_amount',
    'nett_claim_amount'
  
  ])
});


const CONFIG = Object.freeze({
  sectionIndex: CONFIG_SECTION_INDEX,
  spreadsheets: {
    // Single master workbook (all flows write here)
    Master: MASTER_SPREADSHEET_ID,

    // Backward-compat keys (older modules still call these)
    Farhan: MASTER_SPREADSHEET_ID,
    Admin:  MASTER_SPREADSHEET_ID,
    Meilani:MASTER_SPREADSHEET_ID,
    Suci:   MASTER_SPREADSHEET_ID,
    Adi:    MASTER_SPREADSHEET_ID
  },

  // Master workbook identifiers
  masterSpreadsheetId: MASTER_SPREADSHEET_ID,
  masterRawSheetName: MASTER_RAW_SHEET_NAME,

  // Ingestion policies
  emailIngest: EMAIL_INGEST_POLICY,
  subEmailIngest: SUB_EMAIL_INGEST_POLICY,
  subFlow: SUB_FLOW_SPEC,
  formIngestPolicy: FORM_INGEST_POLICY,
  rawDataCustomTailHeaders: RAW_DATA_CUSTOM_TAIL_HEADERS,

  // Mapping Team Claim is deprecated (process all data into master; no PIC routing)
  mappingEnabled: false,

  mappingSpreadsheetId: getPropString_('MAPPING_SPREADSHEET_ID', '1-Und57aVtmLYEovFuNyHQjvIrV8BfpyQVTHiDiWcpvk'),
  mappingSheetName: getPropString_('MAPPING_SHEET_NAME', '[UPDATED] Mapping Team Claim'),

  logSpreadsheetId: getPropString_('LOG_SPREADSHEET_ID', '1TC9YjDo6qxWq17zPYEBqIryhaYbUMqGtaSH0F-G8IwE'),
  logSheetName: getPropString_('LOG_SHEET_NAME', 'Log'),
  detailsSheetName: 'Details',
  detailsLogPolicy: DETAILS_LOG_POLICY,
  mappingErrorLogPolicy: MAPPING_ERROR_LOG_POLICY,
  logPolicy: LOG_POLICY,

  // WebApp movement tracking
  webappProjectSpreadsheetId: WEBAPP_PROJECT_SPREADSHEET_ID,
  webappMovement: WEBAPP_MOVEMENT_POLICY,

  // Status type mapping (mandatory operational column)
  statusTypeByLastStatus: STATUS_TYPE_BY_LAST_STATUS,
  positionByLastStatus: POSITION_BY_LAST_STATUS,

  // Second-year policy (STRICT)
  secondYearMarketValue: SECOND_YEAR_MARKET_VALUE_POLICY,

  // Highlighting + routing aliases used by older modules / maintenance patches
  claimHighlightPolicy: CLAIM_HIGHLIGHT_POLICY,
  CLAIM_HIGHLIGHT_POLICY: CLAIM_HIGHLIGHT_POLICY,
  claimNumberHeaderAliases: CLAIM_HIGHLIGHT_POLICY.CLAIM_NUMBER_HEADER_ALIASES,
  CLAIM_NUMBER_HEADER_ALIASES: CLAIM_HIGHLIGHT_POLICY.CLAIM_NUMBER_HEADER_ALIASES,
  rawDataReorderPolicy: RAW_DATA_REORDER_POLICY,
  RAW_DATA_REORDER_POLICY: RAW_DATA_REORDER_POLICY,

  // Status/routing aliases kept for backward compatibility
  STATUS_TYPE_MAP: STATUS_TYPE_BY_LAST_STATUS,
  statusTypeMap: STATUS_TYPE_BY_LAST_STATUS,
  statusRouting: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET,
  statusRoutingSub: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET,
  SC_FALLBACK_SHEET: OPS_ROUTING_POLICY.SC_FALLBACK_SHEET,
  scFallbackSheet: OPS_ROUTING_POLICY.SC_FALLBACK_SHEET,
  SC_SHEET_ALLOWLISTS: OPS_ROUTING_POLICY.SC_NAME_KEYWORDS,

  // Feature flags (operational knobs)
  features: Object.freeze({
    // Schema validation (fail fast on unexpected header drift)
    strictSchemaValidation: getPropBool_('STRICT_SCHEMA_VALIDATION', false),

    // Run metrics sink (append summary per run)
    enableRunMetrics: getPropBool_('ENABLE_RUN_METRICS', true),

    // Durable task queue mode (experimental)
    useTaskQueue: getPropBool_('USE_TASK_QUEUE', false),

    // Activity Log column is OPTIONAL by user spec; default false means we will NOT auto-add it.
    ensureActivityLogColumn: getPropBool_('ENSURE_ACTIVITY_LOG_COLUMN', false),

    // Idempotency cache (transaction tokens)
    enableTxnIdempotency: getPropBool_('ENABLE_TXN_IDEMPOTENCY', true)
  }),

  validationPolicy: VALIDATION_POLICY,
  linkPolicy: LINK_POLICY,
  checkboxPolicy: CHECKBOX_POLICY,
  fixedSchemaPolicy: FIXED_SCHEMA_POLICY,
  dateCoercionPolicy: DATE_COERCION_POLICY,
  columnTypes: COLUMN_TYPES,
  columnAlignment: COLUMN_ALIGNMENT,
  associateColumnPolicy: ASSOCIATE_COLUMN_POLICY,
  columnPlacementRules: COLUMN_PLACEMENT_RULES,

  opsRouting: OPS_ROUTING_POLICY,
  opsRoutingPolicy: OPS_ROUTING_POLICY,
  opsRoutingPolicyV2: OPS_ROUTING_POLICY,

  /**
   * Workbook profiles determine which sheets are ensured/managed.
   * IMPORTANT: ADMIN must never create optional sheets.
   */
  workbookProfiles: Object.freeze({
    // Profiles kept for backward-compat, but both point to the same master workbook + sheet set.
    [WORKBOOK_PROFILES.PIC]: Object.freeze({
      core: Object.freeze([MASTER_RAW_SHEET_NAME]),
      operational: Object.freeze([
        OPS_ROUTING_POLICY.SHEETS.SUBMISSION,
        OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL,
        OPS_ROUTING_POLICY.SHEETS.OR_OLD,
        OPS_ROUTING_POLICY.SHEETS.START,
        OPS_ROUTING_POLICY.SHEETS.FINISH,
        OPS_ROUTING_POLICY.SHEETS.SC_FARHAN,
        OPS_ROUTING_POLICY.SHEETS.SC_MEILANI,
        OPS_ROUTING_POLICY.SHEETS.SC_IVAN,
        OPS_ROUTING_POLICY.SHEETS.PO,
        OPS_ROUTING_POLICY.SHEETS.EXCLUSION
      ]),
      optional: Object.freeze([]) // deprecated in new master flow
    }),
    [WORKBOOK_PROFILES.ADMIN]: Object.freeze({
      core: Object.freeze([MASTER_RAW_SHEET_NAME]),
      operational: Object.freeze([
        OPS_ROUTING_POLICY.SHEETS.SUBMISSION,
        OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL,
        OPS_ROUTING_POLICY.SHEETS.OR_OLD,
        OPS_ROUTING_POLICY.SHEETS.START,
        OPS_ROUTING_POLICY.SHEETS.FINISH,
        OPS_ROUTING_POLICY.SHEETS.SC_FARHAN,
        OPS_ROUTING_POLICY.SHEETS.SC_MEILANI,
        OPS_ROUTING_POLICY.SHEETS.SC_IVAN,
        OPS_ROUTING_POLICY.SHEETS.PO,
        OPS_ROUTING_POLICY.SHEETS.EXCLUSION
      ]),
      optional: Object.freeze([])
    })
  }),

  /**
   * Backward-compatible sheet buckets used by older code paths.
   * New code should use workbookProfiles above.
   */
  sheetsByPic: {
    // Backward-compatible buckets (new code should use workbookProfiles + OPS_ROUTING_POLICY)
    defaultMain: [MASTER_RAW_SHEET_NAME],
    picOperational: [
      OPS_ROUTING_POLICY.SHEETS.SUBMISSION,
      OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL,
      OPS_ROUTING_POLICY.SHEETS.OR_OLD,
      OPS_ROUTING_POLICY.SHEETS.START,
      OPS_ROUTING_POLICY.SHEETS.FINISH,
      OPS_ROUTING_POLICY.SHEETS.SC_FARHAN,
      OPS_ROUTING_POLICY.SHEETS.SC_MEILANI,
      OPS_ROUTING_POLICY.SHEETS.SC_IVAN,
      OPS_ROUTING_POLICY.SHEETS.PO,
      OPS_ROUTING_POLICY.SHEETS.EXCLUSION
    ],
    adminOperational: [
      OPS_ROUTING_POLICY.SHEETS.SUBMISSION,
      OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL,
      OPS_ROUTING_POLICY.SHEETS.OR_OLD,
      OPS_ROUTING_POLICY.SHEETS.START,
      OPS_ROUTING_POLICY.SHEETS.FINISH,
      OPS_ROUTING_POLICY.SHEETS.SC_FARHAN,
      OPS_ROUTING_POLICY.SHEETS.SC_MEILANI,
      OPS_ROUTING_POLICY.SHEETS.SC_IVAN,
      OPS_ROUTING_POLICY.SHEETS.PO,
      OPS_ROUTING_POLICY.SHEETS.EXCLUSION
    ],
    optional: []
  },

  filePrefixes: {
    main: '_qgp__id__claim_daily_monitoring',
    aging: 'list_of_claims_with_aging',
    agingStd: 'list_of_claims_with_aging__standardization'
  },

  patterns: {
    b2bPartners: [
      'EMG','GDN','Helios','MonsterAR','Buminet','AIMS','AIRMAS','Staffinc',
      'TIGA JEJAK LANGKAH','WPS','PPS','Archor'
    ],
    specialPartners: ['Renewal Qoala', 'Qoala Monsta'],
    evBikePartners: [
      'Best EV','EV SanMoto','Favoriet Ofero Aceh','Franada Ev Bike Batam','GODA',
      'King EV','KingEV','MDM Ofero EV-Bike','Niceral EV Bike','Ofero','Ofero Nusantara',
      'Otobot','Pacific','Pratama Motor','Sahabat EV- Ofero Store','U-Winfly HM Yamin Medan',
      'U-WINFLY MEDAN','Ukka Bike','Ukka Ebike PKY','United Bike'
    ]
  },

  statusRoutingPIC: {
    // Deprecated naming; kept so older modules compile.
    [OPS_ROUTING_POLICY.SHEETS.SUBMISSION]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Submission'],
    [OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Ask Detail'],
    [OPS_ROUTING_POLICY.SHEETS.OR_OLD]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['OR - OLD'],
    [OPS_ROUTING_POLICY.SHEETS.START]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Start'],
    [OPS_ROUTING_POLICY.SHEETS.FINISH]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Finish'],
    [OPS_ROUTING_POLICY.SHEETS.SC_FARHAN]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.SC_MEILANI]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.SC_IVAN]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.PO]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['PO'],
    [OPS_ROUTING_POLICY.SHEETS.EXCLUSION]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Exclusion']
  },

  statusRoutingAdmin: {
    // Deprecated naming; Admin is no longer a separate routing universe in the master workbook.
    [OPS_ROUTING_POLICY.SHEETS.SUBMISSION]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Submission'],
    [OPS_ROUTING_POLICY.SHEETS.ASK_DETAIL]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Ask Detail'],
    [OPS_ROUTING_POLICY.SHEETS.OR_OLD]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['OR - OLD'],
    [OPS_ROUTING_POLICY.SHEETS.START]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Start'],
    [OPS_ROUTING_POLICY.SHEETS.FINISH]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Finish'],
    [OPS_ROUTING_POLICY.SHEETS.SC_FARHAN]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.SC_MEILANI]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.SC_IVAN]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'],
    [OPS_ROUTING_POLICY.SHEETS.PO]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['PO'],
    [OPS_ROUTING_POLICY.SHEETS.EXCLUSION]: OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['Exclusion']
  },

  formFields: {
    // Deprecated: no longer used (single master sheet target).
    extractToSpreadsheet: '',

    // Kept for backward-compat (but new code should use FORM_INGEST_POLICY).
    fileUploadFieldName: FORM_INGEST_POLICY.FILE_UPLOAD_FIELD_NAME,

    // New: choose which pipeline to run from the same Form.
    flowFieldName: FORM_INGEST_POLICY.FLOW_FIELD_NAME,

    // Optional: 2 separate upload questions for SUB (OLD + NEW).
    subOldFileUploadFieldName: FORM_INGEST_POLICY.SUB_OLD_FILE_UPLOAD_FIELD_NAME,
    subNewFileUploadFieldName: FORM_INGEST_POLICY.SUB_NEW_FILE_UPLOAD_FIELD_NAME
  },

  headers: {
    claimNumber: 'claim_number',
    businessPartner: 'business_partner_name',
    productName: 'product_name',
    policyStartDate: 'policy_start_date',
    policyEndDate: 'policy_end_date',
    claimSubmissionDate: 'claim_submission_date',
    claimSubmittedDatetime: 'claim_submitted_datetime',
    lastActivityLogDate: 'last_activity_log_date',
    lastUpdate: 'last_update',
    // NEW: Sub flow datetime sources
    claimLastUpdatedDatetime: 'claim_last_updated_datetime',
    activityLogDatetime: 'activity_log_datetime',
    lastActivityLogDatetime: 'last_activity_log_datetime',
    activityLog: 'activity_log',
    lastActivityLog: 'last_activity_log',
    // NEW: Strict Second-Year source
    monthPolicyAging: 'month_policy_aging',
    daysToEndPolicy: 'days_to_end_policies',
    daysAgingFromSubmission: 'days_aging_from_submission',
    insuranceCode: 'insurance_code',
    lastStatus: 'last_status',
    lastStatusAging: 'last_status_aging',
    activityLogAging: 'activity_log_aging',
    dashboardLink: 'dashboard_link',
    sourceSystem: 'source_system_name',
        // Aliases for backward/optional sheet compatibility
    sourceDb: 'source_system_name',
    serviceCenter: 'sc_name',
    agingFromLastStatus: 'activity_log_aging',
    tat: 'days_aging_from_submission',
    sumInsured: 'sum_insured_amount',
partnerCodeAging: 'insurance_partner_code',
    sumInsuredAmount: 'sum_insured_amount',
    ownRiskAmount: 'claim_own_risk_amount',
    nettClaimAmount: 'nett_claim_amount',
    qLMonths: 'Q-L (Months)',
    associate: 'Associate',
    orColumn: 'OR',
    updateStatus: 'Update Status',
    timestamp: 'Timestamp',
    status: 'Status',
    deviceType: 'device_type',
    scName: 'sc_name',
    partnerName: 'business_partner_name',
    partner: 'business_partner_name',
    product: 'product_name',
    // Alias for operational column Insurance
    insurancePartnerName: 'insurance_partner_name',
    insuranceName: 'insurance_partner_name',
    insurance: 'insurance_partner_name',
    deviceBrand: 'device_brand',
    brand: 'device_brand',
    imeiNumber: 'imei_number',
    // Alias used by some attachments
    serialNumber: 'imei_number',
    claimAmount: 'claim_amount',
    claimOwnRiskAmount: 'claim_own_risk_amount',
    // Classify OLD/NEW
    db: 'DB',
    dbClass: 'DB',
    dbStatus: 'DB',
    ala: 'activity_log_aging',
    ALA: 'activity_log_aging',
    TAT: 'days_aging_from_submission',
    LSA: 'last_status_aging',
    // Alias for raw header candidates
    device_type: 'device_type',
  }
});

/** Runtime state (mutated during a run) */
const RUNTIME = {
  useSubmittedDatetime: false,
  // Optional sheets live in the same master spreadsheet; keep them enabled by default.
  enableEvBike: true,
  enableB2B: true,
  enableSpecialCase: true,
  hasAgingFiles: false,
  detailsAppendedThisRun: 0,

  // Caches (reset every run)
  excludedLastStatusesSet: null,
  evBikeExcludedPolicySet: null,

  // Run metadata
  runStartedAt: null,
  requestId: null,
  flowName: null,
  transactionToken: null
};



/**
 * Reset per-run runtime state.
 * MUST be called as the first line in the entrypoint (01/06) before doing any work.
 */
function resetRuntime_() {
  RUNTIME.useSubmittedDatetime = false;
  // Optional sheets: default ON (can be disabled explicitly by other modules if needed).
  RUNTIME.enableEvBike = true;
  RUNTIME.enableB2B = true;
  RUNTIME.enableSpecialCase = true;
  RUNTIME.hasAgingFiles = false;
  RUNTIME.detailsAppendedThisRun = 0;

  RUNTIME.excludedLastStatusesSet = null;
  RUNTIME.evBikeExcludedPolicySet = null;

  RUNTIME.runStartedAt = new Date();
  RUNTIME.requestId = null;
  RUNTIME.flowName = null;
  RUNTIME.transactionToken = null;
}


/** =========================
 * System introspection utilities
 * ========================= */

/**
 * Produce a JSON summary of current config + registered flows/modules.
 * Best-effort: appends a row to "_SystemCatalog" sheet in Log spreadsheet.
 *
 * @return {Object} summary
 */
function describeSystem_() {
  const snap = (App && App.Registry && typeof App.Registry.snapshot === 'function')
    ? App.Registry.snapshot()
    : { flows: {}, modules: {} };

  const summary = {
    generatedAt: new Date().toISOString(),
    appVersion: (App && App.APP_VERSION) ? App.APP_VERSION : '',
    schemaVersion: (typeof SCHEMA_VERSION !== 'undefined') ? SCHEMA_VERSION : null,
    timezone: (function () { try { return Session.getScriptTimeZone(); } catch (e) { return ''; } })(),
    config: {
      masterSpreadsheetId: (CONFIG && CONFIG.masterSpreadsheetId) ? CONFIG.masterSpreadsheetId : '',
      masterRawSheetName: (CONFIG && CONFIG.masterRawSheetName) ? CONFIG.masterRawSheetName : '',
      logSpreadsheetId: (CONFIG && CONFIG.logSpreadsheetId) ? CONFIG.logSpreadsheetId : '',
      logSheetName: (CONFIG && CONFIG.logSheetName) ? CONFIG.logSheetName : '',
      webappProjectSpreadsheetId: (CONFIG && CONFIG.webappProjectSpreadsheetId) ? CONFIG.webappProjectSpreadsheetId : '',
      features: (CONFIG && CONFIG.features) ? CONFIG.features : {}
    },
    registry: snap
  };

  try { Logger.log(JSON.stringify(summary)); } catch (e0) {}

  try {
    if (!DRY_RUN && CONFIG && CONFIG.logSpreadsheetId) {
      const ss = SpreadsheetApp.openById(String(CONFIG.logSpreadsheetId));
      const name = '_SystemCatalog';
      const sh = ss.getSheetByName(name) || ss.insertSheet(name);
      const header = ['Timestamp', 'App Version', 'Schema Version', 'JSON'];
      if (sh.getLastRow() < 1) {
        sh.getRange(1, 1, 1, header.length).setValues([header]);
        try { sh.setFrozenRows(1); } catch (eF) {}
      }
      const row = [new Date(), summary.appVersion, summary.schemaVersion, JSON.stringify(summary)];
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    }
  } catch (e1) {}

  return summary;
}

/**
 * Quick health check for required spreadsheets and critical sheets.
 * This does NOT mutate business data.
 *
 * @return {Object} result
 */
function healthCheck_() {
  const res = {
    checkedAt: new Date().toISOString(),
    ok: true,
    master: { ok: false, error: '' },
    log: { ok: false, error: '' },
    webapp: { ok: false, error: '' }
  };

  // Master
  try {
    const ss = SpreadsheetApp.openById(String(CONFIG.masterSpreadsheetId));
    const raw = ss.getSheetByName(String(CONFIG.masterRawSheetName || 'Raw Data'));
    res.master.ok = !!raw;
    if (!raw) res.master.error = 'Missing Raw sheet: ' + String(CONFIG.masterRawSheetName || 'Raw Data');
  } catch (eM) {
    res.master.ok = false;
    res.master.error = String(eM && eM.message ? eM.message : eM);
  }

  // Log workbook
  try {
    const ss = SpreadsheetApp.openById(String(CONFIG.logSpreadsheetId));
    const logSh = ss.getSheetByName(String(CONFIG.logSheetName || 'Log'));
    res.log.ok = !!logSh;
    if (!logSh) res.log.error = 'Missing Log sheet: ' + String(CONFIG.logSheetName || 'Log');
  } catch (eL) {
    res.log.ok = false;
    res.log.error = String(eL && eL.message ? eL.message : eL);
  }

  // WebApp Project
  try {
    const ss = SpreadsheetApp.openById(String(CONFIG.webappProjectSpreadsheetId));
    const daily = ss.getSheetByName(WEBAPP_MOVEMENT_POLICY.SHEETS.DAILY);
    const past = ss.getSheetByName(WEBAPP_MOVEMENT_POLICY.SHEETS.PAST);
    res.webapp.ok = !!(daily && past);
    if (!res.webapp.ok) res.webapp.error = 'Missing Daily/Past sheet(s) in WebApp Project';
  } catch (eW) {
    res.webapp.ok = false;
    res.webapp.error = String(eW && eW.message ? eW.message : eW);
  }

  res.ok = !!(res.master.ok && res.log.ok && res.webapp.ok);
  try { Logger.log(JSON.stringify(res)); } catch (e0) {}
  return res;
}
