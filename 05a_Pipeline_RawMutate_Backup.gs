/***************************************
 * 05a_Pipeline_RawMutate_Backup.gs
 * Split from: 05_Pipeline_Routing_Optional.gs
 * Scope:
 *  - Safe fallbacks (FORMATS / OPTIONAL flags / SPECIAL flags)
 *  - RAW write helpers
 *  - Typed date/OR helpers
 *  - Associate mapping + Raw in-memory mutation (aging + QL/ML/MQ + Associate + OR defaulting)
 *  - Backup operational columns -> Raw (in-memory)
 *  - Raw minimal column writeback
 ***************************************/

/** REVISION NOTE (2026-01-26)
 * - Added 'SC - Ivan' to the last-resort operational sheet list in getOperationalSheetsForBackup_(),
 *   aligned with 3-way SC routing (Farhan/Meilani/Ivan).
 */

/***************************************
 * 05_Pipeline_Routing_Optional.gs
 * Enterprise refactor:
 * - Typed date/datetime/timestamp writes (no mixed string dates)
 * - Avoid validation violations (Associate unknown => blank)
 * - Admin skips optional sheets (B2B / Special Case / EV-Bike)
 * - Special Case rules per request (year>=MIN_SUBMISSION_YEAR, Q-L 0/1, exclude done statuses, numeric money)
 * - OR standardized:
 *    - OR        = checkbox (TRUE/FALSE) for operational sheets
 *    - OR Amount = claim_own_risk_amount (numeric) for money usage (Special Case, or any sheet with "OR Amount")
 * - Consistent column formatting & alignment (minimal calls)
 ***************************************/
'use strict';


/** Normalize header cell text: trim + remove BOM + replace NBSP + collapse whitespace. */
function __normalizeHeaderText05a_(v) {
  return String(v == null ? '' : v)
    .replace(/^\uFEFF/, '')       // BOM
    .replace(/\u00A0/g, ' ')      // NBSP
    .replace(/\s+/g, ' ')         // collapse spaces/tabs/newlines
    .trim();
}


/** Case-insensitive exact header matcher (backup/restore resilience). */
function __findHeaderIndex05a_(headerRow, name) {
  try {
    if (typeof findHeaderIndexByCandidates_ === 'function') {
      const idx = findHeaderIndexByCandidates_(headerRow, [name]);
      return (typeof idx === 'number' && idx >= 0) ? idx : -1;
    }
  } catch (e) {}
  const want = __normalizeHeaderText05a_(name).toLowerCase();
  if (!want) return -1;
  const h = Array.isArray(headerRow) ? headerRow : [];
  for (let i = 0; i < h.length; i++) {
    const v = __normalizeHeaderText05a_(h[i]).toLowerCase();
    if (v === want) return i;
  }
  return -1;
}


/** Best-effort resolver for the Raw sheet (used to persist manual fields like Remarks). */
function __resolveRawSheet05a_(ss) {
  if (!ss) return null;

  // Prefer CONFIG if provided
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.sheets) {
      const candidates = []
        .concat(CONFIG.sheets.rawData || [])
        .concat(CONFIG.sheets.raw || [])
        .concat(CONFIG.sheets.raw_sheet || [])
        .concat(CONFIG.sheets.rawSheet || []);
      for (let i = 0; i < candidates.length; i++) {
        const name = candidates[i];
        if (name && typeof name === 'string') {
          const sh = ss.getSheetByName(name);
          if (sh) return sh;
        }
      }
    }
  } catch (e) {}

  // Common defaults
  const names = ['Raw Data', 'RawData', 'RAW DATA', 'Raw', 'RAW'];
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}



// ---- Safe fallbacks (avoid ReferenceError) ----


// ---- Safe fallbacks (avoid ReferenceError / TDZ across files) ----
var __FORMATS = (function () {
  try { if (typeof FORMATS !== 'undefined' && FORMATS) return FORMATS; } catch (e) {}
  try { if (typeof _FORMATS !== 'undefined' && _FORMATS) return _FORMATS; } catch (e) {}
  return Object.freeze({
    DATE: 'd mmm yy',
    DATETIME: 'd mmm yy, HH:mm',
    DATETIME_LONG: 'MMMM d, yyyy, h:mm AM/PM',
    TIMESTAMP: 'd mmm, HH:mm',
    INT: '0',
    MONEY0: '#,##0'
  });
})();

var __OPTIONAL_FLAGS = (function () {
  try { if (typeof OPTIONAL_SHEETS_FLAGS !== 'undefined' && OPTIONAL_SHEETS_FLAGS) return OPTIONAL_SHEETS_FLAGS; } catch (e) {}
  return Object.freeze({});
})();

var __SPECIAL_FLAGS = (function () {
  try { if (typeof SPECIAL_CASE_FLAGS !== 'undefined' && SPECIAL_CASE_FLAGS) return SPECIAL_CASE_FLAGS; } catch (e) {}
  return Object.freeze({ MODE: 'UPSERT', COLORIZE_CLAIM_CELL: false });
})();

var __SPECIAL_RULES = (function () {
  try { if (typeof SPECIAL_CASE_RULES !== 'undefined' && SPECIAL_CASE_RULES) return SPECIAL_CASE_RULES; } catch (e) {}
  return Object.freeze({});
})();

function __getExcludedLastStatuses05a_() {
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.__excludedLastStatuses05a) {
      return RUNTIME.__excludedLastStatuses05a;
    }
  } catch (e0) {}

  let out = null;

  // Prefer per-run builder from 00_Config if available.
  try { if (typeof getExcludedLastStatusesSet_ === 'function') out = getExcludedLastStatusesSet_(); } catch (e) {}
  if (!out) {
    try { if (typeof buildExcludedLastStatusesSet_ === 'function') out = buildExcludedLastStatusesSet_(); } catch (e2) {}
  }

  // Back-compat: older versions exposed a Set directly.
  if (!out) {
    try {
      if (typeof EXCLUDED_LAST_STATUSES !== 'undefined' && EXCLUDED_LAST_STATUSES && EXCLUDED_LAST_STATUSES.has) {
        out = EXCLUDED_LAST_STATUSES;
      }
    } catch (e3) {}
  }

  // Newer versions may expose a base array.
  if (!out) {
    try {
      if (typeof EXCLUDED_LAST_STATUSES_BASE !== 'undefined' && Array.isArray(EXCLUDED_LAST_STATUSES_BASE)) {
        out = new Set(EXCLUDED_LAST_STATUSES_BASE);
      }
    } catch (e4) {}
  }

  if (!out || !out.has) out = new Set();

  try { if (typeof RUNTIME !== 'undefined' && RUNTIME) RUNTIME.__excludedLastStatuses05a = out; } catch (e5) {}
  return out;
}

/** =========================
 *  RAW WRITE (Main -> Raw)
 *  ========================= */

/** Write main data into Raw (never overwrite validated/manual Status col) */
function writeMainDataToRaw_(rawSheet, mainHeader, mainRows, headerIndexRaw) {
  if (!rawSheet || !mainRows || !mainRows.length) return;

  const pairs = [];
  for (let i = 0; i < (mainHeader || []).length; i++) {
    const key = String(mainHeader[i] || '').trim();
    if (!key) continue;
    const keyLc = key.toLowerCase();
    if (key === CONFIG.headers.status || key === 'Status' || keyLc === 'remarks' || keyLc === 'remark') continue; // never overwrite validated/manual cols

    const rawIdx = headerIndexRaw[key];
    if (rawIdx == null) continue;
    pairs.push({ rawIdx, mainIdx: i });
  }
  if (!pairs.length) return;
  pairs.sort((a, b) => a.rawIdx - b.rawIdx);

  // contiguous runs => fewer setValues
  const runs = [];
  let cur = [pairs[0]];
  for (let k = 1; k < pairs.length; k++) {
    const p = pairs[k - 1], n = pairs[k];
    if (n.rawIdx === p.rawIdx + 1) cur.push(n);
    else { runs.push(cur); cur = [n]; }
  }
  runs.push(cur);

  for (let r = 0; r < runs.length; r++) {
    const run = runs[r];
    const startCol1 = run[0].rawIdx + 1;
    const width = run.length;

    const out = new Array(mainRows.length);
    for (let i = 0; i < mainRows.length; i++) {
      const src = mainRows[i];
      const rowOut = new Array(width);
      for (let j = 0; j < width; j++) rowOut[j] = src[run[j].mainIdx];
      out[i] = rowOut;
    }
    safeSetValues_(rawSheet.getRange(2, startCol1, out.length, width), out);
  }
}

/** =========================
 *  Typed date helpers
 *  ========================= */

function coerceDate_(v) {
  const d = normalizeDate_(v);
  return d ? d : null;
}

function coerceDateOnly_(v) {
  const d = coerceDate_(v);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

function coerceDateTime_(v) {
  const d = coerceDate_(v);
  return d ? d : null;
}

/** Timestamp format requested: "15 Dec, 15:17" -> FORMATS.TIMESTAMP */
function coerceTimestamp_(v) {
  return coerceDateTime_(v);
}

/** Submission output rule (source picking) */
function getSubmissionDateForOutput_(rawRow, headerIndexRaw) {
  const h = CONFIG.headers;
  const idxDt = headerIndexRaw[h.claimSubmittedDatetime];
  const dtVal = (idxDt != null) ? rawRow[idxDt] : '';

  // Source of truth: claim_submitted_datetime only.
  if (dtVal) return { val: dtVal, mode: 'datetime' };
  return { val: '', mode: 'date' };
}

function buildSubmissionDateCell_(rawRow, headerIndexRaw) {
  const picked = getSubmissionDateForOutput_(rawRow, headerIndexRaw);
  if (!picked.val) return '';
  const d = coerceDateOnly_(picked.val);
  return d ? d : '';
}

function sanitizeAssociateForWrite_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (s === 'Meilani' || s === 'Farhan' || s === 'Suci' || s === 'Adi') return s;
  return '';
}

/** OR checkbox output (operational): prefer Raw OR checkbox; else derive from OR Amount > 0 */
function getOrCheckboxForOutput_(rawRow, headerIndexRaw) {
  const h = CONFIG.headers;
  const idxOr = headerIndexRaw[h.orColumn];
  const idxOwnRisk = idxAny_(headerIndexRaw, [h.ownRiskAmount, 'Claim Own Risk Amount', 'claim_own_risk_amount', 'Own Risk Amount']);

  // Prefer explicit checkbox value if valid
  if (idxOr != null) {
    const cb = normalizeCheckbox_(rawRow[idxOr]);
    if (cb !== '') return cb; // boolean
  }

  // Fallback: derive from amount
  const amt = (idxOwnRisk != null) ? normalizeNumber_(rawRow[idxOwnRisk]) : null;
  if (amt != null && amt > 0) return true;

  return ''; // keep blank to avoid forcing unchecked on legacy sheets
}

/** OR Amount output (money): prefer claim_own_risk_amount; fallback legacy numeric in Raw OR if exists */
function getOrAmountForOutput_(rawRow, headerIndexRaw) {
  const h = CONFIG.headers;
  const idxOwnRisk = idxAny_(headerIndexRaw, [h.ownRiskAmount, 'Claim Own Risk Amount', 'claim_own_risk_amount', 'Own Risk Amount']);
  const idxOr = headerIndexRaw[h.orColumn];

  const ownRisk = (idxOwnRisk != null) ? normalizeNumber_(rawRow[idxOwnRisk]) : null;
  if (ownRisk != null) return ownRisk;

  // Legacy fallback (only if someone already wrote numeric into Raw OR)
  const rawOrNum = (idxOr != null) ? normalizeNumber_(rawRow[idxOr]) : null;
  return (rawOrNum != null) ? rawOrNum : '';
}

/** =========================
 *  Associate mapping
 *  ========================= */

/**
 * New spec (2026):
 * - Routing is no longer PIC/Associate-based.
 * - The mapping sheet "[UPDATED] Mapping Team Claim" is deprecated.
 * - Associate column in "Raw Data" is treated as a manual/operational field.
 *
 * If you *really* need the legacy mapping for some back-compat run, enable it via Script Properties:
 *   ASSOCIATE_MAPPING_ENABLED = true
 */

// [Refactor] Removed duplicate definition of isAssociateMappingEnabled_ (source of truth in 01_Utils)


/** Convert column letter (e.g. "A", "G", "AA") into 1-based column index. */
function colA1ToIndex_(colA1) {
  const s = String(colA1 || '').trim().toUpperCase();
  if (!s) return null;
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 65 || c > 90) continue; // skip non A-Z
    n = n * 26 + (c - 64);
  }
  return n || null;
}

function loadAssociateMapping_() {
  if (!isAssociateMappingEnabled_()) {
    return { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };
  }
    // Deprecated: Associate mapping is off by default.
  // If someone enables it, keep the pipeline resilient (no hard-fail if mapping doc/sheet is missing).
  let ss, sh;
  try {
    if (!CONFIG || !CONFIG.mappingSpreadsheetId || !CONFIG.mappingSheetName) {
      return { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };
    }
    ss = SpreadsheetApp.openById(CONFIG.mappingSpreadsheetId);
    sh = ss.getSheetByName(CONFIG.mappingSheetName);
    if (!sh) {
      try { Logger && Logger.log && Logger.log('Associate mapping sheet not found (ignored): ' + CONFIG.mappingSheetName); } catch (e) {}
      return { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };
    }
  } catch (e) {
    try { Logger && Logger.log && Logger.log('Associate mapping load failed (ignored): ' + e); } catch (e2) {}
    return { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };
  }

  // Mapping policy (safe fallback when 00_Config is not loaded / renamed).
  const startRow = (function () {
    try {
      if (typeof MAPPING_TEAM_CLAIM_POLICY !== 'undefined' &&
          MAPPING_TEAM_CLAIM_POLICY &&
          MAPPING_TEAM_CLAIM_POLICY.PARTNER_START_ROW) return MAPPING_TEAM_CLAIM_POLICY.PARTNER_START_ROW;
    } catch (e) {}
    return 4;
  })();

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };

  // Existing columns (legacy): A=Meilani, C=Farhan, E=Suci
  const colByAssociate = { Meilani: 1, Farhan: 3, Suci: 5, Adi: 7 }; // Adi default = G

  // Spec: Adi partner list is sourced from "[UPDATED] Mapping Team Claim" starting at G4.
  try {
    if (typeof MAPPING_TEAM_CLAIM_POLICY !== 'undefined' &&
        MAPPING_TEAM_CLAIM_POLICY &&
        MAPPING_TEAM_CLAIM_POLICY.PARTNER_START_COLUMN_BY_PIC &&
        MAPPING_TEAM_CLAIM_POLICY.PARTNER_START_COLUMN_BY_PIC.Adi) {
      const n = colA1ToIndex_(MAPPING_TEAM_CLAIM_POLICY.PARTNER_START_COLUMN_BY_PIC.Adi);
      if (n) colByAssociate.Adi = n;
    }
  } catch (e) {}

  const maxCol = Math.max.apply(null, Object.keys(colByAssociate).map(k => colByAssociate[k]));
  const data = sh.getRange(startRow, 1, lastRow - startRow + 1, maxCol).getValues();

  const map = { Meilani: new Set(), Farhan: new Set(), Suci: new Set(), Adi: new Set() };
  for (let i = 0; i < data.length; i++) {
    Object.keys(colByAssociate).forEach(name => {
      const col0 = colByAssociate[name] - 1;
      const v = (col0 >= 0 && col0 < data[i].length) ? String(data[i][col0] || '').trim() : '';
      if (v) map[name].add(v.toLowerCase());
    });
  }
  return map;
}
/**
 * Single-pass Raw mutation in memory:
 * - Apply aging merge OR compute LSA/ALA if no aging files
 * - Compute derived QL/ML/MQ
 * - Apply Associate mapping (unknown => blank to avoid validation violation)
 * - OR checkbox defaulting: if Raw OR blank/invalid, set TRUE when claim_own_risk_amount > 0
 */
function mutateRawInMemory_(rawValues, headerIndexRaw, agingStdMap, agingMap, assocMap) {
  const h = CONFIG.headers;
  const today = new Date();
  const resolveHdrIdx05a_ = (keys) => {
    if (typeof idxAny_ === 'function') {
      const idx = idxAny_(headerIndexRaw, keys, { enableSnakeCase: true });
      return (idx == null) ? null : idx;
    }
    for (let i = 0; i < (keys || []).length; i++) {
      const k = keys[i];
      if (!k) continue;
      const v = headerIndexRaw[k];
      if (v != null) return v;
    }
    return null;
  };


  const idxClaim = resolveHdrIdx05a_([h.claimNumber, 'Claim Number', 'claim_number']);
  const idxBP = resolveHdrIdx05a_([h.businessPartner, 'Business Partner', 'Partner', 'business_partner']);
  const idxAssoc = resolveHdrIdx05a_([h.associate, 'Associate', 'associate']);
  const idxLastStatus = resolveHdrIdx05a_([h.lastStatus, 'Last Status', 'last_status']);

  const idxInsuranceCodeRaw = idxAny_(headerIndexRaw, [h.insuranceCode, 'Insurance Code', 'insurance_code']);
  const idxLastStatusAgingRaw = idxAny_(headerIndexRaw, ['Last Status Aging', h.lastStatusAging, 'LSA', 'TAT']);
  const idxActLogAgingRaw = idxAny_(headerIndexRaw, ['Activity Log Aging', h.activityLogAging, 'ALA']);
  const idxClaimSubmittedRaw = (headerIndexRaw['claim_submitted_datetime'] != null) ? headerIndexRaw['claim_submitted_datetime'] : null;

  const idxLastUpdate = resolveHdrIdx05a_([h.lastUpdate, 'Last Update', 'last_update']);
  const idxLastActDate = idxAny_(headerIndexRaw, [h.lastActivityLogDate, 'Last Activity Log Date', 'Last Activity Date', 'last_activity_log_date']);

  const idxPolicyStart = resolveHdrIdx05a_([h.policyStartDate, 'Policy Start Date', 'policy_start_date']);
  const idxPolicyEnd = resolveHdrIdx05a_([h.policyEndDate, 'Policy End Date', 'policy_end_date']);
  const idxClaimSubDt = idxClaimSubmittedRaw;

  const idxQL = headerIndexRaw['Q-L (Months)'];
  const idxML = headerIndexRaw['M-L (Months)'];
  const idxMQ = headerIndexRaw['M-Q (Months)'];

  const idxRawOR = idxAny_(headerIndexRaw, [h.orColumn, 'OR', 'Or']);
  const idxOwnRisk = idxAny_(headerIndexRaw, [h.ownRiskAmount, 'Claim Own Risk Amount', 'claim_own_risk_amount', 'Own Risk Amount']);

  const applyAgingFrom_ = (mapObj, row, claim) => {
    if (!mapObj || !mapObj.map) return false;
    const src = mapObj.map[claim];
    if (!src) return false;

    const hA = mapObj.headerIndex;
    const idxInsCodeA = idxAny_(hA, [h.partnerCodeAging, h.insuranceCode, 'Insurance Code', 'insurance_code', 'Partner Code']);
    const idxLSA = idxAny_(hA, ['Last Status Aging', h.lastStatusAging, 'LSA', 'TAT']);
    const idxALA = idxAny_(hA, ['Activity Log Aging', h.activityLogAging, 'ALA']);
    const idxSubDtA = (hA['claim_submitted_datetime'] != null) ? hA['claim_submitted_datetime'] : null;

    if (idxInsuranceCodeRaw != null && idxInsCodeA != null) row[idxInsuranceCodeRaw] = src[idxInsCodeA];
    if (idxLastStatusAgingRaw != null && idxLSA != null) row[idxLastStatusAgingRaw] = src[idxLSA];
    if (idxActLogAgingRaw != null && idxALA != null) row[idxActLogAgingRaw] = src[idxALA];
    if (idxClaimSubmittedRaw != null && idxSubDtA != null) row[idxClaimSubmittedRaw] = src[idxSubDtA];
    return true;
  };

  const unknownRowsAll = [];

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];

    const claimKey = (idxClaim != null) ? String(row[idxClaim] || '').trim() : '';
    if (claimKey) {
      if (agingStdMap) applyAgingFrom_(agingStdMap, row, claimKey);
      if (agingMap) applyAgingFrom_(agingMap, row, claimKey);
    }

    // If no aging files: compute LSA/ALA from dates
    if (!RUNTIME.hasAgingFiles && idxLastStatusAgingRaw != null && idxActLogAgingRaw != null) {
      if (idxLastUpdate != null) {
        const d = coerceDateTime_(row[idxLastUpdate]);
        const dd = diffDays_(d, today);
        if (dd != null && dd >= 0) row[idxLastStatusAgingRaw] = dd;
      }
      if (idxLastActDate != null) {
        const d2 = coerceDateTime_(row[idxLastActDate]);
        const dd2 = diffDays_(d2, today);
        if (dd2 != null && dd2 >= 0) row[idxActLogAgingRaw] = dd2;
      }
    }

    // Derived QL/ML/MQ (compute independently, don't require all dates)
    if (idxClaimSubDt != null) {
  const cs = coerceDateOnly_(row[idxClaimSubDt]);

  const ps = (idxPolicyStart != null) ? coerceDateOnly_(row[idxPolicyStart]) : null;
  const pe = (idxPolicyEnd != null) ? coerceDateOnly_(row[idxPolicyEnd]) : null;

  if (idxQL != null && cs && ps) row[idxQL] = monthDiff_(ps, cs);
  if (idxML != null && ps && pe) row[idxML] = monthDiff_(ps, pe);
  if (idxMQ != null && cs && pe) row[idxMQ] = monthDiff_(cs, pe);
}

    // OR checkbox defaulting: if Raw OR blank/invalid, set TRUE when claim_own_risk_amount > 0
    if (idxRawOR != null && idxOwnRisk != null) {
      const cb = normalizeCheckbox_(row[idxRawOR]);
      if (cb === '') {
        const amt = normalizeNumber_(row[idxOwnRisk]);
        if (amt != null && amt > 0) row[idxRawOR] = true;
      } else {
        // normalize to strict boolean (avoid "TRUE"/"FALSE" strings lingering)
        row[idxRawOR] = cb;
      }
    }
// Associate column is manual/operational in the 2026 routing spec: do not auto-map by default.
// If you need legacy mapping (deprecated), enable Script Properties: ASSOCIATE_MAPPING_ENABLED=true.
if (isAssociateMappingEnabled_() && idxBP != null && idxAssoc != null && assocMap) {
  const currentAssoc = String(row[idxAssoc] || '').trim();
  if (!currentAssoc) {
    const partnerRaw = String(row[idxBP] || '').trim();
    const partner = partnerRaw.toLowerCase();

    let assoc = '';
    if (assocMap.Meilani && assocMap.Meilani.has(partner)) assoc = 'Meilani';
    else if (assocMap.Farhan && assocMap.Farhan.has(partner)) assoc = 'Farhan';
    else if (assocMap.Suci && assocMap.Suci.has(partner)) assoc = 'Suci';
    else if (assocMap.Adi && assocMap.Adi.has(partner)) assoc = 'Adi';

    if (assoc) {
      row[idxAssoc] = assoc; // fill only when blank
    } else {
      unknownRowsAll.push({
        rowNumber: i + 2,
        claim: claimKey,
        partner: partnerRaw,
        lastStatus: (idxLastStatus != null) ? String(row[idxLastStatus] || '').trim() : '',
        submissionDateVal: (idxClaimSubDt != null) ? row[idxClaimSubDt] : ''
      });
    }
  }
}
  }

  const unknownRowsForLog = unknownRowsAll.filter(x => !isLegacyBySubmission_(x.submissionDateVal));
  return { values: rawValues, unknownRowsAll, unknownRowsForLog };
}

/** =========================
 *  Ops -> Raw backup
 *  ========================= */

/**
 * Operational sheets are no longer PIC-scoped in the 2026 routing spec.
 * This helper returns the current operational sheet list, with safe fallbacks.
 */
function getOperationalSheetsForBackup_(pic) { // pic kept for back-compat (ignored)
  // Prefer explicit routing policy (new spec)
  try {
    if (typeof OPS_ROUTING_POLICY !== 'undefined' &&
        OPS_ROUTING_POLICY &&
        Array.isArray(OPS_ROUTING_POLICY.OPERATIONAL_SHEETS)) {
      return OPS_ROUTING_POLICY.OPERATIONAL_SHEETS.slice();
    }
  } catch (e) {}

  // If CONFIG exposes a unified list
  try {
    if (typeof CONFIG !== 'undefined' &&
        CONFIG &&
        Array.isArray(CONFIG.operationalSheetsUnified)) {
      return CONFIG.operationalSheetsUnified.slice();
    }
  } catch (e) {}

  // Fallback to workbook profile PIC operational list
  try {
    if (typeof CONFIG !== 'undefined' &&
        CONFIG &&
        CONFIG.workbookProfiles &&
        CONFIG.workbookProfiles[WORKBOOK_PROFILES.PIC] &&
        Array.isArray(CONFIG.workbookProfiles[WORKBOOK_PROFILES.PIC].operational)) {
      return CONFIG.workbookProfiles[WORKBOOK_PROFILES.PIC].operational.slice();
    }
  } catch (e) {}

  // Legacy union fallback
  try {
    const a = (CONFIG && CONFIG.sheetsByPic && CONFIG.sheetsByPic.picOperational) ? CONFIG.sheetsByPic.picOperational : [];
    const b = (CONFIG && CONFIG.sheetsByPic && CONFIG.sheetsByPic.adminOperational) ? CONFIG.sheetsByPic.adminOperational : [];
    const out = Array.from(new Set([].concat(a, b)));
    if (out.length) return out;
  } catch (e) {}

  // Last-resort list per 2026 spec
  return [
    'Submission',
    'Ask Detail',
    'OR - OLD',
    'Start',
    'Finish',
    'SC - Farhan',
    'SC - Meilani',
    'SC - Ivan',
    'PO',
    'Exclusion'
  ];
}

/**
 * Backup ops -> raw (OR/Update/Timestamp/Status + Asso/Admin tail fields) using in-memory rawValues
 * IMPORTANT: OR is checkbox (TRUE/FALSE). We only accept checkbox-valid values.
 */
function backupOpsToRawInMemory_(ss, rawValues, headerIndexRaw, pic) {
  const h = CONFIG.headers;
  const idxRawClaim = headerIndexRaw[h.claimNumber];
  const idxRawOR = idxAny_(headerIndexRaw, [h.orColumn, 'OR', 'Or']);
  const idxRawUpdate = headerIndexRaw[h.updateStatus];
  const idxRawTs = headerIndexRaw[h.timestamp];
  const idxRawStatus = headerIndexRaw[h.status];
  const idxRawRemarks = idxAny_(headerIndexRaw, ['Remarks', 'Remark', 'remarks', 'remark']);

  // Additional manual/operational tail fields (snapshot per Claim Number)
  const idxRawUpdateAsso = headerIndexRaw['Update Status Asso'];
  const idxRawTsAsso = headerIndexRaw['Timestamp Asso'];
  const idxRawUpdateAdmin = headerIndexRaw['Update Status Admin'];
  const idxRawTsAdmin = headerIndexRaw['Timestamp Admin'];

  if (idxRawClaim == null) return { updated: 0, notes: 'Raw claim_number not found' };
  if (!rawValues || !rawValues.length) return { updated: 0, notes: 'Raw has no rows' };

  const rawMap = {};
  for (let i = 0; i < rawValues.length; i++) {
    const key = String(rawValues[i][idxRawClaim] || '').trim().toUpperCase();
    if (key) rawMap[key] = i;
  }

  const opsList = getOperationalSheetsForBackup_(pic);

  const fieldSet = {};
  const ensureObj = claim => (fieldSet[claim] ? fieldSet[claim] : (fieldSet[claim] = {}));

  let scannedRows = 0;

  for (let s = 0; s < opsList.length; s++) {
    const name = opsList[s];
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr <= 1 || lc <= 0) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v || '').trim());
    const idxClaimOps = header.indexOf('Claim Number');
    if (idxClaimOps === -1) continue;

    const idxOR = __findHeaderIndex05a_(header, 'OR');
    const idxUpdate = __findHeaderIndex05a_(header, 'Update Status');
    const idxTs = __findHeaderIndex05a_(header, 'Timestamp');
    const idxStatus = __findHeaderIndex05a_(header, 'Status');

    const idxUpdateAsso = __findHeaderIndex05a_(header, 'Update Status Asso');
    const idxTsAsso = __findHeaderIndex05a_(header, 'Timestamp Asso');
    const idxUpdateAdmin = __findHeaderIndex05a_(header, 'Update Status Admin');
    const idxTsAdmin = __findHeaderIndex05a_(header, 'Timestamp Admin');

    const idxRemarksTmp = __findHeaderIndex05a_(header, 'Remarks');
    const idxRemarks = (idxRemarksTmp !== -1) ? idxRemarksTmp : __findHeaderIndex05a_(header, 'Remark');

    if (idxOR === -1 && idxUpdate === -1 && idxTs === -1 && idxStatus === -1 &&
        idxUpdateAsso === -1 && idxTsAsso === -1 && idxUpdateAdmin === -1 && idxTsAdmin === -1 && idxRemarks === -1) continue;

    const values = sh.getRange(2, 1, lr - 1, lc).getValues();
    scannedRows += values.length;

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const claim = String(row[idxClaimOps] || '').trim().toUpperCase();
      if (!claim) continue;
      const rawIdx = rawMap[claim];
      if (rawIdx == null) continue;

      const obj = ensureObj(claim);

      if (idxOR !== -1 && obj.or == null) {
        const cb = normalizeCheckbox_(row[idxOR]);
        if (cb !== '') obj.or = cb;
      }
      if (idxUpdate !== -1 && obj.update == null) {
        const v = row[idxUpdate];
        if (v !== '' && v != null) obj.update = v;
      }
      if (idxTs !== -1 && obj.ts == null) {
        const v = row[idxTs];
        if (v !== '' && v != null) obj.ts = v;
      }
      if (idxStatus !== -1 && obj.status == null) {
        const v = normalizeStatusValue_(row[idxStatus]);
        if (v !== '') obj.status = v;
      }

      // Additional manual/ops tail fields (values-only snapshot)
      if (idxUpdateAsso !== -1 && obj.updateAsso == null) {
        const v = row[idxUpdateAsso];
        if (v !== '' && v != null) obj.updateAsso = v;
      }
      if (idxTsAsso !== -1 && obj.tsAsso == null) {
        const v = row[idxTsAsso];
        if (v !== '' && v != null) obj.tsAsso = v;
      }
      if (idxUpdateAdmin !== -1 && obj.updateAdmin == null) {
        const v = row[idxUpdateAdmin];
        if (v !== '' && v != null) obj.updateAdmin = v;
      }
      if (idxTsAdmin !== -1 && obj.tsAdmin == null) {
        const v = row[idxTsAdmin];
        if (v !== '' && v != null) obj.tsAdmin = v;
      }

      if (idxRemarks !== -1 && obj.remarks == null) {
        const v = row[idxRemarks];
        const t = String(v == null ? '' : v).trim();
        if (t) obj.remarks = t;
      }
    }
  }

  let updated = 0;
  let remarksTouched = 0;
  Object.keys(fieldSet).forEach(claim => {
    const i = rawMap[claim];
    if (i == null) return;
    const obj = fieldSet[claim];
    const row = rawValues[i];
    let touched = false;

    if (idxRawOR != null && obj.or != null) { row[idxRawOR] = obj.or; touched = true; }
    if (idxRawUpdate != null && obj.update != null) { row[idxRawUpdate] = obj.update; touched = true; }
    if (idxRawTs != null && obj.ts != null) { row[idxRawTs] = obj.ts; touched = true; }
    if (idxRawStatus != null && obj.status != null) { row[idxRawStatus] = obj.status; touched = true; }

    if (idxRawUpdateAsso != null && obj.updateAsso != null) { row[idxRawUpdateAsso] = obj.updateAsso; touched = true; }
    if (idxRawTsAsso != null && obj.tsAsso != null) { row[idxRawTsAsso] = obj.tsAsso; touched = true; }
    if (idxRawUpdateAdmin != null && obj.updateAdmin != null) { row[idxRawUpdateAdmin] = obj.updateAdmin; touched = true; }
    if (idxRawTsAdmin != null && obj.tsAdmin != null) { row[idxRawTsAdmin] = obj.tsAdmin; touched = true; }

    if (idxRawRemarks != null && obj.remarks != null) { row[idxRawRemarks] = obj.remarks; touched = true; }

    if (idxRawRemarks != null && obj.remarks != null) { row[idxRawRemarks] = obj.remarks; touched = true; remarksTouched++; }

    if (touched) updated++;
  });

  // Best-effort immediate writeback for Remarks, to prevent accidental loss when callers do minimal raw writes.
  if (idxRawRemarks != null && remarksTouched > 0) {
    try {
      const rawSheet = __resolveRawSheet05a_(ss);
      if (rawSheet) writeRawColumns_(rawSheet, rawValues, [idxRawRemarks]);
    } catch (e) {}
  }

  return { updated, notes: 'scannedOpsRows=' + scannedRows + ' | matchedClaims=' + Object.keys(fieldSet).length };
}

/**
 * Backup ops -> raw with full column format + data validation + rich text (Update Status).
 *
 * Why:
 * - Requirement: Status (dropdown+format), OR (checkbox), Update Status (rich text) must be preserved (not value-only).
 * - Claim alignment stays correct even if sheets are sorted differently.
 *
 * Performance:
 * - One values block read per ops sheet + optional rich text column read for Update Status.
 * - One write per raw column + one rich text write for Update Status.
 */
function backupOpsToRawFull_(ss, rawSheet, rawValues, headerIndexRaw, pic) {
  if (DRY_RUN) return { updated: 0, notes: 'DRY_RUN' };
  if (!ss || !rawSheet || !headerIndexRaw) return { updated: 0, notes: 'Missing args' };

  const h = CONFIG.headers;
  const idxRawClaim = headerIndexRaw[h.claimNumber];
  const idxRawOR = idxAny_(headerIndexRaw, [h.orColumn, 'OR', 'Or']);
  const idxRawUpdate = headerIndexRaw[h.updateStatus];
  const idxRawTs = headerIndexRaw[h.timestamp];
  const idxRawStatus = headerIndexRaw[h.status];
  const idxRawRemarks = idxAny_(headerIndexRaw, ['Remarks', 'Remark', 'remarks', 'remark']);

  // Additional manual/operational tail fields (snapshot per Claim Number)
  const idxRawUpdateAsso = headerIndexRaw['Update Status Asso'];
  const idxRawTsAsso = headerIndexRaw['Timestamp Asso'];
  const idxRawUpdateAdmin = headerIndexRaw['Update Status Admin'];
  const idxRawTsAdmin = headerIndexRaw['Timestamp Admin'];


  if (idxRawClaim == null) return { updated: 0, notes: 'Raw claim_number not found' };

  // Ensure rawValues are loaded (rows only, no header)
  const rawRowCount = Math.max(0, rawSheet.getLastRow() - 1);
  const workingRawValues = (rawValues && rawValues.length === rawRowCount)
    ? rawValues
    : (rawRowCount > 0 ? rawSheet.getRange(2, 1, rawRowCount, rawSheet.getLastColumn()).getValues() : []);

  if (!workingRawValues.length) return { updated: 0, notes: 'Raw has no rows' };

  // Build claim -> raw row index map
  const rawMap = {};
  for (let i = 0; i < workingRawValues.length; i++) {
    const key = String(workingRawValues[i][idxRawClaim] || '').trim().toUpperCase();
    if (key) rawMap[key] = i;
  }

  const opsList = getOperationalSheetsForBackup_(pic);
  const fieldSet = {}; // claim -> {or, status, ts, updateRt}
  const ensureObj = claim => (fieldSet[claim] ? fieldSet[claim] : (fieldSet[claim] = {}));

  let scannedRows = 0;
  let formatSource = null; // {sh, idxOR, idxUpdate, idxTs, idxStatus}

  for (let s = 0; s < opsList.length; s++) {
    const name = opsList[s];
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr <= 1 || lc <= 0) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v || '').trim());

    const idxClaimOps = header.indexOf(h.claimNumber) !== -1 ? header.indexOf(h.claimNumber) : header.indexOf('Claim Number');
    if (idxClaimOps === -1) continue;

    const idxOR = __findHeaderIndex05a_(header, 'OR');
    const idxUpdate = __findHeaderIndex05a_(header, 'Update Status');
    const idxTs = __findHeaderIndex05a_(header, 'Timestamp');
    const idxStatus = __findHeaderIndex05a_(header, 'Status');

    const idxUpdateAsso = __findHeaderIndex05a_(header, 'Update Status Asso');
    const idxTsAsso = __findHeaderIndex05a_(header, 'Timestamp Asso');
    const idxUpdateAdmin = __findHeaderIndex05a_(header, 'Update Status Admin');
    const idxTsAdmin = __findHeaderIndex05a_(header, 'Timestamp Admin');

    const idxRemarksTmp = __findHeaderIndex05a_(header, 'Remarks');
    const idxRemarks = (idxRemarksTmp !== -1) ? idxRemarksTmp : __findHeaderIndex05a_(header, 'Remark');

    if (idxOR === -1 && idxUpdate === -1 && idxTs === -1 && idxStatus === -1 &&
        idxUpdateAsso === -1 && idxTsAsso === -1 && idxUpdateAdmin === -1 && idxTsAdmin === -1 && idxRemarks === -1) continue;

    // Prefer using the "Submission" sheet as the formatting/validation source when available.
    const shName = sh.getName();
    if (!formatSource || (formatSource && formatSource.sh && formatSource.sh.getName && formatSource.sh.getName() !== 'Submission' && shName === 'Submission')) {
      formatSource = { sh, idxOR, idxUpdate, idxTs, idxStatus, idxUpdateAsso, idxTsAsso, idxUpdateAdmin, idxTsAdmin, idxRemarks };
    }

    // Read the smallest contiguous block that covers all needed columns
    const needed = [idxClaimOps, idxOR, idxUpdate, idxTs, idxStatus, idxUpdateAsso, idxTsAsso, idxUpdateAdmin, idxTsAdmin, idxRemarks].filter(x => x !== -1);
    const minC = Math.min.apply(null, needed);
    const maxC = Math.max.apply(null, needed);
    const width = maxC - minC + 1;

    const block = sh.getRange(2, minC + 1, lr - 1, width).getValues();
    scannedRows += block.length;

    // Rich text only for Update Status (and Asso/Admin variants, if present)
    let updateRT = null;
    let updateAssoRT = null;
    let updateAdminRT = null;
    if (idxUpdate !== -1) {
      updateRT = sh.getRange(2, idxUpdate + 1, lr - 1, 1).getRichTextValues(); // 2D
    }
    if (idxUpdateAsso !== -1) {
      updateAssoRT = sh.getRange(2, idxUpdateAsso + 1, lr - 1, 1).getRichTextValues(); // 2D
    }
    if (idxUpdateAdmin !== -1) {
      updateAdminRT = sh.getRange(2, idxUpdateAdmin + 1, lr - 1, 1).getRichTextValues(); // 2D
    }

    // Rich text for Remarks (preserve hyperlinks/format)
    let remarksRT = null;
    if (idxRemarks !== -1) {
      try { remarksRT = sh.getRange(2, idxRemarks + 1, lr - 1, 1).getRichTextValues(); } catch (e) { remarksRT = null; }
    }

    for (let r = 0; r < block.length; r++) {
      const row = block[r];
      const claim = String(row[idxClaimOps - minC] || '').trim().toUpperCase();
      if (!claim) continue;

      const obj = ensureObj(claim);

      if (idxOR !== -1 && obj.or == null) {
        const cb = normalizeCheckbox_(row[idxOR - minC]);
        if (cb !== '') obj.or = cb;
      }
      if (idxStatus !== -1 && obj.status == null) {
        const v = normalizeStatusValue_(row[idxStatus - minC]);
        if (v !== '') obj.status = v;
      }
      if (idxTs !== -1 && obj.ts == null) {
        const v = row[idxTs - minC];
        if (v !== '' && v != null) obj.ts = v;
      }
      if (idxUpdate !== -1 && obj.updateRt == null && updateRT) {
        const rt = updateRT[r] && updateRT[r][0];
        const t = (rt && rt.getText) ? String(rt.getText() || '').trim() : '';
        if (t) obj.updateRt = rt;
      }

      if (idxUpdateAsso !== -1 && obj.updateAssoRt == null && updateAssoRT) {
        const rt = updateAssoRT[r] && updateAssoRT[r][0];
        const t = (rt && rt.getText) ? String(rt.getText() || '').trim() : '';
        if (t) obj.updateAssoRt = rt;
      }
      if (idxTsAsso !== -1 && obj.tsAsso == null) {
        const v = row[idxTsAsso - minC];
        if (v !== '' && v != null) obj.tsAsso = v;
      }
      if (idxUpdateAdmin !== -1 && obj.updateAdminRt == null && updateAdminRT) {
        const rt = updateAdminRT[r] && updateAdminRT[r][0];
        const t = (rt && rt.getText) ? String(rt.getText() || '').trim() : '';
        if (t) obj.updateAdminRt = rt;
      }
      if (idxTsAdmin !== -1 && obj.tsAdmin == null) {
        const v = row[idxTsAdmin - minC];
        if (v !== '' && v != null) obj.tsAdmin = v;
      }

      if (idxRemarks !== -1 && obj.remarksRt == null) {
        // Prefer RichText to preserve formatting/hyperlinks/line breaks.
        if (remarksRT && remarksRT[r] && remarksRT[r][0] && remarksRT[r][0].getText) {
          const rt = remarksRT[r][0];
          const txt = String(rt.getText() || '');
          if (txt !== '') {
            obj.remarksRt = rt;
            obj.remarks = txt; // keep plain value too
          }
        } else if (obj.remarks == null) {
          const v = row[idxRemarks - minC];
          const txt = String(v == null ? '' : v);
          if (txt !== '') obj.remarks = txt;
        }
      }
    }
  }

  // Prepare raw Update Status rich text buffers (single batch write)
  let rawUpdateRTs = null;
  let rawUpdateAssoRTs = null;
  let rawUpdateAdminRTs = null;
  if (idxRawUpdate != null) {
    rawUpdateRTs = rawSheet.getRange(2, idxRawUpdate + 1, workingRawValues.length, 1).getRichTextValues();
  }
  if (idxRawUpdateAsso != null) {
    rawUpdateAssoRTs = rawSheet.getRange(2, idxRawUpdateAsso + 1, workingRawValues.length, 1).getRichTextValues();
  }
  if (idxRawUpdateAdmin != null) {
    rawUpdateAdminRTs = rawSheet.getRange(2, idxRawUpdateAdmin + 1, workingRawValues.length, 1).getRichTextValues();
  }

  let rawRemarksRTs = null;
  if (idxRawRemarks != null) {
    try { rawRemarksRTs = rawSheet.getRange(2, idxRawRemarks + 1, workingRawValues.length, 1).getRichTextValues(); } catch (e) { rawRemarksRTs = null; }
  }

  let updated = 0;
  Object.keys(fieldSet).forEach(claim => {
    const i = rawMap[claim];
    if (i == null) return;

    const obj = fieldSet[claim];
    const row = workingRawValues[i];
    let touched = false;

    if (idxRawOR != null && obj.or != null) { row[idxRawOR] = obj.or; touched = true; }
    if (idxRawStatus != null && obj.status != null) { row[idxRawStatus] = obj.status; touched = true; }
    if (idxRawTs != null && obj.ts != null) { row[idxRawTs] = obj.ts; touched = true; }

    if (idxRawTsAsso != null && obj.tsAsso != null) { row[idxRawTsAsso] = obj.tsAsso; touched = true; }
    if (idxRawTsAdmin != null && obj.tsAdmin != null) { row[idxRawTsAdmin] = obj.tsAdmin; touched = true; }

    if (idxRawUpdate != null && obj.updateRt != null && rawUpdateRTs) {
      rawUpdateRTs[i][0] = obj.updateRt;
      row[idxRawUpdate] = obj.updateRt.getText(); // keep values in sync for downstream reads
      touched = true;
    }

    if (idxRawUpdateAsso != null && obj.updateAssoRt != null && rawUpdateAssoRTs) {
      rawUpdateAssoRTs[i][0] = obj.updateAssoRt;
      row[idxRawUpdateAsso] = obj.updateAssoRt.getText();
      touched = true;
    }
    if (idxRawUpdateAdmin != null && obj.updateAdminRt != null && rawUpdateAdminRTs) {
      rawUpdateAdminRTs[i][0] = obj.updateAdminRt;
      row[idxRawUpdateAdmin] = obj.updateAdminRt.getText();
      touched = true;
    }

    if (idxRawRemarks != null) {
      if (obj.remarksRt != null && rawRemarksRTs) {
        rawRemarksRTs[i][0] = obj.remarksRt;
        row[idxRawRemarks] = obj.remarksRt.getText();
        touched = true;
      } else if (obj.remarks != null && String(obj.remarks) !== '') {
        row[idxRawRemarks] = obj.remarks;
        touched = true;
      }
    }

    if (touched) updated++;
  });

  // Apply column-level types/format/validation (cheap, once)
  if (idxRawOR != null) {
    try { rawSheet.getRange(2, idxRawOR + 1, workingRawValues.length, 1).insertCheckboxes(); } catch (e) {}
  }
  if (formatSource) {
    try {
      // Status: preserve dropdown *chip* rule + visuals on Raw (do NOT overwrite values)
      // Rationale: setDataValidation(dv) can drop "chip" display + option colors; copyTo(PASTE_DATA_VALIDATION) preserves it.
      if (idxRawStatus != null && formatSource.idxStatus !== -1) {
        const srcCell = formatSource.sh.getRange(2, formatSource.idxStatus + 1, 1, 1);
        const dstCol  = rawSheet.getRange(2, idxRawStatus + 1, workingRawValues.length, 1);

        try { srcCell.copyTo(dstCol, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); } catch (e) {
          // Fallback (may lose chip styling, but keeps dropdown rule at least)
          try {
            const dv = srcCell.getDataValidation();
            if (dv) dstCol.setDataValidation(dv);
          } catch (e2) {}
        }

        // Visuals (number format, fonts, alignment, borders)
        try { srcCell.copyTo(dstCol, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}

        // Conditional formatting (if the Status column uses CF-based highlighting)
        try { srcCell.copyTo(dstCol, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false); } catch (e) {}
      }

      // Update Status: column format (rich text is written below)
      if (idxRawUpdate != null && formatSource.idxUpdate !== -1) {
        const src2 = formatSource.sh.getRange(2, formatSource.idxUpdate + 1, 1, 1);
        const dst2 = rawSheet.getRange(2, idxRawUpdate + 1, workingRawValues.length, 1);
        src2.copyTo(dst2, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }

      // Remarks: preserve column format + (if any) data validation.
      if (idxRawRemarks != null && formatSource.idxRemarks !== -1) {
        const srcR = formatSource.sh.getRange(2, formatSource.idxRemarks + 1, 1, 1);
        const dstR = rawSheet.getRange(2, idxRawRemarks + 1, workingRawValues.length, 1);
        try { srcR.copyTo(dstR, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
        try { srcR.copyTo(dstR, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); } catch (e2) {}
      }

      // Update Status Asso/Admin: column format (rich text is written below)
      if (idxRawUpdateAsso != null && formatSource.idxUpdateAsso !== -1) {
        const srcA = formatSource.sh.getRange(2, formatSource.idxUpdateAsso + 1, 1, 1);
        const dstA = rawSheet.getRange(2, idxRawUpdateAsso + 1, workingRawValues.length, 1);
        try { srcA.copyTo(dstA, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
      }
      if (idxRawUpdateAdmin != null && formatSource.idxUpdateAdmin !== -1) {
        const srcB = formatSource.sh.getRange(2, formatSource.idxUpdateAdmin + 1, 1, 1);
        const dstB = rawSheet.getRange(2, idxRawUpdateAdmin + 1, workingRawValues.length, 1);
        try { srcB.copyTo(dstB, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
      }

      // Timestamp: number format (optional)
      if (idxRawTs != null && formatSource.idxTs !== -1) {
        const fmt = formatSource.sh.getRange(2, formatSource.idxTs + 1, 1, 1).getNumberFormat();
        if (fmt) rawSheet.getRange(2, idxRawTs + 1, workingRawValues.length, 1).setNumberFormat(fmt);
      }
    } catch (e) {}
  }

  // Write back changed raw columns (batch per contiguous block)
  const colsToWrite = [];
  if (idxRawOR != null) colsToWrite.push(idxRawOR);
  if (idxRawStatus != null) colsToWrite.push(idxRawStatus);
  if (idxRawTs != null) colsToWrite.push(idxRawTs);
  if (idxRawUpdate != null) colsToWrite.push(idxRawUpdate);
  if (idxRawUpdateAsso != null) colsToWrite.push(idxRawUpdateAsso);
  if (idxRawTsAsso != null) colsToWrite.push(idxRawTsAsso);
  if (idxRawUpdateAdmin != null) colsToWrite.push(idxRawUpdateAdmin);
  if (idxRawTsAdmin != null) colsToWrite.push(idxRawTsAdmin);
  if (idxRawRemarks != null) colsToWrite.push(idxRawRemarks);

  writeRawColumns_(rawSheet, workingRawValues, colsToWrite);

  // Rich text write (Update Status + Asso/Admin variants)
  if (idxRawUpdate != null && rawUpdateRTs) {
    try { safeSetRichTextValues_(rawSheet.getRange(2, idxRawUpdate + 1, rawUpdateRTs.length, 1), rawUpdateRTs); } catch (e) {}
  }
  if (idxRawUpdateAsso != null && rawUpdateAssoRTs) {
    try { safeSetRichTextValues_(rawSheet.getRange(2, idxRawUpdateAsso + 1, rawUpdateAssoRTs.length, 1), rawUpdateAssoRTs); } catch (e) {}
  }
  if (idxRawUpdateAdmin != null && rawUpdateAdminRTs) {
    try { safeSetRichTextValues_(rawSheet.getRange(2, idxRawUpdateAdmin + 1, rawUpdateAdminRTs.length, 1), rawUpdateAdminRTs); } catch (e) {}
  }

  if (idxRawRemarks != null && rawRemarksRTs) {
    try { safeSetRichTextValues_(rawSheet.getRange(2, idxRawRemarks + 1, rawRemarksRTs.length, 1), rawRemarksRTs); } catch (e) {}
  }

  return { updated, notes: 'scannedOpsRows=' + scannedRows + ' | matchedClaims=' + Object.keys(fieldSet).length + ' | fullFormat=1' };
}




/** =========================
 *  Raw minimal column writeback
 *  ========================= */

function writeRawColumns_(rawSheet, rawValues, colIdxList0based) {
  if (DRY_RUN) return;
  if (!rawSheet || !rawValues || !rawValues.length) return;

  const cols = Array.from(new Set(colIdxList0based.filter(x => x != null && x >= 0))).sort((a, b) => a - b);
  if (!cols.length) return;

  // Group contiguous columns into runs to minimize Range writes
  const runs = [];
  let cur = [cols[0]];
  for (let i = 1; i < cols.length; i++) {
    const c = cols[i];
    if (c === cur[cur.length - 1] + 1) cur.push(c);
    else { runs.push(cur); cur = [c]; }
  }
  runs.push(cur);

  for (let r = 0; r < runs.length; r++) {
    const run = runs[r];
    const startCol1 = run[0] + 1;
    const width = run.length;

    const out = new Array(rawValues.length);
    for (let i = 0; i < rawValues.length; i++) {
      const src = rawValues[i];
      const rowOut = new Array(width);
      for (let j = 0; j < width; j++) rowOut[j] = src[run[j]];
      out[i] = rowOut;
    }
    safeSetValues_(rawSheet.getRange(2, startCol1, rawValues.length, width), out);
  }
}
