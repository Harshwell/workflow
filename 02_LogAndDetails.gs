/***************************************************************
 * 02_LogAndDetails.gs
 * Enterprise Edition — Semi realtime Log + perf-safe Details
 *
 * Goals:
 * - Log shows progress early in the run (progress + first-line flush)
 * - Writes remain efficient (throttled flush)
 * - Details de-dupe + grouping without becoming a bottleneck
 ***************************************************************/
'use strict';
/** UI flag safe accessor (prevents ReferenceError if UI_FLAGS is unavailable) */
function uiFlag_(key, fallback) {
  try {
    if (typeof UI_FLAGS !== 'undefined' && UI_FLAGS && Object.prototype.hasOwnProperty.call(UI_FLAGS, key)) {
      return !!UI_FLAGS[key];
    }
  } catch (e) {}
  return !!fallback;
}
/** DRY_RUN safe accessor */

// [Refactor] Removed duplicate definition of isDryRun_ (source of truth in 01_Utils)

/** Whether Associate/Partner mapping is enabled (legacy PIC-based mapping). */

// [Refactor] Removed duplicate definition of isAssociateMappingEnabled_ (source of truth in 01_Utils)

/** =========================
 * Layout constants
 * ========================= */
/** Log layout */
const LOG_LAYOUT = Object.freeze({
  PROGRESS_HEADER_ROW: 2,
  PROGRESS_VALUE_ROW: 3,
  DETAIL_HEADER_ROW: 5,
  DETAIL_START_ROW: 6
});
const LOG_COLUMNS = Object.freeze({
  NUMBER: 2,      // B
  TIMESTAMP: 3,   // C
  SEGMENT: 4,     // D
  DURATION: 5,    // E
  METRICS: 6,     // F
  NOTES: 7,       // G
  SEVERITY: 8     // H
});
/** =========================
 * LOG v2 (Structured stage log for MAIN/SUB)
 *
 * Note: This is designed to be backward compatible with the legacy Log table.
 * - Legacy table stays at columns B:H
 * - v2 table is placed starting at column J to avoid breaking existing formulas
 *
 * v2 columns match the spec (Flow/RunID/Stage/Email/Attachment/Counts/Status).
 * Call-sites can gradually migrate to logStageV2_() without breaking old logLine_().
 * ========================= */
const LOGV2_LAYOUT = Object.freeze({
  HEADER_ROW: LOG_LAYOUT.DETAIL_HEADER_ROW, // align with legacy header row
  START_ROW: LOG_LAYOUT.DETAIL_START_ROW,
  START_COL: 10, // J
  COLS: 21
});
const LOGV2_HEADER = Object.freeze([
  'No',
  'Timestamp',
  'Flow',
  'Run ID',
  'Stage',
  'Time Segment',
  'Duration(ms)/(s)',
  'Email From',
  'Email Subject',
  'Thread ID',
  'Attachment Name',
  'Attachment Size',
  'Attachment Type',
  'Raw Rows',
  'Ops Updated',
  'Sub Inserted',
  'Status',
  'Error',
  'Metrics',
  'Notes',
  'Severity'
]);
const DETAILS_LAYOUT = Object.freeze({
  HEADER_ROW: 2,
  START_ROW: 3,
  START_COL: 2, // B
  COLS: 8
});
const DETAILS_COLUMNS = Object.freeze({
  NO: 0,
  TIMESTAMP: 1,
  SUBMISSION_DATE: 2,
  TARGET_SHEET: 3,
  CLAIM_NUMBER: 4,
  PARTNER: 5,
  LAST_STATUS: 6,
  REASON: 7
});
/** =========================
 * Tuning knobs
 * ========================= */
// Flush policy: "semi updated" without killing runtime
const LOG_FLUSH_THROTTLE_MS = 2000;     // max 1 flush / 2 seconds
const LOG_FORCE_FLUSH_ON_FIRST_WRITE = true;
const LOG_FORCE_FLUSH_ON_ERROR = true;
// Details tuning
const DETAILS_DEDUPE_SCAN_LIMIT = 20000;     // scan last N rows for dedupe (perf-safe)
const DETAILS_GROUP_THRESHOLD = 30;          // if many rows, group aggressively
const DETAILS_CLAIM_CELL_MAX = 12;           // max claims shown in cell; full list goes to note
const DETAILS_ROWNUM_CELL_MAX = 20;          // max row numbers shown inline
// Details suppression rule:
// Do NOT log the specific reason below when Submission Date is before Jun 2025.
// Example: 21 May 25 => skip this row if reason matches.
const DETAILS_SUPPRESS_PARTNER_NOT_MAPPED_REASON = 'Partner is not mapped (legacy mapping disabled).';
const DETAILS_SUPPRESS_PARTNER_NOT_MAPPED_CUTOFF = new Date(2025, 5, 1); // 1 Jun 2025 (month is 0-based)
/** Formats fallback-safe */
const DETAILS_FORMATS = Object.freeze({
  DATE: (
    (typeof FORMATS !== 'undefined' && FORMATS.DATE) ? FORMATS.DATE :
    (typeof SHEET_FORMATS !== 'undefined' && SHEET_FORMATS.DATE) ? SHEET_FORMATS.DATE :
    'd mmm yy'
  ),
  DATETIME: (
    (typeof FORMATS !== 'undefined' && FORMATS.DATETIME) ? FORMATS.DATETIME :
    (typeof SHEET_FORMATS !== 'undefined' && SHEET_FORMATS.DATETIME) ? SHEET_FORMATS.DATETIME :
    'd mmm yy, HH:mm'
  )
});
/** =========================
 * Cache
 * ========================= */
const CACHE = {
  log: {
    ss: null,
    sh: null,
    ensured: false,
    nextRow: LOG_LAYOUT.DETAIL_START_ROW,
    segNo: 0,
    didAnyWrite: false,
    lastFlushMs: 0,
    mappingKeySet: new Set(),
    // v2 (structured)
    v2Ensured: false,
    v2NextRow: LOGV2_LAYOUT.START_ROW,
    v2SegNo: 0,
    // run context
    runId: '',
    flow: ''
  },
  details: {
    ss: null,
    sh: null,
    ensured: false,
    formattedOnce: false,
    nextNo: null,
    existingClaimSet: null,   // cached set of claim keys already in Details
    runClaimSet: new Set()    // claim keys appended in current execution
  }
};
function resetLogState_() {
  CACHE.log.segNo = 0;
  CACHE.log.nextRow = LOG_LAYOUT.DETAIL_START_ROW;
  CACHE.log.didAnyWrite = false;
  CACHE.log.lastFlushMs = 0;
  // v2
  CACHE.log.v2SegNo = 0;
  CACHE.log.v2NextRow = LOGV2_LAYOUT.START_ROW;
  CACHE.log.v2Ensured = false;
  // context (re-initialized per run; generated lazily)
  CACHE.log.runId = '';
  CACHE.log.flow = '';
  try { CACHE.log.mappingKeySet = new Set(); } catch (e) {}
}
function resetDetailsState_() {
  CACHE.details.nextNo = null;
  CACHE.details.existingClaimSet = null;
  CACHE.details.runClaimSet = new Set();
  CACHE.details.formattedOnce = false;
}
/** =========================
 * Run context helpers (Flow + Run ID)
 * ========================= */
function normalizeFlowLabel_(flow) {
  const f = String(flow || '').trim();
  if (!f) return '';
  const up = f.toUpperCase();
  if (up === 'MAIN' || up === 'SUB' || up === 'FORM') return up;
  return f;
}
/**
 * Best-effort flow inference from segment/stage id when entrypoints didn't set RUNTIME.flow.
 * This keeps MAIN/SUB log separation usable without forcing changes across all entrypoints.
 */
function inferFlowFromSegId_(segId) {
  const id = String(segId || '').trim().toUpperCase();
  if (!id) return '';
  if (id === 'MAIN' || id.startsWith('MAIL') || id.startsWith('MAIN')) return 'MAIN';
  if (id === 'SUB' || id.startsWith('SUB')) return 'SUB';
  if (id === 'FORM' || id.startsWith('FORM')) return 'FORM';
  return '';
}
/**
 * Allow entrypoints to set flow/runId explicitly.
 * - If not set, we generate runId lazily.
 * - Flow is best-effort (falls back to RUNTIME.flow if present).
 */
function setLogRunContext_(flow, runId) {
  try { CACHE.log.flow = normalizeFlowLabel_(flow); } catch (e) {}
  try { CACHE.log.runId = String(runId || ''); } catch (e2) {}
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME) {
      if (flow != null) RUNTIME.flow = normalizeFlowLabel_(flow);
      if (runId != null) RUNTIME.runId = String(runId || '');
    }
  } catch (e3) {}
}
function getLogFlow_() {
  try {
    if (CACHE.log.flow) return CACHE.log.flow;
  } catch (e0) {}
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.flow) return normalizeFlowLabel_(RUNTIME.flow);
  } catch (e1) {}
  return '';
}
function getRunId_() {
  try {
    if (CACHE.log.runId) return CACHE.log.runId;
  } catch (e0) {}
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.runId) {
      CACHE.log.runId = String(RUNTIME.runId || '');
      if (CACHE.log.runId) return CACHE.log.runId;
    }
  } catch (e1) {}
  // Generate lazily
  let rid = '';
  try { rid = Utilities.getUuid(); } catch (e2) {}
  if (!rid) {
    // Fallback: timestamp + random
    rid = 'RID-' + String(Date.now()) + '-' + String(Math.floor(Math.random() * 1e9));
  }
  try { CACHE.log.runId = rid; } catch (e3) {}
  try { if (typeof RUNTIME !== 'undefined' && RUNTIME) RUNTIME.runId = rid; } catch (e4) {}
  return rid;
}
function msFromSec_(sec) {
  const n = Number(sec);
  if (!isFinite(n)) return '';
  return Math.round(n * 1000);
}
function joinNotes_(metrics, notes) {
  const a = String(metrics || '').trim();
  const b = String(notes || '').trim();
  if (a && b) return a + ' | ' + b;
  return a || b || '';
}
/** =========================
 * Mapping error log helpers (Sheet "Log")
 * ========================= */
/** Whether mapping error logging is enabled */
function isMappingErrorLogEnabled_() {
  try {
    if (typeof MAPPING_ERROR_LOG_POLICY !== 'undefined' && MAPPING_ERROR_LOG_POLICY) {
      return !!MAPPING_ERROR_LOG_POLICY.ENABLE;
    }
  } catch (e) {}
  return true;
}
function formatRelevantDateForLog_(dateVal) {
  const d = normalizeDate_(dateVal);
  if (!d) return '';
  try {
    const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'GMT';
    return Utilities.formatDate(d, tz, 'dd MMM yyyy');
  } catch (e) {}
  try { return String(d); } catch (e2) {}
  return '';
}
/** One-line mapping error to Log sheet: <value> | <claim> | <date> */
function logMappingError_(eventName, value, claimNumber, relevantDateVal) {
  if (isDryRun_()) return false;
  if (!isMappingErrorLogEnabled_()) return false;
  const claim = String(claimNumber || '').trim();
  const val = String(value || '').trim();
  if (!claim || !val) return false;
  const dateStr = formatRelevantDateForLog_(relevantDateVal);
  const msg = val + ' | ' + claim + (dateStr ? (' | ' + dateStr) : '');
  const key = String(eventName || 'MAPPING') + '|' + msg;
  try {
    if (!CACHE.log.mappingKeySet) CACHE.log.mappingKeySet = new Set();
    if (CACHE.log.mappingKeySet.has(key)) return false;
    CACHE.log.mappingKeySet.add(key);
  } catch (e) {}
  const segName = String(eventName || 'MAPPING').replace(/_/g, ' ');
  pushLogRow_('MAP', segName, '', '', msg, 'WARN');
  return true;
}
function isPartnerNotMappedReasonText_(reason) {
  const r = String(reason || '');
  if (!r) return false;
  // Match both legacy and new phrasing (kept for backward compatibility).
  return (r.indexOf(DETAILS_SUPPRESS_PARTNER_NOT_MAPPED_REASON) !== -1) ||
    /partner\s+is\s+not\s+mapped(\s+to\s+any\s+associate)?/i.test(r);
}
function isStatusNotRoutedReasonText_(reason) {
  const r = String(reason || '');
  if (!r) return false;
  return /status\s+is\s+not\s+routed/i.test(r) || /add\s+to\s+routing\s+map/i.test(r);
}
/** Mirror unmapped partner / unmapped status warnings into Log sheet (short entries) */
function mirrorMappingErrorsToLog_(rows) {
  if (!rows || !rows.length) return;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const reason = String(r.reason || '');
    if (!reason) continue;
    const claim = normalizeClaimKey_(r.claim || '') || String(r.claim || '').trim();
    if (!claim) continue;
    // Prefer submission date; fall back to optional status date fields if provided by call-site
    const dateVal =
      (r.submissionDateVal != null && r.submissionDateVal !== '') ? r.submissionDateVal :
      (r.lastStatusDateVal != null ? r.lastStatusDateVal :
       (r.lastStatusDate != null ? r.lastStatusDate :
        (r.statusDateVal != null ? r.statusDateVal : '')));
    if (isPartnerNotMappedReasonText_(reason)) {
      // Legacy associate mapping is disabled in the new master-sheet flow; treat this reason as noise.
      if (!isAssociateMappingEnabled_()) continue;
      const ev = (typeof MAPPING_ERROR_LOG_POLICY !== 'undefined' && MAPPING_ERROR_LOG_POLICY && MAPPING_ERROR_LOG_POLICY.EVENT && MAPPING_ERROR_LOG_POLICY.EVENT.UNMAPPED_PARTNER)
        ? MAPPING_ERROR_LOG_POLICY.EVENT.UNMAPPED_PARTNER
        : 'UNMAPPED_PARTNER';
      logMappingError_(ev, r.partner || '(Partner blank)', claim, dateVal);
      continue;
    }
    if (isStatusNotRoutedReasonText_(reason)) {
      const ev = (typeof MAPPING_ERROR_LOG_POLICY !== 'undefined' && MAPPING_ERROR_LOG_POLICY && MAPPING_ERROR_LOG_POLICY.EVENT && MAPPING_ERROR_LOG_POLICY.EVENT.UNMAPPED_LAST_STATUS)
        ? MAPPING_ERROR_LOG_POLICY.EVENT.UNMAPPED_LAST_STATUS
        : 'UNMAPPED_LAST_STATUS';
      logMappingError_(ev, r.lastStatus || '(Last Status blank)', claim, dateVal);
      continue;
    }
  }
}
/** =========================
 * Locale-safe helpers (formula)
 * ========================= */
function buildSparklineFormula_(ss, ratio01) {
  const r = Math.max(0, Math.min(1, ratio01 || 0));
  // Prefer util from 01_Utils.gs if available
  let argSep = ',';
  try {
    if (typeof getFormulaArgSep_ === 'function') argSep = getFormulaArgSep_(ss);
  } catch (e) {}
  // Array-literal column separator differs by locale; common:
  // - en_* : comma
  // - many non-en locales : backslash
  const colSep = (argSep === ';') ? '\\' : ',';
  const rowSep = ';';
  // {"charttype","bar";"max",1}  (en)
  // {"charttype"\ "bar";"max"\ 1} (many non-en locales)
  const opts = '{"charttype"' + colSep + '"bar"' + rowSep + '"max"' + colSep + '1}';
  return '=SPARKLINE(' + r + argSep + opts + ')';
}
/** =========================
 * Flush policy (semi-realtime)
 * ========================= */
function shouldFlushLogNow_(severity) {
  if (isDryRun_()) return false;
  const sev = String(severity || 'INFO').toUpperCase();
  if (LOG_FORCE_FLUSH_ON_ERROR && (sev === 'ERROR' || sev === 'FATAL')) return true;
  if (LOG_FORCE_FLUSH_ON_FIRST_WRITE && !CACHE.log.didAnyWrite) return true;
  const now = Date.now();
  if (!CACHE.log.lastFlushMs) return true;
  return (now - CACHE.log.lastFlushMs) >= LOG_FLUSH_THROTTLE_MS;
}
function flushLogIfNeeded_(severity) {
  if (!shouldFlushLogNow_(severity)) return;
  try {
    SpreadsheetApp.flush();
    CACHE.log.lastFlushMs = Date.now();
  } catch (e) {}
}
/** =========================
 * Log sheet
 * ========================= */
function getLogSheet_() {
  if (CACHE.log.sh) return CACHE.log.sh;
  const ss = SpreadsheetApp.openById(CONFIG.logSpreadsheetId);
  const sh = ss.getSheetByName(CONFIG.logSheetName) || ss.getSheets()[0];
  CACHE.log.ss = ss;
  CACHE.log.sh = sh;
  if (!CACHE.log.ensured) {
    ensureLogLayout_(sh);
    CACHE.log.ensured = true;
  }
  return sh;
}
function ensureLogLayout_(sh) {
  if (isDryRun_()) return;
  // Minimal + idempotent: don't wipe formatting unless missing
  try {
    const headerRange = sh.getRange(LOG_LAYOUT.PROGRESS_HEADER_ROW, 2, 1, 4);
    const h = headerRange.getValues()[0];
    const expected = ['Progress', '%', 'Current Step', 'Updated At'];
    let needs = false;
    for (let i = 0; i < expected.length; i++) {
      if (String(h[i] || '').trim() !== expected[i]) { needs = true; break; }
    }
    if (needs) {
      safeSetValues_(headerRange, [expected]);
      try { headerRange.setFontWeight('bold'); } catch (e0) {}
    }
    const detailHdrRange = sh.getRange(LOG_LAYOUT.DETAIL_HEADER_ROW, LOG_COLUMNS.NUMBER, 1, 7);
    const dHdr = detailHdrRange.getValues()[0];
    const dExp = ['No','Time','Segment','Duration (s)','Metrics','Notes','Severity'];
    let needsHdr = false;
    for (let i = 0; i < dExp.length; i++) {
      if (String(dHdr[i] || '').trim() !== dExp[i]) { needsHdr = true; break; }
    }
    if (needsHdr) {
      safeSetValues_(detailHdrRange, [dExp]);
      try { detailHdrRange.setFontWeight('bold'); } catch (e1) {}
    }
  } catch (e) {}
  // v2 structured log table (non-breaking, placed at col J)
  try { ensureLogV2Layout_(sh); } catch (e2) {}
}
function ensureLogV2Layout_(sh) {
  if (isDryRun_()) return;
  if (!sh) return;
  try {
    const hdr = sh.getRange(LOGV2_LAYOUT.HEADER_ROW, LOGV2_LAYOUT.START_COL, 1, LOGV2_LAYOUT.COLS);
    const existing = hdr.getValues()[0].map(v => String(v || '').trim());
    let ok = true;
    for (let i = 0; i < LOGV2_HEADER.length; i++) {
      if ((existing[i] || '') !== LOGV2_HEADER[i]) { ok = false; break; }
    }
    if (!ok) {
      safeSetValues_(hdr, [Array.from(LOGV2_HEADER)]);
      try { hdr.setFontWeight('bold'); } catch (e0) {}
    }
    // Formats (bounded to prevent full-sheet formatting overhead)
    const maxRows = Math.max(1, sh.getMaxRows() - LOGV2_LAYOUT.START_ROW + 1);
    // Timestamp
    try {
      sh.getRange(LOGV2_LAYOUT.START_ROW, LOGV2_LAYOUT.START_COL + 1, maxRows, 1)
        .setNumberFormat(DETAILS_FORMATS.DATETIME);
    } catch (e1) {}
    // Duration (ms) integer
    try {
      sh.getRange(LOGV2_LAYOUT.START_ROW, LOGV2_LAYOUT.START_COL + 6, maxRows, 1)
        .setNumberFormat('0');
    } catch (e2) {}
    CACHE.log.v2Ensured = true;
  } catch (e) {}
}
function getNextLogV2Row_(sh) {
  const start = LOGV2_LAYOUT.START_ROW;
  try {
    if (CACHE.log.v2NextRow && CACHE.log.v2NextRow >= start) return CACHE.log.v2NextRow;
  } catch (e0) {}
  const lr = sh.getLastRow ? sh.getLastRow() : 0;
  const next = Math.max(lr + 1, start);
  try { CACHE.log.v2NextRow = next; } catch (e1) {}
  return next;
}
/**
 * Push one structured row into LOG v2 table.
 * Fields are best-effort; call-sites may pass partial.
 */
function pushLogV2Row_(entry) {
  const sh = getLogSheet_();
  try { if (!CACHE.log.v2Ensured) ensureLogV2Layout_(sh); } catch (e0) {}
  const rowNo = getNextLogV2Row_(sh);
  try { CACHE.log.v2SegNo = (CACHE.log.v2SegNo || 0) + 1; } catch (e1) {}
  const no = CACHE.log.v2SegNo || 1;
  const e = entry || {};
  const ts = (e.timestamp instanceof Date) ? e.timestamp : new Date();
  const values = [[
    no,
    ts,
    e.flow || getLogFlow_() || '',
    e.runId || getRunId_() || '',
    e.stage || '',
    e.timeSegment || '',
    (e.durationMs != null && e.durationMs !== '') ? e.durationMs :
      (e.durationSec != null && e.durationSec !== '' ? msFromSec_(e.durationSec) : ''),
    e.emailFrom || '',
    e.emailSubject || '',
    e.threadId || '',
    e.attachmentName || '',
    (e.attachmentSize != null && e.attachmentSize !== '') ? e.attachmentSize : '',
    e.attachmentType || '',
    (e.rawRows != null && e.rawRows !== '') ? e.rawRows : '',
    (e.opsUpdated != null && e.opsUpdated !== '') ? e.opsUpdated : '',
    (e.subInserted != null && e.subInserted !== '') ? e.subInserted : '',
    e.status || '',
    e.error || '',
    e.metrics || '',
    e.notes || '',
    e.severity || ''
  ]];
  safeSetValues_(sh.getRange(rowNo, LOGV2_LAYOUT.START_COL, 1, LOGV2_LAYOUT.COLS), values);
  try { CACHE.log.v2NextRow = rowNo + 1; } catch (e2) {}
  // Keep flushing policy consistent with legacy
  flushLogIfNeeded_(e.status || 'INFO');
}
/**
 * Convenience helper for stage log (recommended new API).
 *
 * @param {string} stage - e.g. SEARCH/ATTACH/CONVERT/EXTRACT/UPDATE/UPSERT/WEBHOOK/CLEANUP/END
 * @param {Object=} meta
 */
function logStageV2_(stage, meta) {
  const m = meta || {};
  pushLogV2Row_({
    timestamp: (m.timestamp instanceof Date) ? m.timestamp : new Date(),
    flow: m.flow || getLogFlow_(),
    runId: m.runId || getRunId_(),
    stage: stage || m.stage || '',
    timeSegment: m.timeSegment || '',
    durationMs: (m.durationMs != null) ? m.durationMs : (m.durationSec != null ? msFromSec_(m.durationSec) : ''),
    durationSec: (m.durationSec != null) ? m.durationSec : '',
    emailFrom: m.emailFrom || '',
    emailSubject: m.emailSubject || '',
    threadId: m.threadId || '',
    attachmentName: m.attachmentName || '',
    attachmentSize: m.attachmentSize || '',
    attachmentType: m.attachmentType || '',
    rawRows: (m.rawRows != null ? m.rawRows : ''),
    opsUpdated: (m.opsUpdated != null ? m.opsUpdated : ''),
    subInserted: (m.subInserted != null ? m.subInserted : ''),
    status: m.status || 'OK',
    error: m.error || '',
    metrics: m.metrics || '',
    notes: m.notes || '',
    severity: m.severity || ''
  });
}
function clearLogSheet_() {
  const sh = getLogSheet_();
  ensureLogLayout_(sh);
  // Clear progress values (content only; keep header)
  const pr = sh.getRange(LOG_LAYOUT.PROGRESS_VALUE_ROW, 2, 1, 4);
  safeClearContents_(pr);
  safeClearNotes_(pr);
  // Clear detail table body
  const maxRows = sh.getMaxRows();
  const detailRows = maxRows - (LOG_LAYOUT.DETAIL_START_ROW - 1);
  if (detailRows > 0) {
    const rng = sh.getRange(LOG_LAYOUT.DETAIL_START_ROW, LOG_COLUMNS.NUMBER, detailRows, 7);
    safeClearContents_(rng);
    safeClearNotes_(rng);
    // Keep format to preserve visual style (per spec)
    // safeClearFormat_(rng);
  }
  // Clear v2 structured log table body (cols J..)
  try {
    const maxRows = sh.getMaxRows();
    const rows = maxRows - (LOGV2_LAYOUT.START_ROW - 1);
    if (rows > 0) {
      const rng2 = sh.getRange(LOGV2_LAYOUT.START_ROW, LOGV2_LAYOUT.START_COL, rows, LOGV2_LAYOUT.COLS);
      safeClearContents_(rng2);
      safeClearNotes_(rng2);
    }
  } catch (e2) {}
  resetLogState_();
  if (uiFlag_('FLUSH_AFTER_LAYOUT_CLEAR', false)) flushNow_();
}
function setProgress_(ratio01, stepLabel) {
const sh = getLogSheet_();
const ss = CACHE.log.ss || SpreadsheetApp.openById(CONFIG.logSpreadsheetId);
const r = Math.max(0, Math.min(1, ratio01 || 0));
const pct = Math.round(r * 100);
const spark = buildSparklineFormula_(ss, r);
// Progress cell:
// - Keeps legacy sparkline visualization.
// - Adds a note containing flow/runId/state for easier debugging.
const progressCell = sh.getRange('B3');
safeSetFormula_(progressCell, spark);
const updatedAt = nowStr_('dd MMM yyyy, HH:mm:ss');
safeSetValues_(sh.getRange('C3:E3'), [[pct + '%', stepLabel || '', updatedAt]]);
try {
  const flow = getLogFlow_() || '';
  const runId = getRunId_() || '';
  const state = (r >= 1) ? 'DONE' : (r <= 0 ? 'START' : 'RUNNING');
  const note = [
    'State: ' + state,
    'Flow: ' + (flow || '-'),
    'RunID: ' + (runId || '-'),
    'Pct: ' + String(pct) + '%',
    'Updated: ' + updatedAt
  ].join('\\n');
  progressCell.setNote(note);
} catch (e0) {}
CACHE.log.didAnyWrite = true;
// "semi realtime": flush early once, then throttle.
if (uiFlag_('REALTIME_PROGRESS', false)) flushNow_();
else flushLogIfNeeded_('INFO');
}
/**
 * Flow-aware progress helper (recommended call-site).
 *
 * This is the canonical entry point to satisfy the requirement:
 * Log sheet must update Progress, %, Current Step, Updated At for MAIN/SUB (and optionally FORM).
 *
 * @param {string} flow - 'main' | 'sub' | 'form' (case-insensitive)
 * @param {number} ratio01 - 0..1
 * @param {string} stepLabel - current step label
 * @param {Object=} opt - {runId?:string, prefixFlowInStep?:boolean}
 */
function setProgressForFlow_(flow, ratio01, stepLabel, opt) {
  const o = opt || {};
  const f = normalizeFlowLabel_(flow || '');
  const rid = (o.runId != null) ? String(o.runId || '') : '';
  try { setLogRunContext_(f, rid || getRunId_()); } catch (e0) {}
  const prefix = (o.prefixFlowInStep === false) ? '' : (f ? '[' + f + '] ' : '');
  setProgress_(ratio01, prefix + (stepLabel || ''));
}
/**
 * Backward-compatible alias used by refactored modules.
 * Keep names stable across files while the project migrates.
 */
function updateLogProgress06_(flow, ratio01, stepLabel, runId) {
  setProgressForFlow_(flow, ratio01, stepLabel, { runId: runId, prefixFlowInStep: true });
}
function startSegment_(id, name) {
  return { id: id || '', name: name || '', start: new Date() };
}
function pushLogRow_(segId, segName, durationSec, metrics, notes, severity) {
  const sev = severity || 'INFO';
  CACHE.log.segNo++;
  const sh = getLogSheet_();
  const timeStr = nowStr_('HH:mm:ss');
  const row = [[
    CACHE.log.segNo,
    timeStr,
    (segId || '') + ' – ' + (segName || ''),
    (durationSec != null && durationSec !== '') ? durationSec : '',
    metrics || '',
    notes || '',
    sev
  ]];
  const r = CACHE.log.nextRow;
  safeSetValues_(sh.getRange(r, LOG_COLUMNS.NUMBER, 1, 7), row);
  safeSetNumberFormat_(sh.getRange(r, LOG_COLUMNS.DURATION, 1, 1), '0.00');
  CACHE.log.nextRow = r + 1;
  CACHE.log.didAnyWrite = true;
  // Mirror into structured log v2 (best-effort, backward compatible)
  try {
    pushLogV2Row_({
      timestamp: new Date(),
      flow: (function(){ const rid = getRunId_(); const f0 = getLogFlow_(); const f = f0 || inferFlowFromSegId_(segId); try { if (!f0 && f) setLogRunContext_(f, rid); } catch (e0) {} return f; })(),
      runId: getRunId_(),
      stage: String(segId || ''),
      timeSegment: String(segName || ''),
      durationSec: (durationSec != null && durationSec !== '') ? durationSec : '',
      durationMs: (durationSec != null && durationSec !== '') ? msFromSec_(durationSec) : '',
      status: String(sev || 'INFO'),
      error: ((String(sev).toUpperCase() === 'ERROR' || String(sev).toUpperCase() === 'FATAL') ? String(notes || '') : ''),
      metrics: String(metrics || ''),
      notes: String(notes || ''),
      severity: String(sev || 'INFO')
    });
  } catch (eV2) {}
  if (uiFlag_('REALTIME_LOG_LINES', false)) flushNow_();
  else flushLogIfNeeded_(sev);
}
function endSegment_(ctx, metrics, notes, severity) {
  const end = new Date();
  const durationSec = (end.getTime() - ctx.start.getTime()) / 1000.0;
  pushLogRow_(ctx.id, ctx.name, durationSec, metrics, notes, severity);
}
function logLine_(id, name, metrics, notes, severity) {
  pushLogRow_(id, name, '', metrics, notes, severity || 'INFO');
}
/** =========================
 * Details sheet
 * ========================= */
function getDetailsSheet_() {
  if (CACHE.details.sh) return CACHE.details.sh;
  const ss = SpreadsheetApp.openById(CONFIG.logSpreadsheetId);
  const sh = ss.getSheetByName(CONFIG.detailsSheetName) || ss.insertSheet(CONFIG.detailsSheetName);
  CACHE.details.ss = ss;
  CACHE.details.sh = sh;
  if (!CACHE.details.ensured) {
    ensureDetailsLayout_(sh);
    CACHE.details.ensured = true;
  }
  return sh;
}
function ensureDetailsLayout_(sh) {
  if (isDryRun_()) return;
  const headerRange = sh.getRange(DETAILS_LAYOUT.HEADER_ROW, DETAILS_LAYOUT.START_COL, 1, DETAILS_LAYOUT.COLS);
  const existing = headerRange.getValues()[0].map(v => String(v || '').trim());
  const expected = ['No.','Timestamp','Submission Date','Target Sheet','Claim Number','Partner','Last Status','Reason'];
  let ok = true;
  for (let i = 0; i < expected.length; i++) {
    if ((existing[i] || '') !== expected[i]) { ok = false; break; }
  }
  if (!ok) safeSetValues_(headerRange, [expected]);
  // Format header once
  try { headerRange.setFontWeight('bold'); } catch (e) {}
  // Avoid formatting entire sheet repeatedly
  if (CACHE.details.formattedOnce) return;
  try {
    const maxRows = Math.max(1, sh.getMaxRows() - DETAILS_LAYOUT.START_ROW + 1);
    // Timestamp col (table col index 1 => absolute col START_COL+1)
    sh.getRange(DETAILS_LAYOUT.START_ROW, DETAILS_LAYOUT.START_COL + DETAILS_COLUMNS.TIMESTAMP, maxRows, 1)
      .setNumberFormat(DETAILS_FORMATS.DATETIME);
    // Submission Date col (table col index 2 => absolute col START_COL+2)
    sh.getRange(DETAILS_LAYOUT.START_ROW, DETAILS_LAYOUT.START_COL + DETAILS_COLUMNS.SUBMISSION_DATE, maxRows, 1)
      .setNumberFormat(DETAILS_FORMATS.DATE);
  } catch (e2) {}
  CACHE.details.formattedOnce = true;
}
function getNextDetailsNo_(detailsSheet) {
  if (CACHE.details.nextNo != null) return CACHE.details.nextNo;
  const lastRow = detailsSheet.getLastRow();
  if (lastRow < DETAILS_LAYOUT.START_ROW) {
    CACHE.details.nextNo = 1;
    return 1;
  }
  const lastNo = normalizeNumber_(detailsSheet.getRange(lastRow, DETAILS_LAYOUT.START_COL).getValue());
  if (lastNo != null && lastNo >= 1) {
    CACHE.details.nextNo = Math.floor(lastNo) + 1;
    return CACHE.details.nextNo;
  }
  const scanStart = Math.max(DETAILS_LAYOUT.START_ROW, lastRow - 2000);
  const vals = detailsSheet.getRange(scanStart, DETAILS_LAYOUT.START_COL, lastRow - scanStart + 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const n = normalizeNumber_(vals[i][0]);
    if (n != null && n >= 1) {
      CACHE.details.nextNo = Math.floor(n) + 1;
      return CACHE.details.nextNo;
    }
  }
  CACHE.details.nextNo = 1;
  return 1;
}
/** Normalize claim key for dedupe */
function normalizeClaimKey_(claim) {
  const s = String(claim || '').trim();
  if (!s) return '';
  return s.toUpperCase();
}
/** Extract plausible claim tokens from a cell value (handles grouped cells/newlines) */
function extractClaimTokens_(cellVal) {
  const s = String(cellVal || '').trim();
  if (!s) return [];
  const tokens = s.split(/[\s,;|]+/).filter(Boolean);
  const out = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = String(tokens[i] || '').trim();
    if (!t) continue;
    if (/^[A-Z0-9][A-Z0-9_-]{3,}$/i.test(t)) out.push(normalizeClaimKey_(t));
  }
  return out;
}
/** Build (cached) set of claims already logged in Details */
function getExistingDetailsClaimSet_(sh) {
  if (CACHE.details.existingClaimSet) return CACHE.details.existingClaimSet;
  const set = new Set();
  const lr = sh.getLastRow();
  if (lr < DETAILS_LAYOUT.START_ROW) {
    CACHE.details.existingClaimSet = set;
    return set;
  }
  // Claim Number absolute col = START_COL + (table index 4)
  const claimCol = DETAILS_LAYOUT.START_COL + DETAILS_COLUMNS.CLAIM_NUMBER;
  const scanStart = Math.max(DETAILS_LAYOUT.START_ROW, lr - DETAILS_DEDUPE_SCAN_LIMIT + 1);
  const count = lr - scanStart + 1;
  if (count <= 0) {
    CACHE.details.existingClaimSet = set;
    return set;
  }
  const vals = sh.getRange(scanStart, claimCol, count, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const tokens = extractClaimTokens_(vals[i][0]);
    for (let k = 0; k < tokens.length; k++) set.add(tokens[k]);
  }
  CACHE.details.existingClaimSet = set;
  return set;
}
/** Pick a representative submission date (earliest valid date-only) */
function pickGroupSubmissionDate_(rows) {
  let best = null;
  for (let i = 0; i < rows.length; i++) {
    const v = rows[i] ? rows[i].submissionDateVal : null;
    const d = normalizeDate_(v);
    if (!d) continue;
    const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
    if (!best || dd.getTime() < best.getTime()) best = dd;
  }
  return best ? best : '';
}
function formatRowNumsInline_(nums) {
  const a = Array.from(new Set((nums || []).filter(n => n != null && n !== ''))).sort((x, y) => x - y);
  if (!a.length) return '';
  const slice = a.slice(0, DETAILS_ROWNUM_CELL_MAX);
  const more = a.length - slice.length;
  return more > 0 ? (slice.join(', ') + ' …(+' + more + ')') : slice.join(', ');
}
function formatClaimsCell_(claims) {
  const a = Array.from(new Set((claims || []).map(normalizeClaimKey_).filter(Boolean)));
  if (!a.length) return { cell: '', note: '' };
  const slice = a.slice(0, DETAILS_CLAIM_CELL_MAX);
  const more = a.length - slice.length;
  const cell = more > 0 ? (slice.join('\n') + '\n…(+' + more + ' more)') : slice.join('\n');
  const note = more > 0 ? a.join('\n') : '';
  return { cell, note };
}
/**
 * Business rule for Details sheet:
 * - If reason indicates "Partner is not mapped" (legacy mapping)
 * - ...and Submission Date is before Jun 2025
 * => suppress this row from being appended into Details.
 */
function shouldSuppressPartnerNotMappedDetail_(submissionDateVal, reason) {
  if (!isAssociateMappingEnabled_()) return false;
  const r = String(reason || '');
  if (!r) return false;
  // exact phrase match first (fast path), fallback to case-insensitive contains
  const hit = (r.indexOf(DETAILS_SUPPRESS_PARTNER_NOT_MAPPED_REASON) !== -1) ||
    /partner\s+is\s+not\s+mapped\s+to\s+any\s+associate/i.test(r);
  if (!hit) return false;
  const d = normalizeDate_(submissionDateVal);
  if (!d) return false; // if date is missing/unparseable, keep the row (safer)
  const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
  return dd.getTime() < DETAILS_SUPPRESS_PARTNER_NOT_MAPPED_CUTOFF.getTime();
}
/**
 * Append to Details:
 * - legacy suppression
 * - dedupe by Claim tokens (existing Details + current run)
 * - optional grouping by (targetSheet, partner, lastStatus, reason)
 */
/** Resolve workbook profile for Details filtering (best-effort, backward compatible). */
function resolveDetailsProfile_(opts) {
  // 1) explicit call-site override
  try {
    if (opts && opts.profile != null && opts.profile !== '') return String(opts.profile);
  } catch (e) {}
  // 2) common globals used across modules (best-effort)
  try { if (typeof ACTIVE_PROFILE !== 'undefined' && ACTIVE_PROFILE) return String(ACTIVE_PROFILE); } catch (e2) {}
  try { if (typeof CURRENT_PROFILE !== 'undefined' && CURRENT_PROFILE) return String(CURRENT_PROFILE); } catch (e3) {}
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME) {
      if (RUNTIME.profile) return String(RUNTIME.profile);
      if (RUNTIME.workbookProfile) return String(RUNTIME.workbookProfile);
    }
  } catch (e4) {}
  try { if (typeof CACHE !== 'undefined' && CACHE && CACHE.profile) return String(CACHE.profile); } catch (e5) {}
  return '';
}
function isAdminProfile_(profile) {
  const p = String(profile || '').trim();
  if (!p) return false;
  return p.toLowerCase() === 'admin';
}
/**
 * Append to Details:
 * - legacy suppression (Details only)
 * - Admin suppression: hide non-routed warning row
 * - dedupe by Claim tokens (existing Details + current run)
 * - per-claim rows (no grouping) to keep Submission Date unambiguous
 *
 * @param {Array<Object>} rows
 * @param {Object=} opts - optional { profile: 'Admin'|'PIC'|... }
 * @return {number} appended row count
 */
function appendDetailsRows_(rows, opts) {
  if (isDryRun_()) return 0;
  if (!rows || !rows.length) return 0;
  // Mirror unmapped partner / status warnings into Log sheet (short, per spec)
  try { mirrorMappingErrorsToLog_(rows); } catch (e0) {}
  const profile = resolveDetailsProfile_(opts);
  const isAdmin = isAdminProfile_(profile);
  // 1) legacy suppression (Details only) + Admin suppression
  const filtered = rows.filter(r => {
    if (!r) return false;
    // Admin: remove the "Status is not routed..." noise row
    if (isAdmin) {
      const reason = String(r.reason || '');
      if (/status\s+is\s+not\s+routed/i.test(reason) || /add\s+to\s+routing\s+map/i.test(reason)) return false;
    }
    // New flow: no Associate/Partner auto-mapping; drop legacy 'Partner not mapped' reasons.
    if (!isAssociateMappingEnabled_() && isPartnerNotMappedReasonText_(r.reason)) return false;
    // Business rule: suppress "Partner is not mapped..." for Submission Date < Jun 2025
    if (shouldSuppressPartnerNotMappedDetail_(r.submissionDateVal, r.reason)) return false;
    const v = r.submissionDateVal;
    if (v == null || v === '') return true;
    try {
      if (typeof isLegacyBySubmission_ === 'function') return !isLegacyBySubmission_(v);
    } catch (eL) {}
    return true;
  });
  if (!filtered.length) return 0;
  const sh = getDetailsSheet_();
  ensureDetailsLayout_(sh);
  // 2) dedupe by claim tokens (skip already logged)
  const existingSet = getExistingDetailsClaimSet_(sh);
  const runSet = CACHE.details.runClaimSet || (CACHE.details.runClaimSet = new Set());
  const deduped = [];
  for (let i = 0; i < filtered.length; i++) {
    const r = filtered[i];
    // Support single claim OR grouped claim text
    const tokens = extractClaimTokens_(r.claim || '');
    if (tokens.length) {
      let anyNew = false;
      for (let k = 0; k < tokens.length; k++) {
        if (!existingSet.has(tokens[k]) && !runSet.has(tokens[k])) { anyNew = true; break; }
      }
      if (!anyNew) continue; // all tokens already known
      for (let k = 0; k < tokens.length; k++) runSet.add(tokens[k]);
    }
    deduped.push(r);
  }
  if (!deduped.length) return 0;
  // 3) stable ordering to visually "group" by Partner + Last Status
  deduped.sort((a, b) => {
    const ap = String(a.partner || '').toLowerCase();
    const bp = String(b.partner || '').toLowerCase();
    if (ap < bp) return -1;
    if (ap > bp) return 1;
    const as = String(a.lastStatus || '').toLowerCase();
    const bs = String(b.lastStatus || '').toLowerCase();
    if (as < bs) return -1;
    if (as > bs) return 1;
    const ac = normalizeClaimKey_(a.claim || '');
    const bc = normalizeClaimKey_(b.claim || '');
    if (ac < bc) return -1;
    if (ac > bc) return 1;
    return 0;
  });
  // 4) append rows (one row per claim to keep submission date unambiguous)
  const ts = new Date();
  const startNo = getNextDetailsNo_(sh);
  const startRow = sh.getLastRow() + 1;
  const out = [];
  const claimNotes = [];
  for (let i = 0; i < deduped.length; i++) {
    const r = deduped[i];
    // Submission Date: normalize to date-only
    let subDate = '';
    const d = normalizeDate_(r.submissionDateVal);
    if (d) subDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
    const claim = normalizeClaimKey_(r.claim || '') || String(r.claim || '');
    const rowNo = (r.rowNumber != null && r.rowNumber !== '') ? String(r.rowNumber) : '';
    const baseReason = String(r.reason || '');
    const reason = baseReason + (rowNo ? (' | Row: ' + rowNo) : '');
    out.push([
      startNo + i,
      ts,
      subDate,
      r.targetSheet || '',
      claim,
      r.partner || '',
      r.lastStatus || '',
      reason
    ]);
    // Notes: if claim contains multiple tokens, keep full list in note (rare)
    try {
      const tokens = extractClaimTokens_(r.claim || '');
      claimNotes.push(tokens && tokens.length > 1 ? tokens.join('\n') : '');
    } catch (eN) {
      claimNotes.push('');
    }
  }
  safeSetValues_(sh.getRange(startRow, DETAILS_LAYOUT.START_COL, out.length, DETAILS_LAYOUT.COLS), out);
  // Date formats: Submission Date col
  try {
    safeSetNumberFormat_(
      sh.getRange(startRow, DETAILS_LAYOUT.START_COL + DETAILS_COLUMNS.SUBMISSION_DATE, out.length, 1),
      DETAILS_FORMATS.DATE
    );
  } catch (e) {}
  // Notes on Claim Number cell
  try {
    const claimColAbs = DETAILS_LAYOUT.START_COL + DETAILS_COLUMNS.CLAIM_NUMBER;
    safeSetNotes_(sh.getRange(startRow, claimColAbs, out.length, 1), claimNotes);
  } catch (e2) {}
  CACHE.details.nextNo = startNo + out.length;
  // RUNTIME is optional across repos; keep this file safe if it doesn't exist.
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME) {
      if (typeof RUNTIME.detailsAppendedThisRun !== 'number') RUNTIME.detailsAppendedThisRun = 0;
      RUNTIME.detailsAppendedThisRun += out.length;
    }
  } catch (eR) {}
  if (uiFlag_('REALTIME_DETAILS_APPEND', false)) flushNow_();
  return out.length;
}
/** =========================
 * Run metrics sink (append-only)
 * ========================= */

const RUN_METRICS_SHEET_NAME = '_RunMetrics';
const RUN_METRICS_HEADER = Object.freeze([
  'Timestamp',
  'Flow',
  'Request ID',
  'Status',
  'Duration(ms)',
  'Processed',
  'Failed',
  'Ops Updated',
  'Rows Moved',
  'Notes'
]);

function ensureRunMetricsSheet_() {
  const enabled = !!(CONFIG && CONFIG.features && CONFIG.features.enableRunMetrics);
  if (!enabled) return null;
  if (DRY_RUN) return null;
  const ss = SpreadsheetApp.openById(String(CONFIG.logSpreadsheetId));
  let sh = ss.getSheetByName(RUN_METRICS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(RUN_METRICS_SHEET_NAME);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, RUN_METRICS_HEADER.length).setValues([RUN_METRICS_HEADER]);
    try { sh.setFrozenRows(1); } catch (e0) {}
  }
  return sh;
}

/**
 * Append a run metric row.
 * metric: {flow, requestId, status, durationMs, processed, failed, opsUpdated, rowsMoved, notes}
 */
function recordRunMetrics_(metric) {
  const enabled = !!(CONFIG && CONFIG.features && CONFIG.features.enableRunMetrics);
  if (!enabled) return false;
  if (DRY_RUN) return false;
  try {
    const sh = ensureRunMetricsSheet_();
    if (!sh) return false;
    const m = metric || {};
    const row = [
      new Date(),
      String(m.flow || ''),
      String(m.requestId || ''),
      String(m.status || ''),
      Number(m.durationMs || 0),
      Number(m.processed || 0),
      Number(m.failed || 0),
      Number(m.opsUpdated || 0),
      Number(m.rowsMoved || 0),
      String(m.notes || '')
    ];
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return true;
  } catch (e) {
    try { Logger.log('recordRunMetrics_ failed: ' + String(e && e.message ? e.message : e)); } catch (e2) {}
    return false;
  }
}
