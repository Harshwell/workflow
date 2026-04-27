/***************************************************************
 * 06a_EntryPoints.gs  (SPLIT FROM 06.gs)
 * Entry points, installers, master-sheet ensure, email ingest
 ***************************************************************/
'use strict';

/** ---------- Global fallbacks stored as properties (no redeclare) ---------- */
(function init06Fallbacks_() {
  // Apps Script global safety: globalThis may exist, but don't assume.
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();

  if (!g.__FORMATS_FALLBACK) {
    g.__FORMATS_FALLBACK = Object.freeze({
      DATE: 'd mmm yy',
      DATETIME: 'd mmm yy, HH:mm',
      DATETIME_LONG: 'MMMM d, yyyy, h:mm AM/PM',
      TIMESTAMP: 'd mmm, HH:mm',
      INT: '0',
      MONEY0: '#,##0'
    });
  }

  if (!g.__OPTIONAL_RULES_FALLBACK) {
    g.__OPTIONAL_RULES_FALLBACK = Object.freeze({
      EVBIKE_ONLY_FOR_PIC: '',
      B2B_ONLY_FOR_PIC: ''
    });
  }


  // Ensure RUNTIME exists to avoid ReferenceError during enable flags assignment.
  // (If another file defines RUNTIME, this does nothing.)
  if (typeof RUNTIME === 'undefined') {
    g.RUNTIME = {};
  }
})();

function _fmt06_() {
  // Prefer real FORMATS; fallback to global object
  try {
    if (typeof FORMATS !== 'undefined' && FORMATS) return FORMATS;
  } catch (e) {}
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();
  return g.__FORMATS_FALLBACK;
}

function _optRules06_() {
  // Prefer real OPTIONAL_SHEETS_RULES; fallback to global object
  try {
    if (typeof OPTIONAL_SHEETS_RULES !== 'undefined' && OPTIONAL_SHEETS_RULES) return OPTIONAL_SHEETS_RULES;
  } catch (e) {}
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();
  return g.__OPTIONAL_RULES_FALLBACK;
}

/** =========================
 * Config validation
 * ========================= */
function resolveSpreadsheetKey_(pic) {
  const p = String(pic || '').trim();
  try {
    if (p && CONFIG && CONFIG.spreadsheets && CONFIG.spreadsheets[p]) return p;
    if (CONFIG && CONFIG.spreadsheets) {
      if (CONFIG.spreadsheets.Master) return 'Master';
      if (CONFIG.spreadsheets.Admin) return 'Admin';
      const keys = Object.keys(CONFIG.spreadsheets);
      return keys.length ? keys[0] : '';
    }
  } catch (e) {}
  return p || 'Master';
}


function validateConfigForPic_(pic) {
  const key = resolveSpreadsheetKey_(pic);
  if (!key) throw new Error('Spreadsheet profile key is empty.');
  if (!CONFIG || !CONFIG.spreadsheets) throw new Error('CONFIG.spreadsheets is missing.');
  if (!CONFIG.spreadsheets[key]) throw new Error('Unknown profile key: ' + key);

  const ssId = String(CONFIG.spreadsheets[key] || '');
  if (ssId.indexOf('PAKAI_SPREADSHEET_ID_') > -1) {
    const propKey = 'SPREADSHEET_ID_' + String(key).toUpperCase().replace(/[^A-Z0-9]+/g, '_');
    throw new Error(
      'Spreadsheet ID for key=' + key + ' is not set. ' +
      'Set Script Properties key "' + propKey + '" with the Spreadsheet ID, or hardcode CONFIG.spreadsheets.' + key + '.'
    );
  }
}

/** =========================
 * Details logging guards
 * ========================= */

/**
 * For Admin profile, suppress unmapped-partner logs for statuses that are not used
 * by Admin core sheets (Submission/OR/Start/Finish).
 */
function filterUnknownPartnerRowsForAdmin_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return [];

  const routingMap = (CONFIG && CONFIG.statusRoutingAdmin) ? CONFIG.statusRoutingAdmin : {};
  // Prefer shared helper from 05b if available; otherwise use local index builder.
  const idx = (typeof compileRoutingIndex_ === 'function')
    ? compileRoutingIndex_(routingMap)
    : buildRoutingIndex06_(routingMap);

  const coreOrSheet = (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY && OPS_ROUTING_POLICY.SHEETS && OPS_ROUTING_POLICY.SHEETS.OR_OLD)
    ? String(OPS_ROUTING_POLICY.SHEETS.OR_OLD || '').trim()
    : 'OR - OLD';
  const CORE = { 'Submission': true, 'OR': true, 'Start': true, 'Finish': true };
  if (coreOrSheet) CORE[coreOrSheet] = true;

  return list.filter(r => {
    const st = String((r && r.lastStatus) ? r.lastStatus : '').trim();
    if (!st) return false;
    const dest = idx[st] || [];
    for (let i = 0; i < dest.length; i++) {
      if (CORE[dest[i]]) return true;
    }
    return false;
  });
}

function buildRoutingIndex06_(routingMap) {
  const map = routingMap || {};
  const idx = {};
  Object.keys(map).forEach(sheetName => {
    const sheetKey = String(sheetName || '').trim();
    // Internal buckets are not real sheets and must not become relocation targets.
    if (!sheetKey || /^__/.test(sheetKey)) return;
    (map[sheetName] || []).forEach(status => {
      const key = String(status || '').trim();
      if (!key) return;
      if (!idx[key]) idx[key] = [];
      idx[key].push(sheetKey);
    });
  });
  return idx;
}

function getUnmappedPartnerMinSubmissionDate06_() {
  // Cache per execution to avoid reparsing the same policy value repeatedly.
  if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME._unmappedPartnerMinSubmissionDate instanceof Date) {
    return RUNTIME._unmappedPartnerMinSubmissionDate;
  }

  const fallback = new Date(2025, 5, 1); // June 1, 2025 (local)
  let raw = null;
  try {
    if (typeof DETAILS_LOG_POLICY !== 'undefined' && DETAILS_LOG_POLICY) {
      raw = DETAILS_LOG_POLICY.UNMAPPED_PARTNER_MIN_SUBMISSION_DATE || null;
    }
    if (!raw && CONFIG && CONFIG.detailsLogPolicy) {
      raw = CONFIG.detailsLogPolicy.UNMAPPED_PARTNER_MIN_SUBMISSION_DATE || null;
    }
  } catch (e) {}

  const d = raw ? __toDate06_(raw) : null;
  const out = d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : fallback;
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME) RUNTIME._unmappedPartnerMinSubmissionDate = out;
  } catch (e2) {}
  return out;
}

function filterUnknownPartnerRowsByMinSubmissionDate06_(rows, minDate) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return [];
  const min = minDate ? new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate()) : null;
  if (!min) return list;

  return list.filter(r => {
    const d = __toDate06_(r && r.submissionDateVal ? r.submissionDateVal : null);
    if (!d) return true; // if missing/invalid date, keep (can't safely exclude)
    const day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return day.getTime() >= min.getTime();
  });
}


/** =========================
 * Lock wrapper
 * ========================= */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (lockErr) {
    // Best-effort logging even when lock fails (but do not crash because of logging).
    try {
      CACHE.log = { ss: null, sh: null, ensured: false, nextRow: LOG_LAYOUT.DETAIL_START_ROW, segNo: 0 };
      CACHE.details = { ss: null, sh: null, ensured: false, nextNo: null, existingClaimSet: null, runClaimSet: new Set() };
      resetLogState_();
      if (PIPELINE_FLAGS.CLEAR_LOG_BEFORE_RUN) clearLogSheet_();
      logLine_('LOCK', 'Lock acquisition failed', '', String(lockErr), 'ERROR');
    } catch (e2) {}
    throw lockErr;
  }

  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

/** =========================
 * Run state reset
 * ========================= */
function resetRunState_() {
  try { if (typeof resetRuntime_ === 'function') resetRuntime_(); } catch (err0) {}
  RUNTIME.detailsAppendedThisRun = 0;
  RUNTIME.useSubmittedDatetime = false;

  // Single-master: optional processors are enabled by default.
  RUNTIME.enableEvBike = true;
  RUNTIME.enableB2B = true;
  RUNTIME.hasAgingFiles = false;

  CACHE.log = { ss: null, sh: null, ensured: false, nextRow: LOG_LAYOUT.DETAIL_START_ROW, segNo: 0 };
  CACHE.details = { ss: null, sh: null, ensured: false, nextNo: null, existingClaimSet: null, runClaimSet: new Set() };
  resetLogState_();
}


/** =========================
 * Overview timing helpers (Pulling Time / Processing Time)
 * =========================
 * Requirement:
 * - Overview!C3: script start timestamp (next to label "Pulling Time" at Overview!B3)
 * - Overview!C4: duration from start to finish (next to label "Processing Time" at Overview!B4)
 * Applies to both flows: onFormSubmit and runEmailIngest.
 */
function __getSpreadsheetIdForKey06_(pic) {
  const key = resolveSpreadsheetKey_(pic);
  let ssId = '';
  try {
    if (CONFIG && CONFIG.spreadsheets && CONFIG.spreadsheets[key]) ssId = String(CONFIG.spreadsheets[key] || '').trim();
  } catch (e) {}

  if (!ssId) {
    try { ssId = String((CONFIG && CONFIG.masterSpreadsheetId) ? CONFIG.masterSpreadsheetId : '').trim(); } catch (e2) {}
  }
  if (!ssId) return '';

  // Support placeholder patterns where spreadsheet IDs are stored in Script Properties.
  if (ssId.indexOf('PAKAI_SPREADSHEET_ID_') > -1) {
    const propKey = 'SPREADSHEET_ID_' + String(key).toUpperCase().replace(/[^A-Z0-9]+/g, '_');
    try {
      const v = PropertiesService.getScriptProperties().getProperty(propKey);
      if (v) ssId = String(v).trim();
    } catch (e3) {}
  }

  return ssId;
}

function __tryOpenSpreadsheetForKey06_(pic) {
  const id = __getSpreadsheetIdForKey06_(pic);
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function __pluralizeUnit06_(n, unit) {
  const num = Number(n) || 0;
  return String(num) + ' ' + unit + (num === 1 ? '' : 's');
}

function __formatProcessingDuration06_(durationMs) {
  const ms = Number(durationMs) || 0;
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const seconds = totalSeconds % 60;
  const totalMinutes = Math.floor(totalSeconds / 60);
  const minutes = totalMinutes % 60;
  const hours = Math.floor(totalMinutes / 60);

  const parts = [];
  if (hours > 0) parts.push(__pluralizeUnit06_(hours, 'Hour'));
  if (minutes > 0 || hours > 0) parts.push(__pluralizeUnit06_(minutes, 'Minute'));
  parts.push(__pluralizeUnit06_(seconds, 'Second'));
  return parts.join(' ');
}

function __formatPullingTimestamp06_(ss, d) {
  const dt = d ? new Date(d) : new Date();

  const months = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'
  ];

  let tz = null;
  try { tz = (ss && ss.getSpreadsheetTimeZone) ? ss.getSpreadsheetTimeZone() : null; } catch (e) {}
  if (!tz) {
    try { tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'GMT'; } catch (e2) { tz = 'GMT'; }
  }

  const dayStr = Utilities.formatDate(dt, tz, 'd');     // 1..31
  const monthStr = Utilities.formatDate(dt, tz, 'M');   // 1..12
  const yearStr = Utilities.formatDate(dt, tz, 'yyyy');
  const hhStr = Utilities.formatDate(dt, tz, 'hh');     // 01..12
  const mmStr = Utilities.formatDate(dt, tz, 'mm');
  let ap = Utilities.formatDate(dt, tz, 'a');           // AM/PM (usually)

  const mIdx = Math.max(1, Math.min(12, parseInt(monthStr, 10) || 1)) - 1;
  const monthName = months[mIdx] || months[0];

  ap = String(ap || '').trim().toUpperCase();
  const day = parseInt(dayStr, 10);
  const dayOut = isNaN(day) ? String(dayStr) : String(day);

  return dayOut + ' ' + monthName + ' ' + yearStr + ', ' + hhStr + ':' + mmStr + ' ' + ap;
}

function __findOverviewLabelRow06_(overviewSheet, label) {
  if (!overviewSheet) return null;
  const want = String(label || '').trim().toLowerCase();
  if (!want) return null;

  const last = Math.max(1, overviewSheet.getLastRow ? overviewSheet.getLastRow() : 1);
  const scan = Math.max(10, Math.min(200, last));

  // Labels are expected on column B.
  const vals = overviewSheet.getRange(1, 2, scan, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim().toLowerCase();
    if (v === want) return i + 1;
  }
  return null;
}

function __writeOverviewValueNextToLabel06_(ss, label, defaultRow, value) {
  if (!ss) return false;

  let sh = null;
  try { sh = ss.getSheetByName('Overview'); } catch (e) {}
  if (!sh) return false;

  let row = null;
  try { row = __findOverviewLabelRow06_(sh, label); } catch (e2) {}
  if (!row) row = defaultRow;
  if (!row) return false;

  try { if (typeof DRY_RUN !== 'undefined' && DRY_RUN) return true; } catch (e3) {}

  try {
    // Values are written in the cell to the right (column C).
    sh.getRange(row, 3).setValue(value);
    return true;
  } catch (e4) {
    return false;
  }
}

function __logOverviewStart06_(pic, startedAt, ssMaybe) {
  const ss = ssMaybe || __tryOpenSpreadsheetForKey06_(pic);
  if (!ss) return null;
  const stamp = __formatPullingTimestamp06_(ss, startedAt);
  __writeOverviewValueNextToLabel06_(ss, 'Pulling Time', 3, stamp);
  return ss;
}

function __logOverviewDuration06_(pic, startedAt, ssMaybe) {
  const ss = ssMaybe || __tryOpenSpreadsheetForKey06_(pic);
  if (!ss || !startedAt) return null;
  const dur = (new Date()).getTime() - (new Date(startedAt)).getTime();
  const text = __formatProcessingDuration06_(dur);
  __writeOverviewValueNextToLabel06_(ss, 'Processing Time', 4, text);
  return ss;
}


/** =========================
 * Trigger entrypoint
 * ========================= */
function onFormSubmit(e) {
  try { if (typeof resetRuntime_ === 'function') resetRuntime_(); } catch (err0) {}
  return withLock_(() => {
    resetRunState_();
    if (PIPELINE_FLAGS.CLEAR_LOG_BEFORE_RUN) clearLogSheet_();

    const startedAt = new Date();
    let ssTiming = null;
    try { ssTiming = __logOverviewStart06_('Master', startedAt); } catch (e0) {}

    const runId = (typeof getRunId_ === 'function') ? getRunId_() : (function() { try { return Utilities.getUuid(); } catch (e) { return String(new Date().getTime()); } })();
    const req = __buildFormRunRequest06a_(e);

    const flowLabel = String(req.flowLabel || '').toUpperCase() || 'MAIN';
    try { setLogRunContext_(flowLabel, runId); } catch (e1) {}
    try {
      if (typeof RUNTIME !== 'undefined' && RUNTIME) {
        RUNTIME.flow = flowLabel;
        RUNTIME.runId = runId;
        RUNTIME.source = 'FORM';
      }
    } catch (e2) {}

    try { setProgressForFlow_(flowLabel, 0, 'Starting...', { runId: runId, prefixFlowInStep: true }); } catch (e3) {}
    try { setProgressForFlow_(flowLabel, 0, 'Starting…', { runId: runId, prefixFlowInStep: true }); } catch (e3) {}
    try { logLine_('FORM', 'Trigger received', 'flow=' + flowLabel + ' runId=' + runId, 'files=' + (req.allFileIds ? req.allFileIds.length : 0), 'INFO'); } catch (e4) {}

    const seg = startSegment_(flowLabel + '_FORM', 'Pipeline run (onFormSubmit)');
    try {
      if (flowLabel === 'SUB') {
        const resSub = runSubFromFormDrive06a_(req, runId);
        endSegment_(seg, 'ok', (resSub && resSub.message) ? resSub.message : 'OK', (resSub && resSub.severity) ? resSub.severity : 'INFO');
        return resSub;
      }

      // Default: MAIN
      const mainIds = Array.isArray(req.mainFileIds) ? req.mainFileIds : [];
      if (!mainIds.length) throw new Error('No uploaded file detected for MAIN run.');

      const resMain = runPipeline_('Master', mainIds, { flow: 'main', source: 'FORM_MAIN', runId: runId });
      endSegment_(seg, 'routed=' + (resMain && resMain.routedTotal ? resMain.routedTotal : 0), (resMain && resMain.message) ? resMain.message : 'OK', (resMain && resMain.severity) ? resMain.severity : 'INFO');
      return resMain;

    } catch (err) {
      endSegment_(seg, '', 'FATAL: ' + err, 'ERROR');
      throw err;
    } finally {
      try { __logOverviewDuration06_('Master', startedAt, ssTiming); } catch (e5) {}
    }
  });
}


/**
 * Build a deterministic run request from a Form submission.
 * Supports:
 * - Flow selector field (MAIN / SUB)
 * - One upload field (multi-file) OR split SUB old/new upload fields
 * - Auto-detect fallback:
 *   - If a "main" dashboard file is detected => MAIN
 *   - Else if both SUB OLD+NEW files are detected => SUB
 *   - Else => MAIN
 */
function __buildFormRunRequest06a_(e) {
  const sub = (typeof getSubmissionContext_ === 'function') ? getSubmissionContext_(e) : { fileIds: [], sub: { oldFileIds: [], newFileIds: [] } };

  const idsPrimary = Array.isArray(sub.fileIds) ? sub.fileIds.map(String).map(s => s.trim()).filter(Boolean) : [];
  const idsOld = (sub.sub && Array.isArray(sub.sub.oldFileIds)) ? sub.sub.oldFileIds.map(String).map(s => s.trim()).filter(Boolean) : [];
  const idsNew = (sub.sub && Array.isArray(sub.sub.newFileIds)) ? sub.sub.newFileIds.map(String).map(s => s.trim()).filter(Boolean) : [];

  const allFileIds = Array.from(new Set(idsPrimary.concat(idsOld).concat(idsNew))).filter(Boolean);

  const reqFlow = (sub && (sub.requestedFlowLabel || sub.flow)) ? String(sub.requestedFlowLabel || sub.flow || '') : '';
  let flowLabel = '';
  try { flowLabel = normalizeFlowLabel_(reqFlow); } catch (e0) { flowLabel = String(reqFlow || '').toUpperCase(); }
  if (flowLabel !== 'MAIN' && flowLabel !== 'SUB') flowLabel = '';

  // Prefer explicit split fields for SUB old/new.
  const pickNewestId = (ids) => {
    if (!ids || !ids.length) return '';
    if (typeof pickNewestFileId_ === 'function') return String(pickNewestFileId_(ids) || '');
    // fallback: pick by updated time if available
    const metas = ids.map(id => (typeof getFileMetaCached_ === 'function') ? getFileMetaCached_(id) : null).filter(Boolean);
    if (!metas.length) return String(ids[0] || '');
    let best = metas[0];
    for (let i = 1; i < metas.length; i++) if ((metas[i].updatedMs || 0) > (best.updatedMs || 0)) best = metas[i];
    return String(best.id || ids[0] || '');
  };

  let oldId = pickNewestId(idsOld);
  let newId = pickNewestId(idsNew);

  // Detect SUB old/new from all uploads (filename heuristics)
  try {
    if ((!oldId || !newId) && typeof pickSubDashboardAttachments04_ === 'function' && typeof getFileMetaCached_ === 'function') {
      const metas = allFileIds.map(id => getFileMetaCached_(id)).filter(Boolean).map(m => ({ id: m.id, name: m.name }));
      const picked = pickSubDashboardAttachments04_(metas);
      if (!oldId && picked && picked.oldAttachment && picked.oldAttachment.id) oldId = String(picked.oldAttachment.id);
      if (!newId && picked && picked.newAttachment && picked.newAttachment.id) newId = String(picked.newAttachment.id);
    }
  } catch (e1) {}

  // Auto-detect MAIN vs SUB (only when not explicitly selected)
  if (!flowLabel) {
    try {
      if (typeof classifyFiles_ === 'function') {
        const buckets = classifyFiles_(allFileIds);
        if (buckets && buckets.main && buckets.main.length) flowLabel = 'MAIN';
      }
    } catch (e2) {}

    if (!flowLabel && oldId && newId) flowLabel = 'SUB';
    if (!flowLabel) flowLabel = 'MAIN';
  }

  // MAIN input: pass all primary uploads (runPipeline_ will classify main/aging)
  const mainFileIds = idsPrimary.length ? idsPrimary : allFileIds;

  return Object.freeze({
    flowLabel: flowLabel,
    mainFileIds: mainFileIds,
    oldFileId: oldId,
    newFileId: newId,
    allFileIds: allFileIds,
    submission: sub
  });
}


/**
 * Manual SUB runner from Drive file IDs (uploaded via Form).
 * Uses the exact same SUB business logic as the email ingest, but without email cleanup.
 */
function runSubFromFormDrive06a_(req, runId) {
  const oldId = String((req && req.oldFileId) ? req.oldFileId : '').trim();
  const newId = String((req && req.newFileId) ? req.newFileId : '').trim();
  if (!oldId || !newId) {
    throw new Error(
      'SUB requires 2 files (OLD + NEW). ' +
      'Tip: upload both "List of Claims with Aging" and "List of Claims with Aging (Standardization)".'
    );
  }

  // Open master workbook
  const masterKey = 'Master';
  validateConfigForPic_(masterKey);
  const masterSs = __tryOpenSpreadsheetForKey06_(masterKey);
  try { if (masterSs && typeof getTzSafe_ === 'function') masterSs.setSpreadsheetTimeZone(getTzSafe_()); } catch (eTzM) {}

  // Resolve SUB spec
  const subFlow = (typeof CONFIG === 'object' && CONFIG && CONFIG.subFlow) ? CONFIG.subFlow : {};
  const rawOldName = String(subFlow.RAW_OLD_SHEET_NAME || 'Raw OLD').trim();
  const rawNewName = String(subFlow.RAW_NEW_SHEET_NAME || 'Raw NEW').trim();

  __ensureSheetByNameSub06a_(masterSs, rawOldName);
  __ensureSheetByNameSub06a_(masterSs, rawNewName);

  // Operational sheets allow-list
  const defaultOpSheets = [
    'Submission',
    'Ask Detail',
    'OR - OLD',
    'SC - Farhan',
    'SC - Meilani',
    'SC - Meindar',
    'Start',
    'Finish',
    'PO',
    'B2B',
    'Special Case'
  ];
  const opSheets = (Array.isArray(subFlow.OPERATIONAL_SHEETS) && subFlow.OPERATIONAL_SHEETS.length)
    ? subFlow.OPERATIONAL_SHEETS.map(s => String(s || '').trim()).filter(Boolean)
    : defaultOpSheets.slice();

  // Ensure SC fallback/quarantine sheet participates in relocate + sort.
  try {
    const fb = __getScFallbackSheet06a_();
    if (fb && opSheets.indexOf(fb) < 0) opSheets.push(fb);
    __ensureSubScFallbackSheetExists06a_(masterSs, fb);
  } catch (eFb) {}

  const sortSpecs = (Array.isArray(subFlow.SORT_SPECS) && subFlow.SORT_SPECS.length) ? subFlow.SORT_SPECS : null;

  // Defer trash of the *uploaded* files until success
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME) {
      RUNTIME.deferTrashUploadedFiles = !!(PIPELINE_FLAGS && PIPELINE_FLAGS.TRASH_UPLOADED_FILES) && (typeof DRY_RUN === 'undefined' || !DRY_RUN);
      if (!Array.isArray(RUNTIME.filesToTrash)) RUNTIME.filesToTrash = [];
      RUNTIME.flow = 'SUB';
      if (runId) RUNTIME.runId = String(runId || '');
      RUNTIME.source = 'FORM_SUB';
    }
  } catch (eR0) {}

  try { if (typeof enqueueTrashFileId_ === 'function') { enqueueTrashFileId_(oldId); enqueueTrashFileId_(newId); } } catch (eQ) {}

  // Load blobs (Drive file => Blob)
  const oldBlob = DriveApp.getFileById(oldId).getBlob();
  const newBlob = DriveApp.getFileById(newId).getBlob();

  const core = __runSubCore06a_(masterSs, oldBlob, newBlob, {
    rawOldName: rawOldName,
    rawNewName: rawNewName,
    opSheets: opSheets,
    sortSpecs: sortSpecs,
    doTrashFlush: true
  });

  return core;
}


/**
 * Core SUB logic shared by different entrypoints.
 * @param {SpreadsheetApp.Spreadsheet} masterSs
 * @param {Blob} oldBlob
 * @param {Blob} newBlob
 * @param {Object} opt {rawOldName, rawNewName, opSheets, sortSpecs, doTrashFlush}
 */
function __runSubCore06a_(masterSs, oldBlob, newBlob, opt) {
  const o = opt || {};
  const rawOldName = String(o.rawOldName || 'Raw OLD').trim();
  const rawNewName = String(o.rawNewName || 'Raw NEW').trim();
  const opSheets = Array.isArray(o.opSheets) ? o.opSheets : [];
  const sortSpecs = o.sortSpecs || null;
  const doTrashFlush = !!o.doTrashFlush;

  const startedAt = new Date();
  try { setProgressForFlow_('SUB', 0.05, 'Snapshot PREV...', { prefixFlowInStep: true }); } catch (e0) {}
  try { setProgressForFlow_('SUB', 0.05, 'Snapshot PREV…', { prefixFlowInStep: true }); } catch (e0) {}

  // WebApp snapshots (best effort)
  try {
    if (typeof webappMovementSnapshotPrevForSub06c_ === 'function') {
      webappMovementSnapshotPrevForSub06c_(masterSs, rawOldName, rawNewName);
    }
  } catch (eWp0) {
    try { logLine_('WEBAPP_SNAP_PREV_ERR', 'Snapshot PREV failed (non-fatal)', String(eWp0), '', 'WARN'); } catch (eWp2) {}
  }

  try { setProgressForFlow_('SUB', 0.20, 'Process OLD...', { prefixFlowInStep: true }); } catch (e1) {}
  try { setProgressForFlow_('SUB', 0.20, 'Process OLD…', { prefixFlowInStep: true }); } catch (e1) {}

  const rOld = __processSubAttachment06a_(masterSs, oldBlob, {
    dbTag: 'OLD',
    rawSheetName: rawOldName,
    operationalSheetNames: opSheets
  });
  if (!rOld || String(rOld.severity || '').toUpperCase() === 'ERROR') {
    try { setProgressForFlow_('SUB', 1, 'Failed (OLD)', { prefixFlowInStep: true }); } catch (e2) {}
    return { severity: 'ERROR', message: 'SUB OLD failed', old: rOld };
  }

  try { setProgressForFlow_('SUB', 0.55, 'Process NEW...', { prefixFlowInStep: true }); } catch (e3) {}
  try { setProgressForFlow_('SUB', 0.55, 'Process NEW…', { prefixFlowInStep: true }); } catch (e3) {}

  const rNew = __processSubAttachment06a_(masterSs, newBlob, {
    dbTag: 'NEW',
    rawSheetName: rawNewName,
    operationalSheetNames: opSheets
  });
  if (!rNew || String(rNew.severity || '').toUpperCase() === 'ERROR') {
    try { setProgressForFlow_('SUB', 1, 'Failed (NEW)', { prefixFlowInStep: true }); } catch (e4) {}
    return { severity: 'ERROR', message: 'SUB NEW failed', old: rOld, new: rNew };
  }

  try { setProgressForFlow_('SUB', 0.78, 'Relocate + sort...', { prefixFlowInStep: true }); } catch (e5) {}
  try { setProgressForFlow_('SUB', 0.78, 'Relocate + sort…', { prefixFlowInStep: true }); } catch (e5) {}

  const relocateSheets = __getSubRelocationSheetNames06a_(opSheets);
  const relocateRes = __relocateOperationalRowsByLastStatusSub06a_(masterSs, relocateSheets);
  const sortRes = __sortOperationalSheetsSub06a_(masterSs, opSheets, sortSpecs);

  // WebApp movement tracking (best effort)
  try {
    if (typeof webappMovementSnapshotCurrAndTrackForSub06c_ === 'function') {
      webappMovementSnapshotCurrAndTrackForSub06c_(masterSs, rawOldName, rawNewName);
    }
  } catch (eWp4) {
    try { logLine_('WEBAPP_MOVE_ERR', 'Movement tracking failed (non-fatal)', String(eWp4), '', 'WARN'); } catch (eWp5) {}
  }

  // Trash uploaded files only after successful SUB
  if (doTrashFlush) {
    try {
      if (PIPELINE_FLAGS && PIPELINE_FLAGS.TRASH_UPLOADED_FILES && (typeof DRY_RUN === 'undefined' || !DRY_RUN)) {
        if (typeof flushTrashQueueBestEffort_ === 'function') flushTrashQueueBestEffort_('SUB');
      }
    } catch (eT) { try { logLine_('WARN', 'Trash flush failed', '', String(eT), 'WARN'); } catch (e2) {} }
  }

  const durMs = new Date().getTime() - startedAt.getTime();
  try { setProgressForFlow_('SUB', 1, 'Done', { prefixFlowInStep: true }); } catch (e6) {}
  try { logLine_('SUB_DONE', 'Completed SUB run', __formatProcessingDuration06_(durMs), '', 'INFO'); } catch (e7) {}

  return {
    severity: 'INFO',
    message: 'SUB completed',
    old: rOld,
    new: rNew,
    relocated: relocateRes,
    sorted: sortRes
  };
}

function __getSubRelocationSheetNames06a_(sheetNames) {
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  // EV-Bike and Exclusion are user-managed optional buckets; SUB relocation must not move/delete their rows.
  const blocked = new Set(['ev-bike', 'exclusion']);
  return names.filter(function (name) {
    const key = String(name || '').trim().toLowerCase();
    return key && !blocked.has(key);
  });
}

function __ensureSubScFallbackSheetExists06a_(ss, fallbackSheetName) {
  const fb = String(fallbackSheetName || '').trim();
  if (!ss || !fb) return;

  function ensureClaimNumberHeader_(sh) {
    if (!sh) return;
    const lc = Math.max(sh.getLastColumn() || 1, 1);
    const hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (v) { return String(v || '').trim().toLowerCase(); });
    if (hdr.indexOf('claim number') >= 0 || hdr.indexOf('claim_number') >= 0 || hdr.indexOf('claim no') >= 0 || hdr.indexOf('claim_no') >= 0) return;
    const tmpl = ss.getSheetByName('SC - Farhan');
    if (tmpl) {
      const tlc = Math.max(tmpl.getLastColumn() || 1, 1);
      tmpl.getRange(1, 1, 1, tlc).copyTo(sh.getRange(1, 1, 1, tlc), { contentsOnly: false });
      try { sh.setFrozenRows(Math.max(tmpl.getFrozenRows() || 0, 1)); } catch (eT) { try { sh.setFrozenRows(1); } catch (eT2) {} }
      return;
    }
    sh.getRange(1, 1).setValue('Claim Number');
    try { sh.setFrozenRows(1); } catch (e3) {}
  }

  try {
    const exists = ss.getSheetByName(fb);
    if (exists) {
      ensureClaimNumberHeader_(exists);
      return;
    }
    if (typeof ensureScFallbackSheet05b_ === 'function') {
      ensureScFallbackSheet05b_(ss, fb, 'SC - Farhan');
      try { ensureClaimNumberHeader_(ss.getSheetByName(fb)); } catch (eH0) {}
      return;
    }
  } catch (e0) {}

  // Local fallback: create minimal sheet to avoid move-skip when SC keyword does not match.
  if (isDryRun_()) return;
  try {
    const sh = ss.insertSheet(fb);
    const tmpl = ss.getSheetByName('SC - Farhan');
    if (tmpl) {
      const lc = Math.max(tmpl.getLastColumn() || 1, 1);
      tmpl.getRange(1, 1, 1, lc).copyTo(sh.getRange(1, 1, 1, lc), { contentsOnly: false });
      try { sh.setFrozenRows(Math.max(tmpl.getFrozenRows() || 0, 1)); } catch (e1) { try { sh.setFrozenRows(1); } catch (e2) {} }
    } else {
      sh.getRange(1, 1).setValue('Claim Number');
      try { sh.setFrozenRows(1); } catch (e3) {}
    }
    ensureClaimNumberHeader_(sh);
  } catch (e4) {}
}



/** =========================
 * Trigger installer
 * ========================= */
function install() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const ss = SpreadsheetApp.openById(RESPONSES_SPREADSHEET_ID);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

/** =========================
 * Master workbook helpers
 * ========================= */
function ensureMasterSheets_(ss) {
  if (!ss) return;

  const rawName = (CONFIG && CONFIG.masterRawSheetName) ? CONFIG.masterRawSheetName : 'Raw Data';

  // Ensure Raw Data exists
  let raw = ss.getSheetByName(rawName);
  if (!raw) {
    if (DRY_RUN) return;
    raw = ss.insertSheet(rawName);
    raw.getRange(1, 1).setValue((CONFIG && CONFIG.headers && CONFIG.headers.claimNumber) ? CONFIG.headers.claimNumber : 'Claim Number');
    try { raw.setFrozenRows(1); } catch (e) {}
  }

  const mustHave = [
    // Operational
    'Submission', 'Ask Detail', 'OR - OLD', 'Start', 'Finish', 'SC - Farhan', 'SC - Meilani', 'SC - Meindar', 'SC - Unmapped', 'PO', 'Exclusion',
    // Optional
    'B2B', 'EV-Bike', 'Special Case'
  ];

  // Prefer 03 templates if available; otherwise create minimal sheets.
  const hasEnsurer = (typeof sv03_ensureSheetWithHeader_ === 'function') && (typeof SV03_TEMPLATES !== 'undefined' && SV03_TEMPLATES);

  mustHave.forEach(name => {
    if (ss.getSheetByName(name)) return;
    if (DRY_RUN) return;

    if (hasEnsurer) {
      if (name === 'PO' && SV03_TEMPLATES.OPS_PIC_PO) {
        sv03_ensureSheetWithHeader_(ss, 'PO', SV03_TEMPLATES.OPS_PIC_PO, 'Master');
      } else if (name === 'B2B' && SV03_TEMPLATES.B2B) {
        sv03_ensureSheetWithHeader_(ss, 'B2B', SV03_TEMPLATES.B2B, 'Master');
      } else if (name === 'EV-Bike' && SV03_TEMPLATES.EV_BIKE) {
        sv03_ensureSheetWithHeader_(ss, 'EV-Bike', SV03_TEMPLATES.EV_BIKE, 'Master');
      } else if (name === 'Special Case' && SV03_TEMPLATES.SPECIAL_CASE) {
        sv03_ensureSheetWithHeader_(ss, 'Special Case', SV03_TEMPLATES.SPECIAL_CASE, 'Master');
      } else if (SV03_TEMPLATES.OPS_PIC_DEFAULT) {
        sv03_ensureSheetWithHeader_(ss, name, SV03_TEMPLATES.OPS_PIC_DEFAULT, 'Master');
      } else {
        ss.insertSheet(name);
      }
    } else {
      ss.insertSheet(name);
    }
  });
}

/** Ensure custom tail columns on Raw Data exist (non-destructive). */
function ensureRawTailColumns06_(rawSheet) {
  if (!rawSheet) return;
  const tail = (CONFIG && Array.isArray(CONFIG.rawDataCustomTailHeaders) && CONFIG.rawDataCustomTailHeaders.length)
    ? CONFIG.rawDataCustomTailHeaders
    : ((typeof RAW_DATA_CUSTOM_TAIL_HEADERS !== 'undefined' && Array.isArray(RAW_DATA_CUSTOM_TAIL_HEADERS) && RAW_DATA_CUSTOM_TAIL_HEADERS.length)
        ? RAW_DATA_CUSTOM_TAIL_HEADERS
        : ['Update Status','Timestamp','Status','Q-L (Months)','M-L (Months)','M-Q (Months)','Update Status Asso','Timestamp Asso','Update Status Admin','Timestamp Admin']);

  // Spec: do NOT add Associate column in Raw Data.
  const tailSafe = (Array.isArray(tail) ? tail : [])
    .map(h => String(h == null ? '' : h).trim())
    .filter(Boolean)
    .filter(h => h !== 'Associate');

  const lc = rawSheet.getLastColumn ? rawSheet.getLastColumn() : 0;
  if (lc < 1) return;
  const header = rawSheet.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v == null ? '' : v).trim());
  const missing = tailSafe.filter(h => header.indexOf(h) === -1);
  if (!missing.length) return;
  if (DRY_RUN) return;
  rawSheet.insertColumnsAfter(lc, missing.length);
  rawSheet.getRange(1, lc + 1, 1, missing.length).setValues([missing]);
}


/** =========================
 * Email ingest flow (Dashboard -> Raw Data)
 * Email ingest flow (Dashboard → Raw Data)
 * ========================= */
function buildDashboardEmailQuery_(policy) {
  // MAIN (daily 08:00) must be QUEUE-based and deterministic.
  // Gmail filter should apply label QUEUED_MAIN to matching emails.
  const p = policy || (CONFIG && CONFIG.emailIngest) || {};

  const queuedLabel = String(p.QUEUED_LABEL || p.QUEUE_LABEL || 'QUEUED_MAIN').trim();
  const from = String(p.FROM || '').trim();
  const subject = String(p.SUBJECT || '').trim();

  const parts = [];
  // Required base query (per spec)
  if (queuedLabel) parts.push('label:' + queuedLabel);
  if (from) parts.push('from:' + from);
  if (subject) parts.push('subject:"' + subject.replace(/"/g, '').replace(/"/g, '') + '"');
  parts.push('has:attachment');
  parts.push('is:unread');
  parts.push('-in:trash');
  parts.push('-in:spam');
  return parts.join(' ');
}

function pickDashboardXlsxAttachment_(message, policy) {
  if (!message) return null;
  const p = policy || (CONFIG && CONFIG.emailIngest) || {};
  const prefix = String(p.ATTACHMENT_NAME_PREFIX || '[QGP][ID] Claim Daily Monitoring').trim();

  const atts = message.getAttachments({ includeInlineImages: false, includeAttachments: true });
  let a = pickFirstAttachmentByPrefix_(atts, prefix, {});
  if (a) return a;

  // Fallback: find first XLSX-like blob
  for (let i = 0; i < atts.length; i++) {
    if (isLikelyXlsx_(atts[i])) return atts[i];
  }
  return null;
}

/**
 * Process dashboard emails (unprocessed) and ingest attachment(s) into the master workbook.
 * After success, email is cleaned up (markRead + remove QUEUED_MAIN label + trash).
 */
function runEmailIngest(maxThreads) {
  // MAIN queue consumer (1 email per run).
  // Success => markRead + removeLabel(QUEUED_MAIN) + moveToTrash
  // Error   => leave as-is in QUEUED_MAIN for automatic retry
  try { if (typeof resetRuntime_ === 'function') resetRuntime_(); } catch (err0) {}
  return withLock_(() => {
    resetRunState_();
    if (PIPELINE_FLAGS.CLEAR_LOG_BEFORE_RUN) clearLogSheet_();

    // Set flow context early so Progress notes and Log metadata are correct.
    try { setLogRunContext_('MAIN', ''); } catch (e0) {}
    try { if (typeof RUNTIME !== 'undefined' && RUNTIME) RUNTIME.flow = 'MAIN'; } catch (e1) {}


    const startedAt = new Date();
    let ssTiming = null;
    try { ssTiming = __logOverviewStart06_('Master', startedAt); } catch (e) {}

    try {
      const policy = (CONFIG && CONFIG.emailIngest) ? CONFIG.emailIngest : {};
      const query = buildDashboardEmailQuery_(policy);

      // Hard rule: 1 email per run (override any provided maxThreads)
      const limit = 1;

      setProgress_(0, 'Searching queued email...');
      logLine_('MAIL', 'MAIN ingest started', 'query=' + query, 'limit=' + limit, 'INFO');

      const threads = GmailApp.search(query, 0, limit);
      if (!threads || !threads.length) {
        setProgress_(1.0, 'No queued emails.');
        logLine_('MAIL', 'No queued emails', '', '', 'INFO');
        return { severity: 'INFO', message: 'No queued MAIN emails to process.', processed: 0, failed: 0 };
      }

      const queuedLabelName = String(policy.QUEUED_LABEL || policy.QUEUE_LABEL || 'QUEUED_MAIN').trim();
      const queuedLabel = getOrCreateGmailLabel_(queuedLabelName);

      let processed = 0;
      let failed = 0;

      // Only 1 thread by design
      const thread = threads[0];
      const msgs = thread.getMessages();

      // Pick newest unread message (deterministic) with an attachment
      let msg = null;
      for (let i = msgs.length - 1; i >= 0; i--) {
        try {
          if (msgs[i].isUnread() && msgs[i].getAttachments({ includeInlineImages: false, includeAttachments: true }).length) {
            msg = msgs[i];
            break;
          }
        } catch (e) {}
      }
      if (!msg) msg = (msgs && msgs.length) ? msgs[msgs.length - 1] : null;

      if (!msg) {
        logLine_('MAIL', 'Queued thread has no messages', '', '', 'WARN');
        return { severity: 'WARN', message: 'Queued thread has no messages.', processed: 0, failed: 1 };
      }
      try {
        if (typeof setLogEventContext_ === 'function') {
          setLogEventContext_({
            emailFrom: (msg && msg.getFrom) ? msg.getFrom() : '',
            emailSubject: (msg && msg.getSubject) ? msg.getSubject() : '',
            threadId: (thread && thread.getId) ? thread.getId() : ''
          });
        }
      } catch (eCtx0) {}

      const att = pickDashboardXlsxAttachment_(msg, policy);
      if (!att) {
        failed++;
        logLine_('MAIL', 'No XLSX attachment found (leave queued)', msg.getSubject(), '', 'ERROR');
        return { severity: 'ERROR', message: 'No XLSX attachment found.', processed: 0, failed: failed };
      }
      try {
        if (typeof setLogEventContext_ === 'function') {
          setLogEventContext_({
            attachmentName: att.getName(),
            attachmentSize: (att.getSize ? att.getSize() : ''),
            attachmentType: (att.getContentType ? att.getContentType() : '')
          });
        }
      } catch (eCtx1) {}
      // Idempotency: prevent duplicate processing of the same queued email.
      try {
        const tok = ['MAIN', thread.getId(), msg.getId(), att.getName(), att.getSize()].join('|');
        const idem = (typeof checkAndMarkTransaction_ === 'function') ? checkAndMarkTransaction_(tok, 12 * 60 * 60 * 1000) : { duplicate: false };
        if (idem && idem.duplicate) {
          try { logLine_('IDEMPOTENT', 'Duplicate MAIN token -> cleanup and skip', tok, '', 'INFO'); } catch (eI) {}
          try { msg.markRead(); } catch (eMR) {}
          try { thread.removeLabel(queuedLabel); } catch (eRL) {}
          try { thread.moveToTrash(); } catch (eTR) {}
          return { severity: 'INFO', message: 'Duplicate MAIN token (skipped).', processed: 0, failed: 0, skipped: true };
        }
      } catch (eId) {}

      let tmpFileId = null;
      try {
        // single conversion per run
        setProgress_(0.15, 'Converting XLSX...');
        const conv = convertXlsxBlobToTempSpreadsheet_(att.copyBlob(), att.getName());
        tmpFileId = conv.fileId;

        setProgress_(0.35, 'Processing pipeline...');
        const res = runPipeline_('Master', [tmpFileId], { flow: 'main', source: 'EMAIL_MAIN', subject: msg.getSubject() });
        try {
          if (typeof setLogEventContext_ === 'function' && res) {
            setLogEventContext_({
              opsUpdated: (res.routedTotal != null ? res.routedTotal : ''),
              rawRows: (res.rawRows != null ? res.rawRows : '')
            });
          }
        } catch (eCtx2) {}

        // Determine success strictly: no exception + not severity ERROR
        const sev = String((res && res.severity) ? res.severity : 'INFO').toUpperCase();
        if (sev === 'ERROR') {
          failed++;
          logLine_('MAIL', 'Pipeline returned ERROR (leave queued)', msg.getSubject(), (res && res.message) ? res.message : 'ERROR', 'ERROR');
          return { severity: 'ERROR', message: (res && res.message) ? res.message : 'Pipeline error.', processed: 0, failed: failed };
        }

        processed++;

        // Cleanup success (per spec)
        setProgress_(0.85, 'Cleaning up email...');
        try { msg.markRead(); } catch (e1) {}
        try { thread.markRead(); } catch (e2) {}
        try { if (queuedLabel) thread.removeLabel(queuedLabel); } catch (e3) {}
        try { thread.moveToTrash(); } catch (e4) {}

        logLine_('MAIL', 'MAIN ingest success', msg.getSubject(), 'cleanup=markRead+removeLabel+trash', 'INFO');
        setProgress_(1.0, 'Done.');
        return { severity: 'INFO', message: 'Processed 1 queued MAIN email.', processed: processed, failed: failed };

      } catch (err) {
        failed++;
        // DO NOT cleanup on error: keep queued for retry
        logLine_('MAIL', 'MAIN ingest failed (leave queued)', msg.getSubject(), String(err), 'ERROR');
        return { severity: 'ERROR', message: String(err), processed: processed, failed: failed };
      } finally {
        try { if (tmpFileId) trashDriveFileById_(tmpFileId); } catch (e) {}
      }

    } finally {
      try { __logOverviewDuration06_('Master', startedAt, ssTiming); } catch (e2) {}
    }
  });
}


/** =========================
 * SUB email ingest (QUEUED_SUB, hourly except 08:00)
 * ========================= */

/**
 * Install hourly trigger for SUB consumer (runs every hour; handler skips at 08:00).
 * Note: time-based triggers cannot express "every hour except 08:00", so the skip is done in code.
 */
function installSubEmailIngestTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'runSubEmailIngest') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Hourly execution; actual skip at 08:00 is enforced inside runSubEmailIngest().
  ScriptApp.newTrigger('runSubEmailIngest')
    .timeBased()
    .everyHours(1)
    .nearMinute(20)
    .create();
}

/**
 * SUB queue consumer (hourly, skip at 08:00).
 * - Reads 1 email thread from label:QUEUED_SUB with subject "Claim Monitoring Operational Dashboard"
 * - Expects 2 XLSX attachments: OLD + NEW
 * - Process order: OLD first, then NEW
 * - For each: convert XLSX -> temp Google Sheet -> copy all data into Raw OLD/Raw NEW
 * - After copy: update operational sheets (excluding EV-Bike & Exclusion) by Claim Number
 * - Extra: append to Submission for:
 *   - OLD: last_status == SUBMITTED
 *   - NEW: last_status == CLAIM_INITIATE
 * - After both succeed: cleanup email (markRead + remove label + trash)
 * - Anti-bentrok: uses tryLock; if lock busy, SKIP without error
 */
function runSubEmailIngest(maxThreads) {
  try { if (typeof resetRuntime_ === 'function') resetRuntime_(); } catch (err0) {}
  try { if (typeof runtimePreflight06f_ === 'function') runtimePreflight06f_('SUB_PIPELINE'); } catch (ePf) {}

  // Skip 08:00 hour to avoid collision with MAIN daily ingest.
  try {
    const tz = (typeof getTzSafe_ === 'function') ? getTzSafe_() : (Session.getScriptTimeZone() || 'Asia/Jakarta');
    const hr = Number(Utilities.formatDate(new Date(), tz, 'H'));
    if (hr === 8) {
      try { resetRunState_(); logLine_('SUB_SKIP', 'Skip SUB at 08:00 to avoid MAIN', '', '', 'INFO'); } catch (e1) {}
      return { severity: 'INFO', message: 'Skip SUB at 08:00', processed: 0, failed: 0, skipped: true };
    }
  } catch (e0) {}

  return __withTryLockSub06a_(() => {
    resetRunState_();

    // Set flow context early so Progress notes and Log metadata are correct.
    try { setLogRunContext_('SUB', ''); } catch (e0) {}
    try { if (typeof RUNTIME !== 'undefined' && RUNTIME) RUNTIME.flow = 'SUB'; } catch (e1) {}


    // Intentionally do NOT clear the log by default for hourly runs (unless explicitly enabled).
    if (PIPELINE_FLAGS.CLEAR_LOG_BEFORE_RUN) clearLogSheet_();

    const startedAt = new Date();
    try { logLine_('SUB_START', 'Start SUB email ingest', '', startedAt.toISOString(), 'INFO'); } catch (e2) {}

    const policy = (CONFIG && CONFIG.subEmailIngest) ? CONFIG.subEmailIngest : {};
    const pMerged = Object.assign(
      { QUEUED_LABEL: 'QUEUED_SUB', SUBJECT: 'Claim Monitoring Operational Dashboard' },
      policy || {}
    );

    const query = buildDashboardEmailQuery_(pMerged);
    const limit = Math.max(1, Math.min(5, Number(maxThreads || 1) || 1));
    try { setProgressForFlow_('SUB', 0.05, 'Searching queued email...', { prefixFlowInStep: true }); } catch (eP0) {}

    const threads = GmailApp.search(query, 0, limit);
    if (!threads || !threads.length) {
      try { setProgressForFlow_('SUB', 1.0, 'No queued emails.', { prefixFlowInStep: true }); } catch (eP1) {}
      try { logLine_('SUB', 'No queued SUB emails', query, '', 'INFO'); } catch (e3) {}
      return { severity: 'INFO', message: 'No queued SUB emails', processed: 0, failed: 0 };
    }

    // Deterministic: process first thread only (per run).
    const thread = threads[0];
    const messages = thread.getMessages();

    // Pick newest unread with attachments (fallback to last).
    let msg = null;
    for (let i = messages.length - 1; i >= 0; i--) {
      try {
        if (messages[i].isUnread()) {
          const atts = messages[i].getAttachments({ includeInlineImages: false, includeAttachments: true });
          if (atts && atts.length) { msg = messages[i]; break; }
        }
      } catch (e4) {}
    }
    if (!msg) msg = messages[messages.length - 1];

    const subject = String((msg && msg.getSubject) ? msg.getSubject() : '');
    try {
      if (typeof setLogEventContext_ === 'function') {
        setLogEventContext_({
          emailFrom: (msg && msg.getFrom) ? msg.getFrom() : '',
          emailSubject: subject,
          threadId: (thread && thread.getId) ? thread.getId() : ''
        });
      }
    } catch (eCtx3) {}
    try { logLine_('SUB_EMAIL', 'Picked SUB email', subject, '', 'INFO'); } catch (e5) {}

    const atts = msg.getAttachments({ includeInlineImages: false, includeAttachments: true });
    const picked = __pickSubOldNewAttachments06a_(atts);
    try {
      if (typeof setLogEventContext_ === 'function') {
        const names = [picked && picked.oldAtt ? picked.oldAtt.getName() : '', picked && picked.newAtt ? picked.newAtt.getName() : ''].filter(Boolean).join(' | ');
        const sizeOld = (picked && picked.oldAtt && picked.oldAtt.getSize) ? Number(picked.oldAtt.getSize() || 0) : 0;
        const sizeNew = (picked && picked.newAtt && picked.newAtt.getSize) ? Number(picked.newAtt.getSize() || 0) : 0;
        setLogEventContext_({
          attachmentName: names,
          attachmentSize: (sizeOld + sizeNew) || '',
          attachmentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
      }
    } catch (eCtx4) {}

    if (!picked.oldAtt || !picked.newAtt) {
      const names = (atts || []).map(a => (a && a.getName) ? a.getName() : '').join(' | ');
      try { setProgressForFlow_('SUB', 1.0, 'Failed (OLD/NEW attachment missing)', { prefixFlowInStep: true }); } catch (eP2) {}
      try { logLine_('SUB_ERR', 'Missing OLD/NEW attachments (expected 2 XLSX)', names, '', 'ERROR'); } catch (e6) {}
      // Leave queued for retry.
      return { severity: 'ERROR', message: 'Missing OLD/NEW attachments', processed: 0, failed: 1 };
    }

    // Open master workbook (same as MAIN).
    const masterKey = 'Master';
    validateConfigForPic_(masterKey);
    const masterSs = __tryOpenSpreadsheetForKey06_(masterKey);
    // Best-effort normalize master spreadsheet timezone to avoid date display shifts.
    try { if (masterSs && typeof getTzSafe_ === 'function') masterSs.setSpreadsheetTimeZone(getTzSafe_()); } catch (eTzM) {}

    // Resolve SUB flow spec (sheet names, op list, sort spec) from CONFIG when available.
    const subFlow = (typeof CONFIG === 'object' && CONFIG && CONFIG.subFlow) ? CONFIG.subFlow : {};

    // Ensure raw target sheets exist.
    try { setProgressForFlow_('SUB', 0.15, 'Preparing workbook...', { prefixFlowInStep: true }); } catch (eP3) {}
    const rawOldName = String(subFlow.RAW_OLD_SHEET_NAME || pMerged.RAW_OLD_SHEET || 'Raw OLD').trim();
    const rawNewName = String(subFlow.RAW_NEW_SHEET_NAME || pMerged.RAW_NEW_SHEET || 'Raw NEW').trim();
    __ensureSheetByNameSub06a_(masterSs, rawOldName);
    __ensureSheetByNameSub06a_(masterSs, rawNewName);

    // Operational sheets to update (explicit allow-list).
    const defaultOpSheets = [
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
    ];
    const opSheets = (Array.isArray(subFlow.OPERATIONAL_SHEETS) && subFlow.OPERATIONAL_SHEETS.length)
      ? subFlow.OPERATIONAL_SHEETS.map(s => String(s || '').trim()).filter(Boolean)
      : (Array.isArray(pMerged.OPERATIONAL_SHEETS) && pMerged.OPERATIONAL_SHEETS.length
        ? pMerged.OPERATIONAL_SHEETS.map(s => String(s || '').trim()).filter(Boolean)
        : defaultOpSheets);

    // Align required operational column layout for SUB updates too.
    try { if (typeof enforceOperationalLayout06_ === 'function') enforceOperationalLayout06_(masterSs); } catch (eLay) {}

    // Ensure SC fallback/quarantine sheet participates in relocate + sort.
    try {
      const fb = __getScFallbackSheet06a_();
      if (fb && opSheets.indexOf(fb) < 0) opSheets.push(fb);
      __ensureSubScFallbackSheetExists06a_(masterSs, fb);
    } catch (eFb) {}


    // Sorting spec (multi-key) after SUB completes.
    const sortSpecs = (Array.isArray(subFlow.SORT_SPECS) && subFlow.SORT_SPECS.length) ? subFlow.SORT_SPECS : null;
    const queuedLabel = getOrCreateGmailLabel_(String(pMerged.QUEUED_LABEL || pMerged.QUEUE_LABEL || 'QUEUED_SUB').trim());

    // Idempotency: prevent duplicate processing of the same queued SUB email.
    try {
      const tok = ['SUB', thread.getId(), msg.getId(), picked.oldAtt.getName(), picked.oldAtt.getSize(), picked.newAtt.getName(), picked.newAtt.getSize()].join('|');
      const idem = (typeof checkAndMarkTransaction_ === 'function') ? checkAndMarkTransaction_(tok, 6 * 60 * 60 * 1000) : { duplicate: false };
      if (idem && idem.duplicate) {
        try { logLine_('IDEMPOTENT', 'Duplicate SUB token -> cleanup and skip', tok, '', 'INFO'); } catch (eI2) {}
        try { msg.markRead(); } catch (eMR2) {}
        try { thread.removeLabel(queuedLabel); } catch (eRL2) {}
        try { thread.moveToTrash(); } catch (eTR2) {}
        return { severity: 'INFO', message: 'Duplicate SUB token (skipped).', processed: 0, failed: 0, skipped: true };
      }
    } catch (eId2) {}




// WebApp Project (Movement Claim Tracking): take PREV snapshots BEFORE SUB overwrites Raw.
try { setProgressForFlow_('SUB', 0.25, 'Snapshot PREV...', { prefixFlowInStep: true }); } catch (eP4) {}
try {
  if (typeof webappMovementSnapshotPrevForSub06c_ === 'function') {
    const snapPrev = webappMovementSnapshotPrevForSub06c_(masterSs, rawOldName, rawNewName);
    try { logLine_('WEBAPP_SNAP_PREV', 'Snapshot PREV (Raw->WebApp)', JSON.stringify(snapPrev || {}), '', 'INFO'); } catch (eWp1) {}
  }
} catch (eWp0) {
  try { logLine_('WEBAPP_SNAP_PREV_ERR', 'Snapshot PREV failed (non-fatal)', String(eWp0), '', 'WARN'); } catch (eWp2) {}
}

    // Process OLD then NEW; cleanup email only if both succeed.
    try { setProgressForFlow_('SUB', 0.40, 'Process OLD...', { prefixFlowInStep: true }); } catch (eP5) {}

    const rOld = __processSubAttachment06a_(masterSs, picked.oldAtt, {
      dbTag: 'OLD',
      rawSheetName: rawOldName,
      operationalSheetNames: opSheets
    });
    if (!rOld || String(rOld.severity || '').toUpperCase() === 'ERROR') {
      try { setProgressForFlow_('SUB', 1.0, 'Failed (OLD)', { prefixFlowInStep: true }); } catch (eP6) {}
      try { logLine_('SUB_FAIL', 'OLD processing failed; leave email queued', '', '', 'ERROR'); } catch (e7) {}
      return { severity: 'ERROR', message: 'SUB OLD failed', processed: 0, failed: 1, details: rOld };
    }

    try { setProgressForFlow_('SUB', 0.62, 'Process NEW...', { prefixFlowInStep: true }); } catch (eP7) {}
    const rNew = __processSubAttachment06a_(masterSs, picked.newAtt, {
      dbTag: 'NEW',
      rawSheetName: rawNewName,
      operationalSheetNames: opSheets
    });
    if (!rNew || String(rNew.severity || '').toUpperCase() === 'ERROR') {
      try { setProgressForFlow_('SUB', 1.0, 'Failed (NEW)', { prefixFlowInStep: true }); } catch (eP8) {}
      try { logLine_('SUB_FAIL', 'NEW processing failed; leave email queued', '', '', 'ERROR'); } catch (e8) {}
      return { severity: 'ERROR', message: 'SUB NEW failed', processed: 0, failed: 1, details: rNew };
    }
    // Relocate rows by Last Status mapping (move FULL row, dedupe by Claim Number).
    try { setProgressForFlow_('SUB', 0.80, 'Relocate + sort...', { prefixFlowInStep: true }); } catch (eP9) {}
    const relocateSheets = __getSubRelocationSheetNames06a_(opSheets);
    const relocateRes = __relocateOperationalRowsByLastStatusSub06a_(masterSs, relocateSheets);
    try {
      if (typeof setLogEventContext_ === 'function') {
        const moved = (relocateRes && relocateRes.moved != null) ? relocateRes.moved : '';
        setLogEventContext_({ opsUpdated: moved, subInserted: moved });
      }
    } catch (eCtx5) {}
    try { logLine_('SUB_MOVE', 'Relocated rows after SUB updates', JSON.stringify(relocateRes || {}), '', 'INFO'); } catch (e9a) {}

    // Final: sort operational sheets (Submission Date -> Last Status Date -> Last Status) preserving filters.
    const sortRes = __sortOperationalSheetsSub06a_(masterSs, opSheets, sortSpecs);
    try { logLine_('SUB_SORT', 'Sorted operational sheets', JSON.stringify(sortRes || {}), '', 'INFO'); } catch (e9) {}

    // Refresh Overview Claim -> Report Base snapshot after SUB updates.
    try {
      if (typeof refreshReportBaseFromOperational06_ === 'function') refreshReportBaseFromOperational06_(masterSs, { incremental: true });
    } catch (eRb) { try { logLine_('SUB_WARN', 'Report Base refresh failed', String(eRb), '', 'WARN'); } catch (eRb2) {} }



// WebApp Project (Movement Claim Tracking): take CURR snapshots AFTER SUB, then emit Daily events (dedup by Event ID).
try { setProgressForFlow_('SUB', 0.92, 'Snapshot CURR + movement...', { prefixFlowInStep: true }); } catch (eP10) {}
try {
  if (typeof webappMovementSnapshotCurrAndTrackForSub06c_ === 'function') {
    const snapCurr = webappMovementSnapshotCurrAndTrackForSub06c_(masterSs, rawOldName, rawNewName);
    try { logLine_('WEBAPP_MOVE', 'Movement tracking emitted to Daily', JSON.stringify(snapCurr || {}), '', 'INFO'); } catch (eWp3) {}
  }
} catch (eWp4) {
  try { logLine_('WEBAPP_MOVE_ERR', 'Movement tracking failed (non-fatal)', String(eWp4), '', 'WARN'); } catch (eWp5) {}
}

    // Cleanup email thread after both succeeded.
    try {
      cleanupQueuedThreadSuccess_(thread, queuedLabel);
      logLine_('SUB_CLEAN', 'Cleaned SUB email thread', queuedLabel, '', 'INFO');
    } catch (e10) {}

    const durMs = new Date().getTime() - startedAt.getTime();
    try { logLine_('SUB_DONE', 'Completed SUB ingest', __formatProcessingDuration06_(durMs), '', 'INFO'); } catch (e11) {}
    try { setProgressForFlow_('SUB', 1.0, 'Done.', { prefixFlowInStep: true }); } catch (eP11) {}

    return {
      severity: 'INFO',
      message: 'SUB ingest completed',
      processed: 1,
      failed: 0,
      old: rOld,
      new: rNew,
      sorted: sortRes
    };
  });
}

/** Try-lock wrapper for SUB (skip when busy). */
function __withTryLockSub06a_(fn) {
  // Prefer shared utility if available (keeps behavior consistent across modules).
  if (typeof withTryScriptLock_ === 'function') {
    const lr = withTryScriptLock_(750, fn);
    if (!lr || !lr.acquired) {
      try {
        resetRunState_();
        logLine_('SUB_SKIP', 'Skip SUB: lock busy (likely MAIN running)', '', '', 'INFO');
      } catch (e1) {}
      return { severity: 'INFO', message: 'Skip SUB (lock busy)', processed: 0, failed: 0, skipped: true };
    }
    return lr.result;
  }

  // Fallback local try-lock.
  const lock = LockService.getScriptLock();
  const ok = lock.tryLock(750); // short try-lock; SUB is hourly and must be non-blocking
  if (!ok) {
    try {
      resetRunState_();
      logLine_('SUB_SKIP', 'Skip SUB: lock busy (likely MAIN running)', '', '', 'INFO');
    } catch (e1) {}
    return { severity: 'INFO', message: 'Skip SUB (lock busy)', processed: 0, failed: 0, skipped: true };
  }
  try {
    return (typeof fn === 'function') ? fn() : null;
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Pick OLD + NEW attachments from the Metabase subscription email. */
function __pickSubOldNewAttachments06a_(attachments) {
  // Prefer shared picker from 04.gs if present (keeps rules in one place).
  if (typeof pickSubDashboardAttachments04_ === 'function') {
    try {
      const r = pickSubDashboardAttachments04_(attachments);
      if (r && (r.oldAttachment || r.newAttachment || r.oldAtt || r.newAtt)) {
        return {
          oldAtt: r.oldAttachment || r.oldAtt || null,
          newAtt: r.newAttachment || r.newAtt || null
        };
      }
    } catch (e0) {}
  }

  const list = Array.isArray(attachments) ? attachments : [];
  let oldAtt = null;
  let newAtt = null;

  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    const name = (a && typeof a.getName === 'function') ? String(a.getName() || '') : '';
    if (!name) continue;

    const isStd = name.indexOf('(Standardization)') > -1;
    const isAging = name.indexOf('List of Claims with Aging') > -1;

    if (isAging && isStd) {
      newAtt = a;
      continue;
    }
    if (isAging && !isStd) {
      oldAtt = a;
      continue;
    }
  }

  return { oldAtt: oldAtt, newAtt: newAtt };
}

/** Ensure a sheet exists by name (create if missing). */
function __ensureSheetByNameSub06a_(ss, name) {
  if (!ss) throw new Error('Spreadsheet is required');
  const n = String(name || '').trim();
  if (!n) throw new Error('Sheet name is empty');

  let sh = ss.getSheetByName(n);
  if (!sh) {
    sh = ss.insertSheet(n);
    try { logLine_('SUB_SHEET', 'Created sheet', n, '', 'INFO'); } catch (e1) {}
  }
  return sh;
}

/**
 * Process one attachment (OLD or NEW):
 * 1) Convert XLSX -> temp spreadsheet
 * 2) Copy values to Raw sheet
 * 3) Update operational sheets (4 fields) by Claim Number
 * 4) Append missing rows to Submission according to rules
 */
function __processSubAttachment06a_(masterSs, attachmentBlob, opts) {
  const o = opts || {};
  const dbTag = String(o.dbTag || '').trim().toUpperCase();
  const rawSheetName = String(o.rawSheetName || '').trim();
  const operationalSheetNames = Array.isArray(o.operationalSheetNames) ? o.operationalSheetNames : [];

  if (!masterSs) throw new Error('Master spreadsheet is required');
  if (!attachmentBlob) throw new Error('Attachment blob is required');
  if (!dbTag) throw new Error('dbTag is required');
  if (!rawSheetName) throw new Error('rawSheetName is required');

  const startedAt = new Date();
  let tmpFileId = null;

  try {
    try { logLine_('SUB_' + dbTag, 'Convert XLSX -> temp spreadsheet', rawSheetName, '', 'INFO'); } catch (e0) {}

    const conv = convertXlsxBlobToTempSpreadsheet_(attachmentBlob.copyBlob(), 'TMP_SUB_' + dbTag + '_' + nowStr_('yyyyMMdd_HHmmss'));
    tmpFileId = conv && conv.fileId;
    const tmpSs = conv && conv.ss;
    if (!tmpSs) throw new Error('Failed to open converted temp spreadsheet');

    const tmpSh = tmpSs.getSheets()[0];
    if (!tmpSh) throw new Error('Temp spreadsheet has no sheets');

    // Avoid getDataRange() bloat (XLSX conversions often carry formatting far down).
    const lr = Math.max(1, tmpSh.getLastRow());
    const lc = Math.max(1, tmpSh.getLastColumn());
    const rng = tmpSh.getRange(1, 1, lr, lc);
    const values = rng.getValues();
    const displayValues = rng.getDisplayValues();
    if (!values || !values.length) throw new Error('Temp spreadsheet is empty');
    // Normalize datetime cells using *display* strings as source of truth (prevents timezone drift).
    try { __coerceSubDatetimeColumnsFromDisplay06a_(values, displayValues); } catch (eC0) {}

    // Copy all data to Raw sheet.
    const rawSh = __ensureSheetByNameSub06a_(masterSs, rawSheetName);
    __overwriteSheetValuesSub06a_(rawSh, values);
    try { __applyRawOldNewDatetimeFormatsSub06a_(rawSh, values[0]); } catch (eFmt0) {}

    try { logLine_('SUB_' + dbTag, 'Copied to raw sheet', rawSheetName, 'rows=' + (values.length - 1), 'INFO'); } catch (e1) {}

    // Build raw map for fast lookup.
    const rawIndex = __buildSubRawIndex06a_(values);
    const map = rawIndex.map;
    const rawHeader = rawIndex.headerIndex;

    // Update operational sheets.
    const upd = __updateOperationalSheetsFromRaw06a_(masterSs, operationalSheetNames, map, {
      claimKeyRaw: rawHeader.claim_number,
      fieldsRaw: rawHeader,
      dbTag: dbTag
    });

    // Extra: append to Submission after updates for this DB.
    const appended = __appendSubmissionFromRawIfMissing06a_(masterSs, map, rawHeader, dbTag);

    const durMs = new Date().getTime() - startedAt.getTime();
    try { logLine_('SUB_' + dbTag, 'Done attachment processing', __formatProcessingDuration06_(durMs), '', 'INFO'); } catch (e2) {}

    return {
      severity: 'INFO',
      db: dbTag,
      rawSheet: rawSheetName,
      rows: values.length - 1,
      updated: upd,
      appendedToSubmission: appended
    };
  } catch (err) {
    try { logLine_('SUB_' + dbTag + '_ERR', 'Attachment processing error', String(err && err.message ? err.message : err), String(err), 'ERROR'); } catch (e3) {}
    return { severity: 'ERROR', db: dbTag, error: String(err && err.message ? err.message : err) };
  } finally {
    try { if (tmpFileId) trashDriveFileById_(tmpFileId); } catch (e4) {}
  }
}

/** Overwrite a sheet with the provided 2D values (clears previous contents). */
function __overwriteSheetValuesSub06a_(sheet, values) {
  if (!sheet) throw new Error('Target sheet is required');
  const v = Array.isArray(values) ? values : [];
  if (!v.length || !Array.isArray(v[0]) || !v[0].length) throw new Error('Values must be a non-empty 2D array');

  const rows = v.length;
  const cols = v[0].length;

  // Clear existing content
  try { sheet.clearContents(); } catch (e1) {}

  const rng = sheet.getRange(1, 1, rows, cols);
  if (typeof safeSetValues_ === 'function') safeSetValues_(rng, v);
  else rng.setValues(v);
}

function __applyRawOldNewDatetimeFormatsSub06a_(rawSh, headerRow) {
  if (!rawSh || !headerRow) return;
  const lr = rawSh.getLastRow();
  if (lr < 2) return;
  const hdr = headerRow.map(h => String(h || '').trim().toLowerCase());
  const n = lr - 1;
  const dtFmt = (typeof FORMATS !== 'undefined' && FORMATS && (FORMATS.DATETIME_LONG || FORMATS.DATETIME))
    ? (FORMATS.DATETIME_LONG || FORMATS.DATETIME)
    : 'MMMM d, yyyy, h:mm AM/PM';

  function idxOfAny(keys) {
    for (let i = 0; i < keys.length; i++) {
      const k = String(keys[i] || '').toLowerCase();
      const j = hdr.indexOf(k);
      if (j >= 0) return j;
    }
    return -1;
  }

  const idxs = [];
  idxs.push(idxOfAny(['claim_last_updated_datetime', 'claim last updated datetime', 'last_status_date', 'last status date']));
  idxs.push(idxOfAny(['activity_log_datetime', 'activity log datetime', 'last_activity_log_datetime', 'last activity log datetime', 'activity_log_timestamp', 'activity log timestamp']));
  idxs.push(idxOfAny(['claim_submitted_datetime', 'claim submitted datetime', 'submitted_datetime', 'submitted datetime']));

  const uniq = Array.from(new Set(idxs.filter(i => i >= 0)));
  for (let k = 0; k < uniq.length; k++) {
    const col = uniq[k] + 1;
    try { rawSh.getRange(2, col, n, 1).setNumberFormat(dtFmt); } catch (e1) {}
  }
}

/**
 * Coerce datetime columns using display strings (source-of-truth) to prevent timezone drift.
 * This runs BEFORE writing to Raw OLD/NEW.
 */
function __coerceSubDatetimeColumnsFromDisplay06a_(values, displayValues) {
  if (!values || !displayValues || !values.length) return;
  if (!Array.isArray(values[0]) || !Array.isArray(displayValues[0])) return;
  if (values.length !== displayValues.length) return;

  const tz = (typeof getTzSafe_ === 'function') ? getTzSafe_() : ((Session && Session.getScriptTimeZone) ? (Session.getScriptTimeZone() || 'Asia/Jakarta') : 'Asia/Jakarta');
  const hdr = values[0].map(h => String(h || '').trim().toLowerCase());
  function idxOfAny(keys) {
    for (let i = 0; i < keys.length; i++) {
      const k = String(keys[i] || '').toLowerCase();
      const j = hdr.indexOf(k);
      if (j >= 0) return j;
    }
    return -1;
  }

  // Datetime columns we care about (SUB + WebApp movement tracking).
  const idxs = [];
  idxs.push(idxOfAny(['claim_last_updated_datetime', 'claim last updated datetime']));
  idxs.push(idxOfAny(['activity_log_timestamp', 'activity log timestamp', 'activity_log_datetime', 'activity log datetime', 'last_activity_log_datetime', 'last activity log datetime']));
  idxs.push(idxOfAny(['claim_submitted_datetime', 'claim submitted datetime', 'submitted_datetime', 'submitted datetime']));

  const dtCols = Array.from(new Set(idxs.filter(i => i >= 0)));
  if (!dtCols.length) return;

  // Coerce cell-by-cell from row 2..N
  for (let r = 1; r < values.length; r++) {
    const rowDisp = displayValues[r];
    const rowVal = values[r];
    if (!rowDisp || !rowVal) continue;
    for (let k = 0; k < dtCols.length; k++) {
      const c = dtCols[k];
      const s0 = (rowDisp[c] == null) ? '' : String(rowDisp[c]);
      const s = s0.replace(/\u00A0/g, ' ').trim();
      if (!s) continue;
      const d = __parseStrictDatetimeFromDisplaySub06a_(s, tz);
      if (d) rowVal[c] = d;
    }
  }
}

function __parseStrictDatetimeFromDisplaySub06a_(s, tz) {
  const t0 = String(s || '').replace(/\u00A0/g, ' ').trim();
  if (!t0) return null;

  // Normalize common noise.
  let t = t0;
  t = t.replace(/\s+(WIB|WITA|WIT)\b/i, '');
  t = t.replace(/\s+GMT[+-]\d{1,2}(:?\d{2})?\b/i, '');
  t = t.replace(/\bat\b/ig, '');
  t = t.replace(/\s+/g, ' ').trim();

  const patterns = [
    // English long
    'MMMM d, yyyy, h:mm a',
    'MMMM d, yyyy, hh:mm a',
    'MMMM d, yyyy, h:mm:ss a',
    'MMMM d, yyyy, hh:mm:ss a',
    'MMM d, yyyy, h:mm a',
    'MMM d, yyyy, hh:mm a',
    'MMM d, yyyy, h:mm:ss a',
    // ISO-ish
    'yyyy-MM-dd HH:mm:ss',
    'yyyy-MM-dd HH:mm',
    'yyyy/MM/dd HH:mm:ss',
    'yyyy/MM/dd HH:mm',
    // Slash formats (common from XLSX conversion)
    'M/d/yyyy h:mm a',
    'M/d/yyyy hh:mm a',
    'M/d/yyyy H:mm',
    'M/d/yyyy HH:mm',
    'd/M/yyyy H:mm',
    'd/M/yyyy HH:mm',
    'dd/MM/yyyy HH:mm:ss',
    'dd/MM/yyyy HH:mm',
    // Short month name variants
    'd MMM yyyy, HH:mm',
    'd MMM yy, HH:mm',
    'd MMMM yyyy, HH:mm'
  ];

  for (let i = 0; i < patterns.length; i++) {
    try {
      const d = Utilities.parseDate(t, tz, patterns[i]);
      if (d && !isNaN(d.getTime())) return d;
    } catch (e) {}
  }

  // As a last attempt, reuse SUB parser (but avoid Date(string)).
  try {
    if (typeof __parseClaimLastUpdatedDatetimeSub06a_ === 'function') {
      const d2 = __parseClaimLastUpdatedDatetimeSub06a_(t);
      if (d2 && !isNaN(d2.getTime())) return d2;
    }
  } catch (e2) {}

  return null;
}


/** Build a map claim_number -> record from raw values. */

function __normalizeClaimKeySub06a_(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function __buildSubRawIndex06a_(values) {
  const v = Array.isArray(values) ? values : [];
  if (!v.length) return { map: new Map(), headerIndex: {} };

  const header = v[0].map(h => String(h || '').trim());
  const norm = header.map(h => h.toLowerCase());

  function idxOfAny(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const c = String(candidates[i] || '').toLowerCase();
      const j = norm.indexOf(c);
      if (j >= 0) return j;
    }
    return -1;
  }

  const idxClaim = idxOfAny(['claim_number', 'claim number', 'claim no', 'claim_no']);

  const idxLSA = idxOfAny(['last_status_aging', 'last status aging', 'lsa']);
  const idxALA = idxOfAny(['activity_log_aging', 'activity log aging', 'ala']);
  const idxLastStatus = idxOfAny(['last_status', 'last status']);
  const idxSc = idxOfAny(['sc_name', 'service center', 'service_center']);

  // NEW: fields requested for SUB enrichment
  const idxClaimLastUpd = idxOfAny(['claim_last_updated_datetime', 'claim last updated datetime', 'last_status_date', 'last status date']);
  const idxActLog = idxOfAny(['activity_log', 'activity log']);
  const idxLastActLog = idxOfAny(['last_activity_log', 'last activity log']);
  const idxActLogDt = idxOfAny([
    'activity_log_datetime', 'activity log datetime', 'activity_log_date_time', 'activity log date time',
    'last_activity_log_datetime', 'last activity log datetime'
  ]);

  // Keep sampling priority: prefer claim_submitted_datetime over claim_submission_date.
  const idxSubmitted = idxOfAny([
    'claim_submitted_datetime',
    'claim submitted datetime',
    'claim_submission_date',
    'claim submission date',
    'submission_date',
    'submission date'
  ]);
  const idxLink = idxOfAny(['dashboard_link', 'db_link', 'dashboard link', 'db link', 'link']);

  const idxPartnerName = idxOfAny(['partner_name', 'partner name', 'partner_code', 'partner code']);
  const idxInsurance = idxOfAny(['insurance_partner_code', 'insurance', 'insurance_code', 'insurance partner code']);
  const idxDeviceType = idxOfAny(['device_type', 'device type']);
  const idxImei = idxOfAny(['device_imei', 'imei/sn', 'imei', 'sn', 'imei/sn']);
  const idxStoreName = idxOfAny(['3. all transaction - qoala_policy_number → outlet_name', 'outlet_name', 'outlet name', 'store_name', 'store name']);
  const idxPaName = idxOfAny(['3. all transaction - qoala_policy_number → pa_name', 'pa_name', 'pa name']);
  const idxSpaName = idxOfAny(['3. all transaction - qoala_policy_number → spa_name', 'spa_name', 'spa name']);

  const map = new Map();
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    if (!row || idxClaim < 0) continue;

    const cn = __normalizeClaimKeySub06a_(row[idxClaim]);
    if (!cn) continue;

    // For SUB (Raw OLD/NEW), Activity Log source should be `activity_log` (fallback `last_activity_log`)
    const act = (idxActLog >= 0 && row[idxActLog] !== '' && row[idxActLog] != null) ? row[idxActLog]
      : (idxLastActLog >= 0 ? row[idxLastActLog] : '');

    map.set(cn, {
      claim_number: cn,
      last_status_aging: idxLSA >= 0 ? row[idxLSA] : '',
      activity_log_aging: idxALA >= 0 ? row[idxALA] : '',
      last_status: idxLastStatus >= 0 ? row[idxLastStatus] : '',
      sc_name: idxSc >= 0 ? row[idxSc] : '',
      claim_last_updated_datetime: idxClaimLastUpd >= 0 ? row[idxClaimLastUpd] : '',
      activity_log: act,
      activity_log_datetime: idxActLogDt >= 0 ? row[idxActLogDt] : '',
      claim_submitted_datetime: idxSubmitted >= 0 ? row[idxSubmitted] : '',
      dashboard_link: idxLink >= 0 ? row[idxLink] : '',
      partner_name: idxPartnerName >= 0 ? row[idxPartnerName] : '',
      insurance: idxInsurance >= 0 ? row[idxInsurance] : '',
      device_type: idxDeviceType >= 0 ? row[idxDeviceType] : '',
      device_imei: idxImei >= 0 ? row[idxImei] : '',
      store_name: idxStoreName >= 0 ? row[idxStoreName] : '',
      pa_name: idxPaName >= 0 ? row[idxPaName] : '',
      spa_name: idxSpaName >= 0 ? row[idxSpaName] : ''
    });
  }

  return {
    map: map,
    headerIndex: {
      claim_number: idxClaim,
      last_status_aging: idxLSA,
      activity_log_aging: idxALA,
      last_status: idxLastStatus,
      sc_name: idxSc,
      claim_last_updated_datetime: idxClaimLastUpd,
      activity_log: idxActLog,
      last_activity_log: idxLastActLog,
      activity_log_datetime: idxActLogDt,
      claim_submitted_datetime: idxSubmitted,
      dashboard_link: idxLink,
      partner_name: idxPartnerName,
      insurance: idxInsurance,
      device_type: idxDeviceType,
      device_imei: idxImei,
      store_name: idxStoreName,
      pa_name: idxPaName,
      spa_name: idxSpaName
    }
  };
}

/** Update operational sheets (4 fields) from raw map by Claim Number. */
/** Update operational sheets (SUB allowed fields only) from raw map by Claim Number. */
function __updateOperationalSheetsFromRaw06a_(ss, sheetNames, rawMap, ctx) {
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  const map = (rawMap && typeof rawMap.get === 'function') ? rawMap : new Map();
  const dbTag = String((ctx && ctx.dbTag) ? ctx.dbTag : '').trim().toUpperCase();
  // Policy lookup (single source of truth: 00.gs)
  const opsPolicy = (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY) ? OPS_ROUTING_POLICY : null;
  const sheetsPolicy = (opsPolicy && opsPolicy.SHEETS) ? opsPolicy.SHEETS : {};
  const scFarhanName = String(sheetsPolicy.SC_FARHAN || 'SC - Farhan');
  const scMeilaniName = String(sheetsPolicy.SC_MEILANI || 'SC - Meilani');
  const scIvanName = String(sheetsPolicy.SC_IVAN || sheetsPolicy.SC_IVAN_NAME || 'SC - Meindar');

  // Type mapping (SC sheets) by Last Status.
  const typePolicy = (opsPolicy && opsPolicy.TYPE_BY_LAST_STATUS) ? opsPolicy.TYPE_BY_LAST_STATUS : null;
  const typeSets = typePolicy ? {
    onRep: new Set(typePolicy['SC - On Rep'] || []),
    waitRep: new Set(typePolicy['SC - Wait Rep'] || []),
    finish: new Set(typePolicy['Finish'] || []),
    orSet: new Set(typePolicy['OR'] || []),
    insurance: new Set(typePolicy['Insurance'] || []),
    est: new Set(typePolicy['SC - Est'] || []),
    rcvd: new Set(typePolicy['SC - Rcvd'] || [])
  } : null;

  // Precedence (specific > general): On Rep / Wait Rep override Insurance.
  const resolveScType = (statusVal) => {
    const s = String(statusVal || '').trim();
    if (!s || !typeSets) return '';
    if (typeSets.onRep && typeSets.onRep.has(s)) return 'SC - On Rep';
    if (typeSets.waitRep && typeSets.waitRep.has(s)) return 'SC - Wait Rep';
    if (typeSets.finish && typeSets.finish.has(s)) return 'Finish';
    if (typeSets.orSet && typeSets.orSet.has(s)) return 'OR';
    if (typeSets.insurance && typeSets.insurance.has(s)) return 'Insurance';
    if (typeSets.est && typeSets.est.has(s)) return 'SC - Est';
    if (typeSets.rcvd && typeSets.rcvd.has(s)) return 'SC - Rcvd';
    return '';
  };


  const summary = { sheets: {}, totalUpdatedRows: 0, totalSheetsTouched: 0 };

  for (let i = 0; i < names.length; i++) {
    const name = String(names[i] || '').trim();
    if (!name) continue;

    const sh = ss.getSheetByName(name);
    if (!sh) {
      try { logLine_('SUB_WARN', 'Operational sheet not found (skip)', name, '', 'WARN'); } catch (e1) {}
      continue;
    }

    let lastRow = sh.getLastRow();
    let lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      summary.sheets[name] = { updatedRows: 0, skipped: 'empty' };
      continue;
    }

    // Mandatory: ensure Status Type header exists (append right).
    try {
      const hdr0 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
      const norm0 = hdr0.map(h => h.toLowerCase());
      if (norm0.indexOf('status type') < 0 && !isDryRun_()) {
        sh.getRange(1, lastCol + 1).setValue('Status Type');
        lastCol++;
      }
    } catch (e0) {}

    // Re-read after possible header append.
    lastRow = sh.getLastRow();
    lastCol = sh.getLastColumn();
    const all = sh.getRange(1, 1, lastRow, lastCol).getValues();

    const hdr = all[0].map(h => String(h || '').trim());
    const norm = hdr.map(h => h.toLowerCase());

    const isScSheet = (name === scFarhanName || name === scMeilaniName || name === scIvanName);

    function idxOfAny(candidates) {
      for (let k = 0; k < candidates.length; k++) {
        const c = String(candidates[k] || '').toLowerCase();
        const j = norm.indexOf(c);
        if (j >= 0) return j;
      }
      return -1;
    }

    const idxClaim = idxOfAny(['claim number', 'claim_number', 'claim no', 'claim_no']);
    const idxLSA = idxOfAny(['last status aging', 'last_status_aging', 'lsa']);
    const idxALA = idxOfAny(['activity log aging', 'activity_log_aging', 'ala']);
    const idxLast = idxOfAny(['last status', 'last_status']);
    const idxSc = idxOfAny(['service center', 'sc_name', 'service_center']);
    const idxActLog = idxOfAny(['activity log', 'activity_log', 'last activity log', 'last_activity_log']);
    const idxLastStatusDate = idxOfAny(['last status date', 'last_status_date', 'claim_last_updated_datetime', 'claim last updated datetime']);
    const idxStatusType = idxOfAny(['status type']);
    const idxType = isScSheet ? idxOfAny(['type']) : -1;
    const idxSubmissionDate = idxOfAny(['submission date', 'claim_submission_date', 'claim submitted datetime', 'submission_date']);
    const idxSubmissionMonth = idxOfAny(['submission by month', 'submission_month']);
    const idxStoreName = idxOfAny(['store name', 'outlet_name', 'outlet name', 'store_name']);
    const idxPaName = idxOfAny(['pa name', 'pa_name']);
    const idxSpaName = idxOfAny(['spa name', 'spa_name']);
    const idxServiceCenterPic = idxOfAny(['service center pic', 'service_center_pic']);
    const idxUpdateStatus = idxOfAny(['update status']);
    const idxTimestamp = idxOfAny(['timestamp']);
    const idxStatus = idxOfAny(['status']);
    const idxRemarks = idxOfAny(['remarks', 'remark']);
    if (idxClaim < 0) {
      try { logLine_('SUB_WARN', 'Sheet missing Claim Number header (skip)', name, '', 'WARN'); } catch (e2) {}
      summary.sheets[name] = { updatedRows: 0, skipped: 'no Claim Number header' };
      continue;
    }

    const numDataRows = all.length - 1;

    // NOTE: Do NOT auto-apply dropdown rules for "Status Type" (derived by script) or "Type".
    // Dropdown styling/colors must be preserved via template-row DV copy; rebuilding list rules can degrade to standard dropdown.


    if (numDataRows <= 0) {
      summary.sheets[name] = { updatedRows: 0, skipped: 'no data rows' };
      continue;
    }

    // Prepare column outputs (only for columns that exist)
    const outLSA = idxLSA >= 0 ? new Array(numDataRows) : null;
    const outALA = idxALA >= 0 ? new Array(numDataRows) : null;
    const outLast = idxLast >= 0 ? new Array(numDataRows) : null;
    const outSc = idxSc >= 0 ? new Array(numDataRows) : null;
    const outActLog = idxActLog >= 0 ? new Array(numDataRows) : null;
    const outLastStatusDate = idxLastStatusDate >= 0 ? new Array(numDataRows) : null;
    const outStatusType = idxStatusType >= 0 ? new Array(numDataRows) : null;
    const outType = (idxType >= 0) ? new Array(numDataRows) : null;
    const outSubmissionDate = idxSubmissionDate >= 0 ? new Array(numDataRows) : null;
    const outSubmissionMonth = idxSubmissionMonth >= 0 ? new Array(numDataRows) : null;
    const outStoreName = idxStoreName >= 0 ? new Array(numDataRows) : null;
    const outPaName = idxPaName >= 0 ? new Array(numDataRows) : null;
    const outSpaName = idxSpaName >= 0 ? new Array(numDataRows) : null;
    const outServiceCenterPic = idxServiceCenterPic >= 0 ? new Array(numDataRows) : null;
    const outUpdateStatus = idxUpdateStatus >= 0 ? new Array(numDataRows) : null;
    const outTimestamp = idxTimestamp >= 0 ? new Array(numDataRows) : null;
    const outStatus = idxStatus >= 0 ? new Array(numDataRows).fill(null) : null;
    const outRemarks = idxRemarks >= 0 ? new Array(numDataRows) : null;

    function isNonEmpty(v) {
      return v !== '' && v != null;
    }
    let updatedRows = 0;

    for (let r = 1; r < all.length; r++) {
      const row = all[r];
      const cn = __normalizeClaimKeySub06a_(row[idxClaim]);
      const o = r - 1;

      // default keep existing values
      if (outLSA) outLSA[o] = [row[idxLSA]];
      if (outALA) outALA[o] = [row[idxALA]];
      if (outLast) outLast[o] = [row[idxLast]];
      if (outSc) outSc[o] = [row[idxSc]];
      if (outActLog) outActLog[o] = [row[idxActLog]];
      if (outLastStatusDate) outLastStatusDate[o] = [row[idxLastStatusDate]];
      if (outStatusType) outStatusType[o] = [row[idxStatusType]];
      if (outType) outType[o] = [row[idxType]];
      if (outSubmissionDate) outSubmissionDate[o] = [row[idxSubmissionDate]];
      if (outSubmissionMonth) outSubmissionMonth[o] = [row[idxSubmissionMonth]];
      if (outStoreName) outStoreName[o] = [row[idxStoreName]];
      if (outPaName) outPaName[o] = [row[idxPaName]];
      if (outSpaName) outSpaName[o] = [row[idxSpaName]];
      if (outServiceCenterPic) outServiceCenterPic[o] = [row[idxServiceCenterPic]];
      if (outUpdateStatus) outUpdateStatus[o] = [row[idxUpdateStatus]];
      if (outTimestamp) outTimestamp[o] = [row[idxTimestamp]];
      if (outStatus) outStatus[o] = null;
      if (outRemarks) outRemarks[o] = [row[idxRemarks]];

      if (!cn) continue;

      const rec = map.get(cn);
      if (!rec) continue;

      // Only update allowed fields. Do NOT touch other columns.
      if (outLSA && isNonEmpty(rec.last_status_aging)) outLSA[o] = [rec.last_status_aging];
      if (outALA && isNonEmpty(rec.activity_log_aging)) outALA[o] = [rec.activity_log_aging];

      if (outLast && isNonEmpty(rec.last_status)) outLast[o] = [rec.last_status];
      if (outSc && isNonEmpty(rec.sc_name)) outSc[o] = [rec.sc_name];

      if (outActLog && isNonEmpty(rec.activity_log)) outActLog[o] = [rec.activity_log];
      if (outSubmissionDate && isNonEmpty(rec.claim_submitted_datetime)) {
        // Align with sampling flow: write source value directly from claim_submitted_datetime.
        // This preserves valid source formats like "31 Dec 24, 00:00" / "31 Dec 24".
        outSubmissionDate[o] = [rec.claim_submitted_datetime];
      }
      if (outSubmissionMonth) {
        const srcSubDate = isNonEmpty(rec.claim_submitted_datetime) ? rec.claim_submitted_datetime : (idxSubmissionDate >= 0 ? row[idxSubmissionDate] : '');
        outSubmissionMonth[o] = [__formatSubmissionMonthSub06a_(srcSubDate)];
      }
      if (outStoreName && isNonEmpty(rec.store_name)) outStoreName[o] = [rec.store_name];
      if (outPaName && isNonEmpty(rec.pa_name)) outPaName[o] = [rec.pa_name];
      if (outSpaName && isNonEmpty(rec.spa_name)) outSpaName[o] = [rec.spa_name];
      if (outServiceCenterPic) {
        const scRaw = isNonEmpty(rec.sc_name) ? rec.sc_name : (idxSc >= 0 ? row[idxSc] : '');
        outServiceCenterPic[o] = [__deriveServiceCenterPicSub06a_(scRaw)];
      }

      if (outLastStatusDate && isNonEmpty(rec.claim_last_updated_datetime)) {
        const d = __parseClaimLastUpdatedDatetimeSub06a_(rec.claim_last_updated_datetime);
        if (d) outLastStatusDate[o] = [d];
      }
      const prevLast = String(idxLast >= 0 ? (row[idxLast] || '') : '').trim();
      const nextLast = String(isNonEmpty(rec.last_status) ? rec.last_status : prevLast).trim();
      // IMPORTANT:
      // Do NOT reset Update Status/Timestamp/Status/Remarks on Last Status change in-place.
      // Reset now only happens when a claim row is relocated across operational sheets.
      void prevLast; void nextLast;

      if (outStatusType) {
        const st = String(isNonEmpty(rec.last_status) ? rec.last_status : row[idxLast] || '').trim();
        outStatusType[o] = [__getStatusTypeSub06a_(st)];
      }

      // SC sheet Type auto-fill (only when header exists)
      if (outType) {
        const st2 = String(isNonEmpty(rec.last_status) ? rec.last_status : (idxLast >= 0 ? (row[idxLast] || '') : '')).trim();
        const typeLabel = resolveScType(st2);
        if (typeLabel) outType[o] = [typeLabel];
      }

      updatedRows++;
    }

    function writeCol(idx, colVals) {
      if (idx < 0 || !colVals) return;
      const rng = sh.getRange(2, idx + 1, numDataRows, 1);
      if (typeof safeSetValues_ === 'function') safeSetValues_(rng, colVals);
      else rng.setValues(colVals);
    }

    if (!isDryRun_()) {
      writeCol(idxLSA, outLSA);
      writeCol(idxALA, outALA);
      writeCol(idxLast, outLast);
      writeCol(idxSc, outSc);
      writeCol(idxActLog, outActLog);
      writeCol(idxLastStatusDate, outLastStatusDate);
      writeCol(idxStatusType, outStatusType);
      writeCol(idxType, outType);
      // Clear stale checkbox/dropdown DV FIRST; otherwise writing date/text into checkbox
      // cells may be coerced into blank/boolean values.
      if (idxSubmissionDate >= 0) {
        try { sh.getRange(2, idxSubmissionDate + 1, numDataRows, 1).clearDataValidations(); } catch (eDvSubPre) {}
      }
      writeCol(idxSubmissionDate, outSubmissionDate);
      writeCol(idxSubmissionMonth, outSubmissionMonth);
      writeCol(idxStoreName, outStoreName);
      writeCol(idxPaName, outPaName);
      writeCol(idxSpaName, outSpaName);
      writeCol(idxServiceCenterPic, outServiceCenterPic);
      writeCol(idxUpdateStatus, outUpdateStatus);
      writeCol(idxTimestamp, outTimestamp);
      if (idxStatus >= 0 && outStatus) {
        for (let rr = 0; rr < outStatus.length; rr++) {
          const cell = outStatus[rr];
          if (!cell) continue;
          const rgStatus = sh.getRange(2 + rr, idxStatus + 1, 1, 1);
          if (typeof safeSetValue_ === 'function') safeSetValue_(rgStatus, cell[0]);
          else rgStatus.setValue(cell[0]);
        }
      }
      writeCol(idxRemarks, outRemarks);

      // Defensive: Submission Date must stay date-type, not checkbox.
      if (idxSubmissionDate >= 0) {
        try { sh.getRange(2, idxSubmissionDate + 1, numDataRows, 1).clearDataValidations(); } catch (eDvSub) {}
      }

      // Enforce SUB datetime format if Last Status Date exists.
      if (idxLastStatusDate >= 0) {
        const isScUniverse = /^SC\s*-\s*/i.test(name || '');
        const dtFmt = isScUniverse
          ? 'mmm d, yyyy, h:mm AM/PM'
          : ((typeof FORMATS !== 'undefined' && FORMATS && (FORMATS.DATETIME_LONG || FORMATS.DATETIME))
            ? (FORMATS.DATETIME_LONG || FORMATS.DATETIME)
            : 'dd MMM yy, HH:mm');
        try { sh.getRange(2, idxLastStatusDate + 1, numDataRows, 1).setNumberFormat(dtFmt); } catch (eFmt) {}
      }
    }

    summary.sheets[name] = { updatedRows: updatedRows, db: dbTag };
    summary.totalUpdatedRows += updatedRows;
    summary.totalSheetsTouched++;

    try { logLine_('SUB_UPD', 'Updated operational sheet (safe)', name, 'updatedRows=' + updatedRows, 'INFO'); } catch (e3) {}
  }

  return summary;
}

/** SUB helpers: safe parsing & status type resolution */
function __deriveServiceCenterPicSub06a_(serviceCenterName) {
  const sc = String(serviceCenterName == null ? '' : serviceCenterName).toLowerCase();
  if (!sc) return '';
  const policy = (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY) ? OPS_ROUTING_POLICY : null;
  const kw = (policy && policy.SC_NAME_KEYWORDS) ? policy.SC_NAME_KEYWORDS : null;
  if (!kw) return '';

  const sheets = ['SC - Farhan', 'SC - Meilani', 'SC - Meindar'];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const list = Array.isArray(kw[sheet]) ? kw[sheet] : [];
    for (let j = 0; j < list.length; j++) {
      const key = String(list[j] == null ? '' : list[j]).toLowerCase().trim();
      if (key && sc.indexOf(key) > -1) return sheet.replace(/^SC\s*-\s*/i, '').trim();
    }
  }
  return '';
}

function __formatSubmissionMonthSub06a_(v) {
  if (v == null || v === '') return '';
  let d = null;
  if (Object.prototype.toString.call(v) === '[object Date]') d = isNaN(v.getTime()) ? null : v;
  if (!d && typeof normalizeDate_ === 'function') {
    try { d = normalizeDate_(v); } catch (e0) {}
  }
  if (!d && typeof tryNativeParseUnambiguousDate_ === 'function') {
    try { d = tryNativeParseUnambiguousDate_(String(v)); } catch (e1) {}
  }
  if (!d || isNaN(d.getTime())) return '';
  const tz = (Session && Session.getScriptTimeZone) ? (Session.getScriptTimeZone() || 'Asia/Jakarta') : 'Asia/Jakarta';
  try { return Utilities.formatDate(d, tz, 'MMMM yyyy'); } catch (e2) {}
  return '';
}

function __inferStatusResetPreferredBySheetSub06a_(sheetName) {
  const nm = String(sheetName || '').trim().toLowerCase();
  if (!nm) return 'Pending Admin';
  if (nm.indexOf('sc') > -1 || nm.indexOf('service center') > -1) return 'Pending SC';
  return 'Pending Admin';
}

function __getStatusResetValueSub06a_(sh, idxStatus) {
  const fallback = __inferStatusResetPreferredBySheetSub06a_(sh && sh.getName ? sh.getName() : '');
  const preferred = String(fallback || '').trim().toLowerCase();
  try {
    if (!sh || idxStatus == null || idxStatus < 0) return fallback;
    const rowProbe = Math.min(Math.max(sh.getLastRow(), 2), 10);
    let firstAllowed = '';
    for (let r = 2; r <= rowProbe; r++) {
      const dv = sh.getRange(r, idxStatus + 1, 1, 1).getDataValidation();
      if (!dv) continue;
      const t = dv.getCriteriaType();
      const vals = dv.getCriteriaValues() || [];
      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST && vals.length && vals[0].length) {
        const options = vals[0] || [];
        for (let i = 0; i < options.length; i++) {
          const opt = String(options[i] || '').trim();
          if (!opt) continue;
          if (!firstAllowed) firstAllowed = opt;
          if (String(opt).toLowerCase() === preferred) return opt;
        }
      }
    }
    if (firstAllowed) return firstAllowed;
  } catch (e) {}
  return fallback;
}

function __parseClaimLastUpdatedDatetimeSub06a_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return (isNaN(v.getTime()) ? null : v);

  // Prefer strict WIB parser from 06c if available.
  if (typeof parseClaimLastUpdatedDatetime06c_ === 'function') {
    try {
      const d0 = parseClaimLastUpdatedDatetime06c_(v);
      if (d0 && !isNaN(d0.getTime())) return d0;
    } catch (e0) {}
  }

  const s = String(v || '').trim();
  if (!s) return null;

  const tz = (Session && Session.getScriptTimeZone) ? (Session.getScriptTimeZone() || 'Asia/Jakarta') : 'Asia/Jakarta';
  const patterns = [
    'MMMM d, yyyy, h:mm a',
    'MMMM d, yyyy, hh:mm a',
    'MMM d, yyyy, h:mm a',
    'MMM d, yyyy, hh:mm a',
    'yyyy-MM-dd HH:mm:ss',
    'yyyy-MM-dd HH:mm',
    'yyyy/MM/dd HH:mm:ss',
    'yyyy/MM/dd HH:mm',
    'M/d/yyyy h:mm a',
    'M/d/yyyy hh:mm a',
    'M/d/yyyy H:mm',
    'M/d/yyyy HH:mm',
    'd/M/yyyy H:mm',
    'd/M/yyyy HH:mm',
    'dd/MM/yyyy HH:mm:ss',
    'dd/MM/yyyy HH:mm'
  ];

  for (let i = 0; i < patterns.length; i++) {
    try {
      const d = Utilities.parseDate(s, tz, patterns[i]);
      if (d && !isNaN(d.getTime())) return d;
    } catch (e1) {}
  }

  // Month-name fallback (explicit, no Date(string)).
  try {
    const m = s.match(/^([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4}),\s*(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (m) {
      const monName = String(m[1] || '').toLowerCase();
      const day = Number(m[2]);
      const year = Number(m[3]);
      let hh = Number(m[4]);
      const mm = Number(m[5]);
      const ap = String(m[6] || '').toUpperCase();

      const monMap = {
        jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4, may:5,
        jun:6, june:6, jul:7, july:7, aug:8, august:8, sep:9, sept:9, september:9,
        oct:10, october:10, nov:11, november:11, dec:12, december:12
      };
      const mon = monMap[monName] || monMap[monName.slice(0, 3)];
      if (mon && year && day && mm >= 0) {
        if (ap === 'PM' && hh < 12) hh += 12;
        if (ap === 'AM' && hh === 12) hh = 0;
        const iso = Utilities.formatString('%04d-%02d-%02d %02d:%02d:00', year, mon, day, hh, mm);
        try {
          const d2 = Utilities.parseDate(iso, tz, 'yyyy-MM-dd HH:mm:ss');
          if (d2 && !isNaN(d2.getTime())) return d2;
        } catch (e2) {}
      }
    }
  } catch (e3) {}

  return null;
}

function __getStatusTypeSub06a_(lastStatus) {
  const st = String(lastStatus || '').trim();
  if (!st) return '';
  if (typeof getStatusType06c_ === 'function') {
    try { return getStatusType06c_(st); } catch (e0) {}
  }
  // Fallback if 06c mapping not loaded.
  return '';
}

/** SUB relocation: ensure each row sits in the operational sheet that matches its Last Status mapping.
 *  - Moves FULL row (all columns) and preserves existing fields.
 *  - Deduplicates by Claim Number (keeps latest by Last Status Date when available).
 *  - Uses Service Center allowlists to prevent SC - Farhan pollution.
 */
function __relocateOperationalRowsByLastStatusSub06a_(ss, sheetNames) {
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  const res = { moved: 0, dedupDeleted: 0, sheets: {} };

  const routingMap = __getSubRoutingMap06a_();
  const routingIdxRaw = (typeof buildRoutingIndex06_ === 'function') ? buildRoutingIndex06_(routingMap) : __buildRoutingIndexLocalSub06a_(routingMap);
  const routingIdx = __normalizeRoutingIndexSub06a_(routingIdxRaw);
  const scAllow = __getScSheetAllowlistsSub06a_();
  const scPolicy = __getScRoutingPolicySub06a_();

  // Preload sheet data
  const data = {};
  names.forEach(n => {
    const name = String(n || '').trim();
    if (!name) return;
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) return;
    const vals = sh.getRange(1, 1, lr, lc).getValues();
    const hdr = vals[0].map(h => String(h || '').trim());
    const norm = hdr.map(h => h.toLowerCase());
    data[name] = { sh, vals, hdr, norm, lr, lc };
  });

  function idxOfAny(norm, candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const k = String(candidates[i] || '').toLowerCase();
      const j = norm.indexOf(k);
      if (j >= 0) return j;
    }
    return -1;
  }

  function normKey(s) {
    return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  function pickDest(status, scName, candidates) {
    const cands = Array.isArray(candidates) ? candidates.slice() : [];
    if (!cands.length) return null;

    // SC routing is keyword-based: pick the SC sheet whose keyword best matches Service Center.
    const sc = __normalizeScString06a_(scName);
    const hasScRules = cands.some(s => Array.isArray(scAllow[s]) && scAllow[s].length);

    if (!hasScRules) return cands[0];

    const fb = __getScFallbackSheet06a_();
    let bestSheet = null;
    let bestScore = 0;

    for (let i = 0; i < cands.length; i++) {
      const sheet = cands[i];
      const score = __scoreScMatch06a_(sc, scAllow[sheet]);
      if (score > bestScore) {
        bestScore = score;
        bestSheet = sheet;
      }
    }

    if (bestSheet && bestScore > 0) return bestSheet;

    // No keyword match -> route to fallback sheet (quarantine), not to a random SC PIC sheet.
    return fb || null;
  }



  // Build moves and dedupe deletes
  const movesBySource = {}; // { [sheetName]: [{row1Based, claim, dest, rowVals, hdr}] }
  const deletesBySheet = {}; // { [sheetName]: Set(row1Based) }

  Object.keys(data).forEach(sheetName => {
    const d = data[sheetName];
    const norm = d.norm;

    const idxClaim = idxOfAny(norm, ['claim number', 'claim_number', 'claim no', 'claim_no']);
    const idxStatus = idxOfAny(norm, ['last status', 'last_status']);
    const idxSc = idxOfAny(norm, ['service center', 'sc_name', 'service_center']);
    const idxLsd = idxOfAny(norm, ['last status date', 'claim_last_updated_datetime', 'claim last updated datetime']);

    if (idxClaim < 0 || idxStatus < 0) {
      res.sheets[sheetName] = { skipped: 'missing Claim Number or Last Status' };
      return;
    }

    const seen = new Map(); // claim -> {row1Based, lsdTime}
    const toDelete = new Set();

    // 1) Dedup within sheet
    for (let r = 1; r < d.vals.length; r++) {
      const row = d.vals[r];
      const claim = String(row[idxClaim] || '').trim();
      if (!claim) continue;

      let t = 0;
      if (idxLsd >= 0) {
        const dd = __parseClaimLastUpdatedDatetimeSub06a_(row[idxLsd]);
        t = dd ? dd.getTime() : 0;
      }

      if (!seen.has(claim)) {
        seen.set(claim, { row1Based: r + 1, lsdTime: t });
      } else {
        const prev = seen.get(claim);
        // Keep latest by Last Status Date when possible; else keep first.
        const keepCurrent = (t && t > (prev.lsdTime || 0));
        if (keepCurrent) {
          toDelete.add(prev.row1Based);
          seen.set(claim, { row1Based: r + 1, lsdTime: t });
        } else {
          toDelete.add(r + 1);
        }
      }
    }

    if (toDelete.size) {
      deletesBySheet[sheetName] = deletesBySheet[sheetName] || new Set();
      toDelete.forEach(x => deletesBySheet[sheetName].add(x));
      res.dedupDeleted += toDelete.size;
    }

    // 2) Movement by status routing
    for (let r = 1; r < d.vals.length; r++) {
      const row = d.vals[r];
      const claim = String(row[idxClaim] || '').trim();
      if (!claim) continue;

      // Skip rows that will be deleted as duplicates
      if (toDelete.has(r + 1)) continue;

      const status = __normalizeRoutingStatusKeySub06a_(row[idxStatus]);
      if (!status) continue;

      let candidates = routingIdx[status] || null;
      // Force SC-universe statuses to be routed by SC keyword split, even if routing map is stale/misaligned.
      if (scPolicy.sharedStatusSet.has(status)) {
        candidates = scPolicy.scSheets.filter(function (n) { return !!ss.getSheetByName(n); });
      }
      if (!candidates || !candidates.length) continue;

      const scName = (idxSc >= 0) ? row[idxSc] : '';
      let dest = pickDest(status, scName, candidates);

      // [Inference] If a row is already in SC - Farhan but Service Center is outside allowlist, push it to SC - Meilani as a safe default.
      // Override by ensuring CONFIG.statusRoutingSub (or CONFIG.statusRoutingAdmin) maps SC-stage statuses to all SC sheets, and/or configure CONFIG.SC_SHEET_ALLOWLISTS.
      if (!dest && sheetName === 'SC - Farhan') {
        const scKey = normKey(scName);
        const allowF = scAllow['SC - Farhan'];
        const inFarhanAllow = Array.isArray(allowF) ? (allowF.indexOf(scKey) > -1) : false;
        if (allowF && scKey && !inFarhanAllow && ss.getSheetByName('SC - Meilani')) {
          dest = 'SC - Meilani';
        }
      }

      if (!dest || dest === sheetName) continue;

      movesBySource[sheetName] = movesBySource[sheetName] || [];
      movesBySource[sheetName].push({ row1Based: r + 1, claim, dest, rowVals: row.slice(), srcHdr: d.hdr });
    }
  });

  // Preload target claim indexes (for dedupe in target)
  const targetIndex = {};
  function getTargetIndex(targetName) {
    if (targetIndex[targetName]) return targetIndex[targetName];
    const sh = ss.getSheetByName(targetName);
    if (!sh) return null;

    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) return null;

    const hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || '').trim());
    const norm = hdr.map(h => h.toLowerCase());
    const idxClaim = idxOfAny(norm, ['claim number', 'claim_number', 'claim no', 'claim_no']);
    if (idxClaim < 0) return null;

    const map = new Map();
    if (lr >= 2) {
      const col = sh.getRange(2, idxClaim + 1, lr - 1, 1).getValues();
      for (let i = 0; i < col.length; i++) {
        const cn = String(col[i][0] || '').trim();
        if (!cn) continue;
        if (!map.has(cn)) map.set(cn, []);
        map.get(cn).push(i + 2);
      }
    }

    targetIndex[targetName] = { sh, hdr, norm, idxClaim, lc, map };
    return targetIndex[targetName];
  }

  function alignRowToTarget(srcHdr, srcRow, tgtHdr, tgtColCount) {
    // Fast path: same header length and same headers
    if (srcHdr && tgtHdr && srcHdr.length === tgtHdr.length) {
      let same = true;
      for (let i = 0; i < tgtHdr.length; i++) {
        if (String(srcHdr[i] || '').trim() !== String(tgtHdr[i] || '').trim()) { same = false; break; }
      }
      if (same) {
        const out = srcRow.slice(0, tgtColCount);
        while (out.length < tgtColCount) out.push('');
        return out;
      }
    }

    const srcNorm = (srcHdr || []).map(h => String(h || '').trim().toLowerCase());
    const out = new Array(tgtColCount).fill('');
    for (let i = 0; i < tgtHdr.length && i < tgtColCount; i++) {
      const key = String(tgtHdr[i] || '').trim().toLowerCase();
      const j = srcNorm.indexOf(key);
      if (j >= 0 && j < srcRow.length) out[i] = srcRow[j];
    }
    return out;
  }

  function getResetColumnIndexesByHeader(tgtHdr) {
    const norm = (tgtHdr || []).map(h => String(h || '').trim().toLowerCase());
    function idxOfAnyLocal(cands) {
      for (let i = 0; i < cands.length; i++) {
        const j = norm.indexOf(String(cands[i] || '').toLowerCase());
        if (j >= 0) return j;
      }
      return -1;
    }
    return {
      updateStatus: idxOfAnyLocal(['update status']),
      timestamp: idxOfAnyLocal(['timestamp']),
      status: idxOfAnyLocal(['status']),
      remarks: idxOfAnyLocal(['remarks', 'remark'])
    };
  }

  function resetMovedRowFieldsByHeader(rowVals, resetIdx) {
    const out = Array.isArray(rowVals) ? rowVals.slice() : [];
    if (!resetIdx) return out;
    const keys = ['updateStatus', 'timestamp', 'status', 'remarks'];
    for (let i = 0; i < keys.length; i++) {
      const ix = resetIdx[keys[i]];
      if (ix != null && ix >= 0 && ix < out.length) out[ix] = '';
    }
    return out;
  }


  function deleteRowsGrouped(sh, rows1Based) {
    const rows = (rows1Based || []).slice().filter(n => Number(n) >= 2);
    if (!rows.length) return;
    rows.sort((a, b) => a - b);

    const groups = [];
    let start = rows[0];
    let prev = rows[0];
    let len = 1;

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r === prev + 1) {
        prev = r;
        len++;
      } else {
        groups.push({ start, len });
        start = prev = r;
        len = 1;
      }
    }
    groups.push({ start, len });

    // Delete bottom-up
    const failed = [];
    for (let i = groups.length - 1; i >= 0; i--) {
      const g = groups[i];
      try { sh.deleteRows(g.start, g.len); } catch (eDel) { failed.push(g); }
    }

    // Fallback: if deleteRows fails (e.g., protections/edge-sheet state), clear row content
    // so claims do not remain duplicated in the wrong sheet.
    if (failed.length) {
      const lc = Math.max(sh.getLastColumn() || 1, 1);
      for (let i = 0; i < failed.length; i++) {
        const g = failed[i];
        try {
          sh.getRange(g.start, 1, g.len, lc).clearContent();
        } catch (eClr) {
          try { logLine_('SUB_WARN', 'Delete/clear failed in relocation source', sh.getName() + ' row=' + g.start + ' len=' + g.len, String(eClr), 'WARN'); } catch (eLog) {}
        }
      }
    }
  }

  function richTextHasAnyLink(rt) {
    try {
      if (!rt) return false;
      if (rt.getLinkUrl && rt.getLinkUrl()) return true;
      if (rt.getRuns) {
        const runs = rt.getRuns();
        for (let i = 0; i < runs.length; i++) {
          const r = runs[i];
          if (r && r.getLinkUrl && r.getLinkUrl()) return true;
        }
      }
    } catch (e0) {}
    return false;
  }


  function applyRichTextLinksToTarget(srcSheetName, srcRow1Based, srcHdr, tgtSheet, tgtHdr, tgtRow1Based, tgtColCount) {
    try {
      const srcData = data[srcSheetName];
      if (!srcData) return;
      const srcLc = srcData.lc;
      if (srcRow1Based < 2) return;
      const srcRt = srcData.sh.getRange(srcRow1Based, 1, 1, srcLc).getRichTextValues()[0];

      const sameSchema = (srcHdr && tgtHdr && srcHdr.length === tgtHdr.length && srcHdr.length > 0)
        ? srcHdr.every((h, i) => String(h || '').trim() === String(tgtHdr[i] || '').trim())
        : false;

      const srcNorm = sameSchema ? null : (srcHdr || []).map(h => String(h || '').trim().toLowerCase());
      for (let i = 0; i < tgtColCount && i < (tgtHdr || []).length; i++) {
        const j = sameSchema ? i : srcNorm.indexOf(String(tgtHdr[i] || '').trim().toLowerCase());
        if (j < 0 || j >= srcRt.length) continue;
        const rt = srcRt[j];
        if (!richTextHasAnyLink(rt)) continue;
        // Copy hyperlink richtext only (avoid converting numbers/formats for other cells).
        tgtSheet.getRange(tgtRow1Based, i + 1).setRichTextValue(rt);
      }
    } catch (e0) {}
  }


  // Execute moves: append to target then delete from source (descending row order per source).
  Object.keys(movesBySource).forEach(srcName => {
    const list = movesBySource[srcName] || [];
    if (!list.length) return;

    // Sort source rows descending for safe deletes later.
    list.sort((a, b) => b.row1Based - a.row1Based);

    // Append first
    for (let i = 0; i < list.length; i++) {
      const mv = list[i];
      const tgt = getTargetIndex(mv.dest);
      if (!tgt || !tgt.sh) {
        res.sheets[srcName] = res.sheets[srcName] || {};
        res.sheets[srcName].errors = (res.sheets[srcName].errors || 0) + 1;
        try { logLine_('SUB_WARN', 'Move skipped: target missing or no Claim Number header', srcName + ' -> ' + mv.dest, mv.claim, 'WARN'); } catch (e0) {}
        continue;
      }

      const aligned = alignRowToTarget(mv.srcHdr, mv.rowVals, tgt.hdr, tgt.lc);
      const resetIdx = getResetColumnIndexesByHeader(tgt.hdr);
      const alignedAfterReset = resetMovedRowFieldsByHeader(aligned, resetIdx);

      // If claim already exists in target, MERGE non-empty cells to avoid data loss and avoid duplicates.
      const existing = tgt.map.get(mv.claim) || [];
      if (existing.length) {
        const keepRow = existing.slice().sort((a, b) => a - b)[0];

        if (!isDryRun_()) {
          const existingRowVals = tgt.sh.getRange(keepRow, 1, 1, tgt.lc).getValues()[0];
          const merged = existingRowVals.slice();
          for (let c = 0; c < tgt.lc; c++) {
            if (alignedAfterReset[c] !== '' && alignedAfterReset[c] != null) merged[c] = alignedAfterReset[c];
          }
          // Always reset the 4 manual columns after cross-sheet movement.
          const mergedAfterReset = resetMovedRowFieldsByHeader(merged, resetIdx);
          tgt.sh.getRange(keepRow, 1, 1, tgt.lc).setValues([mergedAfterReset]);
          applyRichTextLinksToTarget(srcName, mv.row1Based, mv.srcHdr, tgt.sh, tgt.hdr, keepRow, tgt.lc);
        }

        // Delete other duplicates in target (keepRow stays)
        const del = existing.slice().filter(r => r !== keepRow).sort((a, b) => b - a);
        if (del.length && !isDryRun_()) {
          deleteRowsGrouped(tgt.sh, del);
        }
        res.dedupDeleted += Math.max(0, existing.length - 1);

        tgt.map.set(mv.claim, [keepRow]);
      } else {
        // Append to bottom (as requested)
        const appendRow = tgt.sh.getLastRow() + 1;
        if (!isDryRun_()) {
          // Preserve richtext/hyperlinks + formatting when schemas match.
          const srcSh = (data[srcName] && data[srcName].sh) ? data[srcName].sh : null;
          const sameSchema = (mv.srcHdr && tgt.hdr && mv.srcHdr.length === tgt.hdr.length)
            ? mv.srcHdr.every((h, ii) => String(h || '').trim() === String(tgt.hdr[ii] || '').trim())
            : false;

          if (srcSh && sameSchema) {
            srcSh.getRange(mv.row1Based, 1, 1, tgt.lc)
              .copyTo(tgt.sh.getRange(appendRow, 1, 1, tgt.lc), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            const resetRow = resetMovedRowFieldsByHeader(tgt.sh.getRange(appendRow, 1, 1, tgt.lc).getValues()[0], resetIdx);
            tgt.sh.getRange(appendRow, 1, 1, tgt.lc).setValues([resetRow]);
          } else {
            tgt.sh.getRange(appendRow, 1, 1, tgt.lc).setValues([alignedAfterReset]);
            applyRichTextLinksToTarget(srcName, mv.row1Based, mv.srcHdr, tgt.sh, tgt.hdr, appendRow, tgt.lc);
          }
        }
        tgt.map.set(mv.claim, [appendRow]);
      }

      res.moved++;
      res.sheets[srcName] = res.sheets[srcName] || {};
      res.sheets[srcName].moved = (res.sheets[srcName].moved || 0) + 1;
    }

    // Delete sources (including dedupe deletes)
    const delSet = deletesBySheet[srcName] || new Set();
    list.forEach(mv => delSet.add(mv.row1Based));

    const delRows = Array.from(delSet).sort((a, b) => b - a);
    if (!isDryRun_()) {
      deleteRowsGrouped(data[srcName].sh, delRows);
    }
  });

  // Delete duplicate-only sheets (no moves) in descending order
  Object.keys(deletesBySheet).forEach(sheetName => {
    if (movesBySource[sheetName] && movesBySource[sheetName].length) return; // already handled
    const shData = data[sheetName];
    if (!shData) return;
    const delRows = Array.from(deletesBySheet[sheetName]).sort((a, b) => b - a);
    if (!isDryRun_()) {
      deleteRowsGrouped(shData.sh, delRows);
    }
  });

  try { logLine_('SUB_MOVE', 'Relocation summary', 'moved=' + res.moved + ' dedupDeleted=' + res.dedupDeleted, '', 'INFO'); } catch (e9) {}
  return res;
}

function __getSubRoutingMap06a_() {
  // Prefer explicit SUB routing if provided; fallback to the existing routing map used by MAIN/admin.
  try {
    if (CONFIG && CONFIG.statusRoutingSub) return CONFIG.statusRoutingSub;
    if (CONFIG && CONFIG.statusRoutingAdmin) return CONFIG.statusRoutingAdmin;
    if (CONFIG && CONFIG.statusRouting) return CONFIG.statusRouting;
  } catch (e0) {}
  return {};
}

function __buildRoutingIndexLocalSub06a_(routingMap) {
  const map = routingMap || {};
  const idx = {};
  Object.keys(map).forEach(sheetName => {
    const sheetKey = String(sheetName || '').trim();
    if (!sheetKey || /^__/.test(sheetKey)) return;
    (map[sheetName] || []).forEach(status => {
      const key = __normalizeRoutingStatusKeySub06a_(status);
      if (!key) return;
      if (!idx[key]) idx[key] = [];
      idx[key].push(sheetKey);
    });
  });
  return idx;
}

function __getScRoutingPolicySub06a_() {
  const out = {
    scSheets: ['SC - Farhan', 'SC - Meilani', 'SC - Meindar'],
    sharedStatusSet: new Set()
  };

  try {
    if (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY) {
      const s = OPS_ROUTING_POLICY.SHEETS || {};
      out.scSheets = [
        String(s.SC_FARHAN || 'SC - Farhan').trim(),
        String(s.SC_MEILANI || 'SC - Meilani').trim(),
        String((s.SC_IVAN || s.SC_IVAN_NAME || 'SC - Meindar')).trim()
      ].filter(Boolean);

      const shared = (OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET && OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__'])
        ? OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET['__SC_SHARED__']
        : [];
      (shared || []).forEach(function (st) {
        const k = __normalizeRoutingStatusKeySub06a_(st);
        if (k) out.sharedStatusSet.add(k);
      });
    }
  } catch (e0) {}

  if (!out.sharedStatusSet.size) {
    // Minimal fallback for critical SC statuses observed in SUB updates.
    ['SERVICE_CENTER_CLAIM_RECEIVE', 'SERVICE_CENTER_CLAIM_ESTIMATE', 'SERVICE_CENTER_CLAIM_RESUBMIT_ESTIMATE', 'QOALA_CLAIM_RESUBMIT_ESTIMATE']
      .forEach(function (st) { out.sharedStatusSet.add(st); });
  }
  return out;
}

function __normalizeRoutingStatusKeySub06a_(status) {
  return String(status == null ? '' : status)
    .replace(/\u00a0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function __normalizeRoutingIndexSub06a_(routingIdx) {
  const src = routingIdx || {};
  const out = {};
  Object.keys(src).forEach(function (k) {
    const nk = __normalizeRoutingStatusKeySub06a_(k);
    if (!nk) return;
    const arr = Array.isArray(src[k]) ? src[k] : [];
    if (!out[nk]) out[nk] = [];
    arr.forEach(function (name) {
      const sn = String(name || '').trim();
      if (!sn) return;
      if (out[nk].indexOf(sn) === -1) out[nk].push(sn);
    });
  });
  return out;
}

function __getScSheetAllowlistsSub06a_() {
  // Keyword routing by sheetName. Values are arrays of normalized lowercase keywords.
  // Override via CONFIG.SC_SHEET_ALLOWLISTS = { 'SC - Farhan': ['Mitracare', ...], 'SC - Meindar': [...], 'SC - Meilani': [...] }
  const out = {};

  function norm(s) { return __normalizeScString06a_(s); }

  // 1) explicit override (highest priority)
  try {
    if (CONFIG && CONFIG.SC_SHEET_ALLOWLISTS) {
      Object.keys(CONFIG.SC_SHEET_ALLOWLISTS).forEach(k => {
        const arr = CONFIG.SC_SHEET_ALLOWLISTS[k] || [];
        out[k] = arr.map(norm).filter(Boolean);
      });
    }
  } catch (e0) {}

  // 2) default from OPS_ROUTING_POLICY (project source-of-truth)
  try {
    if (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY && OPS_ROUTING_POLICY.SC_NAME_KEYWORDS) {
      Object.keys(OPS_ROUTING_POLICY.SC_NAME_KEYWORDS).forEach(k => {
        if (out[k] && out[k].length) return;
        const arr = OPS_ROUTING_POLICY.SC_NAME_KEYWORDS[k] || [];
        out[k] = arr.map(norm).filter(Boolean);
      });
    }
  } catch (e1) {}

  // 3) hard fallback (should rarely be used)
  if (!out['SC - Farhan'] || !out['SC - Farhan'].length) out['SC - Farhan'] = ['mitracare', 'sitcomtara', 'gsi', 'ibox'].map(norm);
  if (!out['SC - Meindar'] || !out['SC - Meindar'].length) out['SC - Meindar'] = [].map(norm);
  if (!out['SC - Meilani'] || !out['SC - Meilani'].length) out['SC - Meilani'] = [].map(norm);

  return out;
}

function __normalizeScString06a_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function __getScFallbackSheet06a_() {
  try {
    if (CONFIG && CONFIG.SC_FALLBACK_SHEET) return String(CONFIG.SC_FALLBACK_SHEET || '').trim();
    if (CONFIG && CONFIG.scFallbackSheet) return String(CONFIG.scFallbackSheet || '').trim();
  } catch (e0) {}
  try {
    if (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY && OPS_ROUTING_POLICY.SC_FALLBACK_SHEET) {
      return String(OPS_ROUTING_POLICY.SC_FALLBACK_SHEET || '').trim();
    }
  } catch (e1) {}
  return 'SC - Unmapped';
}

function __scoreScMatch06a_(scNorm, keywords) {
  if (!scNorm) return 0;
  const list = Array.isArray(keywords) ? keywords : [];
  let best = 0;
  for (let i = 0; i < list.length; i++) {
    const kw = String(list[i] || '');
    if (!kw) continue;
    if (scNorm.indexOf(kw) >= 0) best = Math.max(best, kw.length);
  }
  return best;
}


/** Append missing rows into Submission (rules differ for OLD vs NEW). */
function __appendSubmissionFromRawIfMissing06a_(ss, rawMap, rawHdrIdx, dbTag) {
  if (typeof __appendSubmissionFromRawIfMissing06e_ === 'function') {
    return __appendSubmissionFromRawIfMissing06e_(ss, rawMap, rawHdrIdx, dbTag);
  }
  throw new Error('__appendSubmissionFromRawIfMissing06e_ is not available.');
}

/** Sort operational sheets by Last Status Aging (Z->A), Last Status (A->Z), DB (A->Z). */
/** Sort operational sheets safely for SUB (preserve filter if present).
 * Order:
 *  1) Submission Date (A->Z)
 *  2) Last Status Date (A->Z)
 *  3) Last Status (A->Z)
 */
function __sortOperationalSheetsSub06a_(ss, sheetNames, sortSpecs) {
  if (typeof __sortOperationalSheetsSub06e_ === 'function') {
    return __sortOperationalSheetsSub06e_(ss, sheetNames, sortSpecs);
  }
  throw new Error('__sortOperationalSheetsSub06e_ is not available.');
}


function installEmailIngestTrigger() {
  // MAIN: daily at 08:00 (script timezone)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'runEmailIngest') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runEmailIngest').timeBased().everyDays(1).atHour(8).create();
}


/**
 * Manual runner:
 * runManual('Farhan', 'id1,id2,id3')
 */
/**
 * Manual runner:
 * - runManual('id1,id2')  // defaults to Master
 * - runManual('Master', 'id1,id2')
 */
function runManual(picOrFileIdsCsv, fileIdsCsvMaybe) {
  try { if (typeof resetRuntime_ === 'function') resetRuntime_(); } catch (err0) {}

  // Backward-compatible signature detection
  let pic = null;
  let fileIdsCsv = null;
  if (fileIdsCsvMaybe == null) {
    fileIdsCsv = String(picOrFileIdsCsv || '');
  } else {
    pic = picOrFileIdsCsv;
    fileIdsCsv = String(fileIdsCsvMaybe || '');
  }

  const key = resolveSpreadsheetKey_(pic);

  return withLock_(() => {
    resetRunState_();
    if (PIPELINE_FLAGS.CLEAR_LOG_BEFORE_RUN) clearLogSheet_();

    const startedAt = new Date();
    let ssTiming = null;
    try { ssTiming = __logOverviewStart06_(key, startedAt); } catch (e) {}

    setProgress_(0, 'Starting (manual)...');
    logLine_('BOOT', 'Manual run started', 'version=' + ((App && App.APP_VERSION) ? App.APP_VERSION : ''), 'profile=' + key, 'INFO');

    const ids = String(fileIdsCsv || '')
      .split(/[\s,;]+/)
      .map(s => s.trim())
      .filter(Boolean)
      .map(extractDriveIdFromUrl_)
      .filter(Boolean);

    if (!ids.length) throw new Error('fileIdsCsv is empty. Provide Drive file IDs/URLs.');

    const seg = startSegment_('MAIN', 'Pipeline run (manual)');
    try {
      const res = runPipeline_(key, ids, { flow: 'main', source: 'MANUAL' });
      endSegment_(
        seg,
        'routed=' + (res && res.routedTotal ? res.routedTotal : 0),
        (res && res.message) || 'OK',
        (res && res.severity) || 'INFO'
      );
      return res;
    } catch (err) {
      endSegment_(seg, '', 'FATAL: ' + err, 'ERROR');
      throw err;
    }
    finally {
      try { __logOverviewDuration06_(key, startedAt, ssTiming); } catch (e2) {}
    }
  });
}
