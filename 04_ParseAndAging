/***************************************************************
 * 04_ParseAndAging.gs
 * Enterprise Edition — Upload context, file parsing, aging join helpers (Split-05 compatible)
 * Refactor focus: fewer Drive calls, typed coercion entrypoints
 ***************************************************************/
'use strict';

/** ---------- Safe global accessors ---------- */
function pipelineFlag_(key, fallback) {
  try {
    if (typeof PIPELINE_FLAGS !== 'undefined' && PIPELINE_FLAGS && Object.prototype.hasOwnProperty.call(PIPELINE_FLAGS, key)) {
      return !!PIPELINE_FLAGS[key];
    }
  } catch (e) {}
  return !!fallback;
}
function isDryRunLocal_() {
  try {
    if (typeof isDryRun_ === 'function') return !!isDryRun_();
    if (typeof DRY_RUN !== 'undefined') return !!DRY_RUN;
  } catch (e) {}
  return false;
}


/** ---------- Segment wrappers (04) ---------- */
function startSeg04_(label, detail) {
  try {
    if (typeof startSegment_ === 'function') return startSegment_(label, detail);
  } catch (e) {}
  return { label: String(label || ''), detail: String(detail || ''), t0: Date.now() };
}

function endSeg04_(seg, meta, msg, level) {
  try {
    if (typeof endSegment_ === 'function') return endSegment_(seg, meta, msg, level);
  } catch (e) {}

  // Fallback logging (keeps runs debuggable even if the segment framework isn't loaded).
  try {
    var dur = (seg && seg.t0) ? (Date.now() - seg.t0) : null;
    var parts = [];
    parts.push('[04]');
    parts.push(level || 'INFO');
    if (seg && seg.label) parts.push(String(seg.label));
    if (dur != null) parts.push('(' + dur + 'ms)');
    if (msg) parts.push(String(msg));
    if (meta) parts.push('| ' + String(meta));
    Logger.log(parts.join(' '));
  } catch (e2) {}
}

/** ---------- Header normalization ---------- */
function cleanHeaderCell_(v) {
  // Remove UTF-8 BOM, normalize whitespace and NBSP; helps prevent header mismatches.
  const s = String(v == null ? '' : v)
    .replace(/^\ufeff/, '')
    .replace(/\u00a0/g, ' ')
    .replace(/\u200b/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  // Optional: canonicalize to snake_case for schema/type registries that expect keys like "claim_number".
  // Default is OFF to avoid breaking sheets that rely on human-readable headers.
  if (pipelineFlag_('CANONICALIZE_HEADERS', false)) {
    return s
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .replace(/_+/g, '_');
  }

  return s;
}
/**
 * cleanHeaderKey04_(v) -> string
 * Canonical comparison key for header matching (NOT for display).
 * Uses cleanHeaderCell_ + extra normalization (zero-width, punctuation).
 */
function cleanHeaderKey04_(v) {
  const base = cleanHeaderCell_(v);
  return String(base == null ? '' : base)
    .replace(/\u200b/g, '')            // zero-width space
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}




/** ---------- Header/table helpers (04) ---------- */
function debugLog04_(msg) {
  try {
    if (typeof log_ === 'function') return log_(msg);
  } catch (e) {}
  try { Logger.log(msg); } catch (e2) {}
}

function dedupeHeader_(header) {
  const seen = Object.create(null);
  const out = new Array(header.length);
  for (let i = 0; i < header.length; i++) {
    const base = cleanHeaderCell_(header[i] == null ? '' : header[i]);
    const key = String(base || '').toLowerCase();
    if (!key) { out[i] = base; continue; }
    if (!seen[key]) { seen[key] = 1; out[i] = base; continue; }
    seen[key] += 1;
    out[i] = base + '__' + seen[key];
  }
  return out;
}

function trimEmptyColumns_(values, startRow, endRow) {
  // Determine last column with any non-empty cell between startRow..endRow (inclusive).
  let lastCol = 0;
  const sr = Math.max(0, startRow || 0);
  const er = Math.min(values.length - 1, (endRow == null ? values.length - 1 : endRow));
  for (let r = sr; r <= er; r++) {
    const row = values[r] || [];
    for (let c = row.length - 1; c >= 0; c--) {
      const v = row[c];
      if (v !== '' && v != null) { if (c + 1 > lastCol) lastCol = c + 1; break; }
    }
  }
  if (lastCol <= 0) return values;
  return values.map(r => (r || []).slice(0, lastCol));
}

function pickHeaderRowIndex_(values) {
  // Heuristic: only shift header row if row-0 is clearly a title row and a later row has much richer headers.
  const maxScan = Math.min(20, values.length);
  const countNonEmpty = row => (row || []).reduce((n, v) => (v !== '' && v != null ? n + 1 : n), 0);

  const c0 = countNonEmpty(values[0] || []);
  let bestIdx = 0;
  let best = c0;

  for (let i = 1; i < maxScan; i++) {
    const c = countNonEmpty(values[i] || []);
    if (c > best) { best = c; bestIdx = i; }
  }

  // shift only when clearly better and row0 looks sparse
  if (bestIdx !== 0 && best >= 5 && (c0 < 5) && (best >= c0 + 2)) return bestIdx;
  return 0;
}

function extractTableFromValues04_(values) {
  if (!values || !values.length) return { header: [], rows: [] };

  const trimmed = trimEmptyColumns_(values, 0, Math.min(values.length - 1, 10));
  const headerRowIndex = pipelineFlag_('SMART_HEADER_DETECT', true) ? pickHeaderRowIndex_(trimmed) : 0;

  const headerRaw = trimmed[headerRowIndex] || [];
  let header = headerRaw.map(cleanHeaderCell_);

  if (pipelineFlag_('DEDUP_HEADERS', true)) header = dedupeHeader_(header);

  const rows = trimmed.slice(headerRowIndex + 1);
  return { header, rows };
}



/** ---------- Drive ID extraction ---------- */
function extractDriveIdFromUrl_(input) {
  if (!input) return '';
  const s = String(input).trim();
  const looksLike_ = (token) => /^[A-Za-z0-9_-]{25,}$/.test(String(token == null ? '' : token).trim());

  if (!s) return '';

  // Prefer explicit URL patterns first (least ambiguous).
  const patterns = [
    /\/d\/([A-Za-z0-9_-]{25,})/i,              // .../d/<id>/...
    /\bid=([A-Za-z0-9_-]{25,})/i,               // ...id=<id>
    /\/folders\/([A-Za-z0-9_-]{25,})/i,        // .../folders/<id>
    /\/file\/d\/([A-Za-z0-9_-]{25,})/i        // .../file/d/<id>
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = s.match(patterns[i]);
    if (m && m[1]) return m[1];
  }

  // Fallback: scan for any plausible token.
  const candidates = s.match(/[A-Za-z0-9_-]{25,}/g) || [];
  for (let i = 0; i < candidates.length; i++) {
    const c = candidates[i];
    if (looksLike_(c)) return c;
  }

  // Raw id pasted alone.
  if (looksLike_(s)) return s;

  return '';
}


function extractFileIdsFromArray_(arr) {
  const ids = [];
  (arr || []).forEach(v => {
    if (!v) return;
    String(v).split(/[,;\n]/).forEach(p => {
      const id = extractDriveIdFromUrl_(p);
      if (id) ids.push(id);
    });
  });
  return Array.from(new Set(ids));
}

/** ---------- Context extraction (Form submit) ---------- */
/**
 * NOTE:
 * - Legacy "Extract to Spreadsheet" / PIC routing is removed.
 * - Form upload is an optional/manual flow that always targets the single master workbook.
 * - This helper only extracts Drive file IDs from the file upload field (default: "Metabase - Upload Claim Data"),
 *   with robust fallbacks scanning the response row values.
 */

/**
 * @deprecated Legacy PIC detection when multiple target spreadsheets were used.
 * Always returns '' in the single-master pipeline.
 */
function detectPicFromRowValues_(rowVals) {
  return '';
}

function detectFileIdsFromRowValues_(rowVals) {
  if (!rowVals) return [];
  return extractFileIdsFromArray_([rowVals.join('\n')]);
}

/**
 * getSubmissionContext_(e) => { fileIds, targetSpreadsheetId, targetRawSheetName, ... }
 * Keeps extractToSpreadsheet for backward compatibility, but it is unused in the new flow.
 */
function getSubmissionContext_(e) {
  const ctx = {
    extractToSpreadsheet: '', // legacy; unused
    targetSpreadsheetId: (typeof MASTER_SPREADSHEET_ID !== 'undefined') ? MASTER_SPREADSHEET_ID : '',
    targetRawSheetName: (typeof MASTER_RAW_SHEET_NAME !== 'undefined') ? MASTER_RAW_SHEET_NAME : '',
    fileIds: [],
    sub: { oldFileIds: [], newFileIds: [] },
    source: 'FORM',
    // Backward-compat: "flow" is best-effort. Prefer ctx.requestedFlowLabel.
    flow: 'MANUAL',
    requestedFlowLabel: ''
  };

  const cfg = (typeof CONFIG !== 'undefined' && CONFIG) ? CONFIG : {};
  const formFields = cfg.formFields || {};

  const uploadFieldName = String(formFields.fileUploadFieldName || 'file_upload').toLowerCase().trim();
  const flowFieldName = String(formFields.flowFieldName || 'flow').toLowerCase().trim();
  const subOldFieldName = String(formFields.subOldFileUploadFieldName || '').toLowerCase().trim();
  const subNewFieldName = String(formFields.subNewFileUploadFieldName || '').toLowerCase().trim();

  const parseFlowLabel = (v) => {
    const s = String(v || '').trim().toUpperCase();
    if (!s) return '';
    if (s.indexOf('SUB') > -1) return 'SUB';
    if (s.indexOf('MAIN') > -1) return 'MAIN';
    return '';
  };

  const addIds = (arr, v) => {
    if (!v) return;
    const ids = extractFileIdsFromArray_([v]);
    if (!ids.length) return;
    for (let i = 0; i < ids.length; i++) arr.push(ids[i]);
  };

  const keyMatch = (k, fieldNameLower) => {
    if (!k || !fieldNameLower) return false;
    return String(k).toLowerCase().indexOf(fieldNameLower) > -1;
  };

  // 1) Prefer namedValues (fast, reliable)
  try {
    const named = e && e.namedValues ? e.namedValues : {};
    Object.keys(named).forEach(k => {
      const rawVal = named[k];
      const arr = Array.isArray(rawVal) ? rawVal : [rawVal];
      arr.forEach(v => {
        if (v == null || v === '') return;

        // Flow selector
        if (flowFieldName && keyMatch(k, flowFieldName) && !ctx.requestedFlowLabel) {
          const f = parseFlowLabel(v);
          if (f) ctx.requestedFlowLabel = f;
          return;
        }

        // Optional split SUB fields
        if (subOldFieldName && keyMatch(k, subOldFieldName)) { addIds(ctx.sub.oldFileIds, v); return; }
        if (subNewFieldName && keyMatch(k, subNewFieldName)) { addIds(ctx.sub.newFileIds, v); return; }

        // Primary: file upload field
        if (uploadFieldName && keyMatch(k, uploadFieldName)) { addIds(ctx.fileIds, v); return; }

        // Secondary: scan other fields for pasted Drive links/IDs
        addIds(ctx.fileIds, v);
      });
    });
  } catch (err) {}

  // Normalize + de-dupe
  ctx.fileIds = Array.from(new Set(ctx.fileIds)).filter(Boolean);
  ctx.sub.oldFileIds = Array.from(new Set(ctx.sub.oldFileIds)).filter(Boolean);
  ctx.sub.newFileIds = Array.from(new Set(ctx.sub.newFileIds)).filter(Boolean);

  if (ctx.requestedFlowLabel) ctx.flow = ctx.requestedFlowLabel;

  if (ctx.fileIds.length || ctx.sub.oldFileIds.length || ctx.sub.newFileIds.length) return ctx;

  // 2) Fallback: read response row (single read)
  try {
    let sheet, rowIndex, lastCol;

    if (e && e.range && e.range.getSheet) {
      sheet = e.range.getSheet();
      rowIndex = e.range.getRow();
      lastCol = sheet.getLastColumn();
    } else {
      const respSs = SpreadsheetApp.openById(RESPONSES_SPREADSHEET_ID);
      sheet = RESPONSES_SHEET_NAME ? respSs.getSheetByName(RESPONSES_SHEET_NAME) : respSs.getSheets()[0];
      rowIndex = sheet.getLastRow();
      lastCol = sheet.getLastColumn();
    }

    if (sheet && rowIndex && lastCol && lastCol > 0) {
      const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeaderCell_);
      const headerLower = header.map(v => cleanHeaderKey04_(v));
      const rowVals = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

      const findHeaderIdx = (pred) => {
        for (let c = 0; c < headerLower.length; c++) if (pred(headerLower[c], c)) return c;
        return -1;
      };

      const idxFlow = flowFieldName
        ? findHeaderIdx((h) => h.indexOf(flowFieldName) > -1)
        : -1;
      if (idxFlow > -1) {
        const f = parseFlowLabel(rowVals[idxFlow]);
        if (f) { ctx.requestedFlowLabel = f; ctx.flow = f; }
      }

      const idxOld = subOldFieldName
        ? findHeaderIdx((h) => h.indexOf(subOldFieldName) > -1)
        : -1;
      if (idxOld > -1) addIds(ctx.sub.oldFileIds, rowVals[idxOld]);

      const idxNew = subNewFieldName
        ? findHeaderIdx((h) => h.indexOf(subNewFieldName) > -1)
        : -1;
      if (idxNew > -1) addIds(ctx.sub.newFileIds, rowVals[idxNew]);

      // find upload column by header match (backward-compatible heuristics)
      let idxFile = -1;
      const fileField = uploadFieldName;
      idxFile = findHeaderIdx((h) => (
        (fileField && h.indexOf(fileField) > -1) ||
        h.indexOf('metabase') > -1 ||
        h.indexOf('upload') > -1 ||
        h.indexOf('file') > -1
      ));

      if (idxFile > -1) addIds(ctx.fileIds, rowVals[idxFile]);

      // brute fallback: scan values
      if (ctx.fileIds.length === 0 && ctx.sub.oldFileIds.length === 0 && ctx.sub.newFileIds.length === 0) {
        ctx.fileIds = detectFileIdsFromRowValues_(rowVals);
      }
    }
  } catch (err2) {}

  ctx.fileIds = Array.from(new Set(ctx.fileIds)).filter(Boolean);
  ctx.sub.oldFileIds = Array.from(new Set(ctx.sub.oldFileIds)).filter(Boolean);
  ctx.sub.newFileIds = Array.from(new Set(ctx.sub.newFileIds)).filter(Boolean);
  if (ctx.requestedFlowLabel) ctx.flow = ctx.requestedFlowLabel;

  return ctx;
}


/** ---------- File metadata cache (reduces Drive I/O) ---------- */
const _FILE_META_CACHE = Object.create(null);

/**
 * getFileMetaCached_(id) => { id, nameLower, name, updatedMs, mimeType } | null
 */
function getFileMetaCached_(id) {
  if (!id) return null;
  if (_FILE_META_CACHE[id] !== undefined) return _FILE_META_CACHE[id];

  try {
    const f = DriveApp.getFileById(id);
    const name = f.getName();
    const meta = {
      id,
      name,
      nameLower: String(name || '').toLowerCase(),
      updatedMs: f.getLastUpdated().getTime(),
      mimeType: (function() { try { return f.getMimeType(); } catch (e) { return ''; } })()
    };
    _FILE_META_CACHE[id] = meta;
    return meta;
  } catch (e) {
    _FILE_META_CACHE[id] = null;
    return null;
  }
}

/** ---------- Trash control (deferred support) ---------- */
/**
 * By default, this pipeline may trash uploaded source files (if PIPELINE_FLAGS.TRASH_UPLOADED_FILES=true).
 * To avoid losing inputs when downstream steps fail, callers can set:
 *   RUNTIME.deferTrashUploadedFiles = true
 * Then this module will queue fileIds into RUNTIME.filesToTrash, and the orchestrator can flush later
 * using flushTrashQueueBestEffort_().
 */
function shouldDeferTrashUploads_() {
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.deferTrashUploadedFiles === true) return true;
  } catch (e) {}
  return false;
}

function enqueueTrashFileId_(fileOrId) {
  if (!fileOrId) return;
  if (isDryRunLocal_()) return;
  if (!pipelineFlag_('TRASH_UPLOADED_FILES', false)) return;

  try {
    if (typeof RUNTIME === 'undefined' || !RUNTIME) return;
    const id = (typeof fileOrId === 'string')
      ? String(fileOrId)
      : (fileOrId.getId ? String(fileOrId.getId()) : '');
    if (!id) return;

    if (!Array.isArray(RUNTIME.filesToTrash)) RUNTIME.filesToTrash = [];
    if (RUNTIME.filesToTrash.indexOf(id) === -1) RUNTIME.filesToTrash.push(id);
  } catch (e) {
    // best effort only
  }
}

/** Trash queued uploads at end of a successful run (best effort). */
function flushTrashQueueBestEffort_(segLabel) {
  if (isDryRunLocal_()) return;
  if (!pipelineFlag_('TRASH_UPLOADED_FILES', false)) return;

  let ids = [];
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && Array.isArray(RUNTIME.filesToTrash)) {
      ids = RUNTIME.filesToTrash.slice();
      RUNTIME.filesToTrash = [];
    }
  } catch (e) {}

  if (!ids.length) return;

  const seg = startSeg04_(segLabel || 'Trash uploads', 'Trash queued uploads');
  let trashed = 0;

  for (let i = 0; i < ids.length; i++) {
    try {
      DriveApp.getFileById(ids[i]).setTrashed(true);
      trashed++;
    } catch (e) {
      // ignore
    }
  }

  endSeg04_(seg, 'count=' + ids.length + ' | trashed=' + trashed, 'Trash queued uploads', 'INFO');
}

function trashFileBestEffort_(fileOrId) {
  if (!fileOrId) return;
  if (isDryRunLocal_()) return;
  if (!pipelineFlag_('TRASH_UPLOADED_FILES', false)) return;
  try {
    if (typeof fileOrId === 'string') DriveApp.getFileById(fileOrId).setTrashed(true);
    else fileOrId.setTrashed(true);
  } catch (e) {}
}

function maybeTrashOrQueueFile_(fileOrId) {
  if (!fileOrId) return;
  if (shouldDeferTrashUploads_()) enqueueTrashFileId_(fileOrId);
  else trashFileBestEffort_(fileOrId);
}

/** ---------- Parse uploaded file ---------- */

function parseUploadedFile_(fileId, segLabel) {
  const seg = startSeg04_(segLabel, 'Parse file ' + fileId);

  try {
    const file = DriveApp.getFileById(fileId);
    const name = file.getName();
    const mime = (function() { try { return file.getMimeType(); } catch (e) { return ''; } })();
    const ext = (function() {
      const parts = String(name || '').split('.');
      return (parts.length > 1) ? String(parts.pop() || '').toLowerCase() : '';
    })();

    let result = null;

    // 1) Google Sheets file (no conversion needed)
    if (mime === MimeType.GOOGLE_SHEETS || mime === 'application/vnd.google-apps.spreadsheet') {
      result = parseGoogleSheetFile_(fileId);
    } else {
      // 2) Blob-based parsing
      const blob = file.getBlob();

      if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm' || ext === 'xltx' || ext === 'xltm') result = parseXlsx_(blob, name);
      else if (mime === 'application/vnd.ms-excel' || mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') result = parseXlsx_(blob, name);
      else if (ext === 'csv' || mime === 'text/csv') result = parseCsv_(blob);
      else if (ext === 'json' || mime === 'application/json') result = parseJson_(blob);
      else {
        // Some exports come as .txt but CSV content
        if (ext === 'txt') result = parseCsv_(blob);
        else {
          endSeg04_(seg, 'ext=' + (ext || '(none)') + ' | mime=' + mime, 'Unsupported file; skipped', 'WARN');
          return null;
        }
      }
    }

    // trash original upload (or queue for end-of-run)
    maybeTrashOrQueueFile_(file);

    endSeg04_(
      seg,
      'rows=' + (result.rows ? result.rows.length : 0) + ' | cols=' + (result.header ? result.header.length : 0),
      name,
      'INFO'
    );

    return { header: result.header, rows: result.rows, name, ext, mimeType: mime, fileId };
  } catch (err) {
    endSeg04_(seg, '', 'Error parsing file: ' + err, 'ERROR');
    throw err;
  }
}

/** ---------- Parse Google Sheets file (already converted) ---------- */
function parseGoogleSheetFile_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheets()[0];
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const values = (lr > 0 && lc > 0) ? sh.getRange(1, 1, lr, lc).getValues() : [[]];

  // Support title rows above the real header (common in XLSX->Sheets conversions).
  return extractTableFromValues04_(values);
}


/** ---------- XLSX parsing (convert to temp sheet) ---------- */
function parseXlsx_(blob, name) {
  let tempFile = null;
  try {
    // Requires Advanced Drive Service enabled: Drive API (v2)
    if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.insert) {
      throw new Error('Advanced Drive Service not enabled (Drive API v2).');
    }

    // Convert XLSX -> Google Sheets
    const resource = { title: 'TMP_' + name, mimeType: MimeType.GOOGLE_SHEETS };
    tempFile = Drive.Files.insert(resource, blob, { convert: true });

    const tempSs = SpreadsheetApp.openById(tempFile.id);
    const sh = tempSs.getSheets()[0];

    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    const values = (lr > 0 && lc > 0) ? sh.getRange(1, 1, lr, lc).getValues() : [[]];

    return extractTableFromValues04_(values);
  } catch (e) {
    throw new Error('XLSX parse failed. Enable Advanced Drive Service "Drive API (v2)". Detail: ' + e);
  } finally {
    try {
      if (tempFile && tempFile.id && !isDryRunLocal_()) DriveApp.getFileById(tempFile.id).setTrashed(true);
    } catch (e2) {}
  }
}

/** ---------- CSV parsing (delimiter auto-detect) ---------- */
function guessCsvDelimiter_(text) {
  // sample first ~5 lines
  const sample = String(text || '').split(/\r?\n/).slice(0, 5).join('\n');

  const score = d => (sample.split(d).length - 1);
  const sComma = score(',');
  const sSemi  = score(';');
  const sTab   = score('\t');

  // choose max
  if (sTab >= sSemi && sTab >= sComma && sTab > 0) return '\t';
  if (sSemi >= sComma && sSemi > 0) return ';';
  return ','; // default
}

function parseCsv_(blob) {
  const text = blob.getDataAsString();
  if (!text) return { header: [], rows: [] };

  const delim = guessCsvDelimiter_(text);
  const allRows = Utilities.parseCsv(text, delim);

  if (!allRows.length) return { header: [], rows: [] };
  return extractTableFromValues04_(allRows);
}

/** ---------- JSON parsing ---------- */
function parseJson_(blob) {
  const text = blob.getDataAsString();
  if (!text) return { header: [], rows: [] };

  const data = JSON.parse(text);
  if (!Array.isArray(data) || !data.length) return { header: [], rows: [] };

  const rawKeys = Object.keys(data[0]);
  const header = rawKeys.map(cleanHeaderCell_);
  const rows = data.map(o => rawKeys.map(k => o[k]));
  return { header, rows };
}

/** ---------- File classification (single-pass, cached meta) ---------- */
function classifyFiles_(fileIds) {
  const uniqIds = Array.from(new Set(fileIds || [])).filter(Boolean);

  const cfg = (typeof CONFIG !== 'undefined' && CONFIG) ? CONFIG : {};
  const pref = cfg.filePrefixes || {};

  const pAgingStd = String(pref.agingStd || '').toLowerCase().trim();
  const pAging    = String(pref.aging || '').toLowerCase().trim();
  const pMain     = String(pref.main || '').toLowerCase().trim();

  const bucket = { main: [], aging: [], agingStd: [], unknown: [] };

  // Heuristic classifiers (only used when prefix-based match fails).
  // Keeps behavior backward-compatible while making email ingestion more resilient to filename variations.
  const heur = {
    agingStd: [/aging\s*std/i, /aging\s*standard/i, /standard\s*aging/i, /aging[_\- ]?standard/i, /aging[_\- ]?std/i],
    aging:    [/\baging\b/i, /\btat\b/i, /\bala\b/i, /\blsa\b/i, /\bsla\b/i,
              /activity\s*log\s*aging/i, /activity[_\- ]?log[_\- ]?aging/i,
              /last\s*status\s*aging/i, /last[_\- ]?status[_\- ]?aging/i],
    main:     [/claim\s*daily\s*monitoring/i, /daily\s*claim\s*monitoring/i, /daily\s*monitoring/i, /claim\s*monitoring/i, /\bdashboard\b/i]
  };

  const classifyByHeur = nameLower => {
    const n = String(nameLower || '');
    for (let i = 0; i < heur.agingStd.length; i++) if (heur.agingStd[i].test(n)) return 'agingStd';
    for (let i = 0; i < heur.aging.length; i++) if (heur.aging[i].test(n)) return 'aging';
    for (let i = 0; i < heur.main.length; i++) if (heur.main[i].test(n)) return 'main';
    return null;
  };

  (uniqIds || []).forEach(id => {
    const meta = getFileMetaCached_(id);
    if (!meta) { bucket.unknown.push(id); return; }

    const name = meta.nameLower || '';

    // 1) Prefix-based classification (preferred)
    if (pAgingStd && name.indexOf(pAgingStd) > -1) bucket.agingStd.push(meta);
    else if (pAging && name.indexOf(pAging) > -1) bucket.aging.push(meta);
    else if (pMain && name.indexOf(pMain) > -1) bucket.main.push(meta);
    else {
      // 2) Heuristic classification
      const t = classifyByHeur(name);
      if (t === 'agingStd') bucket.agingStd.push(meta);
      else if (t === 'aging') bucket.aging.push(meta);
      else if (t === 'main') bucket.main.push(meta);
      else bucket.unknown.push(id);
    }
  });

  // Backward-compatible fallback: single unknown treated as main
  if (bucket.main.length === 0 && bucket.unknown.length === 1) {
    const onlyId = bucket.unknown[0];
    const m = getFileMetaCached_(onlyId);
    if (m) bucket.main.push(m);
    bucket.unknown = [];
  }

  const pickNewestMeta = metas => {
    if (!metas || !metas.length) return null;
    let best = metas[0];
    for (let i = 1; i < metas.length; i++) {
      if ((metas[i].updatedMs || 0) > (best.updatedMs || 0)) best = metas[i];
    }
    return best;
  };

  const mainBest = pickNewestMeta(bucket.main);
  const agingBest = pickNewestMeta(bucket.aging);
  const agingStdBest = pickNewestMeta(bucket.agingStd);

  // Debug summary (helps diagnose why Activity Log Aging/TAT might remain blank due to missing aging files).
  debugLog04_('[04] classifyFiles_: main=' + (mainBest ? mainBest.name : '-') +
             ', aging=' + (agingBest ? agingBest.name : '-') +
             ', agingStd=' + (agingStdBest ? agingStdBest.name : '-') +
             ', unknown=' + (bucket.unknown.length || 0));

  return {
    main: mainBest ? [mainBest.id] : [],
    aging: agingBest ? [agingBest.id] : [],
    agingStd: agingStdBest ? [agingStdBest.id] : [],
    unknown: bucket.unknown
  };
}


/** Backward-compatible helper retained (some callers may still call it) */
function pickNewestFileId_(ids) {
  if (!ids || !ids.length) return null;
  const metas = ids.map(id => getFileMetaCached_(id)).filter(Boolean);
  if (!metas.length) return ids[0] || null;
  let best = metas[0];
  for (let i = 1; i < metas.length; i++) {
    if ((metas[i].updatedMs || 0) > (best.updatedMs || 0)) best = metas[i];
  }
  return best.id;
}

/** ---------- Aging map ---------- */
function buildAgingMap_(agingData, claimHeaderName) {
  if (!agingData || !agingData.header || !agingData.rows) return null;

  // Normalize headers to reduce mismatch risk across XLSX/CSV variants.
  const header = (agingData.header || []).map(cleanHeaderCell_);
  const hMap = buildHeaderIndex_(header);

  let idxClaim = hMap[claimHeaderName];

  // Fallback: case/whitespace-insensitive match (when claimHeaderName is human-readable).
  if (idxClaim == null) {
    const want = cleanHeaderKey04_(claimHeaderName || '');
    for (let i = 0; i < header.length; i++) {
      if (cleanHeaderKey04_(header[i] || '') === want) { idxClaim = i; break; }
    }
  }

  if (idxClaim == null) return null;

  const map = Object.create(null);
  for (let i = 0; i < agingData.rows.length; i++) {
    const r = agingData.rows[i];
    const key = String(r[idxClaim] || '').trim().toUpperCase();
    if (key) map[key] = r;
  }
  return { header, headerIndex: hMap, map };
}


/** ---------- Typed coercion hook (used by later parts of 04/05) ---------- */
/**
 * Normalize RAW type registry to support header renames (LSA/ALA -> Last Status Aging / Activity Log Aging).
 * This keeps coercion stable even when the upstream file uses the newer header labels.
 */
function normalizeTypeRegistryAliases04_(typeRegistry) {
  if (!typeRegistry) return typeRegistry;
  const reg = {};
  // shallow copy
  Object.keys(typeRegistry).forEach(k => { reg[k] = typeRegistry[k]; });

  // Forward aliases
  if (reg['LSA'] && !reg['Last Status Aging']) reg['Last Status Aging'] = reg['LSA'];
  if (reg['LSA'] && !reg['last_status_aging']) reg['last_status_aging'] = reg['LSA'];

  if (reg['ALA'] && !reg['Activity Log Aging']) reg['Activity Log Aging'] = reg['ALA'];
  if (reg['ALA'] && !reg['activity_log_aging']) reg['activity_log_aging'] = reg['ALA'];

  // Reverse aliases (if config updated but older files still use LSA/ALA)
  if (reg['Last Status Aging'] && !reg['LSA']) reg['LSA'] = reg['Last Status Aging'];
  if (reg['Activity Log Aging'] && !reg['ALA']) reg['ALA'] = reg['Activity Log Aging'];

  return reg;
}

/**
 * Coerce parsed dataset using RAW schema types if available.
 * - Only touches columns that match known RAW headers.
 * - Prevents "money becomes date" by coercing MONEY0 as Number and DATE as Date.
 */
function coerceDatasetToRawSchema_(dataset) {
  if (!dataset || !dataset.header || !dataset.rows) return dataset;

  // prefer your global registry (00_Config)
  const typeRegistry0 = (typeof COLUMN_TYPES !== 'undefined' && COLUMN_TYPES && COLUMN_TYPES.RAW) ? COLUMN_TYPES.RAW : null;
  if (!typeRegistry0) return dataset;
  const typeRegistry = normalizeTypeRegistryAliases04_(typeRegistry0);

  const header = dataset.header.map(cleanHeaderCell_);
  const coercedRows = new Array(dataset.rows.length);

  for (let r = 0; r < dataset.rows.length; r++) {
    coercedRows[r] = coerceRowByTypes_(header, dataset.rows[r], typeRegistry);
  }

  const out = { header, rows: coercedRows, name: dataset.name, ext: dataset.ext, mimeType: dataset.mimeType, fileId: dataset.fileId };
  return coerceKnownDatetimeColumns04_(out);
}

/** ---------- SUB dashboard attachment helpers (NEW/OLD) ---------- */
/**
 * SUB flow attachment typing rule (case-insensitive):
 * - NEW: filename contains "(Standardization)" (or just "standardization")
 * - OLD: filename contains "List of Claims with Aging" and does NOT contain "standardization"
 * Returns: 'NEW' | 'OLD' | null
 */
function detectSubDashboardAttachmentType04_(filename) {
  const n = String(filename || '').toLowerCase();
  if (!n) return null;

  // NEW first (priority)
  if (n.indexOf('standardization') > -1) return 'NEW';

  // OLD only if it's the classic "List of Claims with Aging" export
  if (n.indexOf('list of claims with aging') > -1) return 'OLD';

  return null;
}

/**
 * Pick the best attachment for SUB run.
 * Input: array of Gmail attachments (Blob) or meta objects with .getName()/.name
 * Output: { type: 'NEW'|'OLD', attachment: <original> } | null
 */
function pickSubDashboardAttachment04_(attachments) {
  const list = Array.isArray(attachments) ? attachments : [];
  if (!list.length) return null;

  const scored = [];
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    const name = (a && typeof a.getName === 'function') ? a.getName() : (a && a.name ? String(a.name) : '');
    const type = detectSubDashboardAttachmentType04_(name);
    if (!type) continue;
    const pri = (type === 'NEW') ? 0 : 1;
    scored.push({ pri, type, name: String(name || ''), attachment: a });
  }

  if (!scored.length) return null;

  scored.sort((x, y) => {
    if (x.pri !== y.pri) return x.pri - y.pri;
    // deterministic tie-breaker: lexicographic by filename
    return x.name.localeCompare(y.name);
  });

  return { type: scored[0].type, attachment: scored[0].attachment };
}


/** ---------- Known datetime coercion (SUB support) ---------- */
/**
 * Some Metabase exports provide datetime fields as English month strings, e.g.:
 *   "January 24, 2025, 6:06 PM"
 * Downstream flows require these to be true Date objects so the sheet can format:
 * - SUB: dd MMM yy, HH:mm
 * - MAIN/FORM: dd MMM yy
 *
 * This module does NOT assume Raw OLD/Raw NEW are snapshots; it only normalizes types.
 */

function parseEnglishMonthDatetime04_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  if (!s) return null;

  const months = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  // Examples handled:
  // - January 24, 2025, 6:06 PM
  // - Jan 24, 2025 6:06 PM
  // - January 24 2025, 18:06
  const re = /^\s*([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})(?:\s*,?\s*(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2}))?\s*(AM|PM)?)?\s*$/i;
  const m = s.match(re);
  if (!m) return null;

  const monKey = String(m[1] || '').toLowerCase();
  const monthIdx = (months.hasOwnProperty(monKey) ? months[monKey] : null);
  if (monthIdx == null) return null;

  const day = Number(m[2]);
  const year = Number(m[3]);

  let hh = (m[4] != null && m[4] !== '') ? Number(m[4]) : 0;
  const mm = (m[5] != null && m[5] !== '') ? Number(m[5]) : 0;
  const ss = (m[6] != null && m[6] !== '') ? Number(m[6]) : 0;
  const ampm = String(m[7] || '').toUpperCase().trim();

  if (ampm === 'PM' && hh < 12) hh += 12;
  if (ampm === 'AM' && hh === 12) hh = 0;

  const d = new Date(year, monthIdx, day, hh, mm, ss);
  if (isNaN(d.getTime())) return null;
  return d;
}

function parseYmdDatetime04_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  if (!s) return null;

  // 2025-01-24 18:06(:ss)? or 2025/01/24 18:06
  const re = /^\s*(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$/;
  const m = s.match(re);
  if (!m) return null;

  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const da = Number(m[3]);
  const hh = (m[4] != null) ? Number(m[4]) : 0;
  const mi = (m[5] != null) ? Number(m[5]) : 0;
  const se = (m[6] != null) ? Number(m[6]) : 0;

  const d = new Date(y, mo, da, hh, mi, se);
  if (isNaN(d.getTime())) return null;
  return d;
}

function coerceDateValue04_(v) {
  if (v == null || v === '') return v;

  // Date already
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;

  // Excel/Sheets serial (best effort)
  if (typeof v === 'number' && isFinite(v)) {
    // Excel epoch: 1899-12-30 (Google Sheets compatible)
    if (v > 20000 && v < 80000) {
      const ms = Math.round((v - 25569) * 86400 * 1000);
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d;
    }
    return v;
  }

  const s = String(v).trim();
  if (!s) return v;

  const d1 = parseEnglishMonthDatetime04_(s);
  if (d1) return d1;

  const d2 = parseYmdDatetime04_(s);
  if (d2) return d2;

  if (typeof normalizeDate_ === 'function') {
    try {
      const d3 = normalizeDate_(s);
      if (d3 && !isNaN(d3.getTime())) return d3;
    } catch (e3) {}
  }
  if (typeof tryNativeParseUnambiguousDate_ === 'function') {
    try {
      const d4 = tryNativeParseUnambiguousDate_(s);
      if (d4 && !isNaN(d4.getTime())) return d4;
    } catch (e4) {}
  }

  return v;
}

function coerceKnownDatetimeColumns04_(dataset) {
  if (!dataset || !dataset.header || !dataset.rows) return dataset;

  const header = dataset.header || [];
  const keys = header.map(cleanHeaderKey04_);

  const findAnyIdx = (candidates) => {
    const want = (candidates || []).map(c => cleanHeaderKey04_(c));
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (!k) continue;
      for (let j = 0; j < want.length; j++) if (k === want[j]) return i;
    }
    return -1;
  };

  // Primary datetime fields used by SUB flow
  const idxClaimLastUpdated = findAnyIdx([
    'claim_last_updated_datetime',
    'claim last updated datetime',
    'claim last updated date time',
    'claim_last_updated'
  ]);

  const idxActivityLogDt = findAnyIdx([
    'activity_log_datetime',
    'activity log datetime',
    'activity log date time',
    'activity_log_date_time',
    'activity_log_updated_datetime'
  ]);

  const idxLastActivityLogDt = findAnyIdx([
    'last_activity_log_datetime',
    'last activity log datetime',
    'last activity log date time',
    'last_activity_log_date_time',
    'last_activity_datetime',
    'last activity datetime'
  ]);

  if (idxClaimLastUpdated < 0 && idxActivityLogDt < 0 && idxLastActivityLogDt < 0) return dataset;

  const rows = dataset.rows;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;

    if (idxClaimLastUpdated > -1 && row.length > idxClaimLastUpdated) row[idxClaimLastUpdated] = coerceDateValue04_(row[idxClaimLastUpdated]);
    if (idxActivityLogDt > -1 && row.length > idxActivityLogDt) row[idxActivityLogDt] = coerceDateValue04_(row[idxActivityLogDt]);
    if (idxLastActivityLogDt > -1 && row.length > idxLastActivityLogDt) row[idxLastActivityLogDt] = coerceDateValue04_(row[idxLastActivityLogDt]);
  }

  return dataset;
}


/** ---------- SUB attachments: pick OLD + NEW in one pass (shared helper) ---------- */
/**
 * In this project, "OLD" and "NEW" are two different Metabase datasets (not snapshots).
 * - NEW: export contains "standardization" in filename
 * - OLD: export contains "List of Claims with Aging" and does NOT contain "standardization"
 */
function pickSubDashboardAttachments04_(attachments) {
  const list = Array.isArray(attachments) ? attachments : [];
  let oldBest = null;
  let oldName = '';
  let newBest = null;
  let newName = '';

  const getName = (a) => (a && typeof a.getName === 'function') ? String(a.getName() || '') : (a && a.name ? String(a.name || '') : '');

  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    const name = getName(a);
    if (!name) continue;

    const t = detectSubDashboardAttachmentType04_(name);
    if (!t) continue;

    if (t === 'NEW') {
      if (!newBest) { newBest = a; newName = name; continue; }
      if (String(name).localeCompare(String(newName)) < 0) { newBest = a; newName = name; }
      continue;
    }

    if (t === 'OLD') {
      if (!oldBest) { oldBest = a; oldName = name; continue; }
      if (String(name).localeCompare(String(oldName)) < 0) { oldBest = a; oldName = name; }
      continue;
    }
  }

  return { oldAttachment: oldBest, newAttachment: newBest, oldAtt: oldBest, newAtt: newBest };
}
