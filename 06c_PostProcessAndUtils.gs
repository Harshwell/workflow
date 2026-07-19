/***************************************************************
 * 06c_PostProcessAndUtils.gs  (SPLIT FROM 06.gs)
 * Post-route utilities, carry-forward, preflight, schema formats,
 * exclusion recompute, and Raw Data column reordering
 ***************************************************************/
'use strict';

function __claimKey06_(v) {
  let s = String(v == null ? '' : v).trim();
  if (!s) return '';
  // Normalize numeric-looking claim ids (e.g. 12345.0 from Sheets numeric coercion)
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
  return s.toUpperCase();
}

function __groupConsecutiveRows_(nums) {
  if (!nums || !nums.length) return [];
  const a = nums.slice().sort((x, y) => x - y);
  const segments = [];
  let seg = [a[0]];
  for (let i = 1; i < a.length; i++) {
    if (a[i] === a[i - 1] + 1) seg.push(a[i]);
    else { segments.push(seg); seg = [a[i]]; }
  }
  segments.push(seg);
  return segments;
}

/**
 * Apply Update Status rich text from Raw Data into operational sheets.
 * - Does NOT touch optional sheets.
 * - Only writes cells where Raw has non-empty rich text for that claim.
 * This preserves formatting for Update Status (bold/italic/etc).
 */
function applyUpdateStatusRichTextToOperational_(ss, rawSheet, headerIndexRaw, pic) {
  if (DRY_RUN) return;
  if (!ss || !rawSheet || !headerIndexRaw) return;

  const h = CONFIG.headers;
  const idxClaimRaw = headerIndexRaw[h.claimNumber];
  const idxUpdateRaw = headerIndexRaw[h.updateStatus];
  if (idxClaimRaw == null || idxUpdateRaw == null) return;

  const n = rawSheet.getLastRow() - 1;
  if (n <= 0) return;

  const claimVals = rawSheet.getRange(2, idxClaimRaw + 1, n, 1).getValues();
  const rtVals = rawSheet.getRange(2, idxUpdateRaw + 1, n, 1).getRichTextValues();

  const map = Object.create(null);
  for (let i = 0; i < n; i++) {
    const c = __claimKey06_(claimVals[i] && claimVals[i][0]);
    if (!c) continue;
    const rt = (rtVals[i] && rtVals[i][0]) ? rtVals[i][0] : null;
    if (!rt || !rt.getText) continue;
    const t = String(rt.getText() || '');
    if (!t) continue; // only apply when there is text (rich formatting matters)
    map[c] = rt;
  }

  const isAdmin = (pic === 'Admin');
  const ops = isAdmin ? (CONFIG.sheetsByPic && CONFIG.sheetsByPic.adminOperational) : (CONFIG.sheetsByPic && CONFIG.sheetsByPic.picOperational);
  if (!ops || !ops.length) return;

  for (let s = 0; s < ops.length; s++) {
    const name = ops[s];
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    const idxUpdate = __findHeaderIndexFlexible06_(header, 'Update Status');
    if (idxClaim === -1 || idxUpdate === -1) continue;

    const rows = lastRow - 1;
    const claims = sh.getRange(2, idxClaim + 1, rows, 1).getValues();

    const rowNums = [];
    const rowMap = Object.create(null);

    for (let i = 0; i < rows; i++) {
      const claim = __claimKey06_(claims[i] && claims[i][0]);
      if (!claim) continue;
      const rt = map[claim];
      if (!rt) continue;
      const rno = i + 2;
      rowNums.push(rno);
      rowMap[rno] = rt;
    }

    if (!rowNums.length) continue;

    const segments = __groupConsecutiveRows_(rowNums);
    for (let k = 0; k < segments.length; k++) {
      const seg = segments[k];
      const startRow = seg[0];
      const rich2d = seg.map(rn => [rowMap[rn]]);
      sh.getRange(startRow, idxUpdate + 1, rich2d.length, 1).setRichTextValues(rich2d);
    }
  }
}

/**
 * Apply Remarks rich text from Raw Data into operational sheets.
 * - Writes cells where Raw has non-empty rich text for that claim.
 */
function applyRemarksRichTextToOperational_(ss, rawSheet, headerIndexRaw, pic) {
  if (DRY_RUN) return;
  if (!ss || !rawSheet || !headerIndexRaw) return;

  const idxClaimRaw = headerIndexRaw[(CONFIG && CONFIG.headers && CONFIG.headers.claimNumber) ? CONFIG.headers.claimNumber : 'claim_number'];
  const idxRemarksRaw = (typeof idxAny_ === 'function')
    ? idxAny_(headerIndexRaw, ['Remarks', 'Remark', 'remarks', 'remark'])
    : (headerIndexRaw['Remarks'] != null ? headerIndexRaw['Remarks'] : (headerIndexRaw['remarks'] != null ? headerIndexRaw['remarks'] : null));

  if (idxClaimRaw == null || idxRemarksRaw == null) return;

  const n = rawSheet.getLastRow() - 1;
  if (n <= 0) return;

  const claimVals = rawSheet.getRange(2, idxClaimRaw + 1, n, 1).getValues();
  const rtVals = rawSheet.getRange(2, idxRemarksRaw + 1, n, 1).getRichTextValues();

  const map = Object.create(null);
  for (let i = 0; i < n; i++) {
    const c = __claimKey06_(claimVals[i] && claimVals[i][0]);
    if (!c) continue;
    const rt = (rtVals[i] && rtVals[i][0]) ? rtVals[i][0] : null;
    if (!rt || !rt.getText) continue;
    const t = String(rt.getText() || '');
    if (t === '') continue;
    map[c] = rt;
  }

  const isAdmin = (pic === 'Admin');
  const ops = isAdmin ? (CONFIG.sheetsByPic && CONFIG.sheetsByPic.adminOperational) : (CONFIG.sheetsByPic && CONFIG.sheetsByPic.picOperational);
  if (!ops || !ops.length) return;

  for (let s = 0; s < ops.length; s++) {
    const name = ops[s];
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    const idxRemarks = __findHeaderIndexFlexible06_(header, 'Remarks');
    const idxAwb = __findHeaderIndexFlexible06_(header, 'AWB');
    const idxTimestampAwb = __findHeaderIndexFlexible06_(header, 'Timestamp AWB');
    if (idxClaim === -1 || idxRemarks === -1) continue;

    const rows = lastRow - 1;
    const claims = sh.getRange(2, idxClaim + 1, rows, 1).getValues();

    const rowNums = [];
    const rowMap = Object.create(null);

    for (let i = 0; i < rows; i++) {
      const claim = __claimKey06_(claims[i] && claims[i][0]);
      if (!claim) continue;
      const rt = map[claim];
      if (!rt) continue;
      const rno = i + 2;
      rowNums.push(rno);
      rowMap[rno] = rt;
    }

    if (!rowNums.length) continue;

    const segments = __groupConsecutiveRows_(rowNums);
    for (let k = 0; k < segments.length; k++) {
      const seg = segments[k];
      const startRow = seg[0];
      const rich2d = seg.map(rn => [rowMap[rn]]);
      sh.getRange(startRow, idxRemarks + 1, rich2d.length, 1).setRichTextValues(rich2d);
    }
  }
}

/**
 * Snapshot Operational columns that are user-managed and must survive CLR + ROUTE:
 * - Update Status (RichText + wrap)
 * - Timestamp (value + number format)
 * - Status (value)
 * - Remarks (RichText + wrap)
 *
 * Note: this is an in-memory safety net to preserve per-cell formatting (wrap) that is not representable in Raw.
 * It complements (not replaces) ops→raw backup.
 */
function snapshotOpsManualColumnsRich06c_(ss, pic) {
  if (!ss) return { version: 'ops_manual_v1', map: Object.create(null), count: 0 };

  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  const out = Object.create(null);
  let count = 0;

  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;

    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    if (idxClaim === -1) continue;

    const isEvBike = String(sh.getName() || '').trim().toLowerCase() === 'ev-bike';

    const idxUpdate = __findHeaderIndexFlexible06_(header, 'Update Status');
    const idxTs = __findHeaderIndexFlexible06_(header, 'Timestamp');
    const idxStatus = isEvBike ? -1 : __findHeaderIndexFlexible06_(header, 'Status');
    const idxRemarks = __findHeaderIndexFlexible06_(header, 'Remarks');
    const idxAwb = __findHeaderIndexFlexible06_(header, 'AWB');
    const idxTimestampAwb = __findHeaderIndexFlexible06_(header, 'Timestamp AWB');

    if (idxUpdate === -1 && idxTs === -1 && idxStatus === -1 && idxRemarks === -1 && idxAwb === -1 && idxTimestampAwb === -1) continue;

    const n = lr - 1;
    const claims = sh.getRange(2, idxClaim + 1, n, 1).getValues();

    const rngUpdate = (idxUpdate !== -1) ? sh.getRange(2, idxUpdate + 1, n, 1) : null;
    const rngTs = (idxTs !== -1) ? sh.getRange(2, idxTs + 1, n, 1) : null;
    const rngStatus = (idxStatus !== -1) ? sh.getRange(2, idxStatus + 1, n, 1) : null;
    const rngRemarks = (idxRemarks !== -1) ? sh.getRange(2, idxRemarks + 1, n, 1) : null;
    const rngAwb = (idxAwb !== -1) ? sh.getRange(2, idxAwb + 1, n, 1) : null;
    const rngTimestampAwb = (idxTimestampAwb !== -1) ? sh.getRange(2, idxTimestampAwb + 1, n, 1) : null;

    const updRT = rngUpdate ? rngUpdate.getRichTextValues() : null;
    const updWrap = rngUpdate ? rngUpdate.getWrapStrategies() : null;
    const updFormula = rngUpdate ? rngUpdate.getFormulas() : null;

    const tsVals = rngTs ? rngTs.getValues() : null;
    const tsFmt = rngTs ? rngTs.getNumberFormats() : null;
    const tsFormula = rngTs ? rngTs.getFormulas() : null;

    const stVals = rngStatus ? rngStatus.getValues() : null;
    const stFormula = rngStatus ? rngStatus.getFormulas() : null;

    const remRT = rngRemarks ? rngRemarks.getRichTextValues() : null;
    const remWrap = rngRemarks ? rngRemarks.getWrapStrategies() : null;
    const remFormula = rngRemarks ? rngRemarks.getFormulas() : null;
    const awbVals = rngAwb ? rngAwb.getValues() : null;
    const awbFormula = rngAwb ? rngAwb.getFormulas() : null;
    const tsAwbVals = rngTimestampAwb ? rngTimestampAwb.getValues() : null;
    const tsAwbFmt = rngTimestampAwb ? rngTimestampAwb.getNumberFormats() : null;
    const tsAwbFormula = rngTimestampAwb ? rngTimestampAwb.getFormulas() : null;

    for (let i = 0; i < n; i++) {
      const claim = __claimKey06_(claims[i] && claims[i][0]);
      if (!claim) continue;
      const key = claim.toUpperCase();

      // Skip overwrite if already captured from a higher-priority sheet; first-win keeps prior user context.
      if (out[key]) continue;

      const rec = {};
      let any = false;

      if (updRT) {
        const rt = (updRT[i] && updRT[i][0]) ? updRT[i][0] : null;
        const txt = (rt && rt.getText) ? String(rt.getText() || '') : '';
        if (txt !== '') {
          rec.u = { rt: rt, wrap: (updWrap && updWrap[i] && updWrap[i][0]) ? updWrap[i][0] : null, formula: (updFormula && updFormula[i]) ? updFormula[i][0] : '' };
          any = true;
        }
      }

      if (tsVals) {
        const v = tsVals[i] ? tsVals[i][0] : '';
        if (v !== '' && v != null) {
          rec.t = { v: v, fmt: (tsFmt && tsFmt[i] && tsFmt[i][0]) ? tsFmt[i][0] : null, formula: (tsFormula && tsFormula[i]) ? tsFormula[i][0] : '' };
          any = true;
        }
      }

      if (stVals) {
        const v = stVals[i] ? stVals[i][0] : '';
        if (v !== '' && v != null) {
          rec.s = { v: v, formula: (stFormula && stFormula[i]) ? stFormula[i][0] : '' };
          any = true;
        }
      }

      if (remRT) {
        const rt = (remRT[i] && remRT[i][0]) ? remRT[i][0] : null;
        const txt = (rt && rt.getText) ? String(rt.getText() || '') : '';
        if (txt !== '') {
          rec.r = { rt: rt, wrap: (remWrap && remWrap[i] && remWrap[i][0]) ? remWrap[i][0] : null, formula: (remFormula && remFormula[i]) ? remFormula[i][0] : '' };
          any = true;
        }
      }

      if (awbVals) {
        const v = awbVals[i] ? awbVals[i][0] : '';
        const f = (awbFormula && awbFormula[i]) ? awbFormula[i][0] : '';
        if (v !== '' || f) { rec.a = { v: v, formula: f }; any = true; }
      }
      if (tsAwbVals) {
        const v = tsAwbVals[i] ? tsAwbVals[i][0] : '';
        const f = (tsAwbFormula && tsAwbFormula[i]) ? tsAwbFormula[i][0] : '';
        if (v !== '' || f) { rec.ta = { v: v, fmt: (tsAwbFmt && tsAwbFmt[i]) ? tsAwbFmt[i][0] : null, formula: f }; any = true; }
      }

      if (any) {
        out[key] = rec;
        count++;
      }
    }
  }

  return { version: 'ops_manual_v1', map: out, count: count };
}


function auditOpsManualRestore06c_(ss, pic, snapshot) {
  if (!ss || !snapshot || !snapshot.map) return { total: 0, missing: 0 };
  const snap = snapshot.map;
  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  let total = 0;
  let missing = 0;

  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    if (idxClaim === -1) continue;

    const idxUpdate = __findHeaderIndexFlexible06_(header, 'Update Status');
    const idxTs = __findHeaderIndexFlexible06_(header, 'Timestamp');
    const idxStatus = __findHeaderIndexFlexible06_(header, 'Status');
    const idxRemarks = __findHeaderIndexFlexible06_(header, 'Remarks');
    if (idxUpdate === -1 && idxTs === -1 && idxStatus === -1 && idxRemarks === -1) continue;

    const n = lr - 1;
    const claims = sh.getRange(2, idxClaim + 1, n, 1).getValues();
    const ups = (idxUpdate !== -1) ? sh.getRange(2, idxUpdate + 1, n, 1).getDisplayValues() : null;
    const tss = (idxTs !== -1) ? sh.getRange(2, idxTs + 1, n, 1).getDisplayValues() : null;
    const sts = (idxStatus !== -1) ? sh.getRange(2, idxStatus + 1, n, 1).getDisplayValues() : null;
    const rms = (idxRemarks !== -1) ? sh.getRange(2, idxRemarks + 1, n, 1).getDisplayValues() : null;

    for (let i = 0; i < n; i++) {
      const key = __claimKey06_(claims[i] && claims[i][0]);
      if (!key) continue;
      const rec = snap[key];
      if (!rec) continue;
      total++;

      let ok = true;
      if (rec.u && idxUpdate !== -1) ok = ok && String((ups[i] && ups[i][0]) || '').trim() !== '';
      if (rec.t && idxTs !== -1) ok = ok && String((tss[i] && tss[i][0]) || '').trim() !== '';
      if (rec.s && idxStatus !== -1) ok = ok && String((sts[i] && sts[i][0]) || '').trim() !== '';
      if (rec.r && idxRemarks !== -1) ok = ok && String((rms[i] && rms[i][0]) || '').trim() !== '';
      if (!ok) missing++;
    }
  }

  try {
    if (missing > 0) logLine_('RESTORE_AUDIT', 'Manual restore gap detected', 'missing=' + missing, 'total=' + total, 'WARN');
    else logLine_('RESTORE_AUDIT', 'Manual restore audit ok', 'total=' + total, '', 'INFO');
  } catch (e) {}

  return { total: total, missing: missing };
}


function persistOpsManualBackupSheet06c_(ss, pic, snapshot) {
  if (!ss || !snapshot || !snapshot.map) return 0;
  const name = '_OPS_MANUAL_BACKUP';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  try { sh.hideSheet(); } catch (e0) {}

  const header = ['Backup Timestamp','PIC','Claim Number','Update Status','Timestamp','Status','Remarks'];
  const rows = [header];
  const now = new Date();
  const snap = snapshot.map;
  const keys = Object.keys(snap);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const rec = snap[k] || {};
    const upd = (rec.u && rec.u.rt && rec.u.rt.getText) ? String(rec.u.rt.getText() || '') : '';
    const ts = (rec.t && rec.t.v != null) ? rec.t.v : '';
    const st = (rec.s && rec.s.v != null) ? rec.s.v : '';
    const rem = (rec.r && rec.r.rt && rec.r.rt.getText) ? String(rec.r.rt.getText() || '') : '';
    rows.push([now, String(pic || ''), k, upd, ts, st, rem]);
  }

  sh.clearContents();
  sh.getRange(1,1,rows.length,header.length).setValues(rows);
  try { sh.getRange(1,1,1,header.length).setFontWeight('bold'); } catch (e1) {}
  return Math.max(0, rows.length - 1);
}

/**
 * MAIN -> SUB handoff backup (daily one-shot).
 * Snapshot only 6 columns from operational sheets into hidden temp sheet:
 * - Claim Number, Service Center, Update Status, Timestamp, Status, Remarks
 */
function persistOpsManualTempForSub06c_(ss, pic) {
  if (!ss) return { rows: 0, sheet: '' };
  const name = '_OPS_MAIN_SUB_TEMP';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  try { sh.hideSheet(); } catch (e0) {}

  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  const header = ['Backup Timestamp','PIC','Source Sheet','Claim Number','Service Center','Update Status','Timestamp','Status','Remarks'];
  const rows = [header];
  const now = new Date();
  const copyJobs = [];

  for (let si = 0; si < sheetNames.length; si++) {
    const srcName = sheetNames[si];
    const src = ss.getSheetByName(srcName);
    if (!src) continue;
    const lr = src.getLastRow();
    const lc = src.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const hdr = src.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(hdr, 'Claim Number');
    if (idxClaim === -1) continue;
    const idxSc = (function() {
      const cands = ['Service Center', 'Service Center Name', 'SC Name', 'sc_name'];
      for (let i = 0; i < cands.length; i++) {
        const x = __findHeaderIndexFlexible06_(hdr, cands[i]);
        if (x !== -1) return x;
      }
      return -1;
    })();
    const idxUpd = __findHeaderIndexFlexible06_(hdr, 'Update Status');
    const idxTs = __findHeaderIndexFlexible06_(hdr, 'Timestamp');
    const idxSt = __findHeaderIndexFlexible06_(hdr, 'Status');
    const idxRem = __findHeaderIndexFlexible06_(hdr, 'Remarks');

    if (idxUpd === -1 && idxTs === -1 && idxSt === -1 && idxRem === -1) continue;

    // Keep a row-for-row block in the temp sheet. This permits four bulk copyTo calls
    // per source sheet instead of four calls per claim, which avoids MAIN timeouts.
    const vals = src.getRange(2, 1, lr - 1, lc).getValues();
    const tempStartRow = rows.length + 1;
    for (let r = 0; r < vals.length; r++) {
      const row = vals[r] || [];
      rows.push([
        now,
        String(pic || ''),
        srcName,
        String(row[idxClaim] || '').trim(),
        idxSc !== -1 ? row[idxSc] : '',
        idxUpd !== -1 ? row[idxUpd] : '',
        idxTs !== -1 ? row[idxTs] : '',
        idxSt !== -1 ? row[idxSt] : '',
        idxRem !== -1 ? row[idxRem] : ''
      ]);
    }
    copyJobs.push({ src: src, srcStartRow: 2, tempStartRow: tempStartRow, rows: vals.length,
      idxUpd: idxUpd, idxTs: idxTs, idxSt: idxSt, idxRem: idxRem });
  }

  sh.clearContents();
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  try { sh.getRange(1, 1, 1, header.length).setFontWeight('bold'); } catch (e1) {}

  // Preserve styles in bulk (rich text, colors, number format, wrap, DV).
  copyJobs.forEach(function(j) {
    try { if (j.idxUpd !== -1) j.src.getRange(j.srcStartRow, j.idxUpd + 1, j.rows, 1).copyTo(sh.getRange(j.tempStartRow, 6, j.rows, 1), { contentsOnly: false }); } catch (eU) {}
    try { if (j.idxTs !== -1) j.src.getRange(j.srcStartRow, j.idxTs + 1, j.rows, 1).copyTo(sh.getRange(j.tempStartRow, 7, j.rows, 1), { contentsOnly: false }); } catch (eT) {}
    try { if (j.idxSt !== -1) j.src.getRange(j.srcStartRow, j.idxSt + 1, j.rows, 1).copyTo(sh.getRange(j.tempStartRow, 8, j.rows, 1), { contentsOnly: false }); } catch (eS) {}
    try { if (j.idxRem !== -1) j.src.getRange(j.srcStartRow, j.idxRem + 1, j.rows, 1).copyTo(sh.getRange(j.tempStartRow, 9, j.rows, 1), { contentsOnly: false }); } catch (eR) {}
  });

  return { rows: Math.max(0, rows.length - 1), sheet: name };
}

/** SUB one-shot restore from MAIN temp backup. Match key: Claim Number + Service Center. */
function restoreOpsManualFromMainTempForSub06c_(ss, pic, opts) {
  opts = opts || {};
  if (DRY_RUN || !ss) return { restored: 0, rows: 0, skipped: true, reason: 'dry_or_no_ss' };
  const shBak = ss.getSheetByName('_OPS_MAIN_SUB_TEMP');
  if (!shBak || shBak.getLastRow() < 2) return { restored: 0, rows: 0, skipped: true, reason: 'temp_sheet_not_found' };

  const lc = shBak.getLastColumn();
  const hdr = shBak.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  const idxClaim = __findHeaderIndexFlexible06_(hdr, 'Claim Number');
  const idxSc = __findHeaderIndexFlexible06_(hdr, 'Service Center');
  const idxUpd = __findHeaderIndexFlexible06_(hdr, 'Update Status');
  const idxTs = __findHeaderIndexFlexible06_(hdr, 'Timestamp');
  const idxSt = __findHeaderIndexFlexible06_(hdr, 'Status');
  const idxRem = __findHeaderIndexFlexible06_(hdr, 'Remarks');
  if (idxClaim === -1 || idxSc === -1) return { restored: 0, rows: 0, skipped: true, reason: 'invalid_temp_header' };

  const vals = shBak.getRange(2, 1, shBak.getLastRow() - 1, lc).getValues();
  const map = Object.create(null);
  const keyOf = function(claim, sc) { return __claimKey06_(claim) + '|' + String(sc == null ? '' : sc).trim().toUpperCase(); };
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i] || [];
    const claim = row[idxClaim];
    const sc = row[idxSc];
    const ck = __claimKey06_(claim);
    if (!ck) continue;
    map[keyOf(claim, sc)] = {
      rowNo: i + 2,
      u: idxUpd !== -1 ? row[idxUpd] : '',
      t: idxTs !== -1 ? row[idxTs] : '',
      s: idxSt !== -1 ? row[idxSt] : '',
      r: idxRem !== -1 ? row[idxRem] : ''
    };
  }

  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  let restored = 0;
  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;
    const lr = sh.getLastRow(); const lc2 = sh.getLastColumn();
    if (lr < 2 || lc2 < 1) continue;
    const h = sh.getRange(1, 1, 1, lc2).getValues()[0].map(__normalizeHeaderText06_);
    const iC = __findHeaderIndexFlexible06_(h, 'Claim Number');
    if (iC === -1) continue;
    const iSc = (function() {
      const cands = ['Service Center', 'Service Center Name', 'SC Name', 'sc_name'];
      for (let i = 0; i < cands.length; i++) {
        const x = __findHeaderIndexFlexible06_(h, cands[i]);
        if (x !== -1) return x;
      }
      return -1;
    })();
    if (iSc === -1) continue;
    const iU = __findHeaderIndexFlexible06_(h, 'Update Status');
    const iT = __findHeaderIndexFlexible06_(h, 'Timestamp');
    const iS = __findHeaderIndexFlexible06_(h, 'Status');
    const iR = __findHeaderIndexFlexible06_(h, 'Remarks');
    if (iU === -1 && iT === -1 && iS === -1 && iR === -1) continue;

    const n = lr - 1;
    const claims = sh.getRange(2, iC + 1, n, 1).getValues();
    const scs = sh.getRange(2, iSc + 1, n, 1).getValues();
    const outU = iU !== -1 ? sh.getRange(2, iU + 1, n, 1).getValues() : null;
    const outT = iT !== -1 ? sh.getRange(2, iT + 1, n, 1).getValues() : null;
    const outS = iS !== -1 ? sh.getRange(2, iS + 1, n, 1).getValues() : null;
    const outR = iR !== -1 ? sh.getRange(2, iR + 1, n, 1).getValues() : null;
    const styleJobs = [];

    for (let r = 0; r < n; r++) {
      const k = keyOf(claims[r][0], scs[r][0]);
      const rec = map[k];
      if (!rec) continue;
      if (outU && String(outU[r][0] || '').trim() === '' && String(rec.u || '').trim() !== '') { outU[r][0] = rec.u; restored++; }
      if (outT && String(outT[r][0] || '').trim() === '' && String(rec.t || '').trim() !== '') { outT[r][0] = rec.t; restored++; }
      if (outS && String(outS[r][0] || '').trim() === '' && String(rec.s || '').trim() !== '') { outS[r][0] = rec.s; restored++; }
      if (outR && String(outR[r][0] || '').trim() === '' && String(rec.r || '').trim() !== '') { outR[r][0] = rec.r; restored++; }

      // Preserve style 1:1 from temp sheet (rich text, font color, wrap, DV)
      // after value writes. If copyTo runs before setValues(), the later value
      // batch can flatten rich text/font styling and recreate the shifted color
      // issue during MAIN -> SUB restore.
      if (rec.rowNo) {
        if (iU !== -1 && outU && String(outU[r][0] || '').trim() !== '') styleJobs.push({ srcRow: rec.rowNo, srcCol: 6, dstRow: r + 2, dstCol: iU + 1 });
        if (iT !== -1 && outT && String(outT[r][0] || '').trim() !== '') styleJobs.push({ srcRow: rec.rowNo, srcCol: 7, dstRow: r + 2, dstCol: iT + 1 });
        if (iS !== -1 && outS && String(outS[r][0] || '').trim() !== '') styleJobs.push({ srcRow: rec.rowNo, srcCol: 8, dstRow: r + 2, dstCol: iS + 1 });
        if (iR !== -1 && outR && String(outR[r][0] || '').trim() !== '') styleJobs.push({ srcRow: rec.rowNo, srcCol: 9, dstRow: r + 2, dstCol: iR + 1 });
      }
    }

    try { if (outU) sh.getRange(2, iU + 1, n, 1).setValues(outU); } catch (e) {}
    try { if (outT) sh.getRange(2, iT + 1, n, 1).setValues(outT); } catch (e) {}
    try { if (outS) sh.getRange(2, iS + 1, n, 1).setValues(outS); } catch (e) {}
    try { if (outR) sh.getRange(2, iR + 1, n, 1).setValues(outR); } catch (e) {}
    for (let j = 0; j < styleJobs.length; j++) {
      const job = styleJobs[j];
      try { shBak.getRange(job.srcRow, job.srcCol).copyTo(sh.getRange(job.dstRow, job.dstCol), { contentsOnly: false }); } catch (eCopy) {}
    }
  }

  if (opts.deleteAfterRestore !== false) {
    try { ss.deleteSheet(shBak); } catch (eDel) {}
  }
  return { restored: restored, rows: Object.keys(map).length, skipped: false };
}

function restoreOpsManualFromBackupSheet06c_(ss, pic) {
  if (DRY_RUN) return { restored: 0, rows: 0 };
  if (!ss) return { restored: 0, rows: 0 };
  const shBak = ss.getSheetByName('_OPS_MANUAL_BACKUP');
  if (!shBak || shBak.getLastRow() < 2) return { restored: 0, rows: 0 };

  const lc = shBak.getLastColumn();
  const hdr = shBak.getRange(1,1,1,lc).getValues()[0].map(__normalizeHeaderText06_);
  const idxPic = __findHeaderIndexFlexible06_(hdr, 'PIC');
  const idxClaim = __findHeaderIndexFlexible06_(hdr, 'Claim Number');
  const idxUpd = __findHeaderIndexFlexible06_(hdr, 'Update Status');
  const idxTs = __findHeaderIndexFlexible06_(hdr, 'Timestamp');
  const idxSt = __findHeaderIndexFlexible06_(hdr, 'Status');
  const idxRem = __findHeaderIndexFlexible06_(hdr, 'Remarks');
  if (idxPic === -1 || idxClaim === -1) return { restored: 0, rows: 0 };

  const vals = shBak.getRange(2,1,shBak.getLastRow()-1,lc).getValues();
  const map = Object.create(null);
  const picKey = String(pic || '').trim().toUpperCase();
  for (let i=0;i<vals.length;i++) {
    const row=vals[i];
    if (String(row[idxPic] || '').trim().toUpperCase() !== picKey) continue;
    const k = __claimKey06_(row[idxClaim]);
    if (!k) continue;
    map[k] = { u: idxUpd!==-1?row[idxUpd]:'', t: idxTs!==-1?row[idxTs]:'', s: idxSt!==-1?row[idxSt]:'', r: idxRem!==-1?row[idxRem]:'' };
  }

  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  let restored = 0;
  for (let si=0; si<sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;
    const lr = sh.getLastRow(); const lc2 = sh.getLastColumn();
    if (lr < 2 || lc2 < 1) continue;
    const h = sh.getRange(1,1,1,lc2).getValues()[0].map(__normalizeHeaderText06_);
    const iC = __findHeaderIndexFlexible06_(h,'Claim Number');
    if (iC===-1) continue;
    const iU = __findHeaderIndexFlexible06_(h,'Update Status');
    const iT = __findHeaderIndexFlexible06_(h,'Timestamp');
    const iS = __findHeaderIndexFlexible06_(h,'Status');
    const iR = __findHeaderIndexFlexible06_(h,'Remarks');
    const n=lr-1;
    const claims=sh.getRange(2,iC+1,n,1).getValues();
    const outU=iU!==-1?sh.getRange(2,iU+1,n,1).getValues():null;
    const outT=iT!==-1?sh.getRange(2,iT+1,n,1).getValues():null;
    const outS=iS!==-1?sh.getRange(2,iS+1,n,1).getValues():null;
    const outR=iR!==-1?sh.getRange(2,iR+1,n,1).getValues():null;
    
    for (let r=0;r<n;r++) {
      const rec = map[__claimKey06_(claims[r][0])];
      if (!rec) continue;
      if (outU && String(outU[r][0]||'').trim()==='' && String(rec.u||'').trim()!=='') { outU[r][0]=rec.u; restored++; }
      if (outT && String(outT[r][0]||'').trim()==='' && String(rec.t||'').trim()!=='') { outT[r][0]=rec.t; restored++; }
      if (outS && String(outS[r][0]||'').trim()==='' && String(rec.s||'').trim()!=='') { outS[r][0]=rec.s; restored++; }
      if (outR && String(outR[r][0]||'').trim()==='' && String(rec.r||'').trim()!=='') { outR[r][0]=rec.r; restored++; }
    }
    try { if (outU) sh.getRange(2,iU+1,n,1).setValues(outU); } catch(e){}
    try { if (outT) sh.getRange(2,iT+1,n,1).setValues(outT); } catch(e){}
    try { if (outS) sh.getRange(2,iS+1,n,1).setValues(outS); } catch(e){}
    try { if (outR) sh.getRange(2,iR+1,n,1).setValues(outR); } catch(e){}
  }
  return { restored: restored, rows: Object.keys(map).length };
}

function restoreOpsManualColumnsRich06c_(ss, pic, snapshot) {
  if (DRY_RUN) return;
  if (!ss || !snapshot || !snapshot.map) return;

  const snap = snapshot.map;

  const sheetNames = (typeof getOperationalSheetsForBackup_ === 'function')
    ? getOperationalSheetsForBackup_(pic)
    : (CONFIG && CONFIG.sheetsByPic ? Array.from(new Set([].concat(CONFIG.sheetsByPic.picOperational || [], CONFIG.sheetsByPic.adminOperational || []))) : []);

  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;

    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    if (idxClaim === -1) continue;

    const isEvBike = String(sh.getName() || '').trim().toLowerCase() === 'ev-bike';

    const idxUpdate = __findHeaderIndexFlexible06_(header, 'Update Status');
    const idxTs = __findHeaderIndexFlexible06_(header, 'Timestamp');
    const idxStatus = isEvBike ? -1 : __findHeaderIndexFlexible06_(header, 'Status');
    const idxRemarks = __findHeaderIndexFlexible06_(header, 'Remarks');
    const idxAwb = __findHeaderIndexFlexible06_(header, 'AWB');
    const idxTimestampAwb = __findHeaderIndexFlexible06_(header, 'Timestamp AWB');

    if (idxUpdate === -1 && idxTs === -1 && idxStatus === -1 && idxRemarks === -1 && idxAwb === -1 && idxTimestampAwb === -1) continue;

    const n = lr - 1;
    const claims = sh.getRange(2, idxClaim + 1, n, 1).getValues();

    const rowsUpdate = [];
    const mapUpdate = Object.create(null);
    const mapUpdateWrap = Object.create(null);

    const rowsRemarks = [];
    const mapRemarks = Object.create(null);
    const mapRemarksWrap = Object.create(null);

    const rowsTs = [];
    const mapTs = Object.create(null);
    const mapTsFmt = Object.create(null);

    const rowsStatus = [];
    const mapStatus = Object.create(null);
    const formulaJobs = { u: { rows: [], map: Object.create(null), idx: idxUpdate }, t: { rows: [], map: Object.create(null), idx: idxTs }, s: { rows: [], map: Object.create(null), idx: idxStatus }, r: { rows: [], map: Object.create(null), idx: idxRemarks }, a: { rows: [], map: Object.create(null), idx: idxAwb }, ta: { rows: [], map: Object.create(null), idx: idxTimestampAwb } };
    const rowsAwb = [], mapAwb = Object.create(null), rowsTimestampAwb = [], mapTimestampAwb = Object.create(null), mapTimestampAwbFmt = Object.create(null);

    for (let i = 0; i < n; i++) {
      const claim = __claimKey06_(claims[i] && claims[i][0]);
      if (!claim) continue;
      const key = claim.toUpperCase();
      const rec = snap[key];
      if (!rec) continue;

      const rno = i + 2;
      [['u', rec.u], ['t', rec.t], ['s', rec.s], ['r', rec.r], ['a', rec.a], ['ta', rec.ta]].forEach(function(pair) {
        const job = formulaJobs[pair[0]], cell = pair[1];
        if (job.idx !== -1 && cell && cell.formula) { job.rows.push(rno); job.map[rno] = cell.formula; }
      });

      if (idxUpdate !== -1 && rec.u && rec.u.rt) {
        rowsUpdate.push(rno);
        mapUpdate[rno] = rec.u.rt;
        if (rec.u.wrap) mapUpdateWrap[rno] = rec.u.wrap;
      }

      if (idxRemarks !== -1 && rec.r && rec.r.rt) {
        rowsRemarks.push(rno);
        mapRemarks[rno] = rec.r.rt;
        if (rec.r.wrap) mapRemarksWrap[rno] = rec.r.wrap;
      }

      if (idxTs !== -1 && rec.t && rec.t.v != null && rec.t.v !== '') {
        rowsTs.push(rno);
        mapTs[rno] = rec.t.v;
        if (rec.t.fmt) mapTsFmt[rno] = rec.t.fmt;
      }

      if (idxStatus !== -1 && rec.s && rec.s.v != null && rec.s.v !== '') {
        rowsStatus.push(rno);
        mapStatus[rno] = rec.s.v;
      }
      if (idxAwb !== -1 && rec.a && rec.a.v != null && rec.a.v !== '') { rowsAwb.push(rno); mapAwb[rno] = rec.a.v; }
      if (idxTimestampAwb !== -1 && rec.ta && rec.ta.v != null && rec.ta.v !== '') { rowsTimestampAwb.push(rno); mapTimestampAwb[rno] = rec.ta.v; if (rec.ta.fmt) mapTimestampAwbFmt[rno] = rec.ta.fmt; }
    }

    const applyRichSeg = (idxCol0, rowNums, rowMapObj) => {
      if (!rowNums.length) return;
      const segs = __groupConsecutiveRows_(rowNums);
      for (let k = 0; k < segs.length; k++) {
        const seg = segs[k];
        const startRow = seg[0];
        const rich2d = seg.map(rn => [rowMapObj[rn]]);
        try { sh.getRange(startRow, idxCol0 + 1, rich2d.length, 1).setRichTextValues(rich2d); } catch (e) {}
      }
    };

    const applyWrapSeg = (idxCol0, rowNums, rowWrapObj) => {
      if (!rowNums.length) return;
      const segs = __groupConsecutiveRows_(rowNums);
      for (let k = 0; k < segs.length; k++) {
        const seg = segs[k];
        const startRow = seg[0];
        const wraps2d = seg.map(rn => [rowWrapObj[rn] || 'WRAP']);
        try { sh.getRange(startRow, idxCol0 + 1, wraps2d.length, 1).setWrapStrategies(wraps2d); } catch (e) {}
      }
    };

    const applyValSeg = (idxCol0, rowNums, rowValObj) => {
      if (!rowNums.length) return;
      const segs = __groupConsecutiveRows_(rowNums);
      for (let k = 0; k < segs.length; k++) {
        const seg = segs[k];
        const startRow = seg[0];
        const vals2d = seg.map(rn => [rowValObj[rn]]);
        try { sh.getRange(startRow, idxCol0 + 1, vals2d.length, 1).setValues(vals2d); } catch (e) {}
      }
    };

    const applyFormulaSeg = (idxCol0, rowNums, rowFormulaObj) => {
      if (idxCol0 === -1 || !rowNums.length) return;
      __groupConsecutiveRows_(rowNums).forEach(function(seg) {
        try { sh.getRange(seg[0], idxCol0 + 1, seg.length, 1).setFormulas(seg.map(rn => [rowFormulaObj[rn]])); } catch (e) {}
      });
    };

    const applyNumFmtSeg = (idxCol0, rowNums, rowFmtObj) => {
      if (!rowNums.length) return;
      const segs = __groupConsecutiveRows_(rowNums);
      for (let k = 0; k < segs.length; k++) {
        const seg = segs[k];
        const startRow = seg[0];
        const fmt2d = seg.map(rn => [rowFmtObj[rn] || sh.getRange(startRow, idxCol0 + 1).getNumberFormat()]);
        try { sh.getRange(startRow, idxCol0 + 1, fmt2d.length, 1).setNumberFormats(fmt2d); } catch (e) {}
      }
    };

    if (idxUpdate !== -1) {
      applyRichSeg(idxUpdate, rowsUpdate, mapUpdate);
      if (Object.keys(mapUpdateWrap).length) applyWrapSeg(idxUpdate, rowsUpdate, mapUpdateWrap);
    }

    if (idxRemarks !== -1) {
      applyRichSeg(idxRemarks, rowsRemarks, mapRemarks);
      if (Object.keys(mapRemarksWrap).length) applyWrapSeg(idxRemarks, rowsRemarks, mapRemarksWrap);
    }

    if (idxTs !== -1) {
      applyValSeg(idxTs, rowsTs, mapTs);
      if (Object.keys(mapTsFmt).length) applyNumFmtSeg(idxTs, rowsTs, mapTsFmt);
    }

    if (idxStatus !== -1) {
      applyValSeg(idxStatus, rowsStatus, mapStatus);
    }
    if (idxAwb !== -1) applyValSeg(idxAwb, rowsAwb, mapAwb);
    if (idxTimestampAwb !== -1) { applyValSeg(idxTimestampAwb, rowsTimestampAwb, mapTimestampAwb); if (Object.keys(mapTimestampAwbFmt).length) applyNumFmtSeg(idxTimestampAwb, rowsTimestampAwb, mapTimestampAwbFmt); }
    // Formula restore is deliberately last: formulas recalculate from the newly routed row.
    Object.keys(formulaJobs).forEach(function(key) {
      const job = formulaJobs[key];
      applyFormulaSeg(job.idx, job.rows, job.map);
    });
  }
}


/**
 * Apply template DV + formatting to operational sheets.
 * Template row: row 2 (must be pre-configured with correct dropdown chips, formats, checkbox, etc.)
 * Destination: rows 2..(lastRow + buffer)
 */
function applyTemplateRowToOperationalSheets_(ss, pic) {
  if (DRY_RUN) return;
  if (!ss) return;

  const isAdmin = (pic === 'Admin');
  const sheets = isAdmin ? (CONFIG.sheetsByPic.adminOperational || []) : (CONFIG.sheetsByPic.picOperational || []);
  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastCol = sh.getLastColumn();
    if (lastCol <= 0) return;

    // Spec: Operational sheet header (row 1) must be centered and vertically aligned to middle.
    // Apply only to header row (bounded), to avoid mass-formatting entire sheets.
    try {
      sh.getRange(1, 1, 1, lastCol)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    } catch (e0) {}

    const maxRows = sh.getMaxRows();
    const lastRow = sh.getLastRow();

    // Keep DV extended for usability, but DO NOT extend formats beyond existing data rows
    // (prevents "pink leak" in Claim Number and other template backgrounds).
    const buffer = (name === 'Ask Detail') ? 1500 : 400;

    const dvTargetLastRow = Math.min(maxRows, Math.max(2, lastRow) + buffer);
    const dvRowCount = dvTargetLastRow - 1; // rows starting at 2
    const fmtTargetLastRow = Math.min(maxRows, Math.max(2, lastRow));
    const fmtRowCount = fmtTargetLastRow - 1;

    const src = sh.getRange(2, 1, 1, lastCol);

    if (dvRowCount > 0) {
      const dstDv = sh.getRange(2, 1, dvRowCount, lastCol);
      try { src.copyTo(dstDv, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); } catch (e) {}
      // Do NOT apply DV to "Status Type" (derived by script; must remain non-dropdown).
      // Do NOT apply DV to "Submission Date" (must stay date/plain, never checkbox).
      try {
        const hdr1 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);
        const idxStatusType = __findHeaderIndexFlexible06_(hdr1, 'Status Type');
        if (idxStatusType !== -1) {
          sh.getRange(2, idxStatusType + 1, dvRowCount, 1).clearDataValidations();
        }
        const idxSubmissionDate = __findHeaderIndexFlexible06_(hdr1, 'Submission Date');
        if (idxSubmissionDate !== -1) {
          sh.getRange(2, idxSubmissionDate + 1, dvRowCount, 1).clearDataValidations();
        }
      } catch (e2) {}
    }

    if (fmtRowCount > 0) {
      const dstFmt = sh.getRange(2, 1, fmtRowCount, lastCol);
      try { src.copyTo(dstFmt, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
    }
  });
}

/**
 * Defensive validation sanitizer for known problematic columns.
 * - Submission Date: must never carry checkbox/dropdown validation.
 * - EV-Bike.Last Status: user-managed free text; ignore dropdown validation.
 */
function sanitizeProblematicDataValidations06_(ss, pic) {
  if (DRY_RUN) return;
  if (!ss) return;

  const isAdmin = (pic === 'Admin');
  const opsSheets = isAdmin ? (CONFIG.sheetsByPic.adminOperational || []) : (CONFIG.sheetsByPic.picOperational || []);

  // 1) Operational sheets: clear DV on Submission Date
  opsSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastCol = sh.getLastColumn();
    const maxRows = sh.getMaxRows();
    const lastRow = Math.max(2, sh.getLastRow());
    const rows = Math.min(maxRows - 1, (lastRow - 1) + 400);
    if (lastCol <= 0 || maxRows <= 1 || rows <= 0) return;

    try {
      const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      const idxSubmissionDate = __findHeaderIndexFlexible06_(header, 'Submission Date');
      if (idxSubmissionDate !== -1) {
        const subDateRng = sh.getRange(2, idxSubmissionDate + 1, rows, 1);
        // Clear any stale checkbox / dropdown DV
        subDateRng.clearDataValidations();
        // FIX: also enforce date number format so cells cannot display as checkbox/serial
        // number regardless of what the template row previously had.
        try {
          const dateFmt = (typeof FORMATS !== 'undefined' && FORMATS && FORMATS.DATE)
            ? FORMATS.DATE
            : 'd mmm yy';
          subDateRng.setNumberFormat(dateFmt);
        } catch (_eFmtSub) {}
      }
    } catch (e0) {}
  });

  // 2) EV-Bike/Doss: clear DV on Last Status to avoid validation violations when writing statuses
  ['EV-Bike', 'Doss'].forEach(function(sheetName) {
    try {
      const ev = ss.getSheetByName(sheetName);
      if (!ev) return;
      const lc = ev.getLastColumn();
      const mr = ev.getMaxRows();
      if (lc <= 0 || mr <= 1) return;

      const headerEv = ev.getRange(1, 1, 1, lc).getValues()[0];
      const idxLastStatusEv = __findHeaderIndexFlexible06_(headerEv, 'Last Status');
      if (idxLastStatusEv !== -1) {
        ev.getRange(2, idxLastStatusEv + 1, mr - 1, 1).clearDataValidations();
      }
    } catch (e1) {}
  });
}

/**
 * Restore operational fields from Raw backup after CLR + ROUTE.
 * Restores:
 * - Status (values)
 * - OR (checkbox + values)
 * - Timestamp (values)
 *
 * Update Status rich text is handled by applyUpdateStatusRichTextToOperational_().
 */
function restoreOpsFieldsFromRawBackup_(ss, rawSheet, headerIndexRaw, pic) {
  if (DRY_RUN) return;
  if (!ss || !rawSheet || !headerIndexRaw) return;

  const h = CONFIG.headers;
  const idxClaimRaw = headerIndexRaw[h.claimNumber];
  if (idxClaimRaw == null) return;

  const idxStatusRaw = headerIndexRaw[h.status];
  const idxOrRaw = headerIndexRaw[h.orColumn];
  const idxTsRaw = headerIndexRaw[h.timestamp];

  const idxUpdateRaw = headerIndexRaw[h.updateStatus];
  const idxRemarksRaw = (typeof idxAny_ === 'function') ? idxAny_(headerIndexRaw, ['Remarks','Remark','remarks','remark']) : (headerIndexRaw['Remarks'] != null ? headerIndexRaw['Remarks'] : null);

  // Deprecated Asso/Admin tail fields are no longer restored.
  const idxUpdateAssoRaw = null;
  const idxTsAssoRaw = null;
  const idxUpdateAdminRaw = null;
  const idxTsAdminRaw = null;

  const n = rawSheet.getLastRow() - 1;
  if (n <= 0) return;

  const rawClaims = rawSheet.getRange(2, idxClaimRaw + 1, n, 1).getValues();
  const rawStatus = (idxStatusRaw != null) ? rawSheet.getRange(2, idxStatusRaw + 1, n, 1).getValues() : null;
  const rawOR = (idxOrRaw != null) ? rawSheet.getRange(2, idxOrRaw + 1, n, 1).getValues() : null;
  const rawTs = (idxTsRaw != null) ? rawSheet.getRange(2, idxTsRaw + 1, n, 1).getValues() : null;
  const rawUpdate = (idxUpdateRaw != null) ? rawSheet.getRange(2, idxUpdateRaw + 1, n, 1).getValues() : null;
  const rawRemarks = (idxRemarksRaw != null) ? rawSheet.getRange(2, idxRemarksRaw + 1, n, 1).getValues() : null;

  const rawUpAsso = (idxUpdateAssoRaw != null) ? rawSheet.getRange(2, idxUpdateAssoRaw + 1, n, 1).getValues() : null;
  const rawTsAsso = (idxTsAssoRaw != null) ? rawSheet.getRange(2, idxTsAssoRaw + 1, n, 1).getValues() : null;
  const rawUpAdmin = (idxUpdateAdminRaw != null) ? rawSheet.getRange(2, idxUpdateAdminRaw + 1, n, 1).getValues() : null;
  const rawTsAdmin = (idxTsAdminRaw != null) ? rawSheet.getRange(2, idxTsAdminRaw + 1, n, 1).getValues() : null;

  const rawMap = Object.create(null);
  for (let i = 0; i < rawClaims.length; i++) {
    const key = __claimKey06_(rawClaims[i][0]);
    if (key) rawMap[key] = i;
  }

  const isAdmin = (pic === 'Admin');
  const sheets = isAdmin ? (CONFIG.sheetsByPic.adminOperational || []) : (CONFIG.sheetsByPic.picOperational || []);

  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastCol = sh.getLastColumn();
    if (lastCol <= 0) return;
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);

    const idxClaimOps = __findHeaderIndexFlexible06_(header, 'Claim Number');
    if (idxClaimOps === -1) return;

    const isEvBike = String(name || '').trim().toLowerCase() === 'ev-bike';
    const idxStatusOps = isEvBike ? -1 : __findHeaderIndexFlexible06_(header, 'Status');
    const idxOrOps = __findHeaderIndexFlexible06_(header, 'OR');
    const idxTsOps = __findHeaderIndexFlexible06_(header, 'Timestamp');

    const idxUpdateOps = __findHeaderIndexFlexible06_(header, 'Update Status');
    const idxRemarksOps = __findHeaderIndexFlexible06_(header, 'Remarks');

    const idxUpAssoOps = -1;
    const idxTsAssoOps = -1;
    const idxUpAdminOps = -1;
    const idxTsAdminOps = -1;

    const lr = sh.getLastRow();
    if (lr <= 1) return;

    const m = lr - 1;
    const claimsOps = sh.getRange(2, idxClaimOps + 1, m, 1).getValues();

    const outStatus = (idxStatusOps !== -1 && rawStatus) ? new Array(m) : null;
    const outOR = (idxOrOps !== -1 && rawOR) ? new Array(m) : null;
    const outTs = (idxTsOps !== -1 && rawTs) ? new Array(m) : null;

    const outUpdate = (idxUpdateOps !== -1 && rawUpdate) ? new Array(m) : null;
    const outRemarks = (idxRemarksOps !== -1 && rawRemarks) ? new Array(m) : null;

    const outUpAsso = (idxUpAssoOps !== -1 && rawUpAsso) ? new Array(m) : null;
    const outTsAsso = (idxTsAssoOps !== -1 && rawTsAsso) ? new Array(m) : null;
    const outUpAdmin = (idxUpAdminOps !== -1 && rawUpAdmin) ? new Array(m) : null;
    const outTsAdmin = (idxTsAdminOps !== -1 && rawTsAdmin) ? new Array(m) : null;

    for (let r = 0; r < m; r++) {
      const claimKey = __claimKey06_(claimsOps[r][0]);
      const ri = claimKey ? rawMap[claimKey] : null;

      if (outStatus) outStatus[r] = [ (ri != null ? rawStatus[ri][0] : '') ];
      if (outOR) outOR[r] = [ (ri != null ? normalizeCheckbox_(rawOR[ri][0]) : '') ];
      if (outTs) outTs[r] = [ (ri != null ? rawTs[ri][0] : '') ];

      if (outUpdate) outUpdate[r] = [ (ri != null ? rawUpdate[ri][0] : '') ];
      if (outRemarks) outRemarks[r] = [ (ri != null ? rawRemarks[ri][0] : '') ];

      if (outUpAsso) outUpAsso[r] = [ (ri != null ? rawUpAsso[ri][0] : '') ];
      if (outTsAsso) outTsAsso[r] = [ (ri != null ? rawTsAsso[ri][0] : '') ];
      if (outUpAdmin) outUpAdmin[r] = [ (ri != null ? rawUpAdmin[ri][0] : '') ];
      if (outTsAdmin) outTsAdmin[r] = [ (ri != null ? rawTsAdmin[ri][0] : '') ];
    }

    // Apply checkbox validation first (then values)
    if (idxOrOps !== -1 && outOR) {
      try { sh.getRange(2, idxOrOps + 1, m, 1).insertCheckboxes(); } catch (e) {}
      try { sh.getRange(2, idxOrOps + 1, m, 1).setValues(outOR); } catch (e) {}
    }

    if (idxStatusOps !== -1 && outStatus) {
      try { sh.getRange(2, idxStatusOps + 1, m, 1).setValues(outStatus); } catch (e) {}
    }

    if (idxTsOps !== -1 && outTs) {
      try { sh.getRange(2, idxTsOps + 1, m, 1).setValues(outTs); } catch (e) {}
    }

    // IMPORTANT:
    // Do not write plain Update Status values here.
    // RichText restoration is handled separately by applyUpdateStatusRichTextToOperational_().
    // Rewriting this column with setValues() can retrigger sheet formulas/automation tied to
    // Update Status and unexpectedly refresh Timestamp for many rows in one run.

    if (idxRemarksOps !== -1 && outRemarks) {
      try { sh.getRange(2, idxRemarksOps + 1, m, 1).setValues(outRemarks); } catch (e) {}
    }

    // Tail/manual columns (best-effort)
    if (idxUpAssoOps !== -1 && outUpAsso) {
      try { sh.getRange(2, idxUpAssoOps + 1, m, 1).setValues(outUpAsso); } catch (e) {}
    }
    if (idxTsAssoOps !== -1 && outTsAsso) {
      try { sh.getRange(2, idxTsAssoOps + 1, m, 1).setValues(outTsAsso); } catch (e) {}
    }
    if (idxUpAdminOps !== -1 && outUpAdmin) {
      try { sh.getRange(2, idxUpAdminOps + 1, m, 1).setValues(outUpAdmin); } catch (e) {}
    }
    if (idxTsAdminOps !== -1 && outTsAdmin) {
      try { sh.getRange(2, idxTsAdminOps + 1, m, 1).setValues(outTsAdmin); } catch (e) {}
    }
  });
}
function buildRawValuesFromMain_(rawHeader, headerIndexRaw, mainHeader, mainRows) {
  const rows = mainRows || [];
  const out = new Array(rows.length);

  // Build map mainIdx -> rawIdx
  const pairs = [];
  for (let i = 0; i < (mainHeader || []).length; i++) {
    const key = String(mainHeader[i] || '').trim();
    if (!key) continue;

    const rawIdx = headerIndexRaw[key];
    if (rawIdx == null) continue;

    // never carry into Raw "Status" from main file (manual+validated)
    if (key === CONFIG.headers.status || key === 'Status') continue;

    pairs.push({ rawIdx: rawIdx, mainIdx: i });
  }

  for (let r = 0; r < rows.length; r++) {
    const row = new Array(rawHeader.length).fill('');
    const src = rows[r];

    for (let k = 0; k < pairs.length; k++) {
      row[pairs[k].rawIdx] = src[pairs[k].mainIdx];
    }
    out[r] = row;
  }
  return out;
}

/** Carry-forward map from existing Raw (Claim -> {OR, Update, Timestamp, Status}) */
function buildRawCarryForwardMap_(rawSheet, headerIndexRaw) {
  const h = CONFIG.headers;

  const idxClaim = headerIndexRaw[h.claimNumber];
  const idxOR = headerIndexRaw[h.orColumn];
  const idxUpdate = headerIndexRaw[h.updateStatus];
  const idxTs = headerIndexRaw[h.timestamp];
  const idxStatus = headerIndexRaw[h.status];
  const idxRemarks = (typeof idxAny_ === 'function') ? idxAny_(headerIndexRaw, ['Remarks','Remark','remarks','remark']) : (headerIndexRaw['Remarks'] != null ? headerIndexRaw['Remarks'] : null);

  // Custom tail columns (manual)
  const idxAssoc = (headerIndexRaw[h.associate] != null) ? headerIndexRaw[h.associate] : headerIndexRaw['Associate'];
  const idxUpdateAsso = null;
  const idxTsAsso = null;
  const idxUpdateAdmin = null;
  const idxTsAdmin = null;

  const lr = rawSheet.getLastRow();
  if (lr < 2 || idxClaim == null) return { map: Object.create(null), count: 0 };

  const n = lr - 1;

  const colVals = (c0) => (c0 == null ? null : rawSheet.getRange(2, c0 + 1, n, 1).getValues());
  const claims = colVals(idxClaim);

  const ors = colVals(idxOR);
  const ups = colVals(idxUpdate);
  const tss = colVals(idxTs);
  const sts = colVals(idxStatus);
  const rms = colVals(idxRemarks);

  const assocs = colVals(idxAssoc);
  const upAsso = colVals(idxUpdateAsso);
  const tsAsso = colVals(idxTsAsso);
  const upAdmin = colVals(idxUpdateAdmin);
  const tsAdmin = colVals(idxTsAdmin);

  const map = Object.create(null);
  let count = 0;

  for (let i = 0; i < n; i++) {
    const claim = String(claims[i][0] || '').trim().toUpperCase();
    if (!claim) continue;

    map[claim] = {
      assocVal: assocs ? assocs[i][0] : null,
      orVal: ors ? ors[i][0] : null,
      updateVal: ups ? ups[i][0] : null,
      tsVal: tss ? tss[i][0] : null,
      statusVal: sts ? sts[i][0] : null,
      remarksVal: rms ? rms[i][0] : null,
      updateAssoVal: upAsso ? upAsso[i][0] : null,
      tsAssoVal: tsAsso ? tsAsso[i][0] : null,
      updateAdminVal: upAdmin ? upAdmin[i][0] : null,
      tsAdminVal: tsAdmin ? tsAdmin[i][0] : null
    };
    count++;
  }

  return { map: map, count: count };
}


function applyCarryForwardToRawValues_(rawValues, headerIndexRaw, carry) {
  if (!carry || !carry.map) return;

  const h = CONFIG.headers;
  const idxClaim = headerIndexRaw[h.claimNumber];
  if (idxClaim == null) return;

  const idxAssoc = (headerIndexRaw[h.associate] != null) ? headerIndexRaw[h.associate] : headerIndexRaw['Associate'];
  const idxOR = headerIndexRaw[h.orColumn];
  const idxUpdate = headerIndexRaw[h.updateStatus];
  const idxTs = headerIndexRaw[h.timestamp];
  const idxStatus = headerIndexRaw[h.status];
  const idxRemarks = (typeof idxAny_ === 'function') ? idxAny_(headerIndexRaw, ['Remarks','Remark','remarks','remark']) : (headerIndexRaw['Remarks'] != null ? headerIndexRaw['Remarks'] : null);

  const idxUpdateAsso = null;
  const idxTsAsso = null;
  const idxUpdateAdmin = null;
  const idxTsAdmin = null;

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];
    const claim = String(row[idxClaim] || '').trim().toUpperCase();
    if (!claim) continue;

    const saved = carry.map[claim];
    if (!saved) continue;

    // Preserve manual tail columns (only when current is blank)
    if (idxAssoc != null && (row[idxAssoc] == null || String(row[idxAssoc]).trim() === '') && saved.assocVal != null && String(saved.assocVal).trim() !== '') {
      row[idxAssoc] = saved.assocVal;
    }

    if (idxOR != null && saved.orVal != null && saved.orVal !== '') row[idxOR] = saved.orVal;
    if (idxUpdate != null && saved.updateVal != null && String(saved.updateVal).trim() !== '') row[idxUpdate] = saved.updateVal;
    if (idxTs != null && saved.tsVal != null && saved.tsVal !== '') row[idxTs] = saved.tsVal;
    if (idxStatus != null && saved.statusVal != null && String(saved.statusVal).trim() !== '') row[idxStatus] = saved.statusVal;

    if (idxRemarks != null && saved.remarksVal != null && String(saved.remarksVal) !== '') row[idxRemarks] = saved.remarksVal;

    if (idxUpdateAsso != null && saved.updateAssoVal != null && String(saved.updateAssoVal).trim() !== '') row[idxUpdateAsso] = saved.updateAssoVal;
    if (idxTsAsso != null && saved.tsAssoVal != null && saved.tsAssoVal !== '') row[idxTsAsso] = saved.tsAssoVal;
    if (idxUpdateAdmin != null && saved.updateAdminVal != null && String(saved.updateAdminVal).trim() !== '') row[idxUpdateAdmin] = saved.updateAdminVal;
    if (idxTsAdmin != null && saved.tsAdminVal != null && saved.tsAdminVal !== '') row[idxTsAdmin] = saved.tsAdminVal;
  }
}


/** Status dropdown sanitization (prevents "violates data validation rules") */
function sanitizeRawStatusDropdownInMemory_(rawValues, headerIndexRaw) {
  const idx = headerIndexRaw[CONFIG.headers.status];
  if (idx == null) return;

  const allowed = getAllowedStatusSet_();
  for (let i = 0; i < rawValues.length; i++) {
    const v = normalizeStatusValue_(rawValues[i][idx]);
    if (!v) {
      rawValues[i][idx] = '';
      continue;
    }
    rawValues[i][idx] = allowed.has(v) ? v : '';
  }
}

function getAllowedStatusSet_() {
  const list =
    (typeof VALIDATION_LISTS !== 'undefined' && VALIDATION_LISTS && VALIDATION_LISTS.UPDATE_STATUS)
      ? VALIDATION_LISTS.UPDATE_STATUS
      : [
          'Pending Admin','Pending SC','Pending Partner','DONE','Pending Insurance',
          'Pending TO','Pending Finance','Pending Meilani','Pending Cust'
        ];

  const set = new Set();
  for (let i = 0; i < list.length; i++) {
    set.add(String(list[i] || '').trim());
  }
  return set;
}

/** Safety preflight: count routable rows BEFORE clearing ops sheets (PARITY with routing suppression) */
function preflightFilterScTargets06_(targets, scNameVal, scFarhanName, scMeilaniName, scIvanName, kwFarhan, kwMeilani, kwIvan) {
  const t = (targets || []).slice();
  if (!t.length) return t;

  const s = __normalizeScKeywordText06c_(scNameVal);

  const hasFarhan = (kwFarhan || []).some(k => {
    const nk = __normalizeScKeywordText06c_(k);
    return nk && s.indexOf(nk) > -1;
  });
  const hasMeilani = (kwMeilani || []).some(k => {
    const nk = __normalizeScKeywordText06c_(k);
    return nk && s.indexOf(nk) > -1;
  });
  const hasIvan = (kwIvan || []).some(k => {
    const nk = __normalizeScKeywordText06c_(k);
    return nk && s.indexOf(nk) > -1;
  });

  return t.filter(x => {
    if (x === scFarhanName) return hasFarhan;
    if (x === scMeilaniName) return hasMeilani;
    if (scIvanName && x === scIvanName) return hasIvan;
    return true;
  });
}

function __normalizeScKeywordText06c_(v) {
  return String(v == null ? '' : v)
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function preflightRoutableCount_(rawValues, headerIndexRaw, pic) {
  const h = CONFIG.headers;
  const idxStatus = headerIndexRaw[h.lastStatus];
  if (idxStatus == null) return { total: 0, reason: 'missing_last_status_header' };

  const routingMap = (CONFIG && (CONFIG.statusRoutingAdmin || CONFIG.statusRoutingPIC)) ? (CONFIG.statusRoutingAdmin || CONFIG.statusRoutingPIC) : {};
  const routingIndex = compileRoutingIndex_(routingMap);

  const opsPolicy = (CONFIG && (CONFIG.opsRouting || CONFIG.opsRoutingPolicy || CONFIG.OPS_ROUTING_POLICY)) || null;
  const scFarhanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_FARHAN) ? opsPolicy.SHEETS.SC_FARHAN : 'SC - Farhan';
  const scMeilaniName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_MEILANI) ? opsPolicy.SHEETS.SC_MEILANI : 'SC - Meilani';
  const scIvanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_IVAN) ? opsPolicy.SHEETS.SC_IVAN : 'SC - Meindar';

  const scKeywords = (opsPolicy && opsPolicy.SC_NAME_KEYWORDS) ? opsPolicy.SC_NAME_KEYWORDS : {};
  const kwFarhan = scKeywords[scFarhanName] || scKeywords['SC - Farhan'] || [];
  const kwMeilani = scKeywords[scMeilaniName] || scKeywords['SC - Meilani'] || [];
  const kwIvan = scKeywords[scIvanName] || scKeywords['SC - Meindar'] || [];

  const idxScName = headerIndexRaw[h.scName];

  let total = 0;
  for (let i = 0; i < (rawValues || []).length; i++) {
    const row = rawValues[i];
    const st = String(row[idxStatus] || '').trim();
    if (!st) continue;

    let targets = (routingIndex[st] || []).slice();
    if (!targets.length) continue;

    if (idxScName != null && (targets.indexOf(scFarhanName) > -1 || targets.indexOf(scMeilaniName) > -1 || targets.indexOf(scIvanName) > -1)) {
      const scNameVal = row[idxScName];
      if (typeof filterScTargets05b_ === 'function') {
        targets = filterScTargets05b_(targets, scNameVal, scFarhanName, scMeilaniName, scIvanName, kwFarhan, kwMeilani, kwIvan);
      } else {
        targets = preflightFilterScTargets06_(targets, scNameVal, scFarhanName, scMeilaniName, scIvanName, kwFarhan, kwMeilani, kwIvan);
      }
    }

    if (targets.length) total++;
  }

  return { total: total, reason: total ? 'ok' : 'no_routable_rows' };
}


/** =========================
 * SC sheet enrichments: Branch autofill + Finish type tagging
 * ========================= */

function __getScSheetNames06_() {
  const opsPolicy = (CONFIG && (CONFIG.opsRouting || CONFIG.opsRoutingPolicy || CONFIG.OPS_ROUTING_POLICY)) || null;
  const scFarhanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_FARHAN) ? opsPolicy.SHEETS.SC_FARHAN : 'SC - Farhan';
  const scMeilaniName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_MEILANI) ? opsPolicy.SHEETS.SC_MEILANI : 'SC - Meilani';
  const scIvanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_IVAN) ? opsPolicy.SHEETS.SC_IVAN : 'SC - Meindar';
  return [scFarhanName, scMeilaniName, scIvanName].filter(Boolean);
}

function __getBranchFromServiceCenter06_(serviceCenter) {
  const s = __normalizeHeaderText06_(serviceCenter).toLowerCase();
  if (!s) return '';
  if (s.indexOf('ez care') > -1 || s.indexOf('ezcare') > -1) return 'EzCare';

  const rules = [
    ['Mitracare', 'mitracare'],
    ['Sitcomtara', 'sitcomtara'],
    ['iBox', 'ibox'],
    ['GSI', 'gsi'],
    ['Rejeki Seluler', 'rejeki seluler'],
    ['Rejeki Seluler', 'rejeki seluller'],
    ['Andalas', 'andalas'],
    ['Klikcare', 'klikcare'],
    ['J-Bros', 'j-bros'],
    ['Makmur Era Abadi', 'makmur era abadi'],
    ['Manado Mitra Bersama', 'manado mitra bersama'],
    ['CV Kayu Awet Sejahtera', 'cv kayu awet sejahtera'],
    ['GH Store', 'gh store'],
    ['Unicom', 'unicom'],
    ['Samsung Authorized Service Centre by Unicom', 'samsung authorized service centre by unicom'],
    ['Xiaomi Authorized', 'xiaomi authorized'],
    ['Samsung Exclusive', 'samsung exclusive'],
    ['Carlcare', 'carlcare'],
    ['CV Berkah', 'cv berkah athallah'],
    ['CV Berkah', 'cv berkah'],
    ['B-Store', 'b-store'],
    ['MDP', 'mdp'],
    ['Deltafone', 'deltasindo']
  ];

  for (let i = 0; i < rules.length; i++) {
    if (s.indexOf(rules[i][1]) > -1) return rules[i][0];
  }
  return '';
}

function __getMiddlePicFromServiceCenter06_(serviceCenter) {
  const sc = __normalizeScKeywordText06c_(serviceCenter);
  if (!sc) return 'Unknown';

  const map = [
    ['Farhan', ['mitracare', 'sitcomtara', 'ibox', 'rejeki seluler', 'rejeki seluller']],
    ['Meilani', ['unicom', 'xiaomi authorized', 'samsung exclusive', 'carlcare', 'andalas', 'gsi']],
    ['Meindar', ['klikcare', 'j-bros', 'makmur era abadi', 'manado mitra bersama', 'cv kayu awet sejahtera', 'gh store', 'mdp', 'deltasindo', 'ezcare', 'ez care', 'b-store']],
  ];

  for (let i = 0; i < map.length; i++) {
    const pic = map[i][0];
    const keys = map[i][1];
    for (let j = 0; j < keys.length; j++) {
      if (sc.indexOf(keys[j]) > -1) return pic;
    }
  }

  return 'Unknown';
}

function __getReportBasePicFromPosition06_(position, serviceCenter) {
  const pos = String(position || '').trim().toLowerCase();
  if (pos === 'front') return 'Adi & Yudha';
  if (pos === 'expedition') return 'Adit';
  if (pos === 'back') return 'Suci & Detha';

  const middlePic = __getMiddlePicFromServiceCenter06_(serviceCenter);
  if (pos === 'middle') return middlePic;

  // Fallback: when position is missing/unmapped but SC keyword is known,
  // still assign middle PIC instead of leaving Unknown.
  if (middlePic !== 'Unknown') return middlePic;
  return 'Unknown';
}

function autofillBranchInScSheets06_(ss) {
  if (DRY_RUN) return 0;
  if (!ss) return 0;

  const sheetNames = __getScSheetNames06_();
  let filled = 0;

  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);
    const idxSc = __findHeaderIndexFlexible06_(header, 'Service Center');
    const idxBranch = __findHeaderIndexFlexible06_(header, 'Branch');
    if (idxSc === -1 || idxBranch === -1) continue;

    const n = lastRow - 1;
    const scVals = sh.getRange(2, idxSc + 1, n, 1).getValues();
    const branchVals = sh.getRange(2, idxBranch + 1, n, 1).getValues();

    let changed = false;
    const out = new Array(n);

    for (let r = 0; r < n; r++) {
      const cur = __normalizeHeaderText06_(branchVals[r][0]);
      if (cur) {
        out[r] = [branchVals[r][0]];
        continue;
      }

      const br = __getBranchFromServiceCenter06_(scVals[r][0]);
      if (br) {
        out[r] = [br];
        filled++;
        changed = true;
      } else {
        out[r] = [''];
      }
    }

    if (changed) {
      try { sh.getRange(2, idxBranch + 1, n, 1).setValues(out); } catch (e) {}
    }
  }

  return filled;
}

function __getReportBaseSourceSheets06_(ss) {
  const out = [];
  const seen = Object.create(null);
  const base = (CONFIG && CONFIG.sheetsByPic && Array.isArray(CONFIG.sheetsByPic.adminOperational))
    ? CONFIG.sheetsByPic.adminOperational.slice()
    : ['Submission','Ask Detail','OR - OLD','Start','Finish','Reject Claim','SC - Farhan','SC - Meilani','SC - Meindar','PO','Exclusion'];
  const extras = ['SC - Unmapped', 'B2B', 'EV-Bike', 'Special Case'];
  const all = base.concat(extras);
  for (let i = 0; i < all.length; i++) {
    const name = String(all[i] || '').trim();
    if (!name || seen[name]) continue;
    if (name === 'Raw Data') continue;
    if (ss && !ss.getSheetByName(name)) continue;
    seen[name] = true;
    out.push(name);
  }
  return out;
}

function __parseAnyDateReportBase06_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  const s = String(v || '').trim();
  if (!s) return null;
  if (typeof normalizeDate_ === 'function') {
    try {
      const d0 = normalizeDate_(s);
      if (d0 && !isNaN(d0.getTime())) return d0;
    } catch (e0) {}
  }
  if (typeof parseClaimLastUpdatedDatetime06c_ === 'function') {
    try {
      const d1 = parseClaimLastUpdatedDatetime06c_(s);
      if (d1 && !isNaN(d1.getTime())) return d1;
    } catch (e1) {}
  }
  if (typeof tryNativeParseUnambiguousDate_ === 'function') {
    try {
      const d2 = tryNativeParseUnambiguousDate_(s);
      if (d2 && !isNaN(d2.getTime())) return d2;
    } catch (e2) {}
  }
  return null;
}

function __formatSubmissionMonthReportBase06_(submissionDateVal) {
  const d = __parseAnyDateReportBase06_(submissionDateVal);
  if (!d) return '';
  return new Date(d.getFullYear(), d.getMonth(), 1, 0, 0, 0, 0);
}

function __ensureHeaderAtColumn06_(sh, headerName, targetCol) {
  if (!sh || !headerName || !targetCol || targetCol < 1) return false;
  if (DRY_RUN) return false;
  let lastCol = Math.max(sh.getLastColumn() || 0, 1);
  while (lastCol < targetCol - 1) {
    sh.insertColumnAfter(lastCol);
    lastCol = sh.getLastColumn();
  }
  const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), targetCol)).getValues()[0].map(v => String(v || '').trim());
  const curIdx = hdr.indexOf(String(headerName).trim()) + 1;
  if (curIdx === targetCol) return false;

  if (curIdx > 0) {
    sh.moveColumns(sh.getRange(1, curIdx, sh.getMaxRows(), 1), targetCol);
  } else {
    sh.insertColumnBefore(targetCol);
    sh.getRange(1, targetCol).setValue(headerName);
  }
  return true;
}

function __removeHeaderColumns06_(sh, headersToRemove, keepFirstByHeader) {
  if (!sh || DRY_RUN) return 0;
  const lc = sh.getLastColumn();
  if (lc < 1) return 0;
  const hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v || '').trim());
  const normalizedKeep = Object.create(null);
  Object.keys(keepFirstByHeader || {}).forEach(function(k) {
    normalizedKeep[__normalizeHeaderText06_(k)] = !!keepFirstByHeader[k];
  });
  const removeSet = new Set((headersToRemove || []).map(function(h) { return __normalizeHeaderText06_(h); }).filter(Boolean));
  const firstSeen = Object.create(null);
  const toDelete = [];
  for (let i = 0; i < hdr.length; i++) {
    const key = __normalizeHeaderText06_(hdr[i]);
    if (!key || !removeSet.has(key)) continue;
    if (normalizedKeep[key]) {
      if (!firstSeen[key]) {
        firstSeen[key] = true;
        continue;
      }
    }
    toDelete.push(i + 1);
  }
  for (let i = toDelete.length - 1; i >= 0; i--) {
    try { sh.deleteColumn(toDelete[i]); } catch (e) {}
  }
  return toDelete.length;
}

function __renameHeaderColumns06_(sh, renameMap) {
  if (!sh || DRY_RUN) return 0;
  const lc = sh.getLastColumn();
  if (lc < 1) return 0;
  const hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v || '').trim());
  const map = renameMap || {};
  let touched = 0;
  for (let i = 0; i < hdr.length; i++) {
    const key = __normalizeHeaderText06_(hdr[i]);
    const next = map[key];
    if (!next || hdr[i] === next) continue;
    try {
      sh.getRange(1, i + 1).setValue(next);
      touched++;
    } catch (e) {}
  }
  return touched;
}

function __fillBranchFromServiceCenter06_(sh) {
  if (!sh || DRY_RUN) return 0;
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return 0;
  const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  const idxSc = __findHeaderIndexFlexible06_(header, 'Service Center');
  const idxBranch = __findHeaderIndexFlexible06_(header, 'Branch');
  if (idxSc === -1 || idxBranch === -1) return 0;
  const n = lr - 1;
  const scVals = sh.getRange(2, idxSc + 1, n, 1).getValues();
  const branchVals = sh.getRange(2, idxBranch + 1, n, 1).getValues();
  const out = new Array(n);
  let touched = 0;
  for (let i = 0; i < n; i++) {
    const cur = String(branchVals[i][0] || '').trim();
    if (cur) { out[i] = [branchVals[i][0]]; continue; }
    const fill = __getBranchFromServiceCenter06_(scVals[i][0]);
    out[i] = [fill || ''];
    if (fill) touched++;
  }
  if (touched > 0) {
    try { sh.getRange(2, idxBranch + 1, n, 1).setValues(out); } catch (e) {}
  }
  return touched;
}

function __normalizeSubmissionByMonthColumn06_(sh) {
  if (!sh || DRY_RUN) return 0;
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return 0;
  const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  const idx = __findHeaderIndexFlexible06_(header, 'Submission by Month');
  if (idx === -1) return 0;

  const n = lr - 1;
  const rg = sh.getRange(2, idx + 1, n, 1);
  const vals = rg.getValues();
  let touched = 0;
  for (let i = 0; i < vals.length; i++) {
    const cur = vals[i][0];
    const norm = (typeof toSubmissionMonthDate06b_ === 'function')
      ? toSubmissionMonthDate06b_(cur)
      : '';
    const curTs = (Object.prototype.toString.call(cur) === '[object Date]' && !isNaN(cur.getTime()))
      ? new Date(cur.getFullYear(), cur.getMonth(), 1, 0, 0, 0, 0).getTime()
      : null;
    const normTs = (Object.prototype.toString.call(norm) === '[object Date]' && !isNaN(norm.getTime())) ? norm.getTime() : null;
    if (curTs !== normTs) {
      vals[i][0] = norm || '';
      touched++;
    }
  }
  if (touched > 0) {
    try { rg.setValues(vals); } catch (e) {}
  }
  try { rg.setNumberFormat('MMM yy'); } catch (e2) {}
  return touched;
}

function enforceOperationalLayout06_(ss) {
  if (!ss || DRY_RUN) return { touched: 0 };
  try {
    const oldExpired = ss.getSheetByName('Claim Expired');
    const newExpired = ss.getSheetByName('Expired Claim');
    if (oldExpired && !newExpired) oldExpired.setName('Expired Claim');
  } catch (eRenameExpired) {}
  const monthSheets = ['Submission', 'Ask Detail', 'Start', 'SC - Farhan', 'SC - Meilani', 'SC - Meindar', 'Finish', 'Expired Claim', 'Reject Claim', 'PO', 'Exclusion'];
  let touched = 0;
  for (let i = 0; i < monthSheets.length; i++) {
    const sh = ss.getSheetByName(monthSheets[i]);
    if (!sh) continue;
    if (__ensureHeaderAtColumn06_(sh, 'Submission by Month', 2)) touched++;
    touched += __normalizeSubmissionByMonthColumn06_(sh);
  }
  ['Start', 'Finish', 'Expired Claim', 'Reject Claim', 'PO'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    if (__ensureHeaderAtColumn06_(sh, 'Service Center PIC', 14)) touched++;
    if (__ensureHeaderAtColumn06_(sh, 'Branch', 15)) touched++;
  });

  const deprecatedEverywhere = ['DB', 'Status Type', 'Update Status Asso', 'Timestamp Asso', 'Update Status Admin', 'Timestamp Admin'];
  const financeExcluded = ['Claim Amount', 'Claim Own Risk Amount', 'Nett Claim Amount', '% Approval'];
  const evDossB2bDeprecated = ['Status Type', 'Start Date', 'End Date', 'Details'];
  const stageRename = {
    'Service Type': 'Claim Type',
    'Aging Position': 'Stage Aging',
    'Aging Post.': 'Stage Aging',
    'Aging Post': 'Stage Aging'
  };

  const allCleanupSheets = ['Submission', 'Ask Detail', 'OR - OLD', 'Start', 'Finish', 'Expired Claim', 'Reject Claim', 'SC - Farhan', 'SC - Meilani', 'SC - Meindar', 'SC - Unmapped', 'PO', 'Exclusion', 'B2B', 'EV-Bike', 'Doss', 'Special Case'];
  allCleanupSheets.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    touched += __removeHeaderColumns06_(sh, deprecatedEverywhere, {});
    touched += __renameHeaderColumns06_(sh, stageRename);
  });

  ['Submission', 'Ask Detail', 'Start', 'Finish', 'Expired Claim', 'Reject Claim'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    touched += __removeHeaderColumns06_(sh, financeExcluded, {});
    if (name === 'Submission') touched += __removeHeaderColumns06_(sh, ['Start Date', 'End Date', 'Details', 'Submission Date', 'Stage Aging', 'Aging Position', 'Aging Post.', 'Aging Post'], { 'Submission Date': true });
  });

  ['EV-Bike', 'Doss', 'B2B'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    touched += __removeHeaderColumns06_(sh, evDossB2bDeprecated, {});
  });

  ['Start', 'Finish', 'Expired Claim', 'Reject Claim'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh) touched += __fillBranchFromServiceCenter06_(sh);
  });

  return { touched: touched };
}

function refreshReportBaseFromOperational06_(ss, opts) {
  if (!ss) return { written: 0, skipped: 'missing spreadsheet' };
  const sh = ss.getSheetByName('Daily Report Base') || ss.getSheetByName('Report Base');
  if (!sh) return { written: 0, skipped: 'Daily Report Base / Report Base not found' };

  // Hard reset filter before rewrite/upsert to avoid stale hidden-row behavior during writes.
  try {
    const f0 = sh.getFilter ? sh.getFilter() : null;
    if (f0) f0.remove();
  } catch (eF0) {}
  if (DRY_RUN) return { written: 0, skipped: 'DRY_RUN' };
  const incremental = !!(opts && opts.incremental);

  const headers = [
    'Submission Date',
    'Submission by Month',
    'Claim Number',
    'Last Status',
    'Last Status Date',
    'Service Center',
    'Branch',
    'Position',
    'PIC',
    'Position Detail',
    'Position Detail Order',
    'Status Aging Days',
    'Status Aging Bucket',
    'Submission Aging Days',
    'Submission Aging Bucket'
  ];

  const srcSheets = __getReportBaseSourceSheets06_(ss);
  const byClaim = Object.create(null);
  function hasVal_(v) { return String(v == null ? '' : v).trim() !== ''; }
  function pickValue_(primary, fallback) { return hasVal_(primary) ? primary : fallback; }

  for (let si = 0; si < srcSheets.length; si++) {
    const name = srcSheets[si];
    const src = ss.getSheetByName(name);
    if (!src) continue;
    const lr = src.getLastRow();
    const lc = src.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const hdr = src.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxClaim = __findHeaderIndexFlexible06_(hdr, 'Claim Number');
    if (idxClaim === -1) continue;

    const idxSubDate = (function() {
      const cands = ['Submission Date', 'claim_submission_date', 'Claim Submission Date', 'Claim Submitted Datetime'];
      for (let ci = 0; ci < cands.length; ci++) {
        const ix = __findHeaderIndexFlexible06_(hdr, cands[ci]);
        if (ix !== -1) return ix;
      }
      return -1;
    })();
    const idxLast = __findHeaderIndexFlexible06_(hdr, 'Last Status');
    const idxLastDate = __findHeaderIndexFlexible06_(hdr, 'Last Status Date');
    const idxSc = (function() {
      const cands = ['Service Center', 'Service Center Name', 'SC Name', 'sc_name'];
      for (let ci = 0; ci < cands.length; ci++) {
        const ix = __findHeaderIndexFlexible06_(hdr, cands[ci]);
        if (ix !== -1) return ix;
      }
      return -1;
    })();
    const idxLsa = (function() {
      const cands = ['Last Status Aging', 'LSA'];
      for (let ci = 0; ci < cands.length; ci++) {
        const ix = __findHeaderIndexFlexible06_(hdr, cands[ci]);
        if (ix !== -1) return ix;
      }
      return -1;
    })();
    const idxTat = __findHeaderIndexFlexible06_(hdr, 'TAT');
    const vals = src.getRange(2, 1, lr - 1, lc).getValues();

    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      const claim = String(row[idxClaim] || '').trim();
      if (!claim) continue;

      const lastStatus = (idxLast !== -1) ? String(row[idxLast] || '').trim() : '';
      const subDateVal = (idxSubDate !== -1) ? row[idxSubDate] : '';
      const lastDateVal = (idxLastDate !== -1) ? row[idxLastDate] : '';
      const scVal = (idxSc !== -1) ? row[idxSc] : '';
      const lsaVal = (idxLsa !== -1) ? row[idxLsa] : '';
      const tatVal = (idxTat !== -1) ? row[idxTat] : '';

      const lastDateObj = __parseAnyDateReportBase06_(lastDateVal);
      const lastDateTs = lastDateObj ? lastDateObj.getTime() : -1;
      const key = claim.toUpperCase();
      const prev = byClaim[key];
      if (prev && prev.lastDateTs > lastDateTs) continue;

      let position = '';
      if (name === 'Exclusion') position = 'Closed';
      else if (typeof getPositionFromLastStatus_ === 'function') {
        try { position = getPositionFromLastStatus_(lastStatus); } catch (eP) { position = ''; }
      }

      const candidate = {
        subDate: __parseAnyDateReportBase06_(subDateVal) || subDateVal || '',
        subMonth: __formatSubmissionMonthReportBase06_(subDateVal),
        claim: claim,
        lastStatus: lastStatus,
        lastDate: lastDateObj || lastDateVal || '',
        serviceCenter: String(scVal || '').trim(),
        branch: __getBranchFromServiceCenter06_(scVal || ''),
        position: position || '',
        pic: __getReportBasePicFromPosition06_(position || '', scVal || ''),
        statusAgingDays: __toNonNegativeIntReportBase06_(lsaVal),
        submissionAgingDays: __toNonNegativeIntReportBase06_(tatVal),
        lastDateTs: lastDateTs
      };

      if (prev) {
        candidate.subDate = pickValue_(candidate.subDate, prev.subDate);
        candidate.subMonth = pickValue_(candidate.subMonth, prev.subMonth);
        candidate.serviceCenter = pickValue_(candidate.serviceCenter, prev.serviceCenter);
        candidate.branch = pickValue_(candidate.branch, prev.branch);
        candidate.position = pickValue_(candidate.position, prev.position);
        candidate.pic = pickValue_(candidate.pic, prev.pic);
        candidate.statusAgingDays = pickValue_(candidate.statusAgingDays, prev.statusAgingDays);
        candidate.submissionAgingDays = pickValue_(candidate.submissionAgingDays, prev.submissionAgingDays);
      }

      candidate.pic = __getReportBasePicFromPosition06_(candidate.position, candidate.serviceCenter);

      byClaim[key] = candidate;
    }
  }

  const rows = Object.keys(byClaim).map(function(k) {
    const x = byClaim[k];
    const positionDetail = __buildPositionDetailReportBase06_(x.position, x.pic);
    return [
      x.subDate,
      x.subMonth,
      x.claim,
      x.lastStatus,
      x.lastDate,
      x.serviceCenter,
      x.branch,
      x.position,
      x.pic,
      positionDetail,
      __getPositionDetailOrderReportBase06_(positionDetail),
      (x.statusAgingDays == null) ? '' : x.statusAgingDays,
      __getStatusAgingBucketReportBase06_(x.statusAgingDays),
      (x.submissionAgingDays == null) ? '' : x.submissionAgingDays,
      __getSubmissionAgingBucketReportBase06_(x.submissionAgingDays)
    ];
  });

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  try { sh.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center').setVerticalAlignment('middle'); } catch (e0) {}

  if (!incremental) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
      try { sh.getRange(2, 1, rows.length, 1).setNumberFormat('dd MMM yy'); } catch (e1) {}
      try { sh.getRange(2, 2, rows.length, 1).setNumberFormat('MMM yy'); } catch (eM1) {}
      try { sh.getRange(2, 5, rows.length, 1).setNumberFormat('dd MMM yy, HH:mm'); } catch (e2) {}
    }
    return { written: rows.length, sheets: srcSheets.length, mode: 'full-rewrite' };
  }

  // Incremental mode: upsert by Claim Number and keep untouched historical rows.
  const lc = Math.max(sh.getLastColumn(), headers.length);
  const lr = sh.getLastRow();
  let existing = [];
  if (lr >= 2 && lc >= headers.length) {
    existing = sh.getRange(2, 1, lr - 1, headers.length).getValues();
  }
  const idxByClaim = Object.create(null);
  for (let i = 0; i < existing.length; i++) {
    const cn = String(existing[i][2] || '').trim(); // Claim Number column
    if (cn && idxByClaim[cn.toUpperCase()] == null) idxByClaim[cn.toUpperCase()] = i;
  }

  let upserted = 0;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const cn = String(row[2] || '').trim();
    if (!cn) continue;
    const k = cn.toUpperCase();
    if (idxByClaim[k] != null) existing[idxByClaim[k]] = row;
    else {
      idxByClaim[k] = existing.length;
      existing.push(row);
    }
    upserted++;
  }

  if (existing.length) {
    sh.getRange(2, 1, existing.length, headers.length).setValues(existing);
    try { sh.getRange(2, 1, existing.length, 1).setNumberFormat('dd MMM yy'); } catch (e3) {}
    try { sh.getRange(2, 2, existing.length, 1).setNumberFormat('MMM yy'); } catch (eM2) {}
    try { sh.getRange(2, 5, existing.length, 1).setNumberFormat('dd MMM yy, HH:mm'); } catch (e4) {}
  }
  try { __expandSheetFilterToUsedRange06_(sh); } catch (eF3) {}
  return { written: upserted, totalRows: existing.length, sheets: srcSheets.length, mode: 'incremental-upsert' };
}

function extractSnapshotDateFromFileName_(sourceFileName) {
  const s = String(sourceFileName || '').trim();
  if (!s) return null;
  const m = s.match(/(\d{4}-\d{2}-\d{2})T/);
  return m ? m[1] : null;
}

function fillWeeklyReportBase(snapshotDateOverride, sourceFileName, ssOverride) {
  let ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    try {
      const fallbackId = String((CONFIG && CONFIG.masterSpreadsheetId) ? CONFIG.masterSpreadsheetId : '').trim();
      if (fallbackId) ss = SpreadsheetApp.openById(fallbackId);
    } catch (e0) {}
  }
  if (!ss) throw new Error('fillWeeklyReportBase: Spreadsheet tidak ditemukan.');
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Jakarta';
  const daily = ss.getSheetByName('Daily Report Base');
  if (!daily) throw new Error('fillWeeklyReportBase: sheet "Daily Report Base" tidak ditemukan.');

  const weeklyName = 'Weekly Report Base';
  const weekly = ss.getSheetByName(weeklyName) || ss.insertSheet(weeklyName);
  const headers = [
    'Snapshot Run At','Snapshot Date','Submission by Month','Branch','PIC','Position','Position Detail',
    'Position Detail Order','Last Status','Count','Previous Snapshot Date','Previous Count','Daily Change',
    'Is Last 7 Days','Position Order','Source File Name'
  ];

  const now = new Date();
  const snapshotKey = (function() {
    const fromOverride = String(snapshotDateOverride || '').trim();
    if (fromOverride) return fromOverride;
    const fromFile = extractSnapshotDateFromFileName_(sourceFileName);
    if (fromFile) return fromFile;
    return Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  })();
  const snapshotDate = new Date(snapshotKey + 'T00:00:00');
  if (isNaN(snapshotDate.getTime())) throw new Error('fillWeeklyReportBase: Snapshot Date invalid: ' + snapshotKey);

  const lr = daily.getLastRow();
  const lc = daily.getLastColumn();
  if (lr < 2 || lc < 1) throw new Error('fillWeeklyReportBase: Daily Report Base kosong.');
  const h = daily.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  function req(header) {
    const i = __findHeaderIndexFlexible06_(h, header);
    if (i === -1) throw new Error('fillWeeklyReportBase: kolom wajib tidak ditemukan di Daily Report Base: ' + header);
    return i;
  }
  const idxClaim = req('Claim Number');
  const idxSubMonth = req('Submission by Month');
  const idxBranch = req('Branch');
  const idxPic = req('PIC');
  const idxPos = req('Position');
  const idxLast = req('Last Status');
  const idxPosDet = __findHeaderIndexFlexible06_(h, 'Position Detail');
  const idxPosDetOrder = __findHeaderIndexFlexible06_(h, 'Position Detail Order');
  const idxPosOrder = __findHeaderIndexFlexible06_(h, 'Position Order');

  const vals = daily.getRange(2, 1, lr - 1, lc).getValues();
  const agg = new Map();
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const claim = String(r[idxClaim] || '').trim();
    if (!claim) continue;
    const subMonth = __formatSubmissionMonthReportBase06_(r[idxSubMonth]);
    if (!subMonth) continue;
    const branch = String(r[idxBranch] || '').trim();
    const pic = __toTitleCaseReportBase06_(r[idxPic]);
    const posRaw = String(r[idxPos] || '').trim();
    const pos = __toTitleCaseReportBase06_(posRaw);
    const lastStatus = String(r[idxLast] || '').trim();
    if (!lastStatus) continue;
    const posDetailRaw = (idxPosDet !== -1) ? String(r[idxPosDet] || '').trim() : '';
    const posDetail = posDetailRaw || __buildPositionDetailReportBase06_(pos, pic);
    const pdoRaw = (idxPosDetOrder !== -1) ? r[idxPosDetOrder] : '';
    const pdoNum = (pdoRaw === '' || pdoRaw == null || !isFinite(Number(pdoRaw))) ? __getPositionDetailOrderReportBase06_(posDetail) : Number(pdoRaw);
    const poRaw = (idxPosOrder !== -1) ? r[idxPosOrder] : '';
    const poNum = (poRaw === '' || poRaw == null || !isFinite(Number(poRaw))) ? __getPositionOrderWeekly06_(pos) : Number(poRaw);
    const subMonthKey = Utilities.formatDate(subMonth, tz, 'yyyy-MM-dd');
    const key = [snapshotKey, subMonthKey, branch, pic, pos, posDetail, lastStatus].join('|');
    if (!agg.has(key)) {
      agg.set(key, { subMonth: subMonth, branch: branch, pic: pic, pos: pos, posDetail: posDetail, pdo: pdoNum, lastStatus: lastStatus, count: 0, po: poNum });
    }
    agg.get(key).count += 1;
  }

  const wr = weekly.getLastRow();
  const wc = Math.max(weekly.getLastColumn(), headers.length);
  let existing = [];
  if (wr >= 2) existing = weekly.getRange(2, 1, wr - 1, wc).getValues();
  const idxW = {};
  const wh = (wr >= 1) ? weekly.getRange(1, 1, 1, wc).getValues()[0] : [];
  for (let i = 0; i < headers.length; i++) idxW[headers[i]] = i;
  if (wh.length < headers.length || headers.some((x, i) => String(wh[i] || '').trim() !== x)) {
    weekly.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  function dKey(v) { const d = __parseAnyDateReportBase06_(v); return d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; }
  const keep = existing.filter(r => dKey(r[idxW['Snapshot Date']]) !== snapshotKey);
  const allDates = Array.from(new Set(keep.map(r => dKey(r[idxW['Snapshot Date']])).filter(Boolean))).sort();
  const prevDateKey = allDates.filter(d => d < snapshotKey).slice(-1)[0] || '';

  const prevMap = new Map();
  if (prevDateKey) {
    keep.forEach(function(r) {
      if (dKey(r[idxW['Snapshot Date']]) !== prevDateKey) return;
      const k = [
        dKey(r[idxW['Submission by Month']]),
        String(r[idxW['Branch']] || '').trim(),
        String(r[idxW['PIC']] || '').trim(),
        String(r[idxW['Position']] || '').trim(),
        String(r[idxW['Position Detail']] || '').trim(),
        String(r[idxW['Last Status']] || '').trim()
      ].join('|');
      prevMap.set(k, r);
    });
  }

  const runAt = now;
  const srcName = String(sourceFileName || '');
  const currentRows = [];
  agg.forEach(function(v) {
    currentRows.push([runAt, snapshotDate, v.subMonth, v.branch, v.pic, v.pos, v.posDetail, v.pdo, v.lastStatus, v.count, '', '', '', false, v.po, srcName]);
  });
  if (prevDateKey) {
    const curKeys = new Set(currentRows.map(r => [Utilities.formatDate(r[2], tz, 'yyyy-MM-dd'), r[3], r[4], r[5], r[6], r[8]].join('|')));
    prevMap.forEach(function(prevRow, k) {
      if (curKeys.has(k)) return;
      currentRows.push([runAt, snapshotDate, __parseAnyDateReportBase06_(prevRow[idxW['Submission by Month']]), prevRow[idxW['Branch']], prevRow[idxW['PIC']], prevRow[idxW['Position']], prevRow[idxW['Position Detail']], Number(prevRow[idxW['Position Detail Order']] || 99), prevRow[idxW['Last Status']], 0, '', '', '', false, Number(prevRow[idxW['Position Order']] || 99), srcName]);
    });
  }

  const all = keep.concat(currentRows);
  __recalculateWeeklyHelpers06_(all, tz, idxW);
  all.sort(function(a, b) {
    return __cmpWeeklySort06_(a, b, idxW, tz);
  });

  weekly.clearContents();
  weekly.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (all.length) weekly.getRange(2, 1, all.length, headers.length).setValues(all.map(r => r.slice(0, headers.length)));
  __formatWeeklyReportBase06_(weekly, Math.max(2, all.length + 1));
  try { __expandSheetFilterToUsedRange06_(weekly); } catch (eFwk) {}
}

function __expandSheetFilterToUsedRange06_(sh) {
  if (!sh || typeof sh.getFilter !== 'function') return false;
  const filter = sh.getFilter();
  if (!filter || !filter.getRange) return false;
  const oldRange = filter.getRange();
  const oldRow = oldRange.getRow();
  const oldCol = oldRange.getColumn();
  const oldCols = oldRange.getNumColumns();
  const oldRows = oldRange.getNumRows();
  if (oldRow !== 1 || oldCols < 1) return false;

  const criteriaByAbsCol = {};
  for (let rel = 1; rel <= oldCols; rel++) {
    const c = filter.getColumnFilterCriteria(rel);
    if (c) criteriaByAbsCol[oldCol + rel - 1] = c;
  }

  const lastCol = Math.max(sh.getLastColumn(), oldCol + oldCols - 1);
  const lastRow = Math.max(sh.getLastRow(), 1);
  const needResize = (oldCol !== 1) || (oldCols !== lastCol) || (oldRows !== lastRow);
  if (!needResize) return false;

  filter.remove();
  sh.getRange(1, 1, lastRow, lastCol).createFilter();
  const nf = sh.getFilter();
  if (!nf) return true;
  Object.keys(criteriaByAbsCol).forEach(function(absStr) {
    const abs = Number(absStr);
    if (!isFinite(abs) || abs < 1 || abs > lastCol) return;
    try { nf.setColumnFilterCriteria(abs, criteriaByAbsCol[abs]); } catch (e) {}
  });
  return true;
}

function __expandWorkbookFiltersToUsedRange06_(ss, sheetNames) {
  if (!ss) return { touched: 0, checked: 0 };
  const names = [];
  const seen = Object.create(null);
  if (Array.isArray(sheetNames) && sheetNames.length) {
    sheetNames.forEach(function(name) {
      const n = String(name || '').trim();
      if (!n || seen[n]) return;
      seen[n] = true;
      names.push(n);
    });
  } else if (typeof ss.getSheets === 'function') {
    ss.getSheets().forEach(function(sh) {
      const n = sh && sh.getName ? sh.getName() : '';
      if (!n || seen[n]) return;
      seen[n] = true;
      names.push(n);
    });
  }

  let touched = 0;
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    try {
      if (__expandSheetFilterToUsedRange06_(sh)) touched++;
    } catch (e) {}
  }
  return { touched: touched, checked: names.length };
}

function __getPositionOrderWeekly06_(position) {
  const p = String(position || '').trim().toLowerCase();
  if (p === 'front') return 1;
  if (p === 'expedition') return 2;
  if (p === 'middle') return 3;
  if (p === 'back') return 4;
  if (p === 'closed') return 5;
  return 99;
}

function __recalculateWeeklyHelpers06_(rows, tz, idxW) {
  const dateSet = Array.from(new Set(rows.map(r => Utilities.formatDate(__parseAnyDateReportBase06_(r[idxW['Snapshot Date']]), tz, 'yyyy-MM-dd')).filter(Boolean))).sort();
  const latest = dateSet[dateSet.length - 1] || '';
  const prevByDate = {};
  for (let i = 0; i < dateSet.length; i++) prevByDate[dateSet[i]] = i > 0 ? dateSet[i - 1] : '';

  const countByDateKey = new Map();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const d = Utilities.formatDate(__parseAnyDateReportBase06_(r[idxW['Snapshot Date']]), tz, 'yyyy-MM-dd');
    const k = [
      Utilities.formatDate(__parseAnyDateReportBase06_(r[idxW['Submission by Month']]), tz, 'yyyy-MM-dd'),
      String(r[idxW['Branch']] || '').trim(),
      String(r[idxW['PIC']] || '').trim(),
      String(r[idxW['Position']] || '').trim(),
      String(r[idxW['Position Detail']] || '').trim(),
      String(r[idxW['Last Status']] || '').trim()
    ].join('|');
    countByDateKey.set(d + '|' + k, Number(r[idxW['Count']] || 0));
  }

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const curDate = Utilities.formatDate(__parseAnyDateReportBase06_(r[idxW['Snapshot Date']]), tz, 'yyyy-MM-dd');
    const prevDate = prevByDate[curDate] || '';
    r[idxW['Previous Snapshot Date']] = prevDate ? new Date(prevDate + 'T00:00:00') : '';
    if (!prevDate) {
      r[idxW['Previous Count']] = '';
      r[idxW['Daily Change']] = '';
    } else {
      const key = [
        Utilities.formatDate(__parseAnyDateReportBase06_(r[idxW['Submission by Month']]), tz, 'yyyy-MM-dd'),
        String(r[idxW['Branch']] || '').trim(),
        String(r[idxW['PIC']] || '').trim(),
        String(r[idxW['Position']] || '').trim(),
        String(r[idxW['Position Detail']] || '').trim(),
        String(r[idxW['Last Status']] || '').trim()
      ].join('|');
      const prevCount = Number(countByDateKey.get(prevDate + '|' + key) || 0);
      r[idxW['Previous Count']] = prevCount;
      r[idxW['Daily Change']] = Number(r[idxW['Count']] || 0) - prevCount;
    }
    if (!latest) r[idxW['Is Last 7 Days']] = false;
    else {
      const d = new Date(curDate + 'T00:00:00');
      const l = new Date(latest + 'T00:00:00');
      const min = new Date(l.getTime() - (6 * 24 * 60 * 60 * 1000));
      r[idxW['Is Last 7 Days']] = d.getTime() >= min.getTime();
    }
  }
}

function __cmpWeeklySort06_(a, b, idxW, tz) {
  function ds(v) { return Utilities.formatDate(__parseAnyDateReportBase06_(v), tz, 'yyyy-MM-dd'); }
  const cands = [
    [ds(a[idxW['Snapshot Date']]), ds(b[idxW['Snapshot Date']])],
    [ds(a[idxW['Submission by Month']]), ds(b[idxW['Submission by Month']])],
    [String(a[idxW['Branch']] || ''), String(b[idxW['Branch']] || '')],
    [String(a[idxW['PIC']] || ''), String(b[idxW['PIC']] || '')],
    [Number(a[idxW['Position Detail Order']] || 99), Number(b[idxW['Position Detail Order']] || 99)],
    [String(a[idxW['Last Status']] || ''), String(b[idxW['Last Status']] || '')]
  ];
  for (let i = 0; i < cands.length; i++) {
    if (cands[i][0] < cands[i][1]) return -1;
    if (cands[i][0] > cands[i][1]) return 1;
  }
  return 0;
}

function __formatWeeklyReportBase06_(sh, lastRow) {
  try { sh.getRange(1, 1, 1, 16).setFontWeight('bold'); } catch (e0) {}
  try { sh.setFrozenRows(1); } catch (e1) {}
  if (lastRow < 2) return;
  try { sh.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); } catch (e2) {}
  try { sh.getRange(2, 2, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd'); } catch (e3) {}
  try { sh.getRange(2, 3, lastRow - 1, 1).setNumberFormat('mmm yyyy'); } catch (e4) {}
  try { sh.getRange(2, 8, lastRow - 1, 1).setNumberFormat('0.00'); } catch (e5) {}
  try { sh.getRange(2, 10, lastRow - 1, 1).setNumberFormat('#,##0'); } catch (e6) {}
  try { sh.getRange(2, 11, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd'); } catch (e7) {}
  try { sh.getRange(2, 12, lastRow - 1, 1).setNumberFormat('#,##0'); } catch (e8) {}
  try { sh.getRange(2, 13, lastRow - 1, 1).setNumberFormat('+#,##0;-#,##0;0'); } catch (e9) {}
  try { sh.getRange(2, 15, lastRow - 1, 1).setNumberFormat('0'); } catch (e10) {}
  try { sh.autoResizeColumns(1, 16); } catch (e11) {}
}

const POSITION_DETAIL_ORDER_MAP_06_ = Object.freeze({
  'Front': 1,
  'Expedition': 2,
  'Middle - Farhan': 3.1,
  'Middle - Meilani': 3.2,
  'Middle - Meindar': 3.3,
  'Back': 4,
  'Closed': 5
});

function __toNonNegativeIntReportBase06_(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  if (!isFinite(n)) return null;
  const out = Math.floor(n);
  return out < 0 ? null : out;
}

function __toTitleCaseReportBase06_(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (!s) return '';
  return s.split(/\s+/).map(function(part) {
    return part ? (part.charAt(0).toUpperCase() + part.slice(1)) : '';
  }).join(' ');
}

function __buildPositionDetailReportBase06_(position, pic) {
  const posNorm = String(position == null ? '' : position).trim().toLowerCase();
  if (!posNorm) return 'Unknown';
  if (posNorm === 'middle') {
    const picNorm = String(pic == null ? '' : pic).trim().toLowerCase();
    const canonicalPic = ({
      'farhan': 'Farhan',
      'meilani': 'Meilani',
      'meindar': 'Meindar'
    })[picNorm];
    const picClean = canonicalPic || __toTitleCaseReportBase06_(pic);
    return 'Middle - ' + (picClean || 'Unassigned');
  }
  return __toTitleCaseReportBase06_(position);
}

function __getPositionDetailOrderReportBase06_(positionDetail) {
  const pd = String(positionDetail == null ? '' : positionDetail).trim();
  if (!pd) return 99;
  if (POSITION_DETAIL_ORDER_MAP_06_[pd] != null) return POSITION_DETAIL_ORDER_MAP_06_[pd];
  if (pd.indexOf('Middle - ') === 0) {
    if (pd === 'Middle - Unassigned') return 3.99;
    return 3.9;
  }
  return 99;
}

function __getStatusAgingBucketReportBase06_(days) {
  if (days == null || days === '') return '';
  const d = Number(days);
  if (!isFinite(d) || d < 0) return '';
  if (d <= 1) return '01. 0-1 days';
  if (d <= 3) return '02. 2-3 days';
  if (d <= 7) return '03. 4-7 days';
  if (d <= 14) return '04. 8-14 days';
  if (d <= 30) return '05. 15-30 days';
  return '06. >30 days';
}

function __getSubmissionAgingBucketReportBase06_(days) {
  if (days == null || days === '') return '';
  const d = Number(days);
  if (!isFinite(d) || d < 0) return '';
  if (d <= 7) return '01. 0-7 days';
  if (d <= 14) return '02. 8-14 days';
  if (d <= 30) return '03. 15-30 days';
  if (d <= 60) return '04. 31-60 days';
  return '05. >60 days';
}

function __isFinishStatus06_(status) {
  const s = String(status || '').trim(); // may contain trailing spaces
  if (!s) return false;

  try {
    if (typeof FINISH_STATUSES !== 'undefined' && Array.isArray(FINISH_STATUSES)) {
      return FINISH_STATUSES.indexOf(s) !== -1;
    }
  } catch (e) {}

  // Fallback
  return (
    s === 'DONE_REPAIR' ||
    s === 'WAITING_WALKIN_FINISH' ||
    s === 'COURIER_PICKED_UP' ||
    s === 'WAITING_COURIER_FINISH' ||
    s === 'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH'
  );
}

function applyFinishTypeInScSheets06_(ss) {
  // Backward-compatible entrypoint:
  // Historically this only set Type=Finish for finish statuses.
  // As of 2026-02 policy, this fills Type for SC sheets based on Last Status mapping.
  if (DRY_RUN) return 0;
  if (!ss) return 0;

  const sheetNames = __getScSheetNames06_();
  let updated = 0;

  const opsPolicy =
    (CONFIG && (CONFIG.opsRouting || CONFIG.opsRoutingPolicy || CONFIG.OPS_ROUTING_POLICY)) ||
    (typeof OPS_ROUTING_POLICY !== 'undefined' ? OPS_ROUTING_POLICY : null);

  const typePolicy = (opsPolicy && opsPolicy.TYPE_BY_LAST_STATUS) ? opsPolicy.TYPE_BY_LAST_STATUS : null;
  if (!typePolicy) return 0;

  const sets = {
    onRep: new Set(typePolicy['SC - On Rep'] || []),
    waitRep: new Set(typePolicy['SC - Wait Rep'] || []),
    finish: new Set(typePolicy['Finish'] || []),
    orSet: new Set(typePolicy['OR'] || []),
    insurance: new Set(typePolicy['Insurance'] || []),
    est: new Set(typePolicy['SC - Est'] || []),
    rcvd: new Set(typePolicy['SC - Rcvd'] || [])
  };

  const resolveType = (statusVal) => {
    const s = String(statusVal || '').trim();
    if (!s) return '';
    if (sets.onRep.has(s)) return 'SC - On Rep';
    if (sets.waitRep.has(s)) return 'SC - Wait Rep';
    if (sets.finish.has(s)) return 'Finish';
    if (sets.orSet.has(s)) return 'OR';
    if (sets.insurance.has(s)) return 'Insurance';
    if (sets.est.has(s)) return 'SC - Est';
    if (sets.rcvd.has(s)) return 'SC - Rcvd';
    return '';
  };

  // Ensure Type dropdown exists (preserve dropdown-chip/colors by copying DV from a template cell).
  let typeDvTemplateCell = null;
  try {
    for (let ti = 0; ti < sheetNames.length; ti++) {
      const shT = ss.getSheetByName(sheetNames[ti]);
      if (!shT) continue;
      const lcT = shT.getLastColumn();
      if (lcT < 1) continue;
      const headerT = shT.getRange(1, 1, 1, lcT).getValues()[0].map(__normalizeHeaderText06_);
      const idxTypeT = __findHeaderIndexFlexible06_(headerT, 'Type');
      if (idxTypeT === -1) continue;
      const cell = shT.getRange(2, idxTypeT + 1, 1, 1);
      const dv = cell.getDataValidation();
      if (dv) { typeDvTemplateCell = cell; break; }
    }
  } catch (e) {}

  // Fallback list-based rule (may lose dropdown-chip styling if chips are required).
  const fallbackTypeOpts = (typeof getScTypeDropdownOptions_ === 'function')
    ? getScTypeDropdownOptions_()
    : ['SC - Rcvd','SC - Est','Insurance','OR','Finish','SC - Wait Rep','SC - On Rep'];

  const fallbackDvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(fallbackTypeOpts, true)
    .setAllowInvalid(true)
    .build();

  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);

    let idxStatus = __findHeaderIndexFlexible06_(header, 'Last Status');
    if (idxStatus === -1) idxStatus = __findHeaderIndexFlexible06_(header, 'Status'); // fallback
    const idxType = __findHeaderIndexFlexible06_(header, 'Type');
    if (idxStatus === -1 || idxType === -1) continue;

    // Add DV if missing (do not override)
    try {
      const dv = sh.getRange(2, idxType + 1).getDataValidation();
      if (!dv) {
        const buffer = 600;
        const maxRows = sh.getMaxRows();
        const endRow = Math.min(maxRows, Math.max(2, lastRow) + buffer);
        const rows = Math.max(0, endRow - 1);
        if (rows > 0) {
          const dst = sh.getRange(2, idxType + 1, rows, 1);
          if (typeDvTemplateCell) {
            try { typeDvTemplateCell.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); } catch (e1) {}
            try { typeDvTemplateCell.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e2) {}
          } else {
            try { dst.setDataValidation(fallbackDvRule); } catch (e3) {}
          }
        }
      }
    } catch (e) {}

    const n = lastRow - 1;
    const stVals = sh.getRange(2, idxStatus + 1, n, 1).getValues();
    const typeVals = sh.getRange(2, idxType + 1, n, 1).getValues();

    let changed = false;
    const out = new Array(n);

    for (let r = 0; r < n; r++) {
      const st = stVals[r][0];
      const want = resolveType(st);
      const cur = String(typeVals[r][0] || '');
      if (want && cur !== want) {
        out[r] = [want];
        updated++;
        changed = true;
      } else {
        out[r] = [typeVals[r][0]];
      }
    }

    if (changed) {
      try { sh.getRange(2, idxType + 1, n, 1).setValues(out); } catch (e) {}
    }
  }

  return updated;
}

/** Apply number formats for RAW columns by schema (bounded rows) */
function applyRawSchemaFormats_(rawSheet, rawHeader, nRows, buffer) {
  if (DRY_RUN) return;
  if (!rawSheet || !rawHeader) return;

  const typeReg =
    (typeof COLUMN_TYPES !== 'undefined' && COLUMN_TYPES && COLUMN_TYPES.RAW)
      ? COLUMN_TYPES.RAW
      : null;

  if (!typeReg) return;

  const F = _fmt06_();
  const idx = buildHeaderIndex_(rawHeader);
  const idxCanon = __buildCanonicalHeaderIndex06_(rawHeader);

  const maxDataRows = Math.max(1, (rawSheet.getMaxRows ? (rawSheet.getMaxRows() - 1) : 1));
  const want = Math.max(1, (nRows || 0) + (buffer || 0));
  const rows = Math.min(maxDataRows, want);

  Object.keys(typeReg).forEach(colName => {
    let c0 = idx[colName];
    if (c0 == null) {
      const k = __canonicalHeaderKey06_(colName);
      if (k && idxCanon[k] != null) c0 = idxCanon[k];
    }
    if (c0 == null) return;

    const t = String(typeReg[colName] || '').toUpperCase();
    let fmt = null;

    if (t === 'DATE') fmt = F.DATE;
    else if (t === 'DATETIME') fmt = F.DATETIME;
    else if (t === 'TIMESTAMP') fmt = F.TIMESTAMP;
    else if (t === 'INT') fmt = F.INT;
    else if (t === 'MONEY0') fmt = F.MONEY0;

    if (!fmt) return;

    try {
      rawSheet.getRange(2, c0 + 1, rows, 1).setNumberFormat(fmt);
    } catch (e) {}
  });
}

/**
 * Recompute Exclusion.TAT as the number of calendar days between Submission Date and Last Status Date.
 * - TAT = (Last Status Date) - (Submission Date)
 * - If either date is missing or the computed diff is negative, TAT is cleared.
 *
 * Returns the number of rows where TAT was filled.
 */
function recomputeExclusionTat_(ss, pic) {
  if (DRY_RUN) return 0;
  if (!ss) return 0;

  const sh = ss.getSheetByName('Exclusion');
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(__normalizeHeaderText06_);

  const idx = Object.create(null);
  for (let c = 0; c < header.length; c++) {
    const k = String(header[c] || '').toUpperCase();
    if (k) idx[k] = c;
  }

  function findCol_(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const k = String(candidates[i] || '').trim().toUpperCase();
      if (!k) continue;
      if (idx[k] != null) return idx[k];
    }
    return null;
  }

  const cSubmission = findCol_(['Submission Date', 'SUBMISSION DATE', 'Submission date', 'SubmissionDate']);
  const cLast = findCol_(['Last Status Date', 'LAST STATUS DATE', 'Last status date', 'LastStatusDate']);
  const cTat = findCol_(['TAT', 'TAT (Days)', 'Tat', 'Tat (Days)', 'TAT Days']);

  if (cSubmission == null || cLast == null || cTat == null) return 0;

  const n = lastRow - 1;
  const subVals = sh.getRange(2, cSubmission + 1, n, 1).getValues();
  const lastVals = sh.getRange(2, cLast + 1, n, 1).getValues();

  const out = new Array(n);
  let filled = 0;

  const F = _fmt06_();

  for (let r = 0; r < n; r++) {
    const sd = subVals[r][0];
    const ld = lastVals[r][0];

    let d = null;
    try {
      if (typeof diffDays_ === 'function') d = diffDays_(sd, ld);
      else d = __diffDays06_(sd, ld);
    } catch (e) {
      d = null;
    }

    if (d == null || isNaN(d) || d < 0) {
      out[r] = [''];
    } else {
      out[r] = [d];
      filled++;
    }
  }

  try { sh.getRange(2, cTat + 1, n, 1).setValues(out); } catch (e1) {}
  try { sh.getRange(2, cTat + 1, n, 1).setNumberFormat(F.INT); } catch (e2) {}

  return filled;
}

// Local fallback if diffDays_ is not available.
function __diffDays06_(d1, d2) {
  const a = __toDate06_(d1);
  const b = __toDate06_(d2);
  if (!a || !b) return null;
  const t1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const t2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.floor((t2 - t1) / 86400000);
}

function __toDate06_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;
  if (typeof normalizeDate_ === 'function') {
    try {
      const d0 = normalizeDate_(v);
      if (d0 && !isNaN(d0.getTime())) return d0;
    } catch (e0) {}
  }
  if (typeof parseClaimLastUpdatedDatetime06c_ === 'function') {
    try {
      const d1 = parseClaimLastUpdatedDatetime06c_(v);
      if (d1 && !isNaN(d1.getTime())) return d1;
    } catch (e1) {}
  }
  if (typeof v === 'number') {
    const dNum = new Date(Math.round((v - 25569) * 86400000));
    return isNaN(dNum.getTime()) ? null : dNum;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    if (typeof tryNativeParseUnambiguousDate_ === 'function') {
      try {
        const dStr = tryNativeParseUnambiguousDate_(s);
        if (dStr && !isNaN(dStr.getTime())) return dStr;
      } catch (e2) {}
    }
  }
  return null;
}

/** =========================
 * Header normalization helpers (resilient to renamed/aliased columns)
 * ========================= */

function __normalizeHeaderText06_(v) {
  return String(v == null ? '' : v)
    .replace(/^\uFEFF/, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function __normalizeHeaderKey06_(v) {
  const s = __normalizeHeaderText06_(v);
  if (!s) return '';
  return s
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function __canonicalHeaderKey06_(v) {
  const n = __normalizeHeaderKey06_(v);
  // Support legacy abbreviations -> new canonical headers
  if (n === 'lsa') return 'last_status_aging';
  if (n === 'ala') return 'activity_log_aging';
  return n;
}

function __buildCanonicalHeaderIndex06_(headerRow) {
  const map = Object.create(null);
  const arr = Array.isArray(headerRow) ? headerRow : [];
  for (let c = 0; c < arr.length; c++) {
    const k = __canonicalHeaderKey06_(arr[c]);
    if (!k) continue;
    if (map[k] == null) map[k] = c;
  }
  return map;
}

function __findHeaderIndexFlexible06_(headers, desiredHeader) {
  if (!headers || !headers.length) return -1;
  const want = __normalizeHeaderText06_(desiredHeader);
  if (!want) return -1;

  // 0) shared utility exact normalized lookup
  try {
    if (typeof findHeaderIndexByCandidates_ === 'function') {
      const ix = findHeaderIndexByCandidates_(headers, [want]);
      if (ix > -1) return ix;
    }
  } catch (e0) {}

  // 1) exact match
  let idx0 = headers.indexOf(want);
  if (idx0 > -1) return idx0;

  // 2) case-insensitive exact match
  const wantLower = want.toLowerCase();
  idx0 = headers.findIndex(h => __normalizeHeaderText06_(h).toLowerCase() === wantLower);
  if (idx0 > -1) return idx0;

  // 3) canonical match (handles renamed labels like ALA -> Activity Log Aging)
  const wantKey = __canonicalHeaderKey06_(want);
  if (!wantKey) return -1;
  idx0 = headers.findIndex(h => __canonicalHeaderKey06_(h) === wantKey);
  return idx0;
}
/** =========================
 * Raw Data column reordering (end-of-run normalization)
 * ========================= */
function getRawDataReorderPriority06_() {
  try {
    if (typeof RAW_DATA_REORDER_POLICY !== 'undefined' && RAW_DATA_REORDER_POLICY && RAW_DATA_REORDER_POLICY.PRIORITY_HEADERS) {
      return RAW_DATA_REORDER_POLICY.PRIORITY_HEADERS;
    }
  } catch (e) {}
  try {
    if (CONFIG && CONFIG.RAW_DATA_REORDER_POLICY && CONFIG.RAW_DATA_REORDER_POLICY.PRIORITY_HEADERS) {
      return CONFIG.RAW_DATA_REORDER_POLICY.PRIORITY_HEADERS;
    }
  } catch (e2) {}
  return [
    'qoala_policy_number',
    'source_system_name',
    'claim_number',
    'claim_submission_date',
    'last_status',
    'last_update',
    'last_status_aging',
    'activity_log_aging',
    'business_partner_name',
    'insurance_partner_name',
    'dashboard_link',
    'Associate',
    'Update Status',
    'Timestamp',
    'Status'
  ];
}


function isRawDataReorderDisabled06_() {
  try {
    if (typeof RAW_DATA_REORDER_POLICY !== 'undefined' && RAW_DATA_REORDER_POLICY && RAW_DATA_REORDER_POLICY.DISABLED) return true;
  } catch (e) {}
  try {
    if (CONFIG && CONFIG.RAW_DATA_REORDER_POLICY && CONFIG.RAW_DATA_REORDER_POLICY.DISABLED) return true;
  } catch (e2) {}
  return false;
}

function reorderRawDataColumns06_(rawSheet) {
  if (isRawDataReorderDisabled06_()) return 0;
  if (!rawSheet) return 0;
  const headerRow = 1;
  const lastCol = rawSheet.getLastColumn ? rawSheet.getLastColumn() : 0;
  if (!lastCol || lastCol < 2) return 0;

  const headers = rawSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(__normalizeHeaderText06_);

  const desired = getRawDataReorderPriority06_();
  return reorderColumnsByHeaderPriority06_(rawSheet, headers, desired);
}

function reorderColumnsByHeaderPriority06_(sheet, headers, desiredHeaders) {
  if (!sheet || !headers || !desiredHeaders || !desiredHeaders.length) return 0;

  let moves = 0;
  let dest = 1;

  for (let i = 0; i < desiredHeaders.length; i++) {
    const h = String(desiredHeaders[i] || '').trim();
    if (!h) continue;

    const idx0 = __findHeaderIndexFlexible06_(headers, h);
    if (idx0 < 0) continue;

    const currentCol = idx0 + 1;
    if (currentCol === dest) {
      dest++;
      continue;
    }

    // Move entire column (all rows) to destination position.
    const maxRows = sheet.getMaxRows ? sheet.getMaxRows() : Math.max(1, sheet.getLastRow ? sheet.getLastRow() : 1);
    sheet.moveColumns(sheet.getRange(1, currentCol, maxRows, 1), dest);

    // Update local header array to reflect move.
    const [moved] = headers.splice(idx0, 1);
    headers.splice(dest - 1, 0, moved);

    moves++;
    dest++;
  }

  return moves;
}


/** =========================
 * NEW (Enterprise refactor additions - Feb 2026)
 *
 * Implemented as utilities so other flow files (00–06b) can call them.
 * - Status Type (mandatory) utilities
 * - Last Status Date parsing/formatting (Sub flow: dd MMM yy, HH:mm)
 * - Sorting operational sheets while preserving filters
 *
 * NOTE: This file does NOT assume Raw OLD/Raw NEW are snapshots.
 * =========================
 */

/** Simple structured log helper (opt-in usage). */
function logJson06c_(level, event, data) {
  try {
    const payload = {
      ts: new Date().toISOString(),
      level: String(level || 'INFO').toUpperCase(),
      event: String(event || 'event'),
      data: data || {}
    };
    Logger.log(JSON.stringify(payload));
  } catch (e) {
    // swallow
  }
}

/**
 * Ensure a header column exists in row 1. If missing, appends to the right.
 * Returns 0-based column index.
 */
function ensureHeaderColumn06c_(sheet, headerName) {
  if (!sheet) return -1;
  const name = __normalizeHeaderText06_(headerName);
  if (!name) return -1;

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0].map(__normalizeHeaderText06_);
  let idx = __findHeaderIndexFlexible06_(header, name);
  if (idx !== -1) return idx;

  // Append new header at the far right.
  idx = header.length;
  try { sheet.getRange(1, idx + 1).setValue(name); } catch (e) {}
  return idx;
}

/**
 * Optional header: return index if present, else -1 (no mutation).
 */
function findOptionalHeaderColumn06c_(sheet, headerName) {
  if (!sheet) return -1;
  const name = __normalizeHeaderText06_(headerName);
  if (!name) return -1;
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0].map(__normalizeHeaderText06_);
  return __findHeaderIndexFlexible06_(header, name);
}

/**
 * Status Type mapping (mandatory column on operational sheets).
 * Lookup key is normalized (trim + upper).
 */
function getStatusTypeMap06c_() {
  // Allow override via CONFIG / source-of-truth constants when present.
  try {
    if (CONFIG && (CONFIG.STATUS_TYPE_MAP || CONFIG.statusTypeMap || CONFIG.statusTypeByLastStatus)) {
      return CONFIG.STATUS_TYPE_MAP || CONFIG.statusTypeMap || CONFIG.statusTypeByLastStatus;
    }
  } catch (e) {}
  try {
    if (typeof STATUS_TYPE_BY_LAST_STATUS !== 'undefined' && STATUS_TYPE_BY_LAST_STATUS) {
      return STATUS_TYPE_BY_LAST_STATUS;
    }
  } catch (e2) {}

  // Strict fallback: source-of-truth only.
  return {};
}

function getStatusType06c_(lastStatus) {
  const s = String(lastStatus || '').trim().toUpperCase();
  if (!s) return '';
  const map = getStatusTypeMap06c_();
  return map[s] || '';
}

/**
 * Update (or fill) Status Type column in a single sheet.
 * - Ensures Status Type header exists.
 * - If Last Status column doesn't exist, does nothing.
 */
function upsertStatusTypeColumnInOperationalSheet06c_(sheet) {
  if (DRY_RUN) return 0;
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText06_);
  let idxLastStatus = __findHeaderIndexFlexible06_(header, 'Last Status');
  if (idxLastStatus === -1) idxLastStatus = __findHeaderIndexFlexible06_(header, 'Status');
  if (idxLastStatus === -1) return 0;

  const idxStatusType = (function() {
    const idx = __findHeaderIndexFlexible06_(header, 'Status Type');
    if (idx !== -1) return idx;
    return ensureHeaderColumn06c_(sheet, 'Status Type');
  })();

  const n = lastRow - 1;
  const stVals = sheet.getRange(2, idxLastStatus + 1, n, 1).getValues();
  const curVals = sheet.getRange(2, idxStatusType + 1, n, 1).getValues();

  const out = new Array(n);
  let changed = 0;
  for (let r = 0; r < n; r++) {
    const want = getStatusType06c_(stVals[r][0]);
    if (String(curVals[r][0] || '') !== want) {
      out[r] = [want];
      changed++;
    } else {
      out[r] = [curVals[r][0]];
    }
  }

  if (changed) {
    try { sheet.getRange(2, idxStatusType + 1, n, 1).setValues(out); } catch (e) {}
  }
  return changed;
}

/**
 * Parse Raw "claim_last_updated_datetime" string into a Date object.
 * Expected sample: "January 24, 2025, 6:06 PM".
 * Returns null if parsing fails.
 */
function parseClaimLastUpdatedDatetime06c_(v) {
  try {
    if (typeof parseClaimLastUpdatedDatetime_ === 'function') return parseClaimLastUpdatedDatetime_(v);
  } catch (e) {}
  return null;
}

/**
 * Apply number format for Last Status Date based on flow type.
 * - sub: dd MMM yy, HH:mm
 * - main/form: dd MMM yy
 */
function applyLastStatusDateFormat06c_(sheet, idxLastStatusDate, flow) {
  if (DRY_RUN) return;
  if (!sheet || idxLastStatusDate == null || idxLastStatusDate < 0) return;
  const lr = sheet.getLastRow();
  if (lr < 2) return;

  const n = lr - 1;
  const isSub = String(flow || '').toLowerCase() === 'sub';
  const fmt = isSub ? 'dd MMM yy, HH:mm' : 'dd MMM yy';
  try { sheet.getRange(2, idxLastStatusDate + 1, n, 1).setNumberFormat(fmt); } catch (e) {}
}

/**
 * Sort data by Last Status Date then Last Status, without removing an existing filter.
 */
function sortOperationalSheetPreserveFilter06c_(sheet) {
  if (DRY_RUN) return;
  if (!sheet) return;
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return;

  const header = sheet.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  const idxDate = __findHeaderIndexFlexible06_(header, 'Last Status Date');
  const idxStatus = __findHeaderIndexFlexible06_(header, 'Last Status');
  if (idxDate === -1 || idxStatus === -1) return;

  try {
    try { __expandSheetFilterToUsedRange06_(sheet); } catch (eF) {}
    const filter = sheet.getFilter ? sheet.getFilter() : null;
    if (filter && filter.getRange) {
      const fr = filter.getRange();
      const sr = fr.getRow();
      const sc = fr.getColumn();
      const nr = fr.getNumRows();
      const nc = fr.getNumColumns();
      if (nr <= 1) return;
      const absDate = idxDate + 1;
      const absStatus = idxStatus + 1;
      const frColEnd = sc + nc - 1;
      if (absDate < sc || absDate > frColEnd || absStatus < sc || absStatus > frColEnd) return;
      sheet.getRange(sr + 1, sc, nr - 1, nc).sort([
        { column: absDate - sc + 1, ascending: true },
        { column: absStatus - sc + 1, ascending: true }
      ]);
      return;
    }
  } catch (e0) {}

  try {
    sheet.getRange(2, 1, lr - 1, lc).sort([
      { column: idxDate + 1, ascending: true },
      { column: idxStatus + 1, ascending: true }
    ]);
  } catch (e1) {}
}

/**
 * Strict Second-Year (Market Value) detector.
 * Source must be month_policy_aging from Raw Data.
 */
function isSecondYearMarketValue06c_(monthPolicyAging) {
  const n = (typeof monthPolicyAging === 'number') ? monthPolicyAging : parseFloat(String(monthPolicyAging || '').trim());
  return !isNaN(n) && n > 12;
}


/***************************************************************
 * Consolidated from legacy 06d_IntegratedMaintenance.gs,
 * 06e_SubHelpers.gs, and 06f_RuntimeAssertions.gs
 * to keep runtime modules in 00-06c only.
 ***************************************************************/


// ---- BEGIN MERGED: 06d_IntegratedMaintenance.gs ----
/***************************************************************
 * 06d_IntegratedMaintenance.gs
 * Integrated maintenance module.
 *
 * Absorbs previous patch-layer files:
 * - 97_SelfCheck
 * - 98_PipelineCoreHardening
 * - 99_RuntimeFixes
 *
 * Goal:
 * - keep transition logic centralized,
 * - reduce patch-layer sprawl,
 * - keep self-check + maintenance utilities centralized.
 ***************************************************************/
'use strict';

(function bootstrapIntegratedMaintenance06d_() {
  // Runtime hardening that previously lived in this bootstrap has been merged back
  // into the source modules (05b/06b/06c). Keep this bootstrap as an explicit
  // no-op so file load order stays stable without re-shadowing upstream helpers.
})();

function runSelfCheck_() {
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();
  const startedAt = new Date();
  const report = {
    checkedAt: startedAt.toISOString(),
    ok: true,
    errors: [],
    warnings: [],
    summary: {}
  };

  function pushErr(msg) {
    report.ok = false;
    report.errors.push(String(msg || 'Unknown error'));
  }
  function pushWarn(msg) {
    report.warnings.push(String(msg || 'Unknown warning'));
  }
  function resolveSymbol_(name) {
    try {
      // eval can resolve top-level const/let that do not become properties on globalThis.
      return eval(name);
    } catch (e0) {}
    try {
      return g[name];
    } catch (e1) {}
    return undefined;
  }
  function hasFn(name) {
    try { return typeof resolveSymbol_(name) === 'function'; } catch (e) { return false; }
  }
  function hasObj(name) {
    try {
      const v = resolveSymbol_(name);
      return typeof v !== 'undefined' && v != null;
    } catch (e) { return false; }
  }
  function uniqSorted(arr) {
    const seen = Object.create(null);
    const out = [];
    (arr || []).forEach(function (x) {
      const s = String(x || '').trim();
      if (!s || seen[s]) return;
      seen[s] = 1;
      out.push(s);
    });
    out.sort();
    return out;
  }

  try {
    if (hasFn('healthCheck_')) {
      const hc = healthCheck_();
      report.summary.health = hc;
      if (!hc || hc.ok !== true) pushErr('healthCheck_ reported failure.');
    } else {
      pushWarn('healthCheck_ is not available.');
    }
  } catch (e0) {
    pushErr('healthCheck_ threw: ' + (e0 && e0.message ? e0.message : e0));
  }

  const criticalFns = [
    'runPipeline_',
    'runEmailIngest',
    'runSubEmailIngest',
    'onFormSubmit',
    'runManual',
    'parseUploadedFile_',
    'mutateRawInMemory_',
    'routeRawToOperationalSheetsInMemory_',
    'enrichOperationalSheetsFromRaw06_',
    'processB2B_',
    'processSpecialCase_',
    'processEVBike_',
    'processDoss_',
    '__expandSheetFilterToUsedRange06_',
    '__expandWorkbookFiltersToUsedRange06_',
    'sortOperationalSheetsPreserveFilter06b_',
    'sortOperationalSheetPreserveFilter06c_'
  ];
  const missingFns = criticalFns.filter(function (name) { return !hasFn(name); });
  report.summary.missingCriticalFunctions = missingFns;
  if (missingFns.length) pushErr('Missing critical functions: ' + missingFns.join(', '));

  const criticalObjects = [
    'CONFIG',
    'OPS_ROUTING_POLICY',
    'STATUS_TYPE_BY_LAST_STATUS',
    'POSITION_BY_LAST_STATUS',
    'FINISH_STATUSES'
  ];
  const missingObjects = criticalObjects.filter(function (name) { return !hasObj(name); });
  report.summary.missingCriticalObjects = missingObjects;
  if (missingObjects.length) pushErr('Missing critical config objects: ' + missingObjects.join(', '));

  try {
    const routed = [];
    const bySheet = (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY && OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET)
      ? OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET
      : {};
    Object.keys(bySheet || {}).forEach(function (sheetName) {
      (bySheet[sheetName] || []).forEach(function (status) {
        const s = String(status || '').trim();
        if (s) routed.push(s);
      });
    });
    const routedStatuses = uniqSorted(routed);
    report.summary.routedStatusCount = routedStatuses.length;

    const statusTypeMap = (typeof STATUS_TYPE_BY_LAST_STATUS !== 'undefined' && STATUS_TYPE_BY_LAST_STATUS)
      ? STATUS_TYPE_BY_LAST_STATUS
      : ((typeof CONFIG !== 'undefined' && CONFIG && CONFIG.statusTypeByLastStatus) ? CONFIG.statusTypeByLastStatus : {});
    const positionMap = (typeof POSITION_BY_LAST_STATUS !== 'undefined' && POSITION_BY_LAST_STATUS)
      ? POSITION_BY_LAST_STATUS
      : ((typeof CONFIG !== 'undefined' && CONFIG && CONFIG.positionByLastStatus) ? CONFIG.positionByLastStatus : {});

    const missingStatusType = routedStatuses.filter(function (s) { return statusTypeMap[s] == null || String(statusTypeMap[s]).trim() === ''; });
    const missingPosition = routedStatuses.filter(function (s) { return positionMap[s] == null || String(positionMap[s]).trim() === ''; });

    report.summary.missingStatusTypeMapping = missingStatusType;
    report.summary.missingPositionMapping = missingPosition;

    if (missingStatusType.length) pushErr('Missing Status Type mapping for: ' + missingStatusType.join(', '));
    if (missingPosition.length) pushErr('Missing Position mapping for: ' + missingPosition.join(', '));
  } catch (e1) {
    pushErr('Routing coverage check failed: ' + (e1 && e1.message ? e1.message : e1));
  }

  try {
    const kwBySheet = (typeof OPS_ROUTING_POLICY !== 'undefined' && OPS_ROUTING_POLICY && OPS_ROUTING_POLICY.SC_NAME_KEYWORDS)
      ? OPS_ROUTING_POLICY.SC_NAME_KEYWORDS
      : {};
    const kwIndex = Object.create(null);
    Object.keys(kwBySheet || {}).forEach(function (sheet) {
      (kwBySheet[sheet] || []).forEach(function (kw) {
        const k = String(kw || '').trim().toLowerCase();
        if (!k) return;
        if (!kwIndex[k]) kwIndex[k] = [];
        kwIndex[k].push(sheet);
      });
    });
    const overlaps = [];
    Object.keys(kwIndex).forEach(function (k) {
      const owners = uniqSorted(kwIndex[k]);
      if (owners.length > 1) overlaps.push(k + ' => ' + owners.join(' / '));
    });
    report.summary.scKeywordOverlap = overlaps;
    if (overlaps.length) pushWarn('SC keyword overlap detected: ' + overlaps.join(' | '));
  } catch (e2) {
    pushWarn('SC keyword overlap check failed: ' + (e2 && e2.message ? e2.message : e2));
  }


  try {
    const part9Checks = {
      manualSnapshotFlow: hasFn('clearOperationalSheets_'),
      b2bFallbackFlow: hasFn('processB2B_'),
      evBikeOverlayFlow: hasFn('processEVBike_'),
      reportBaseSyncFlow: hasFn('refreshReportBaseFromOperational06_')
    };
    const missingPart9 = Object.keys(part9Checks).filter(function (k) { return !part9Checks[k]; });
    report.summary.part9HardeningReadiness = {
      checks: part9Checks,
      missing: missingPart9,
      ready: missingPart9.length === 0
    };
    if (missingPart9.length) pushWarn('Part 9 hardening readiness missing: ' + missingPart9.join(', '));
  } catch (ePart9) {
    pushWarn('Part 9 hardening readiness check failed: ' + (ePart9 && ePart9.message ? ePart9.message : ePart9));
  }

  try {
    const tails = (typeof RAW_DATA_CUSTOM_TAIL_HEADERS !== 'undefined' && RAW_DATA_CUSTOM_TAIL_HEADERS) ? RAW_DATA_CUSTOM_TAIL_HEADERS : [];
    const counts = Object.create(null);
    (tails || []).forEach(function (h) {
      const s = String(h || '').trim();
      if (!s) return;
      counts[s] = (counts[s] || 0) + 1;
    });
    const duplicates = Object.keys(counts).filter(function (k) { return counts[k] > 1; });
    report.summary.duplicateRawTailHeaders = duplicates;
    if (duplicates.length) pushWarn('Duplicate raw tail headers: ' + duplicates.join(', '));
  } catch (e3) {
    pushWarn('Raw tail duplicate check failed: ' + (e3 && e3.message ? e3.message : e3));
  }

  report.summary.durationMs = new Date().getTime() - startedAt.getTime();
  try { Logger.log(JSON.stringify(report)); } catch (e5) {}
  return report;
}

function summarizeSelfCheck_(report) {
  const r = report || runSelfCheck_();
  const parts = [];
  parts.push(r.ok ? 'SELF CHECK: OK' : 'SELF CHECK: FAIL');
  parts.push('errors=' + ((r.errors || []).length));
  parts.push('warnings=' + ((r.warnings || []).length));
  parts.push('durationMs=' + (((r.summary || {}).durationMs) || 0));
  if (r.errors && r.errors.length) parts.push('firstError=' + r.errors[0]);
  if (r.warnings && r.warnings.length) parts.push('firstWarning=' + r.warnings[0]);
  return parts.join(' | ');
}

function runSelfCheckAndDescribe_() {
  const out = {
    selfCheck: runSelfCheck_(),
    system: (typeof describeSystem_ === 'function') ? describeSystem_() : null
  };
  try { Logger.log(JSON.stringify(out)); } catch (e) {}
  return out;
}
// ---- END MERGED: 06d_IntegratedMaintenance.gs ----


// ---- BEGIN MERGED: 06e_SubHelpers.gs ----
/***************************************************************
 * 06e_SubHelpers.gs
 * Extracted SUB helper implementations from 06a_EntryPoints.
 ***************************************************************/
'use strict';

function __appendSubmissionFromRawIfMissing06e_(ss, rawMap, rawHdrIdx, dbTag) {
  const sh = ss.getSheetByName('Submission');
  if (!sh) {
    try { logLine_('SUB_WARN', 'Submission sheet not found (skip append)', '', '', 'WARN'); } catch (e1) {}
    return { appended: 0, skipped: true };
  }

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 1) return { appended: 0, skipped: true };

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const norm = header.map(h => h.toLowerCase());

  function idxOf(name) {
    const i = norm.indexOf(String(name || '').toLowerCase());
    return i >= 0 ? i : -1;
  }

  const idxClaim = idxOf('claim number');
  if (idxClaim < 0) {
    try { logLine_('SUB_WARN', 'Submission missing Claim Number header (skip append)', '', '', 'WARN'); } catch (e2) {}
    return { appended: 0, skipped: true };
  }

  // Existing claim set
  const existing = new Set();
  if (lastRow >= 2) {
    const colVals = sh.getRange(2, idxClaim + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < colVals.length; i++) {
      const cn = String(colVals[i][0] || '').trim();
      if (cn) existing.add(cn);
    }
  }

  const idxSubDate = idxOf('submission date');
  const idxSubMonth = idxOf('submission by month');
  const idxDbLink = idxOf('db link');
  const idxPartner = idxOf('partner name');
  const idxIns = idxOf('insurance');
  const idxDeviceType = idxOf('device type');
  const idxImei = idxOf('imei/sn');
  const idxLast = idxOf('last status');
  const idxSc = idxOf('service center');
  const idxLSA = idxOf('last status aging');
  const idxALA = idxOf('activity log aging');
  const idxTat = idxOf('tat');
  const idxActLog = idxOf('activity log');

  const flow = (typeof CONFIG === 'object' && CONFIG && CONFIG.subFlow) ? CONFIG.subFlow : null;
  let dbValue = String(dbTag || '').trim().toUpperCase();
  let wantStatus = (dbValue === 'OLD') ? 'SUBMITTED' : 'CLAIM_INITIATE';

  // Prefer configured DB tag + trigger status mapping when available.
  try {
    if (flow && flow.SUBMISSION_RULES) {
      if (dbValue === 'OLD' && flow.SUBMISSION_RULES.OLD) {
        if (flow.SUBMISSION_RULES.OLD.DB_VALUE) {
          dbValue = String(flow.SUBMISSION_RULES.OLD.DB_VALUE || '').trim().toUpperCase();
        }
        if (flow.SUBMISSION_RULES.OLD.TRIGGER_RAW_LAST_STATUS) {
          wantStatus = String(flow.SUBMISSION_RULES.OLD.TRIGGER_RAW_LAST_STATUS || '').trim().toUpperCase();
        }
      } else if (dbValue === 'NEW' && flow.SUBMISSION_RULES.NEW) {
        if (flow.SUBMISSION_RULES.NEW.DB_VALUE) {
          dbValue = String(flow.SUBMISSION_RULES.NEW.DB_VALUE || '').trim().toUpperCase();
        }
        if (flow.SUBMISSION_RULES.NEW.TRIGGER_RAW_LAST_STATUS) {
          wantStatus = String(flow.SUBMISSION_RULES.NEW.TRIGGER_RAW_LAST_STATUS || '').trim().toUpperCase();
        }
      }
    }
  } catch (e0) {}

  const rowsToAppend = [];
  const richLinks = []; // { rowOffset, url }

  rawMap.forEach(rec => {
    const cn = String(rec.claim_number || '').trim();
    if (!cn) return;

    const st = String(rec.last_status || '').trim().toUpperCase();
    if (st !== wantStatus) return;
    if (existing.has(cn)) return;

    const row = new Array(lastCol).fill('');

    if (idxSubDate >= 0) row[idxSubDate] = rec.claim_submitted_datetime || '';
    if (idxSubMonth >= 0) {
      row[idxSubMonth] = (typeof formatSubmissionMonthShort_ === 'function')
        ? formatSubmissionMonthShort_(rec.claim_submitted_datetime)
        : '';
    }
    row[idxClaim] = cn;

    if (idxDbLink >= 0) {
      row[idxDbLink] = 'LINK';
      const url = String(rec.dashboard_link || '').trim() || ((typeof buildDashboardLinkFromClaimNumber_ === 'function') ? buildDashboardLinkFromClaimNumber_(cn) : '');
      if (url) richLinks.push({ rowOffset: rowsToAppend.length, url: url });
    }
    if (idxPartner >= 0) row[idxPartner] = rec.partner_name || '';
    if (idxIns >= 0) row[idxIns] = rec.insurance || '';
    if (idxDeviceType >= 0) row[idxDeviceType] = rec.device_type || '';
    if (idxImei >= 0) row[idxImei] = rec.device_imei || '';
    if (idxLast >= 0) row[idxLast] = rec.last_status || '';
    if (idxSc >= 0) row[idxSc] = rec.sc_name || '';
    if (idxLSA >= 0) row[idxLSA] = rec.last_status_aging || '';
    if (idxALA >= 0) row[idxALA] = rec.activity_log_aging || '';
    if (idxTat >= 0) row[idxTat] = (typeof diffDaysDecimalFromNow_ === 'function') ? diffDaysDecimalFromNow_(rec.claim_submitted_datetime) : (rec.tat || '');
    if (idxActLog >= 0) row[idxActLog] = rec.activity_log || '';

    rowsToAppend.push(row);
    existing.add(cn);
  });

  if (!rowsToAppend.length) {
    try { logLine_('SUB_APP', 'No new Submission rows to append', dbTag, '', 'INFO'); } catch (e3) {}
    return { appended: 0 };
  }

  const startRow = sh.getLastRow() + 1;
  const rng = sh.getRange(startRow, 1, rowsToAppend.length, lastCol);
  if (typeof safeSetValues_ === 'function') safeSetValues_(rng, rowsToAppend);
  else rng.setValues(rowsToAppend);

  // Apply RichText hyperlink to DB Link column if possible.
  if (idxDbLink >= 0 && richLinks.length) {
    try {
      const rich = [];
      for (let i = 0; i < rowsToAppend.length; i++) rich.push([SpreadsheetApp.newRichTextValue().setText('LINK').build()]);

      for (let i = 0; i < richLinks.length; i++) {
        const r = richLinks[i].rowOffset;
        const url = richLinks[i].url;
        rich[r][0] = SpreadsheetApp.newRichTextValue().setText('LINK').setLinkUrl(url).build();
      }

      const linkRange = sh.getRange(startRow, idxDbLink + 1, rowsToAppend.length, 1);
      if (typeof safeSetRichTextValues_ === 'function') safeSetRichTextValues_(linkRange, rich);
      else linkRange.setRichTextValues(rich);
    } catch (e4) {
      try { logLine_('SUB_WARN', 'Failed to apply RichText hyperlinks for DB Link', String(e4), '', 'WARN'); } catch (e5) {}
    }
  }

  try { logLine_('SUB_APP', 'Appended rows to Submission', dbTag, 'count=' + rowsToAppend.length, 'INFO'); } catch (e6) {}

  return { appended: rowsToAppend.length };
}

function __sortOperationalSheetsSub06e_(ss, sheetNames, sortSpecs) {
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  const out = { sheets: {}, sortedSheets: 0, skippedSheets: 0 };

  // Default SUB sort order:
  // 1) Submission Date (A->Z)
  // 2) Last Status Date (A->Z)
  // 3) Last Status (A->Z)
  const defaultSpecs = [
    { headerCandidates: ['Submission Date', 'claim_submission_date', 'claim submitted datetime'], ascending: true },
    { headerCandidates: ['Last Status Date', 'claim_last_updated_datetime', 'claim last updated datetime'], ascending: true },
    { headerCandidates: ['Last Status', 'last_status', 'last status'], ascending: true }
  ];

  const normalizedSpecs = (Array.isArray(sortSpecs) && sortSpecs.length)
    ? sortSpecs.map(s => ({
        headerCandidates: Array.isArray(s && s.headerCandidates)
          ? s.headerCandidates
          : (s && s.header ? [s.header] : []),
        ascending: (s && typeof s.ascending === 'boolean') ? s.ascending : true
      }))
    : defaultSpecs;

  for (let i = 0; i < names.length; i++) {
    const name = String(names[i] || '').trim();
    if (!name) continue;

    const sh = ss.getSheetByName(name);
    if (!sh) { out.sheets[name] = { skipped: 'not found' }; out.skippedSheets++; continue; }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 3 || lastCol < 2) { out.sheets[name] = { skipped: 'too few rows' }; out.skippedSheets++; continue; }

    try {
      const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
      const norm = hdr.map(h => h.toLowerCase());

      function idxByCandidates(cands) {
        for (let k = 0; k < cands.length; k++) {
          const key = String(cands[k] || '').toLowerCase();
          const j = norm.indexOf(key);
          if (j >= 0) return j;
        }
        return -1;
      }

      const sortCriteria = [];
      normalizedSpecs.forEach(s => {
        const idx = idxByCandidates(s.headerCandidates);
        if (idx >= 0) sortCriteria.push({ column: idx + 1, ascending: !!s.ascending });
      });

      if (!sortCriteria.length) {
        out.sheets[name] = { skipped: 'missing headers for sort' };
        out.skippedSheets++;
        continue;
      }

      // Preserve filter criteria, but first expand the range so hidden/out-of-range rows are included.
      try { __expandSheetFilterToUsedRange06_(sh); } catch (eF) {}
      const filter = sh.getFilter ? sh.getFilter() : null;
      if (filter && filter.getRange) {
        const fr = filter.getRange();
        const sr = fr.getRow();
        const sc = fr.getColumn();
        const nr = fr.getNumRows();
        const nc = fr.getNumColumns();

        if (nr <= 1) { out.sheets[name] = { skipped: 'filter range too small' }; out.skippedSheets++; continue; }

        const body = sh.getRange(sr + 1, sc, nr - 1, nc);
        if (!isDryRun_()) body.sort(sortCriteria);
      } else {
        const body = sh.getRange(2, 1, lastRow - 1, lastCol);
        if (!isDryRun_()) body.sort(sortCriteria);
      }

      out.sheets[name] = { sorted: true, criteria: sortCriteria.length };
      out.sortedSheets++;
    } catch (e0) {
      out.sheets[name] = { error: String(e0 && e0.message ? e0.message : e0) };
      out.skippedSheets++;
      try { logLine_('SUB_WARN', 'Sort failed', name, String(e0), 'WARN'); } catch (e1) {}
    }
  }

  return out;
}
// ---- END MERGED: 06e_SubHelpers.gs ----


// ---- BEGIN MERGED: 06f_RuntimeAssertions.gs ----
/***************************************************************
 * 06f_RuntimeAssertions.gs
 * Runtime preflight assertions (non-fatal by default)
 ***************************************************************/
'use strict';

function runtimePreflight06f_(contextTag) {
  const tag = String(contextTag || 'PIPELINE').trim().toUpperCase();
  const issues = [];

  function check(name, ok, note) {
    if (ok) return;
    issues.push(name + (note ? ' (' + note + ')' : ''));
  }

  try { check('__appendSubmissionFromRawIfMissing06e_', typeof __appendSubmissionFromRawIfMissing06e_ === 'function', '06e split helper missing'); } catch (e0) { issues.push('__appendSubmissionFromRawIfMissing06e_ check failed'); }
  try { check('__sortOperationalSheetsSub06e_', typeof __sortOperationalSheetsSub06e_ === 'function', '06e split helper missing'); } catch (e1) { issues.push('__sortOperationalSheetsSub06e_ check failed'); }
  try { check('findHeaderIndexByCandidates_', typeof findHeaderIndexByCandidates_ === 'function', 'shared header matcher missing'); } catch (e2) { issues.push('findHeaderIndexByCandidates_ check failed'); }
  try { check('computeDbValueFromClaimNumber_', typeof computeDbValueFromClaimNumber_ === 'function', 'shared DB classifier missing'); } catch (e3) { issues.push('computeDbValueFromClaimNumber_ check failed'); }
  try { check('parseClaimLastUpdatedDatetime_', typeof parseClaimLastUpdatedDatetime_ === 'function', 'shared datetime parser missing'); } catch (e4) { issues.push('parseClaimLastUpdatedDatetime_ check failed'); }

  if (!issues.length) return { ok: true, issues: [] };

  try {
    if (typeof logLine_ === 'function') {
      logLine_('WARN', 'PREFLIGHT_' + tag, 'Runtime preflight found missing symbols', issues.join(' | '), 'WARN');
    }
  } catch (e5) {}

  return { ok: false, issues: issues };
}
// ---- END MERGED: 06f_RuntimeAssertions.gs ----

/** Restore named Raw backup fields to any operational sheet that exposes them. */
function restoreNamedOpsFieldsFromRaw06c_(ss, rawSheet, headerIndexRaw, pic, names) {
  if (DRY_RUN || !ss || !rawSheet) return 0;
  const idxClaimRaw = headerIndexRaw[CONFIG.headers.claimNumber];
  if (idxClaimRaw == null) return 0;
  const n = rawSheet.getLastRow() - 1; if (n < 1) return 0;
  const raw = rawSheet.getRange(2, 1, n, rawSheet.getLastColumn()).getValues();
  const map = Object.create(null);
  raw.forEach(function(row) { const key = __claimKey06_(row[idxClaimRaw]); if (key) map[key] = row; });
  let restored = 0;
  getOperationalSheetsForBackup_(pic).forEach(function(name) {
    const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return;
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const iClaim = __findHeaderIndexFlexible06_(hdr, 'Claim Number'); if (iClaim === -1) return;
    const count = sh.getLastRow() - 1, rows = sh.getRange(2, 1, count, sh.getLastColumn()).getValues();
    (names || []).forEach(function(field) {
      const rawIdx = headerIndexRaw[field], opsIdx = __findHeaderIndexFlexible06_(hdr, field);
      if (rawIdx == null || opsIdx === -1) return;
      const out = rows.map(function(row) { const saved = map[__claimKey06_(row[iClaim])]; const v = saved ? saved[rawIdx] : ''; if ((row[opsIdx] === '' || row[opsIdx] == null) && v !== '' && v != null) { restored++; return [v]; } return [row[opsIdx]]; });
      sh.getRange(2, opsIdx + 1, count, 1).setValues(out);
    });
  });
  return restored;
}
