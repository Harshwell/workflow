/***************************************************************
 * 06c_PostProcessAndUtils.gs  (SPLIT FROM 06.gs)
 * Post-route utilities, carry-forward, preflight, schema formats,
 * exclusion recompute, and Raw Data column reordering
 ***************************************************************/
'use strict';
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
    const c = String((claimVals[i] && claimVals[i][0]) || '').trim();
    if (!c) continue;
    const rt = (rtVals[i] && rtVals[i][0]) ? rtVals[i][0] : null;
    if (!rt || !rt.getText) continue;
    const t = String(rt.getText() || '');
    if (!t) continue; // only apply when there is text (rich formatting matters)
    map[c.toUpperCase()] = rt;
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
      const claim = String((claims[i] && claims[i][0]) || '').trim();
      if (!claim) continue;
      const rt = map[claim.toUpperCase()];
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
    const c = String((claimVals[i] && claimVals[i][0]) || '').trim();
    if (!c) continue;
    const rt = (rtVals[i] && rtVals[i][0]) ? rtVals[i][0] : null;
    if (!rt || !rt.getText) continue;
    const t = String(rt.getText() || '');
    if (t === '') continue;
    map[c.toUpperCase()] = rt;
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
    if (idxClaim === -1 || idxRemarks === -1) continue;

    const rows = lastRow - 1;
    const claims = sh.getRange(2, idxClaim + 1, rows, 1).getValues();

    const rowNums = [];
    const rowMap = Object.create(null);

    for (let i = 0; i < rows; i++) {
      const claim = String((claims[i] && claims[i][0]) || '').trim();
      if (!claim) continue;
      const rt = map[claim.toUpperCase()];
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

    if (idxUpdate === -1 && idxTs === -1 && idxStatus === -1 && idxRemarks === -1) continue;

    const n = lr - 1;
    const claims = sh.getRange(2, idxClaim + 1, n, 1).getValues();

    const rngUpdate = (idxUpdate !== -1) ? sh.getRange(2, idxUpdate + 1, n, 1) : null;
    const rngTs = (idxTs !== -1) ? sh.getRange(2, idxTs + 1, n, 1) : null;
    const rngStatus = (idxStatus !== -1) ? sh.getRange(2, idxStatus + 1, n, 1) : null;
    const rngRemarks = (idxRemarks !== -1) ? sh.getRange(2, idxRemarks + 1, n, 1) : null;

    const updRT = rngUpdate ? rngUpdate.getRichTextValues() : null;
    const updWrap = rngUpdate ? rngUpdate.getWrapStrategies() : null;

    const tsVals = rngTs ? rngTs.getValues() : null;
    const tsFmt = rngTs ? rngTs.getNumberFormats() : null;

    const stVals = rngStatus ? rngStatus.getValues() : null;

    const remRT = rngRemarks ? rngRemarks.getRichTextValues() : null;
    const remWrap = rngRemarks ? rngRemarks.getWrapStrategies() : null;

    for (let i = 0; i < n; i++) {
      const claim = String((claims[i] && claims[i][0]) || '').trim();
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
          rec.u = { rt: rt, wrap: (updWrap && updWrap[i] && updWrap[i][0]) ? updWrap[i][0] : null };
          any = true;
        }
      }

      if (tsVals) {
        const v = tsVals[i] ? tsVals[i][0] : '';
        if (v !== '' && v != null) {
          rec.t = { v: v, fmt: (tsFmt && tsFmt[i] && tsFmt[i][0]) ? tsFmt[i][0] : null };
          any = true;
        }
      }

      if (stVals) {
        const v = stVals[i] ? stVals[i][0] : '';
        if (v !== '' && v != null) {
          rec.s = { v: v };
          any = true;
        }
      }

      if (remRT) {
        const rt = (remRT[i] && remRT[i][0]) ? remRT[i][0] : null;
        const txt = (rt && rt.getText) ? String(rt.getText() || '') : '';
        if (txt !== '') {
          rec.r = { rt: rt, wrap: (remWrap && remWrap[i] && remWrap[i][0]) ? remWrap[i][0] : null };
          any = true;
        }
      }

      if (any) {
        out[key] = rec;
        count++;
      }
    }
  }

  return { version: 'ops_manual_v1', map: out, count: count };
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

    if (idxUpdate === -1 && idxTs === -1 && idxStatus === -1 && idxRemarks === -1) continue;

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

    for (let i = 0; i < n; i++) {
      const claim = String((claims[i] && claims[i][0]) || '').trim();
      if (!claim) continue;
      const key = claim.toUpperCase();
      const rec = snap[key];
      if (!rec) continue;

      const rno = i + 2;

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

  // 2) EV-Bike: clear DV on Last Status to avoid validation violations when writing statuses
  try {
    const ev = ss.getSheetByName('EV-Bike');
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
}

/**
 * Restore operational fields from Raw backup after CLR + ROUTE.
 * Restores:
 * - Status (values)
 * - OR (checkbox + values)
 * - Timestamp (values)
 * - Update Status Asso + Timestamp Asso
 * - Update Status Admin + Timestamp Admin
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

  // Manual tail columns (carry-forward / restore)
  const idxUpdateAssoRaw = headerIndexRaw['Update Status Asso'];
  const idxTsAssoRaw = headerIndexRaw['Timestamp Asso'];
  const idxUpdateAdminRaw = headerIndexRaw['Update Status Admin'];
  const idxTsAdminRaw = headerIndexRaw['Timestamp Admin'];

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
    const key = String(rawClaims[i][0] || '').trim().toUpperCase();
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

    const idxUpAssoOps = __findHeaderIndexFlexible06_(header, 'Update Status Asso');
    const idxTsAssoOps = __findHeaderIndexFlexible06_(header, 'Timestamp Asso');
    const idxUpAdminOps = __findHeaderIndexFlexible06_(header, 'Update Status Admin');
    const idxTsAdminOps = __findHeaderIndexFlexible06_(header, 'Timestamp Admin');

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
      const claimKey = String(claimsOps[r][0] || '').trim().toUpperCase();
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

    if (idxUpdateOps !== -1 && outUpdate) {
      try { sh.getRange(2, idxUpdateOps + 1, m, 1).setValues(outUpdate); } catch (e) {}
    }

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
  const idxUpdateAsso = headerIndexRaw['Update Status Asso'];
  const idxTsAsso = headerIndexRaw['Timestamp Asso'];
  const idxUpdateAdmin = headerIndexRaw['Update Status Admin'];
  const idxTsAdmin = headerIndexRaw['Timestamp Admin'];

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

  const idxUpdateAsso = headerIndexRaw['Update Status Asso'];
  const idxTsAsso = headerIndexRaw['Timestamp Asso'];
  const idxUpdateAdmin = headerIndexRaw['Update Status Admin'];
  const idxTsAdmin = headerIndexRaw['Timestamp Admin'];

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

  const s = String(scNameVal || '').toLowerCase();

  const hasFarhan = (kwFarhan || []).some(k => k && s.indexOf(String(k).toLowerCase()) > -1);
  const hasMeilani = (kwMeilani || []).some(k => k && s.indexOf(String(k).toLowerCase()) > -1);
  const hasIvan = (kwIvan || []).some(k => k && s.indexOf(String(k).toLowerCase()) > -1);

  return t.filter(x => {
    if (x === scFarhanName) return hasFarhan;
    if (x === scMeilaniName) return hasMeilani;
    if (scIvanName && x === scIvanName) return hasIvan;
    return true;
  });
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
    ['B-Store', 'b-store']
  ];

  for (let i = 0; i < rules.length; i++) {
    if (s.indexOf(rules[i][1]) > -1) return rules[i][0];
  }
  return '';
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
    : ['Submission','Ask Detail','OR - OLD','Start','Finish','SC - Farhan','SC - Meilani','SC - Meindar','PO','Exclusion'];
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
  const tz = (Session && Session.getScriptTimeZone) ? (Session.getScriptTimeZone() || 'Asia/Jakarta') : 'Asia/Jakarta';
  try { return Utilities.formatDate(d, tz, 'MMMM yyyy'); } catch (e) {}
  return '';
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

function enforceOperationalLayout06_(ss) {
  if (!ss || DRY_RUN) return { touched: 0 };
  const monthSheets = ['Submission', 'Ask Detail', 'Start', 'SC - Farhan', 'SC - Meilani', 'SC - Meindar', 'Finish', 'PO', 'Exclusion'];
  let touched = 0;
  for (let i = 0; i < monthSheets.length; i++) {
    const sh = ss.getSheetByName(monthSheets[i]);
    if (!sh) continue;
    if (__ensureHeaderAtColumn06_(sh, 'Submission by Month', 2)) touched++;
  }
  ['Start', 'Finish'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    if (__ensureHeaderAtColumn06_(sh, 'Service Center PIC', 14)) touched++;
  });
  return { touched: touched };
}

function refreshReportBaseFromOperational06_(ss, opts) {
  if (!ss) return { written: 0, skipped: 'missing spreadsheet' };
  const sh = ss.getSheetByName('Report Base');
  if (!sh) return { written: 0, skipped: 'Report Base not found' };
  if (DRY_RUN) return { written: 0, skipped: 'DRY_RUN' };
  const incremental = !!(opts && opts.incremental);

  const headers = [
    'Submission Date',
    'Submission by Month',
    'Claim Number',
    'Last Status',
    'Last Status Date',
    'Branch',
    'Position'
  ];

  const srcSheets = __getReportBaseSourceSheets06_(ss);
  const byClaim = Object.create(null);

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

    const idxSubDate = __findHeaderIndexFlexible06_(hdr, 'Submission Date');
    const idxLast = __findHeaderIndexFlexible06_(hdr, 'Last Status');
    const idxLastDate = __findHeaderIndexFlexible06_(hdr, 'Last Status Date');
    const idxSc = __findHeaderIndexFlexible06_(hdr, 'Service Center');
    const vals = src.getRange(2, 1, lr - 1, lc).getValues();

    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      const claim = String(row[idxClaim] || '').trim();
      if (!claim) continue;

      const lastStatus = (idxLast !== -1) ? String(row[idxLast] || '').trim() : '';
      const subDateVal = (idxSubDate !== -1) ? row[idxSubDate] : '';
      const lastDateVal = (idxLastDate !== -1) ? row[idxLastDate] : '';
      const scVal = (idxSc !== -1) ? row[idxSc] : '';

      const lastDateObj = __parseAnyDateReportBase06_(lastDateVal);
      const lastDateTs = lastDateObj ? lastDateObj.getTime() : -1;
      const key = claim.toUpperCase();
      const prev = byClaim[key];
      if (prev && prev.lastDateTs > lastDateTs) continue;

      let position = '';
      if (name === 'Exclusion') position = 'Exclusion';
      else if (typeof getPositionFromLastStatus_ === 'function') {
        try { position = getPositionFromLastStatus_(lastStatus); } catch (eP) { position = ''; }
      }

      byClaim[key] = {
        subDate: __parseAnyDateReportBase06_(subDateVal) || subDateVal || '',
        subMonth: __formatSubmissionMonthReportBase06_(subDateVal),
        claim: claim,
        lastStatus: lastStatus,
        lastDate: lastDateObj || lastDateVal || '',
        branch: __getBranchFromServiceCenter06_(scVal || ''),
        position: position || '',
        lastDateTs: lastDateTs
      };
    }
  }

  const rows = Object.keys(byClaim).map(k => byClaim[k]).map(x => [
    x.subDate,
    x.subMonth,
    x.claim,
    x.lastStatus,
    x.lastDate,
    x.branch,
    x.position
  ]);

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  try { sh.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center').setVerticalAlignment('middle'); } catch (e0) {}

  if (!incremental) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
      try { sh.getRange(2, 1, rows.length, 1).setNumberFormat('dd MMM yy'); } catch (e1) {}
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
    try { sh.getRange(2, 5, existing.length, 1).setNumberFormat('dd MMM yy, HH:mm'); } catch (e4) {}
  }
  return { written: upserted, totalRows: existing.length, sheets: srcSheets.length, mode: 'incremental-upsert' };
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
 * - WebApp Project movement tracking (Daily/Past) with snapshot sheets
 *
 * NOTE: This file does NOT assume Raw OLD/Raw NEW are snapshots.
 * Snapshot baseline is stored ONLY inside WebApp Project spreadsheet.
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


/** =========================
 * WebApp Project Movement Tracking (Daily/Past)
 * Baseline snapshots are stored inside WebApp Project spreadsheet only.
 *
 * Expected caller supplies "current state" rows (already post-Sub updates)
 * as an array of objects:
 * {
 *   claimNumber, lastStatus, lastUpdateDatetime, activityLog, activityLogDatetime,
 *   position, statusType, branch
 * }
 * =========================
 */

function __ensureSheetWithHeader06c_(ss, name, header, opts) {
  const options = opts || {};
  const strict = !!options.strict; // strict=true is safe for internal snapshot sheets

  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    try { sh.getRange(1, 1, 1, header.length).setValues([header]); } catch (e0) {}
    return sh;
  }

  const want = (header || []).map(__normalizeHeaderText06_);
  const lc = sh.getLastColumn();
  const probe = Math.max(want.length, Math.max(1, lc));
  const cur = sh.getRange(1, 1, 1, probe).getValues()[0].map(__normalizeHeaderText06_);
  const samePrefix = want.length && want.every((h, i) => cur[i] === h);

  // If header is empty (new/blank sheet), set it.
  const isHeaderBlank = cur.every(v => !String(v || '').trim());
  if (isHeaderBlank) {
    try { sh.getRange(1, 1, 1, want.length).setValues([want]); } catch (e1) {}
    return sh;
  }

  // Daily/Past should never be wiped. Only snapshot sheets can be strict.
  if (!samePrefix && strict) {
    try {
      sh.clearContents();
      sh.getRange(1, 1, 1, want.length).setValues([want]);
    } catch (e2) {}
  }

  return sh;
}

function __readSnapshotMap06c_(sheet) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return Object.create(null);

  const header = sheet.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
  if (idxClaim === -1) return Object.create(null);

  const n = lr - 1;
  const vals = sheet.getRange(2, 1, n, lc).getValues();
  const map = Object.create(null);
  for (let i = 0; i < n; i++) {
    const claim = String(vals[i][idxClaim] || '').trim().toUpperCase();
    if (!claim) continue;
    map[claim] = vals[i];
  }
  map.__header = header;
  return map;
}

function __hashMovementRow06c_(obj) {
  // Deterministic compact signature
  const parts = [
    obj.claimNumber || '',
    obj.lastStatus || '',
    String(obj.lastUpdateDatetime || ''),
    obj.activityLog || '',
    String(obj.activityLogDatetime || ''),
    obj.position || '',
    obj.statusType || '',
    obj.branch || ''
  ];
  return parts.map(x => String(x)).join('||');
}

/**
 * Track movements to WebApp Project spreadsheet.
 *
 * opts:
 * - spreadsheetId (required)
 * - currentRows (required) array<object>
 * - dailySheetName (default: 'Daily')
 * - pastSheetName (default: 'Past')
 * - snapshotPrevName (default: '_SNAPSHOT_PREV')
 * - snapshotCurrName (default: '_SNAPSHOT_CURR')
 */
function trackMovementsToWebAppProject06c_(opts) {
  if (DRY_RUN) return { appended: 0, rolled: false };

  const spreadsheetId = opts && opts.spreadsheetId;
  const currentRows = (opts && opts.currentRows) ? opts.currentRows : [];
  if (!spreadsheetId) return { appended: 0, rolled: false, error: 'missing_spreadsheetId' };
  if (!Array.isArray(currentRows)) return { appended: 0, rolled: false, error: 'currentRows_not_array' };

  const dailyName = (opts && opts.dailySheetName) || 'Daily';
  const pastName = (opts && opts.pastSheetName) || 'Past';
  const prevName = (opts && opts.snapshotPrevName) || '_SNAPSHOT_PREV';
  const currName = (opts && opts.snapshotCurrName) || '_SNAPSHOT_CURR';

  const lock = LockService.getScriptLock();
  try { lock.tryLock(25000); } catch (e) { return { appended: 0, rolled: false, error: 'lock_failed' }; }

  let ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e0) {
    try { lock.releaseLock(); } catch (e) {}
    return { appended: 0, rolled: false, error: 'open_failed' };
  }

  // Ensure sheets exist + headers.
  const dailyHeader = ['Timestamp','Claim Number','Last Status','Last Update Datetime','Activity Log','Activity Log Datetime','Position','Status Type','Branch'];
  const snapHeader = ['Claim Number','Last Status','Last Update Datetime','Activity Log','Activity Log Datetime','Position','Status Type','Branch','_hash'];

  const shDaily = __ensureSheetWithHeader06c_(ss, dailyName, dailyHeader, { strict: false });
  const shPast = __ensureSheetWithHeader06c_(ss, pastName, dailyHeader, { strict: false });
  const shPrev = __ensureSheetWithHeader06c_(ss, prevName, snapHeader, { strict: true });
  const shCurr = __ensureSheetWithHeader06c_(ss, currName, snapHeader, { strict: true });

  // Daily rollover by date (Asia/Jakarta is expected via project settings).
  const props = PropertiesService.getScriptProperties();
  const key = 'WEBAPP_MOVEMENT_LAST_DATE__' + spreadsheetId;
  const tz = (opts && opts.timezone) ? String(opts.timezone) : Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const last = String(props.getProperty(key) || '').trim();
  let rolled = false;
  if (last && last !== today) {
    const lrD = shDaily.getLastRow();
    if (lrD > 1) {
      const data = shDaily.getRange(2, 1, lrD - 1, dailyHeader.length).getValues();
      const dest = shPast.getLastRow() + 1;
      try { shPast.getRange(dest, 1, data.length, dailyHeader.length).setValues(data); } catch (e1) {}
      try { shDaily.getRange(2, 1, lrD - 1, dailyHeader.length).clearContent(); } catch (e2) {}
    }
    rolled = true;
  }
  if (!last) props.setProperty(key, today);
  if (rolled) props.setProperty(key, today);

  // Read previous snapshot map.
  const prevMap = __readSnapshotMap06c_(shPrev);

  // Build snapshot curr values + detect diffs.
  const appendedRows = [];
  const snapVals = [];

  for (let i = 0; i < currentRows.length; i++) {
    const r = currentRows[i] || {};
    const claim = String(r.claimNumber || '').trim();
    if (!claim) continue;

    const obj = {
      claimNumber: claim,
      lastStatus: r.lastStatus || '',
      lastUpdateDatetime: r.lastUpdateDatetime || '',
      activityLog: r.activityLog || '',
      activityLogDatetime: r.activityLogDatetime || '',
      position: r.position || '',
      statusType: r.statusType || '',
      branch: r.branch || ''
    };
    const hash = __hashMovementRow06c_(obj);

    snapVals.push([
      obj.claimNumber,
      obj.lastStatus,
      obj.lastUpdateDatetime,
      obj.activityLog,
      obj.activityLogDatetime,
      obj.position,
      obj.statusType,
      obj.branch,
      hash
    ]);

    const prev = prevMap[String(claim).toUpperCase()];
    const prevHash = prev ? String(prev[8] || '') : '';
    if (!prev || prevHash !== hash) {
      appendedRows.push([
        new Date(),
        obj.claimNumber,
        obj.lastStatus,
        obj.lastUpdateDatetime,
        obj.activityLog,
        obj.activityLogDatetime,
        obj.position,
        obj.statusType,
        obj.branch
      ]);
    }
  }

  // Append to Daily.
  if (appendedRows.length) {
    const dest = shDaily.getLastRow() + 1;
    try { shDaily.getRange(dest, 1, appendedRows.length, dailyHeader.length).setValues(appendedRows); } catch (e3) {}
  }

  // Write curr snapshot.
  try {
    const lr = shCurr.getLastRow();
    if (lr > 1) shCurr.getRange(2, 1, lr - 1, snapHeader.length).clearContent();
  } catch (e4) {}
  if (snapVals.length) {
    try { shCurr.getRange(2, 1, snapVals.length, snapHeader.length).setValues(snapVals); } catch (e5) {}
  }

  // Replace prev = curr (overwrite), then clear curr to keep only one baseline.
  try {
    const lrP = shPrev.getLastRow();
    if (lrP > 1) shPrev.getRange(2, 1, lrP - 1, snapHeader.length).clearContent();
  } catch (e6) {}
  if (snapVals.length) {
    try { shPrev.getRange(2, 1, snapVals.length, snapHeader.length).setValues(snapVals); } catch (e7) {}
  }

  try { lock.releaseLock(); } catch (e8) {}
  return { appended: appendedRows.length, rolled: rolled };
}

/**
 * Build movement-tracking current rows from operational sheets (post-run state).
 * Best-effort: only reads columns that exist.
 *
 * Output is compatible with trackMovementsToWebAppProject06c_({ currentRows }).
 */
function buildMovementCurrentRowsFromOperationalSheets06c_(ss, pic) {
  if (!ss) return [];

  const isAdmin = (pic === 'Admin');
  const ops = isAdmin ? (CONFIG.sheetsByPic && CONFIG.sheetsByPic.adminOperational) : (CONFIG.sheetsByPic && CONFIG.sheetsByPic.picOperational);
  const sheetNames = (ops || []).slice().filter(Boolean);
  if (!sheetNames.length) return [];

  const out = [];

  for (let i = 0; i < sheetNames.length; i++) {
    const name = sheetNames[i];
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);

    const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
    if (idxClaim === -1) continue;

    let idxLastStatus = __findHeaderIndexFlexible06_(header, 'Last Status');
    if (idxLastStatus === -1) idxLastStatus = __findHeaderIndexFlexible06_(header, 'Status');

    const idxLastStatusDate = __findHeaderIndexFlexible06_(header, 'Last Status Date');
    const idxActivityLog = __findHeaderIndexFlexible06_(header, 'Activity Log'); // optional
    const idxActivityLogDt = __findHeaderIndexFlexible06_(header, 'Activity Log Datetime'); // optional
    const idxStatusType = __findHeaderIndexFlexible06_(header, 'Status Type');
    const idxBranch = __findHeaderIndexFlexible06_(header, 'Branch');

    const n = lr - 1;
    const vals = sh.getRange(2, 1, n, lc).getValues();

    for (let r = 0; r < n; r++) {
      const row = vals[r];
      const claim = String(row[idxClaim] || '').trim();
      if (!claim) continue;

      const lastStatus = (idxLastStatus !== -1) ? String(row[idxLastStatus] || '').trim() : '';
      const statusType = (idxStatusType !== -1) ? String(row[idxStatusType] || '').trim() : getStatusType06c_(lastStatus);

      out.push({
        claimNumber: claim,
        lastStatus: lastStatus,
        lastUpdateDatetime: (idxLastStatusDate !== -1) ? row[idxLastStatusDate] : '',
        activityLog: (idxActivityLog !== -1) ? String(row[idxActivityLog] || '') : '',
        activityLogDatetime: (idxActivityLogDt !== -1) ? row[idxActivityLogDt] : '',
        position: name,
        statusType: statusType,
        branch: (idxBranch !== -1) ? String(row[idxBranch] || '') : ''
      });
    }
  }

  return out;
}

/** Convenience wrapper: build from operational sheets then write to WebApp Project. */
function trackMovementsFromOperationalToWebAppProject06c_(mainSpreadsheet, pic, webappSpreadsheetId, opts) {
  if (!mainSpreadsheet || !webappSpreadsheetId) return { appended: 0, rolled: false, error: 'missing_inputs' };
  const currentRows = buildMovementCurrentRowsFromOperationalSheets06c_(mainSpreadsheet, pic);
  return trackMovementsToWebAppProject06c_(Object.assign({}, opts || {}, {
    spreadsheetId: webappSpreadsheetId,
    currentRows: currentRows
  }));
}




/** =====================================================================
 * Movement Claim Tracking (WebApp Project) — Snapshot-from-Raw implementation
 * Spec (Feb 2026):
 * - PREV snapshots are taken from Raw OLD/Raw NEW BEFORE SUB overwrites Raw
 * - CURR snapshots are taken AFTER SUB completes
 * - Daily rows are emitted on:
 *    - STATUS change (last_status differs) OR new claim (prev missing)
 *    - ACTIVITY change (activity_log differs AND activity_log_datetime is newer)
 * - If both occur for a claim, emit STATUS row first, then ACTIVITY row.
 * - Dedup is enforced via deterministic Event ID (hidden rightmost column).
 * ===================================================================== */

/** Coerce Date/string/number to millis; returns NaN if invalid. */
function __coerceMillis06c_(v) {
  if (v == null || v === '') return NaN;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    const t = v.getTime();
    return isNaN(t) ? NaN : t;
  }
  if (typeof v === 'number') return isNaN(v) ? NaN : v;
  const s = String(v).trim();
  if (!s) return NaN;
  try {
    if (typeof parseClaimLastUpdatedDatetime06c_ === 'function') {
      const d0 = parseClaimLastUpdatedDatetime06c_(s);
      if (d0 && !isNaN(d0.getTime())) return d0.getTime();
    }
  } catch (e0) {}

  // Controlled fallback:
  // Only allow native Date parsing for unambiguous strings that contain a timezone.
  // This avoids silent WIB shifts when the source string has no TZ indicator.
  if (/[zZ]$/.test(s) || /[+-]\d{2}:?\d{2}/.test(s) || /\bGMT\b/i.test(s)) {
    const d = new Date(s);
    const t = d.getTime();
    return isNaN(t) ? NaN : t;
  }
  return NaN;
}

/** Minutes -> "1d 12h 12m" (compact). */
function __formatGapDhm06c_(minutes) {
  const m0 = Number(minutes);
  if (!isFinite(m0) || m0 < 0) return '';
  const m = Math.floor(m0);
  const d = Math.floor(m / 1440);
  const h = Math.floor((m % 1440) / 60);
  const mm = m % 60;
  const parts = [];
  if (d) parts.push(d + 'd');
  if (h || d) parts.push(h + 'h');
  parts.push(mm + 'm');
  return parts.join(' ');
}

/** SHA-256 hex (short) */
function __hashHexShort06c_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s || ''), Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = (bytes[i] < 0) ? bytes[i] + 256 : bytes[i];
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex.slice(0, 16); // compact
}

function __webappOpenSs06c_() {
  try {
    if (typeof WEBAPP_MOVEMENT_POLICY === 'undefined' || !WEBAPP_MOVEMENT_POLICY) return null;
    if (!WEBAPP_MOVEMENT_POLICY.ENABLE) return null;
    const id = WEBAPP_MOVEMENT_POLICY.SPREADSHEET_ID;
    if (!id) return null;
    return SpreadsheetApp.openById(id);
  } catch (e) {
    return null;
  }
}

/**
 * Ensure sheet exists and required headers are present (append missing headers).
 * Returns { sheet, header, idxByName } where idxByName maps required header -> 1-based column index.
 */
function __ensureHeaders06c_(ss, sheetName, requiredHeaders) {
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lc = Math.max(1, sh.getLastColumn());
  const existing = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
  let header = existing.slice();

  const missing = [];
  for (let i = 0; i < requiredHeaders.length; i++) {
    const h = requiredHeaders[i];
    const idx = __findHeaderIndexFlexible06_(header, h);
    if (idx === -1) missing.push(h);
  }
  if (!header.filter(Boolean).length) {
    header = requiredHeaders.slice();
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  } else if (missing.length) {
    const startCol = header.length + 1;
    header = header.concat(missing);
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }

  const idxByName = Object.create(null);
  for (let i = 0; i < requiredHeaders.length; i++) {
    const h = requiredHeaders[i];
    const idx0 = __findHeaderIndexFlexible06_(header, h);
    if (idx0 > -1) idxByName[h] = idx0 + 1; // 1-based
  }

  return { sheet: sh, header: header, idxByName: idxByName };
}

/** Overwrite a sheet as a clean table (header + rows). */
function __overwriteTable06c_(ss, sheetName, headers, rows) {
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();
  const all = [headers].concat(rows || []);
  if (!all.length) return 0;
  sh.getRange(1, 1, all.length, headers.length).setValues(all);
  return (all.length - 1);
}

/** Build snapshot rows (2D) from a Raw sheet. */
function __buildSnapshotRowsFromRaw06c_(rawSheet) {
  if (!rawSheet) return [];
  const lr = rawSheet.getLastRow();
  const lc = rawSheet.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const header = rawSheet.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);

  function idxAny(cands) {
    for (let i = 0; i < cands.length; i++) {
      const idx = __findHeaderIndexFlexible06_(header, cands[i]);
      if (idx > -1) return idx;
    }
    return -1;
  }

  const idxClaim = idxAny(['claim_number', 'Claim Number', 'claim number']);
  const idxStatus = idxAny(['last_status', 'Last Status', 'Status']);
  const idxLastUpd = idxAny(['claim_last_updated_datetime', 'Claim Last Updated Datetime', 'claim_last_updated', 'last_update', 'Last Update']);
  const idxAct = idxAny(['activity_log', 'Activity Log', 'last_activity_log', 'Last Activity Log']);
  const idxActDt = idxAny(['activity_log_timestamp', 'Activity Log Datetime', 'last_activity_log_date', 'Last Activity Log Date', 'last_activity_log_timestamp']);
  const idxSc = idxAny(['sc_name', 'Service Center Name', 'SC Name', 'Service Center']);

  if (idxClaim === -1) return [];

  const n = lr - 1;
  const vals = rawSheet.getRange(2, 1, n, lc).getValues();
  const out = [];

  for (let i = 0; i < n; i++) {
    const claim = String(vals[i][idxClaim] || '').trim();
    if (!claim) continue;

    const lastStatus = (idxStatus > -1) ? String(vals[i][idxStatus] || '').trim() : '';
    let lastUpd = (idxLastUpd > -1) ? vals[i][idxLastUpd] : '';
    try { const p = (typeof parseClaimLastUpdatedDatetime06c_ === 'function') ? parseClaimLastUpdatedDatetime06c_(lastUpd) : null; if (p) lastUpd = p; } catch (eD1) {}
    const act = (idxAct > -1) ? String(vals[i][idxAct] || '').trim() : '';
    let actDt = (idxActDt > -1) ? vals[i][idxActDt] : '';
    try { const p2 = (typeof parseClaimLastUpdatedDatetime06c_ === 'function') ? parseClaimLastUpdatedDatetime06c_(actDt) : null; if (p2) actDt = p2; } catch (eD2) {}
    const scName = (idxSc > -1) ? String(vals[i][idxSc] || '').trim() : '';

    let branch = '';
    try { branch = __getBranchFromServiceCenter06_(scName); } catch (eB) { branch = ''; }

    let position = '';
    try { position = (typeof getPositionFromLastStatus_ === 'function') ? getPositionFromLastStatus_(lastStatus) : ''; } catch (eP) { position = ''; }

    let statusType = '';
    try {
      if (typeof getStatusTypeFromLastStatus_ === 'function') statusType = getStatusTypeFromLastStatus_(lastStatus);
      else if (typeof getStatusType06c_ === 'function') statusType = getStatusType06c_(lastStatus);
      else if (typeof getStatusTypeFromLastStatus06b_ === 'function') statusType = getStatusTypeFromLastStatus06b_(lastStatus);
    } catch (eS) { statusType = ''; }
    if (!statusType) statusType = '';

    out.push([
      claim,
      lastStatus,
      lastUpd,
      act,
      actDt,
      scName,
      branch,
      position,
      statusType
    ]);
  }

  return out;
}

/**
 * PREV snapshot: called BEFORE SUB overwrites Raw.
 * Overwrites SNAPSHOT_PREV_OLD and SNAPSHOT_PREV_NEW inside WebApp Project spreadsheet.
 */
function webappMovementSnapshotPrevForSub06c_(masterSs, rawOldName, rawNewName) {
  const webappSs = __webappOpenSs06c_();
  if (!webappSs || !masterSs) return { ok: false, reason: 'webapp_or_master_missing' };

  const policy = WEBAPP_MOVEMENT_POLICY;
  const sheets = policy.SHEETS;

  const rawOld = masterSs.getSheetByName(String(rawOldName || SUB_FLOW_SPEC.RAW_OLD_SHEET_NAME || 'Raw OLD'));
  const rawNew = masterSs.getSheetByName(String(rawNewName || SUB_FLOW_SPEC.RAW_NEW_SHEET_NAME || 'Raw NEW'));

  const rowsOld = __buildSnapshotRowsFromRaw06c_(rawOld);
  const rowsNew = __buildSnapshotRowsFromRaw06c_(rawNew);

  __overwriteTable06c_(webappSs, sheets.SNAPSHOT_PREV_OLD, policy.SNAPSHOT_COLUMNS, rowsOld);
  __overwriteTable06c_(webappSs, sheets.SNAPSHOT_PREV_NEW, policy.SNAPSHOT_COLUMNS, rowsNew);

  return { ok: true, prevOld: rowsOld.length, prevNew: rowsNew.length };
}

/**
 * CURR snapshot + movement diff:
 * - Overwrites SNAPSHOT_CURR_OLD / SNAPSHOT_CURR_NEW
 * - Appends movement events into Daily (dedup by Event ID)
 */
function webappMovementSnapshotCurrAndTrackForSub06c_(masterSs, rawOldName, rawNewName) {
  const webappSs = __webappOpenSs06c_();
  if (!webappSs || !masterSs) return { ok: false, reason: 'webapp_or_master_missing' };

  const policy = WEBAPP_MOVEMENT_POLICY;
  const sheets = policy.SHEETS;

  const rawOld = masterSs.getSheetByName(String(rawOldName || SUB_FLOW_SPEC.RAW_OLD_SHEET_NAME || 'Raw OLD'));
  const rawNew = masterSs.getSheetByName(String(rawNewName || SUB_FLOW_SPEC.RAW_NEW_SHEET_NAME || 'Raw NEW'));

  const rowsOld = __buildSnapshotRowsFromRaw06c_(rawOld);
  const rowsNew = __buildSnapshotRowsFromRaw06c_(rawNew);

  __overwriteTable06c_(webappSs, sheets.SNAPSHOT_CURR_OLD, policy.SNAPSHOT_COLUMNS, rowsOld);
  __overwriteTable06c_(webappSs, sheets.SNAPSHOT_CURR_NEW, policy.SNAPSHOT_COLUMNS, rowsNew);

  const existing = __loadExistingEventIds06c_(webappSs, sheets.DAILY, sheets.PAST);

  const now = new Date();

  const r1 = __diffSnapshotsAndAppendDaily06c_(webappSs, {
    db: 'OLD',
    prevSheetName: sheets.SNAPSHOT_PREV_OLD,
    currSheetName: sheets.SNAPSHOT_CURR_OLD,
    dailySheetName: sheets.DAILY,
    requiredDailyHeaders: policy.DAILY_COLUMNS,
    existingEventIds: existing,
    now: now
  });

  const r2 = __diffSnapshotsAndAppendDaily06c_(webappSs, {
    db: 'NEW',
    prevSheetName: sheets.SNAPSHOT_PREV_NEW,
    currSheetName: sheets.SNAPSHOT_CURR_NEW,
    dailySheetName: sheets.DAILY,
    requiredDailyHeaders: policy.DAILY_COLUMNS,
    existingEventIds: existing,
    now: now
  });

  return {
    ok: true,
    currOld: rowsOld.length,
    currNew: rowsNew.length,
    appendedOld: r1.appended,
    appendedNew: r2.appended,
    skippedOld: r1.skipped,
    skippedNew: r2.skipped
  };
}

function __getPastEventScanMaxRows06c_() {
  try {
    if (typeof WEBAPP_MOVEMENT_POLICY !== 'undefined' && WEBAPP_MOVEMENT_POLICY) {
      const n = Number(WEBAPP_MOVEMENT_POLICY.PAST_EVENT_SCAN_MAX_ROWS || 0);
      if (isFinite(n) && n > 0) return Math.floor(n);
    }
  } catch (e) {}
  return 5000;
}

/** Load existing Event IDs from Daily and Past. */
function __loadExistingEventIds06c_(webappSs, dailyName, pastName) {
  const set = Object.create(null);

  function loadOne(name, maxRows) {
    const sh = webappSs.getSheetByName(name);
    if (!sh) return;
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) return;
    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(__normalizeHeaderText06_);
    const idxEvent = __findHeaderIndexFlexible06_(header, 'Event ID');
    if (idxEvent === -1) return;
    const nAll = lr - 1;
    const n = (maxRows && maxRows > 0) ? Math.min(nAll, maxRows) : nAll;
    if (n <= 0) return;
    const startRow = lr - n + 1;
    const vals = sh.getRange(startRow, idxEvent + 1, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      const v = String(vals[i][0] || '').trim();
      if (v) set[v] = true;
    }
  }

  loadOne(dailyName, 0);
  loadOne(pastName, __getPastEventScanMaxRows06c_());
  return set;
}

/**
 * Diff prev/curr snapshot sheets and append events to Daily.
 * opts:
 * - db: 'OLD'|'NEW'
 * - prevSheetName, currSheetName
 * - dailySheetName
 * - requiredDailyHeaders (policy order)
 * - existingEventIds: object-as-set (mutated)
 * - now: Date
 */
function __diffSnapshotsAndAppendDaily06c_(webappSs, opts) {
  const db = String(opts.db || '').trim().toUpperCase();
  const prevSh = webappSs.getSheetByName(opts.prevSheetName);
  const currSh = webappSs.getSheetByName(opts.currSheetName);
  if (!prevSh || !currSh) return { appended: 0, skipped: 0, reason: 'missing_snapshot' };

  const prevMap = __readSnapshotMap06c_(prevSh);
  const currMap = __readSnapshotMap06c_(currSh);

  const header = (currMap && currMap.__header) ? currMap.__header : (currSh.getRange(1, 1, 1, currSh.getLastColumn()).getValues()[0].map(__normalizeHeaderText06_));
  const idxClaim = __findHeaderIndexFlexible06_(header, 'Claim Number');
  const idxLastStatus = __findHeaderIndexFlexible06_(header, 'Last Status');
  const idxLastUpd = __findHeaderIndexFlexible06_(header, 'Last Update Datetime');
  const idxAct = __findHeaderIndexFlexible06_(header, 'Activity Log');
  const idxActDt = __findHeaderIndexFlexible06_(header, 'Activity Log Datetime');
  const idxSc = __findHeaderIndexFlexible06_(header, 'Service Center Name');
  const idxBranch = __findHeaderIndexFlexible06_(header, 'Branch');
  const idxPos = __findHeaderIndexFlexible06_(header, 'Position');
  const idxSt = __findHeaderIndexFlexible06_(header, 'Status Type');

  // Ensure Daily headers exist.
  const dailyMeta = __ensureHeaders06c_(webappSs, opts.dailySheetName, opts.requiredDailyHeaders);
  const dailySh = dailyMeta.sheet;
  const idxDaily = dailyMeta.idxByName;

  const existing = opts.existingEventIds || Object.create(null);

  function s(v) { return String(v == null ? '' : v).trim(); }
  function upper(v) { return s(v).toUpperCase(); }

  const rowsToAppend = [];
  let skipped = 0;

  // Iterate curr keys (ignore deletions).
  Object.keys(currMap).forEach(k => {
    if (k === '__header') return;
    const claimKey = k; // already upper
    const afterRow = currMap[claimKey];
    if (!afterRow) return;

    const beforeRow = prevMap[claimKey];

    const afterStatus = (idxLastStatus > -1) ? s(afterRow[idxLastStatus]) : '';
    const afterLastUpd = (idxLastUpd > -1) ? afterRow[idxLastUpd] : '';
    const afterAct = (idxAct > -1) ? s(afterRow[idxAct]) : '';
    const afterActDt = (idxActDt > -1) ? afterRow[idxActDt] : '';
    const afterSc = (idxSc > -1) ? s(afterRow[idxSc]) : '';
    const afterBranch = (idxBranch > -1) ? s(afterRow[idxBranch]) : '';
    const afterPos = (idxPos > -1) ? s(afterRow[idxPos]) : '';
    const afterStatusType = (idxSt > -1) ? s(afterRow[idxSt]) : '';

    let statusChanged = false;
    let activityChanged = false;

    let beforeStatus = '';
    let beforeLastUpd = '';
    let beforeAct = '';
    let beforeActDt = '';

    if (!beforeRow) {
      // New claim in current snapshot => treat as STATUS event
      statusChanged = true;
    } else {
      beforeStatus = (idxLastStatus > -1) ? s(beforeRow[idxLastStatus]) : '';
      beforeLastUpd = (idxLastUpd > -1) ? beforeRow[idxLastUpd] : '';
      beforeAct = (idxAct > -1) ? s(beforeRow[idxAct]) : '';
      beforeActDt = (idxActDt > -1) ? beforeRow[idxActDt] : '';

      statusChanged = upper(afterStatus) !== upper(beforeStatus);

      if (upper(afterAct) !== upper(beforeAct)) {
        const tAfter = __coerceMillis06c_(afterActDt);
        const tBefore = __coerceMillis06c_(beforeActDt);
        // Guard: require "newer" when both parse; otherwise accept change if after has datetime and before doesn't.
        if (isFinite(tAfter) && isFinite(tBefore)) activityChanged = tAfter > tBefore;
        else if (isFinite(tAfter) && !isFinite(tBefore)) activityChanged = true;
        else activityChanged = true; // last-resort, still record change
      }
    }

    // Emit STATUS row
    if (statusChanged) {
      const bMs = __coerceMillis06c_(beforeLastUpd);
      const aMs = __coerceMillis06c_(afterLastUpd);
      const gapMin = (isFinite(bMs) && isFinite(aMs)) ? Math.max(0, Math.floor((aMs - bMs) / 60000)) : '';
      const gapDisp = (gapMin === '') ? '' : __formatGapDhm06c_(gapMin);

      const base = [db, claimKey, 'STATUS', upper(beforeStatus), String(isFinite(bMs) ? bMs : ''), upper(afterStatus), String(isFinite(aMs) ? aMs : '')].join('|');
      const eventId = 'E' + __hashHexShort06c_(base);

      if (existing[eventId]) {
        skipped++;
      } else {
        existing[eventId] = true;
        const row = new Array(dailyMeta.header.length).fill('');
        row[idxDaily['Timestamp'] - 1] = opts.now;
        row[idxDaily['DB'] - 1] = db;
        row[idxDaily['Claim Number'] - 1] = claimKey;
        row[idxDaily['Change Type'] - 1] = 'STATUS';

        row[idxDaily['Last Status (Before)'] - 1] = beforeStatus || '';
        row[idxDaily['Last Update Datetime (Before)'] - 1] = beforeLastUpd || '';
        row[idxDaily['Last Status (After)'] - 1] = afterStatus || '';
        row[idxDaily['Last Update Datetime (After)'] - 1] = afterLastUpd || '';
        row[idxDaily['Gap Time Status (Minutes)'] - 1] = (gapMin === '' ? '' : gapMin);
        row[idxDaily['Gap Time Status'] - 1] = gapDisp;

        // Activity fields blank (strict)
        row[idxDaily['Activity Log'] - 1] = '';
        row[idxDaily['Activity Log Datetime'] - 1] = '';

        // Context from AFTER
        row[idxDaily['Service Center Name'] - 1] = afterSc;
        row[idxDaily['Branch'] - 1] = afterBranch;
        row[idxDaily['Position'] - 1] = afterPos;
        row[idxDaily['Status Type'] - 1] = afterStatusType;

        row[idxDaily['Event ID'] - 1] = eventId;
        rowsToAppend.push(row);
      }
    }

    // Emit ACTIVITY row
    if (activityChanged) {
      const aMs = __coerceMillis06c_(afterActDt);
      const base = [db, claimKey, 'ACTIVITY', afterAct, String(isFinite(aMs) ? aMs : '')].join('|');
      const eventId = 'E' + __hashHexShort06c_(base);

      if (existing[eventId]) {
        skipped++;
      } else {
        existing[eventId] = true;
        const row = new Array(dailyMeta.header.length).fill('');
        row[idxDaily['Timestamp'] - 1] = opts.now;
        row[idxDaily['DB'] - 1] = db;
        row[idxDaily['Claim Number'] - 1] = claimKey;
        row[idxDaily['Change Type'] - 1] = 'ACTIVITY';

        // Status fields blank (strict)
        row[idxDaily['Last Status (Before)'] - 1] = '';
        row[idxDaily['Last Update Datetime (Before)'] - 1] = '';
        row[idxDaily['Last Status (After)'] - 1] = '';
        row[idxDaily['Last Update Datetime (After)'] - 1] = '';
        row[idxDaily['Gap Time Status (Minutes)'] - 1] = '';
        row[idxDaily['Gap Time Status'] - 1] = '';

        row[idxDaily['Activity Log'] - 1] = afterAct;
        row[idxDaily['Activity Log Datetime'] - 1] = afterActDt;

        // Context from AFTER
        row[idxDaily['Service Center Name'] - 1] = afterSc;
        row[idxDaily['Branch'] - 1] = afterBranch;
        row[idxDaily['Position'] - 1] = afterPos;
        row[idxDaily['Status Type'] - 1] = afterStatusType;

        row[idxDaily['Event ID'] - 1] = eventId;
        rowsToAppend.push(row);
      }
    }
  });

  if (!rowsToAppend.length) return { appended: 0, skipped: skipped };

  // Sort emitted rows: STATUS before ACTIVITY for same claim, then timestamp stable.
  rowsToAppend.sort((a, b) => {
    const ca = String(a[idxDaily['Claim Number'] - 1] || '');
    const cb = String(b[idxDaily['Claim Number'] - 1] || '');
    if (ca < cb) return -1;
    if (ca > cb) return 1;
    const ta = String(a[idxDaily['Change Type'] - 1] || '');
    const tb = String(b[idxDaily['Change Type'] - 1] || '');
    if (ta === tb) return 0;
    if (ta === 'STATUS') return -1;
    if (tb === 'STATUS') return 1;
    return 0;
  });

  const startRow = dailySh.getLastRow() + 1;
  dailySh.getRange(startRow, 1, rowsToAppend.length, dailyMeta.header.length).setValues(rowsToAppend);
  return { appended: rowsToAppend.length, skipped: skipped };
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
  const idxDbLink = idxOf('db link');
  const idxDb = idxOf('db');
  const idxPartner = idxOf('partner name');
  const idxIns = idxOf('insurance');
  const idxDeviceType = idxOf('device type');
  const idxImei = idxOf('imei/sn');
  const idxLast = idxOf('last status');
  const idxSc = idxOf('service center');
  const idxLSA = idxOf('last status aging');
  const idxALA = idxOf('activity log aging');

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
    row[idxClaim] = cn;

    if (idxDbLink >= 0) {
      row[idxDbLink] = 'LINK';
      const url = String(rec.dashboard_link || '').trim();
      if (url) richLinks.push({ rowOffset: rowsToAppend.length, url: url });
    }
    if (idxDb >= 0) row[idxDb] = dbValue;

    if (idxPartner >= 0) row[idxPartner] = rec.partner_name || '';
    if (idxIns >= 0) row[idxIns] = rec.insurance || '';
    if (idxDeviceType >= 0) row[idxDeviceType] = rec.device_type || '';
    if (idxImei >= 0) row[idxImei] = rec.device_imei || '';
    if (idxLast >= 0) row[idxLast] = rec.last_status || '';
    if (idxSc >= 0) row[idxSc] = rec.sc_name || '';
    if (idxLSA >= 0) row[idxLSA] = rec.last_status_aging || '';
    if (idxALA >= 0) row[idxALA] = rec.activity_log_aging || '';

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

      // Preserve filter: sort only the data body, not the header.
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
