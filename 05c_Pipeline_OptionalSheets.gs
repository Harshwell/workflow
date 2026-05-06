/***************************************
 * 05c_Pipeline_OptionalSheets.gs
 * Split from: 05_Pipeline_Routing_Optional.gs
 * Scope:
 *  - Optional sheets processors: B2B / Special Case / EV-Bike
 ***************************************/
'use strict';


/** Local DRY_RUN guard to avoid load-order ReferenceError. */
function __isDryRun05c__() {
  try { if (typeof isDryRun_ === 'function') return !!isDryRun_(); } catch (e) {}
  try { return !!DRY_RUN; } catch (e2) { return false; }
}


/**
 * Normalize header text for robust schema checks:
 * - trim, collapse whitespace
 * - replace NBSP (\u00A0) with space
 * - strip BOM (\uFEFF) and zero-width space (\u200B)
 */
function __normalizeHeaderText05c_(v) {
  return String(v == null ? '' : v)
    .replace(/\uFEFF/g, '')
    .replace(/\u200B/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Read normalized header row safely (supports empty/new sheets). */
function __getHeaderRow05c_(sh) {
  if (!sh) return [];
  const lastCol = Math.max((sh.getLastColumn && sh.getLastColumn()) || 0, 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(__normalizeHeaderText05c_);
}

/** DB classifier from Claim Number (SFP/SFX/SMR -> OLD, VVMAR/GADLD -> NEW) */
function computeDbValueFromClaimNumber05c_(claimNumber) {
  try {
    if (typeof computeDbValueFromClaimNumber_ === 'function') return computeDbValueFromClaimNumber_(claimNumber);
  } catch (e) {}
  const s = String(claimNumber == null ? '' : claimNumber).trim().toUpperCase();
  if (!s) return '';
  if (s.indexOf('SFP') > -1 || s.indexOf('SFX') > -1 || s.indexOf('SMR') > -1) return 'OLD';
  if (s.indexOf('VVMAR') > -1 || s.indexOf('GADLD') > -1) return 'NEW';
  return '';
}



/** =========================
 *  Enterprise helpers (local)
 *  ========================= */

function __getRuntimeFlow05c_() {
  try {
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.flowName) return String(RUNTIME.flowName);
    if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.flow) return String(RUNTIME.flow);
  } catch (e) {}
  return 'main';
}

function __getStatusType05c_(lastStatus) {
  const st = String(lastStatus == null ? '' : lastStatus).trim();
  if (!st) return '';
  try {
    if (typeof getStatusType06c_ === 'function') return getStatusType06c_(st);
  } catch (e) {}
  // Fallback (strict): unknown mapping.
  return '';
}

/** Append a column header at the far-right if missing; return updated normalized header array. */
function __ensureAppendColumnIfMissing05c_(sh, headerArr, colName) {
  const name = __normalizeHeaderText05c_(colName);
  const idx = buildHeaderIndex_(headerArr);
  if (idx[name] != null) return headerArr;
  if (__isDryRun05c__()) return headerArr.concat([name]);
  const col = headerArr.length + 1;
  try { sh.getRange(1, col).setValue(name); } catch (e) {}
  return headerArr.concat([name]);
}

/** Pick Activity Log value per flow contract: Raw Data -> last_activity_log; Raw OLD/NEW -> activity_log. */
function __pickActivityLogValue05c_(row, headerIndexRaw) {
  const flow = (__getRuntimeFlow05c_() || 'main').toLowerCase();
  const idxLast = (headerIndexRaw && headerIndexRaw['last_activity_log'] != null) ? headerIndexRaw['last_activity_log'] : null;
  const idxAct  = (headerIndexRaw && headerIndexRaw['activity_log'] != null) ? headerIndexRaw['activity_log'] : null;

  // Strict per request: main/form use last_activity_log; sub uses activity_log.
  const preferLast = (flow === 'main' || flow === 'form');
  const v = preferLast
    ? ((idxLast != null) ? row[idxLast] : ((idxAct != null) ? row[idxAct] : ''))
    : ((idxAct != null) ? row[idxAct] : ((idxLast != null) ? row[idxLast] : ''));
  return (v == null) ? '' : v;
}

function __pickActivityLogDatetimeValue05c_(row, headerIndexRaw) {
  if (!headerIndexRaw) return '';
  // Common candidates across datasets
  const candidates = [
    'activity_log_datetime',
    'last_activity_log_datetime',
    'last_activity_datetime',
    'activity_datetime'
  ];
  for (let i = 0; i < candidates.length; i++) {
    const k = candidates[i];
    if (headerIndexRaw[k] != null) return row[headerIndexRaw[k]];
  }
  return '';
}

function __parseClaimLastUpdatedDatetime05c_(v) {
  try {
    if (typeof parseClaimLastUpdatedDatetime06c_ === 'function') {
      const d0 = parseClaimLastUpdatedDatetime06c_(v);
      if (d0 && !isNaN(d0.getTime())) return d0;
    }
  } catch (e) {}
  try {
    if (typeof parseClaimLastUpdatedDatetime06b_ === 'function') {
      const d1 = parseClaimLastUpdatedDatetime06b_(v);
      if (d1 && !isNaN(d1.getTime())) return d1;
    }
  } catch (e1) {}
  try {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    const s = String(v == null ? '' : v).trim();
    if (!s) return null;
    if (typeof normalizeDate_ === 'function') {
      const d2 = normalizeDate_(s);
      if (d2 && !isNaN(d2.getTime())) return d2;
    }
    if (typeof tryNativeParseUnambiguousDate_ === 'function') {
      const d3 = tryNativeParseUnambiguousDate_(s);
      if (d3 && !isNaN(d3.getTime())) return d3;
    }
    return null;
  } catch (e2) { return null; }
}


/** =========================
 *  Small helpers (local)
 *  ========================= */

function __blankRichText_() {
  return SpreadsheetApp.newRichTextValue().setText('').build();
}

function __setDbLinkRichTextRange_(sh, colIndex0, startRow, urls) {
  if (__isDryRun05c__()) return;
  if (colIndex0 == null || colIndex0 < 0) return;
  if (!urls || !urls.length) return;
  const r = sh.getRange(startRow, colIndex0 + 1, urls.length, 1);
  const rich2d = urls.map(u => [u ? makeRichTextHyperlink_(String(u || ''), 'LINK') : __blankRichText_()]);
  r.setRichTextValues(rich2d);
}

function __groupConsecutive_(nums) {
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

function __writeRowSegments_(sh, rowNums, rowMap, headerLen) {
  if (!rowNums || !rowNums.length) return;
  const segments = __groupConsecutive_(rowNums);
  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];
    const startRow = seg[0];
    const vals = seg.map(rn => rowMap[rn]);
    safeSetValues_(sh.getRange(startRow, 1, vals.length, headerLen), vals);
  }
}

function __setBgSegments_(sh, colIndex0, rowNums, colorMap) {
  if (__isDryRun05c__()) return;
  if (colIndex0 == null || colIndex0 < 0) return;
  if (!rowNums || !rowNums.length) return;
  const segments = __groupConsecutive_(rowNums);
  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];
    const startRow = seg[0];
    const bgs = seg.map(rn => [colorMap[rn] || '']);
    sh.getRange(startRow, colIndex0 + 1, bgs.length, 1).setBackgrounds(bgs);
  }
}

function __setNotesSegments_(sh, colIndex0, rowNums, noteMap) {
  if (__isDryRun05c__()) return;
  if (colIndex0 == null || colIndex0 < 0) return;
  if (!rowNums || !rowNums.length) return;
  const segments = __groupConsecutive_(rowNums);
  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];
    const startRow = seg[0];
    const notes = seg.map(rn => [noteMap[rn] || '']);
    try { sh.getRange(startRow, colIndex0 + 1, notes.length, 1).setNotes(notes); } catch (e) {}
  }
}

function __setDbLinkRichTextSegments_(sh, colIndex0, rowNums, urlMap) {
  if (__isDryRun05c__()) return;
  if (colIndex0 == null || colIndex0 < 0) return;
  if (!rowNums || !rowNums.length) return;
  const segments = __groupConsecutive_(rowNums);
  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];
    const startRow = seg[0];
    const rich2d = seg.map(rn => [urlMap[rn] ? makeRichTextHyperlink_(String(urlMap[rn] || ''), 'LINK') : __blankRichText_()]);
    sh.getRange(startRow, colIndex0 + 1, rich2d.length, 1).setRichTextValues(rich2d);
  }
}

/** =========================
 *  Optional sheets
 *  ========================= */

/** Optional: B2B — enabled by default. Set RUNTIME.enableB2B = false to disable. */
function processB2B_(ss, rawValues, headerIndexRaw, pic) { // `pic` kept for backward compatibility
  // Default ON. Only skip when explicitly disabled.
  if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.enableB2B === false) return 0;

  const patterns = (CONFIG.patterns.b2bPartners || []).map(s => String(s || '').toLowerCase());
  const claimToken = String(
    (typeof B2B_CLAIM_NUMBER_SUBSTRING !== 'undefined' && B2B_CLAIM_NUMBER_SUBSTRING)
    || (CONFIG && CONFIG.B2B_CLAIM_NUMBER_SUBSTRING)
    || 'SMR'
  ).trim().toUpperCase();
  const h = CONFIG.headers;
  const sh = ss.getSheetByName('B2B');
  if (!sh) return 0;
  if (!rawValues || !rawValues.length) return 0;

  const OPTIONAL_FLAGS = (typeof __OPTIONAL_FLAGS !== 'undefined' && __OPTIONAL_FLAGS) ? __OPTIONAL_FLAGS : {};
  const EXCLUDED_LAST_STATUSES = getSpecialCaseExcludedStatuses_();

  const idxBP = headerIndexRaw[h.businessPartner];
  const idxClaim = headerIndexRaw[h.claimNumber];
  const idxLastStatus = headerIndexRaw[h.lastStatus];
  const idxAssociate = headerIndexRaw[h.associate];
  const idxDashboard = headerIndexRaw[h.dashboardLink];
  const idxClaimLastUpdated = (headerIndexRaw['claim_last_updated_datetime'] != null) ? headerIndexRaw['claim_last_updated_datetime'] : null;

  const idxIns = headerIndexRaw[h.insuranceCode];
  const idxInsPartner = (h.insurancePartnerName && headerIndexRaw[h.insurancePartnerName] != null)
    ? headerIndexRaw[h.insurancePartnerName]
    : headerIndexRaw['insurance_partner_name'];
  const idxSource =
    (headerIndexRaw[h.sourceSystem] != null) ? headerIndexRaw[h.sourceSystem]
    : (headerIndexRaw['source_system_name'] != null) ? headerIndexRaw['source_system_name']
    : (headerIndexRaw['source_db'] != null) ? headerIndexRaw['source_db']
    : null;

  const idxDevice = headerIndexRaw[h.deviceType];
  const idxSvc =
    (headerIndexRaw[h.scName] != null) ? headerIndexRaw[h.scName]
    : (headerIndexRaw['sc_name'] != null) ? headerIndexRaw['sc_name']
    : (headerIndexRaw[h.serviceCenter] != null) ? headerIndexRaw[h.serviceCenter]
    : (headerIndexRaw['service_center'] != null) ? headerIndexRaw['service_center']
    : null;

  const idxLSA =
    (headerIndexRaw['Last Status Aging'] != null) ? headerIndexRaw['Last Status Aging']
    : (headerIndexRaw['LSA'] != null) ? headerIndexRaw['LSA']
    : (headerIndexRaw[h.lastStatusAging] != null) ? headerIndexRaw[h.lastStatusAging]
    : (headerIndexRaw['last_status_aging'] != null) ? headerIndexRaw['last_status_aging']
    : null;
  const idxALA =
    (headerIndexRaw['Activity Log Aging'] != null) ? headerIndexRaw['Activity Log Aging']
    : (headerIndexRaw['ALA'] != null) ? headerIndexRaw['ALA']
    : (headerIndexRaw[h.activityLogAging] != null) ? headerIndexRaw[h.activityLogAging]
    : (headerIndexRaw['activity_log_aging'] != null) ? headerIndexRaw['activity_log_aging']
    : (headerIndexRaw[h.agingFromLastStatus] != null) ? headerIndexRaw[h.agingFromLastStatus]
    : null;
  const idxTAT =
    (headerIndexRaw[h.daysAgingFromSubmission] != null) ? headerIndexRaw[h.daysAgingFromSubmission]
    : (headerIndexRaw['days_aging_from_submission'] != null) ? headerIndexRaw['days_aging_from_submission']
    : (headerIndexRaw[h.tat] != null) ? headerIndexRaw[h.tat]
    : null;

  const idxQL = headerIndexRaw['Q-L (Months)'];

  const idxProduct = headerIndexRaw[h.productName];

  const idxSumInsured = headerIndexRaw[h.sumInsured];
  const idxOwnRisk = headerIndexRaw[h.ownRiskAmount];
  const idxNett = headerIndexRaw[h.nettClaimAmount];

  let headerRow = __getHeaderRow05c_(sh);

  // Fix: kalau template B2B punya leading blank column (A kosong, header mulai di B),
  // shift schema ke kolom A supaya data nulis dari A (bukan B).
  const firstNonEmpty = headerRow.findIndex(v => !!v);
  if (firstNonEmpty > 0) {
    const shifted = headerRow.slice(firstNonEmpty);
    try {
      // Overwrite header mulai kolom A.
      sh.getRange(1, 1, 1, shifted.length).setValues([shifted]);
      // Clear sisa header lama biar nggak ada "ghost header".
      const rest = sh.getLastColumn() - shifted.length;
      if (rest > 0) sh.getRange(1, shifted.length + 1, 1, rest).clearContent();
    } catch (e) {}
    headerRow = shifted;
  }

  let header = headerRow;
  let idxH = buildHeaderIndex_(header);

  // Ensure mandatory Status Type only when this sheet schema includes Last Status.
  if (idxH['Last Status'] != null) {
    header = __ensureAppendColumnIfMissing05c_(sh, header, 'Status Type');
    idxH = buildHeaderIndex_(header);
  }
  // Schema guard + self-heal for required computed columns.
  try {
    const need = ['Claim Number','Start Date','End Date','Details'].map(__normalizeHeaderText05c_);
    const missing = need.filter(n => idxH[n] == null);
    if (missing.length) {
      for (let mi = 0; mi < missing.length; mi++) {
        header = __ensureAppendColumnIfMissing05c_(sh, header, missing[mi]);
      }
      idxH = buildHeaderIndex_(header);
      if (typeof logLine_ === 'function') {
        logLine_('WARN', 'B2B_SCHEMA_HEAL', 'Auto-added columns: ' + missing.join(', '), '', '');
      }
    }
  } catch (e) {}


  const dbLinkCol0 = idxH['DB Link'];
  const rows = [];
  const dbUrls = [];
  const seenClaims = new Set();
  let rawMatchedCount = 0;
  let submissionFallbackCount = 0;
  let skippedExcludedRawCount = 0;
  let skippedExcludedSubmissionCount = 0;

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];

    const partnerLower = String((idxBP != null) ? row[idxBP] : '' || '').toLowerCase();
    const claimUp = String((idxClaim != null) ? row[idxClaim] : '' || '').toUpperCase();
    const lastStatus = String((idxLastStatus != null) ? row[idxLastStatus] : '' || '').trim();
    const lastStatusKey = lastStatus.toUpperCase();

    if (OPTIONAL_FLAGS.B2B_SKIP_EXCLUDED_LAST_STATUSES && EXCLUDED_LAST_STATUSES.has(lastStatusKey)) {
      skippedExcludedRawCount++;
      continue;
    }

    const matchPartner = patterns.some(p => p && partnerLower.indexOf(p) > -1);
    const matchClaim = claimToken ? (claimUp.indexOf(claimToken) > -1) : false;
    if (!matchPartner && !matchClaim) continue;

    const out = new Array(header.length).fill('');
    const set = (k, v) => { const j = idxH[k]; if (j != null) out[j] = v; };

    set('Submission Date', buildSubmissionDateCell_(row, headerIndexRaw));
    set('Claim Number', (idxClaim != null) ? row[idxClaim] : '');
    if (claimUp) seenClaims.add(claimUp);

    const dbUrl = (idxDashboard != null) ? row[idxDashboard] : '';
    set('DB Link', dbUrl ? 'LINK' : '');
    dbUrls.push(dbUrl);

    const src = String((idxSource != null) ? row[idxSource] : '' || '').toUpperCase();
    const dbFromClaim = computeDbValueFromClaimNumber05c_(claimUp);
    set('DB', dbFromClaim || ((src.indexOf('OLD') > -1) ? 'OLD' : (src.indexOf('NEW') > -1 ? 'NEW' : (src || ''))));

    set('Partner Name', (idxBP != null) ? row[idxBP] : '');
    set('Insurance', mapInsuranceShort_((idxInsPartner != null) ? row[idxInsPartner] : ''));
    set('Device Type', (idxDevice != null) ? row[idxDevice] : '');
    set('Service Center', (idxSvc != null) ? row[idxSvc] : '');

    // Optional columns (if present in sheet): Last Status, Activity Log, Last Status Date, Status Type
    set('Last Status', lastStatus);
    if (idxH['Activity Log'] != null) set('Activity Log', __pickActivityLogValue05c_(row, headerIndexRaw));
    if (idxH['Activity Log Datetime'] != null) set('Activity Log Datetime', __pickActivityLogDatetimeValue05c_(row, headerIndexRaw));
    if (idxH['Last Status Date'] != null) {
      const d0 = __parseClaimLastUpdatedDatetime05c_((idxClaimLastUpdated != null) ? row[idxClaimLastUpdated] : null);
      set('Last Status Date', d0 ? d0 : '');
    }
    if (idxH['Status Type'] != null) set('Status Type', __getStatusType05c_(lastStatus));

    const lsa = (idxLSA != null) ? normalizeInt_(row[idxLSA]) : null;
    const ala = (idxALA != null) ? normalizeInt_(row[idxALA]) : null;
    const tat = (idxTAT != null) ? normalizeInt_(row[idxTAT]) : null;

    set('Last Status Aging', (lsa != null) ? lsa : '');
    set('LSA', (lsa != null) ? lsa : '');
    set('Activity Log Aging', (ala != null) ? ala : '');
    set('ALA', (ala != null) ? ala : '');
    set('TAT', (tat != null) ? tat : '');

    set('Q-L (Months)', (idxQL != null) ? (normalizeInt_(row[idxQL]) ?? '') : '');

    // Product
    set('Product', (idxProduct != null) ? row[idxProduct] : '');

    // Money fields (force numeric)
    const sumIns = (idxSumInsured != null) ? normalizeNumber_(row[idxSumInsured]) : null;
    const orAmt = (idxOwnRisk != null) ? normalizeNumber_(row[idxOwnRisk]) : null;
    const nett = (idxNett != null) ? normalizeNumber_(row[idxNett]) : null;

    set('Sum Insured', (sumIns != null) ? sumIns : '');
    if (idxH['OR Amount'] != null) set('OR Amount', (orAmt != null) ? orAmt : '');
    set('Nett Claim Amount', (nett != null) ? nett : '');

    rows.push(out);
    rawMatchedCount++;
  }

  // Fallback source: Submission sheet (for claims missing in current Raw pull window).
  try {
    const subSh = ss.getSheetByName('Submission');
    if (subSh && subSh.getLastRow() > 1 && subSh.getLastColumn() > 1) {
      const subHeader = __getHeaderRow05c_(subSh);
      const subIdx = buildHeaderIndex_(subHeader);
      const subVals = subSh.getRange(2, 1, subSh.getLastRow() - 1, subHeader.length).getValues();
      const sClaim = subIdx['Claim Number'];
      const sPartner = (subIdx['Partner Name'] != null) ? subIdx['Partner Name'] : subIdx['Partner'];
      const sSubDate = subIdx['Submission Date'];
      const sDb = subIdx['DB'];
      const sDbLink = subIdx['DB Link'];
      const sInsurance = subIdx['Insurance'];
      const sDevice = subIdx['Device Type'];
      const sSc = (subIdx['Service Center'] != null) ? subIdx['Service Center'] : subIdx['Service Center Name'];
      const sLastStatus = subIdx['Last Status'];

      if (sClaim != null) {
        for (let i = 0; i < subVals.length; i++) {
          const r = subVals[i] || [];
          const claim = String(r[sClaim] || '').trim();
          if (!claim) continue;
          const claimUp = claim.toUpperCase();
          if (seenClaims.has(claimUp)) continue;
          const partner = String((sPartner != null) ? r[sPartner] : '').trim().toLowerCase();
          const lastStatus = String((sLastStatus != null) ? r[sLastStatus] : '').trim();
          const lastStatusKey = lastStatus.toUpperCase();
          const matchPartner = patterns.some(p => p && partner.indexOf(p) > -1);
          const matchClaim = claimToken ? (claimUp.indexOf(claimToken) > -1) : false;
          if (!matchPartner && !matchClaim) continue;
          if (OPTIONAL_FLAGS.B2B_SKIP_EXCLUDED_LAST_STATUSES && EXCLUDED_LAST_STATUSES.has(lastStatusKey)) {
            skippedExcludedSubmissionCount++;
            continue;
          }

          const out = new Array(header.length).fill('');
          const set = (k, v) => { const j = idxH[k]; if (j != null) out[j] = v; };
          set('Submission Date', (sSubDate != null) ? r[sSubDate] : '');
          set('Claim Number', claim);
          set('DB', (sDb != null) ? r[sDb] : computeDbValueFromClaimNumber05c_(claimUp));
          set('Partner Name', (sPartner != null) ? r[sPartner] : '');
          set('Insurance', (sInsurance != null) ? r[sInsurance] : '');
          set('Device Type', (sDevice != null) ? r[sDevice] : '');
          set('Service Center', (sSc != null) ? r[sSc] : '');
          set('Last Status', lastStatus);
          if (idxH['Status Type'] != null) set('Status Type', __getStatusType05c_(lastStatus));
          const dbUrl = (sDbLink != null) ? r[sDbLink] : '';
          set('DB Link', dbUrl ? 'LINK' : '');
          dbUrls.push(dbUrl);
          rows.push(out);
          seenClaims.add(claimUp);
          submissionFallbackCount++;
        }
      }
    }
  } catch (eSub) {}

  if (!rows.length) return 0;
  // Rebuild list only when we have replacement rows.
  // Prevent accidental "header-only" sheet when source window is temporarily empty.
  clearSheetDataHard_(sh, { bufferRows: 1200 });
  safeSetValues_(sh.getRange(2, 1, rows.length, header.length), rows);

  // DB Link as RichText (avoid #ERROR!, display "LINK")
  __setDbLinkRichTextRange_(sh, dbLinkCol0, 2, dbUrls);

  applyOperationalColumnSchema_(sh, header, 2, rows.length, { orIsMoney: false });
  applyDbLinkFormatting_(sh, header, rows.length, 2);
  try {
    if (typeof logLine_ === 'function') {
      logLine_(
        'INFO',
        'B2B_METRICS',
        'rows=' + rows.length
          + ' raw=' + rawMatchedCount
          + ' sub_fallback=' + submissionFallbackCount
          + ' skip_excluded_raw=' + skippedExcludedRawCount
          + ' skip_excluded_sub=' + skippedExcludedSubmissionCount,
        '',
        'INFO'
      );
    }
  } catch (eM) {}
  return rows.length;
}

/** Special Case excluded statuses (ongoing only) */
function getSpecialCaseExcludedStatuses_() {
  // Prefer central per-run Set builder (00/01) if present.
  try {
    if (typeof getExcludedLastStatusesSet_ === 'function') return getExcludedLastStatusesSet_();
  } catch (e) {}
  try {
    if (typeof buildExcludedLastStatusesSet_ === 'function') return buildExcludedLastStatusesSet_();
  } catch (e) {}
  try {
    if (typeof __getExcludedLastStatuses05a_ === 'function') return __getExcludedLastStatuses05a_();
  } catch (e) {}

  // Backward-compatible fallbacks.
  try {
    if (typeof EXCLUDED_LAST_STATUSES !== 'undefined' && EXCLUDED_LAST_STATUSES && EXCLUDED_LAST_STATUSES.size) {
      return EXCLUDED_LAST_STATUSES;
    }
  } catch (e) {}
  try {
    if (typeof EXCLUDED_LAST_STATUSES_BASE !== 'undefined' && Array.isArray(EXCLUDED_LAST_STATUSES_BASE)) {
      return new Set(EXCLUDED_LAST_STATUSES_BASE);
    }
  } catch (e) {}

  // Legacy hardcoded list (fallback of last resort)
  return new Set([
    'DONE_REJECTED','DONE','DONE_REPLACED','DONE_REJECT',
    'QOALA_REQUEST_SALVAGE','QOALA_CLAIM_REJECT',
    'SERVICE_CENTER_CLAIM_DONE_REJECT','SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT',
    'INSURANCE_CLAIM_WAITING_PAID_REPAIR','INSURANCE_CLAIM_PAID_REPAIR',
    'INSURANCE_CLAIM_WAITING_PAID_REPLACE','INSURANCE_CLAIM_PAID_REPLACE',
    'CUSTOMER_RECEIVE_REPLACE'
  ]);
}

function getMinSpecialCaseSubmissionYear_() {
  const rules = (typeof __SPECIAL_RULES !== 'undefined' && __SPECIAL_RULES) ? __SPECIAL_RULES : null;
  const y = Number(rules && rules.MIN_SUBMISSION_YEAR);

  // Default: no hard year gate unless explicitly configured.
  // Rationale: menghindari Special Case "kosong" kalau data historis masih relevan.
  return (Number.isFinite(y) && y >= 2000) ? y : 0;
}

/** Optional: Special Case — enabled by default. Set RUNTIME.enableSpecialCase = false to disable. */
function processSpecialCase_(ss, rawValues, headerIndexRaw, pic) { // `pic` kept for backward compatibility
  // Default ON. Only skip when explicitly disabled.
  if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.enableSpecialCase === false) {
    return { count: 0, metrics: { status: 'disabled' } };
  }

  const h = CONFIG.headers;
  const sh = ss.getSheetByName('Special Case');
  if (!sh) return { count: 0, metrics: { status: 'sheet_missing' } };
  if (!rawValues || !rawValues.length) return { count: 0, metrics: { status: 'raw_empty' } };

  // Safe defaults to avoid load-order ReferenceError.
  const SPECIAL_RULES = (typeof __SPECIAL_RULES !== 'undefined' && __SPECIAL_RULES) ? __SPECIAL_RULES : {};
  const SPECIAL_FLAGS = (typeof __SPECIAL_FLAGS !== 'undefined' && __SPECIAL_FLAGS) ? __SPECIAL_FLAGS : { MODE: 'UPSERT', COLORIZE_CLAIM_CELL: true };

  const idxBP = headerIndexRaw[h.businessPartner];
  const idxClaim = headerIndexRaw[h.claimNumber];
  const idxProduct = headerIndexRaw[h.productName];
  const idxAssociate = headerIndexRaw[h.associate];
  const idxDashboard = headerIndexRaw[h.dashboardLink];
  const idxClaimLastUpdated = (headerIndexRaw['claim_last_updated_datetime'] != null) ? headerIndexRaw['claim_last_updated_datetime'] : null;
  const idxInsurance = headerIndexRaw[h.insuranceCode];
  const idxInsurancePartner = (h.insurancePartnerName && headerIndexRaw[h.insurancePartnerName] != null)
    ? headerIndexRaw[h.insurancePartnerName]
    : headerIndexRaw['insurance_partner_name'];
  const idxDevice = headerIndexRaw[h.deviceType];
  const idxLastStatus = headerIndexRaw[h.lastStatus];
  const idxSource =
    (headerIndexRaw[h.sourceSystem] != null) ? headerIndexRaw[h.sourceSystem]
    : (headerIndexRaw['source_system_name'] != null) ? headerIndexRaw['source_system_name']
    : (headerIndexRaw['source_db'] != null) ? headerIndexRaw['source_db']
    : null;
  const idxSvc =
    (headerIndexRaw[h.scName] != null) ? headerIndexRaw[h.scName]
    : (headerIndexRaw['sc_name'] != null) ? headerIndexRaw['sc_name']
    : (headerIndexRaw[h.serviceCenter] != null) ? headerIndexRaw[h.serviceCenter]
    : (headerIndexRaw['service_center'] != null) ? headerIndexRaw['service_center']
    : null;

  const idxLSA =
    (headerIndexRaw['Last Status Aging'] != null) ? headerIndexRaw['Last Status Aging']
    : (headerIndexRaw['LSA'] != null) ? headerIndexRaw['LSA']
    : (headerIndexRaw[h.lastStatusAging] != null) ? headerIndexRaw[h.lastStatusAging]
    : (headerIndexRaw['last_status_aging'] != null) ? headerIndexRaw['last_status_aging']
    : null;
  const idxALA =
    (headerIndexRaw['Activity Log Aging'] != null) ? headerIndexRaw['Activity Log Aging']
    : (headerIndexRaw['ALA'] != null) ? headerIndexRaw['ALA']
    : (headerIndexRaw[h.activityLogAging] != null) ? headerIndexRaw[h.activityLogAging]
    : (headerIndexRaw['activity_log_aging'] != null) ? headerIndexRaw['activity_log_aging']
    : (headerIndexRaw[h.agingFromLastStatus] != null) ? headerIndexRaw[h.agingFromLastStatus]
    : null;
  const idxTAT =
    (headerIndexRaw[h.daysAgingFromSubmission] != null) ? headerIndexRaw[h.daysAgingFromSubmission]
    : (headerIndexRaw['days_aging_from_submission'] != null) ? headerIndexRaw['days_aging_from_submission']
    : (headerIndexRaw[h.tat] != null) ? headerIndexRaw[h.tat]
    : null;

  const idxDaysAgingFromSubmission = headerIndexRaw[h.daysAgingFromSubmission];

  // Sum insured: enforce raw column `sum_insured_amount` when present.
  const idxSumInsuredRaw = (headerIndexRaw['sum_insured_amount'] != null)
    ? headerIndexRaw['sum_insured_amount']
    : headerIndexRaw[h.sumInsured];

  const idxQL = headerIndexRaw['Q-L (Months)'];
  const idxMQ = headerIndexRaw['M-Q (Months)'];
  const idxPolicyStart = headerIndexRaw[h.policyStartDate];
  const idxPolicyEnd = headerIndexRaw[h.policyEndDate];
  const idxMonthPolicyAging = (headerIndexRaw['month_policy_aging'] != null) ? headerIndexRaw['month_policy_aging'] : null;
  const idxClaimSubmittedDt = (headerIndexRaw['claim_submitted_datetime'] != null) ? headerIndexRaw['claim_submitted_datetime'] : null;

  let header = __getHeaderRow05c_(sh);
  let idxH = buildHeaderIndex_(header);

  // EV-Bike cleanup: remove deprecated columns if present.
  try {
    const dropCols = new Set(['Start Date', 'End Date', 'Details'].map(__normalizeHeaderText05c_));
    const toDelete = [];
    for (let i = 0; i < header.length; i++) {
      const hk = __normalizeHeaderText05c_(header[i]);
      if (dropCols.has(hk)) toDelete.push(i + 1);
    }
    for (let i = toDelete.length - 1; i >= 0; i--) {
      sh.deleteColumn(toDelete[i]);
    }
    if (toDelete.length) {
      header = __getHeaderRow05c_(sh);
      idxH = buildHeaderIndex_(header);
    }
  } catch (eDrop) {}

  // Ensure mandatory Status Type only when this sheet schema includes Last Status.
  if (idxH['Last Status'] != null) {
    header = __ensureAppendColumnIfMissing05c_(sh, header, 'Status Type');
    idxH = buildHeaderIndex_(header);
  }
  // Schema guard (no auto-add columns). Missing columns are logged and simply skipped.
  try {
    const need = ['Claim Number','Start Date','End Date','Details'].map(__normalizeHeaderText05c_);
    const missing = need.filter(n => idxH[n] == null);
    if (missing.length && typeof logLine_ === 'function') {
      logLine_('ERROR', 'SPECIAL_CASE_SCHEMA_MISSING', 'Missing columns: ' + missing.join(', '), '', '');
    }
  } catch (e) {}


  const claimCol0 = (idxH['Claim Number'] != null) ? idxH['Claim Number'] : -1;
  const dbLinkCol0 = (idxH['DB Link'] != null) ? idxH['DB Link'] : -1;

  // Mode handling
  if (SPECIAL_FLAGS.MODE === 'REBUILD') clearSheetDataHard_(sh, { bufferRows: 1200 });

  // Existing claims map (for APPEND_NEW_ONLY / UPSERT)
  let existingRowCount = Math.max(sh.getLastRow() - 1, 0);
  const existingClaims = new Set();
  const existingMap = {};
  let existingValsAll = (existingRowCount > 0) ? sh.getRange(2, 1, existingRowCount, header.length).getValues() : [];

  if (claimCol0 > -1 && SPECIAL_FLAGS.MODE !== 'REBUILD') {
    for (let i = 0; i < existingValsAll.length; i++) {
      const c = String(existingValsAll[i][claimCol0] || '').trim().toUpperCase();
      if (!c) continue;
      existingClaims.add(c);
      existingMap[c] = (2 + i); // sheet row number
    }
  }

  const excluded = getSpecialCaseExcludedStatuses_();

  // Auto-prune: if a claim already in Special Case becomes excluded, delete its row (no blanks).
  const pruneOnExcluded = (function () {
    try {
      if (typeof SPECIAL_CASE_WRITER_POLICY !== 'undefined' && SPECIAL_CASE_WRITER_POLICY) {
        const v =
          (SPECIAL_CASE_WRITER_POLICY.PRUNE_ON_EXCLUDED != null) ? SPECIAL_CASE_WRITER_POLICY.PRUNE_ON_EXCLUDED :
          (SPECIAL_CASE_WRITER_POLICY.PRUNE_WHEN_EXCLUDED != null) ? SPECIAL_CASE_WRITER_POLICY.PRUNE_WHEN_EXCLUDED :
          null;
        if (v != null) return !!v;
      }
    } catch (e) {}
    return true;
  })();

  const skipDonePrune = (SPECIAL_RULES) ? !!SPECIAL_RULES.SKIP_EXCLUDED_LAST_STATUSES : true;

  let deletedExcluded = 0;
  if (pruneOnExcluded && skipDonePrune && SPECIAL_FLAGS.MODE === 'UPSERT' && claimCol0 > -1 && idxLastStatus != null && idxClaim != null && existingRowCount > 0) {
    const excludedByClaim = {};
    for (let i = 0; i < rawValues.length; i++) {
      const row = rawValues[i];
      const c0 = String(row[idxClaim] || '').trim();
      if (!c0) continue;
      const st0 = String((row[idxLastStatus] != null) ? row[idxLastStatus] : '').trim();
      if (st0 && excluded.has(st0)) excludedByClaim[c0.toUpperCase()] = true;
    }

    const rowsToDelete = [];
    for (const c in existingMap) {
      if (excludedByClaim[c]) rowsToDelete.push(existingMap[c]);
    }

    if (rowsToDelete.length) {
      try {
        if (typeof safeDeleteRowsDescending_ === 'function') safeDeleteRowsDescending_(sh, rowsToDelete);
        else {
          rowsToDelete.sort((a, b) => b - a);
          for (let i = 0; i < rowsToDelete.length; i++) sh.deleteRow(rowsToDelete[i]);
        }
      } catch (e) {
        // best-effort: log and continue without pruning
        try {
          if (typeof logLine_ === 'function') logLine_('WARN', 'SPECIAL_CASE_DELETE_FAILED', String(e), '', '');
        } catch (e2) {}
      }
      deletedExcluded = rowsToDelete.length;

      // Refresh existing cache after deletes (row numbers shift).
      existingRowCount = Math.max(sh.getLastRow() - 1, 0);
      existingValsAll = (existingRowCount > 0) ? sh.getRange(2, 1, existingRowCount, header.length).getValues() : [];

      existingClaims.clear();
      for (const k in existingMap) delete existingMap[k];

      if (claimCol0 > -1) {
        for (let i = 0; i < existingValsAll.length; i++) {
          const c = String(existingValsAll[i][claimCol0] || '').trim().toUpperCase();
          if (!c) continue;
          existingClaims.add(c);
          existingMap[c] = (2 + i);
        }
      }
    }
  }

  const specialPatterns = (CONFIG.patterns.specialPartners || []).map(s => String(s || '').toLowerCase());
  const minYear = getMinSpecialCaseSubmissionYear_();

  // Metrics
  let candidates = 0;
  let skippedAssoc = 0;
  let skippedExisting = 0;
  let skippedDoneStatus = 0;
  let skippedYearGate = 0;
  let skippedMissingSubmissionDate = 0;

  let matchedFlex = 0;
  let matchedOver12 = 0;
  let matchedMQ = 0;
  let matchedQL01 = 0; // First-Month Policy (day-based)

  // Controlled columns (UPSERT should not wipe other manual columns)
  const controlledNames = [
    'Submission Date','Claim Number','DB','DB Link','Partner Name','Insurance','Device Type','Last Status',
    'Service Center',
    // Column rename support
    'Last Status Aging','LSA',
    'Activity Log Aging','ALA',
    'TAT','Q-L (Months)','Product',
    // Special Case renamed monetary columns
    'Sum Insured Amount','Sum Insured',
    'Claim Amount','Repair/Replace Amount','Nett Claim Amount',
    'OR','Claim Own Risk Amount','OR Amount',
    'Selisih',
    'Start Date','End Date','Details','Reason'
  ];
  const controlledIdx = [];
  for (let k = 0; k < controlledNames.length; k++) {
    const j = idxH[controlledNames[k]];
    if (j != null) controlledIdx.push(j);
  }

  const rowsOut = [];
  const claimColors = [];
  const claimNotes = [];
  const dbUrls = [];

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];

    // Submission date source of truth: claim_submitted_datetime.
    const subRaw = (idxClaimSubmittedDt != null) ? row[idxClaimSubmittedDt] : null;
    const subDate = coerceDateOnly_(subRaw);

    // Year gate (optional): only apply when Submission Date tersedia.
    // Kalau Submission Date kosong, jangan drop data (biar Special Case tidak "kosong").
    if (!subDate) { skippedMissingSubmissionDate++; }
    if (subDate && minYear > 0 && subDate.getFullYear && subDate.getFullYear() < minYear) { skippedYearGate++; continue; }

    const claimRaw = String((idxClaim != null) ? row[idxClaim] : '' || '').trim();
    const claimKey = claimRaw.toUpperCase();
    if (!claimKey) continue;

    const lastStatus = String((idxLastStatus != null) ? row[idxLastStatus] : '' || '').trim();
    const skipDone = (SPECIAL_RULES) ? !!SPECIAL_RULES.SKIP_EXCLUDED_LAST_STATUSES : true;
    if (skipDone && excluded && excluded.has && excluded.has(lastStatus)) { skippedDoneStatus++; continue; }

    if (SPECIAL_FLAGS.MODE === 'APPEND_NEW_ONLY' && existingClaims.has(claimKey)) { skippedExisting++; continue; }

    const partnerRaw = String((idxBP != null) ? row[idxBP] : '' || '').trim();
    const partner = partnerRaw.toLowerCase();
    const productRaw = String((idxProduct != null) ? row[idxProduct] : '' || '').trim();
    const product = productRaw.toLowerCase();

    // Strict Second-Year (Market Value) detection: month_policy_aging (Raw Data) > 12
    const monthPolicyAging = (idxMonthPolicyAging != null) ? normalizeInt_(row[idxMonthPolicyAging]) : null;

    // Rules (combined; reason can be multiple)
    let isFlex = false;
    if (SPECIAL_RULES && SPECIAL_RULES.ENABLE_FLEX_RULE) {
      isFlex =
        specialPatterns.some(p => p && partner.indexOf(p) > -1) ||
        claimKey.indexOf('SFX') > -1 ||
        product.indexOf('flex') > -1;
    }
    let over12 = false;
    if (SPECIAL_RULES && SPECIAL_RULES.ENABLE_Q_L_OVER_12_RULE) {
      // STRICT per requirement: only month_policy_aging is allowed for Second-Year (Market Value)
      over12 = (monthPolicyAging != null && monthPolicyAging > 12);
    }

    let mqFlag = false;
    if (SPECIAL_RULES && SPECIAL_RULES.ENABLE_POLICY_REMAINING_LE_1_RULE) {
      // New rule: Policy Remaining ≤ 1 Month (day-based)
      // policy_end_date - claim_submission_date < 30 days (and non-negative)
      const pe = (idxPolicyEnd != null) ? coerceDateOnly_(row[idxPolicyEnd]) : null;
      if (pe && subDate) {
        const daysBefore = Math.floor((pe.getTime() - subDate.getTime()) / (24 * 60 * 60 * 1000));
        mqFlag = (daysBefore >= 0 && daysBefore < 30);
      }
    }

    let qlLe1 = false;
    if (!SPECIAL_RULES || SPECIAL_RULES.ENABLE_Q_L_0_1_RULE) {
      // New rule: First-Month Policy (day-based)
      // claim_submission_date - policy_start_date < 30 days (and non-negative)
      const ps = (idxPolicyStart != null) ? coerceDateOnly_(row[idxPolicyStart]) : null;
      if (ps && subDate) {
        const daysAfter = Math.floor((subDate.getTime() - ps.getTime()) / (24 * 60 * 60 * 1000));
        qlLe1 = (daysAfter >= 0 && daysAfter < 30);
      }
    }

    if (!isFlex && !over12 && !mqFlag && !qlLe1) continue;
    candidates++;

    const out = new Array(header.length).fill('');
    const set = (k, v) => { const j = idxH[k]; if (j != null) out[j] = v; };

    // Common fields
    set('Submission Date', subDate ? subDate : '');
    set('Claim Number', claimRaw);

    const src = String((idxSource != null) ? row[idxSource] : '' || '').toUpperCase();
    const dbFromClaim = computeDbValueFromClaimNumber05c_(claimKey);
    set('DB', dbFromClaim || ((src.indexOf('OLD') > -1) ? 'OLD' : (src.indexOf('NEW') > -1 ? 'NEW' : (src || ''))));

    const dbUrl = (idxDashboard != null) ? row[idxDashboard] : '';
    set('DB Link', dbUrl ? 'LINK' : '');

    set('Partner Name', partnerRaw);
    set('Insurance', mapInsuranceShort_((idxInsurancePartner != null) ? row[idxInsurancePartner] : ''));
    set('Device Type', (idxDevice != null) ? row[idxDevice] : '');

    set('Service Center', (idxSvc != null) ? row[idxSvc] : '');

    // Optional columns (if present in sheet): Last Status, Activity Log, Last Status Date, Status Type
    set('Last Status', lastStatus);
    if (idxH['Activity Log'] != null) set('Activity Log', __pickActivityLogValue05c_(row, headerIndexRaw));
    if (idxH['Activity Log Datetime'] != null) set('Activity Log Datetime', __pickActivityLogDatetimeValue05c_(row, headerIndexRaw));
    if (idxH['Last Status Date'] != null) {
      const d0 = __parseClaimLastUpdatedDatetime05c_((idxClaimLastUpdated != null) ? row[idxClaimLastUpdated] : null);
      set('Last Status Date', d0 ? d0 : '');
    }
    if (idxH['Status Type'] != null) set('Status Type', __getStatusType05c_(lastStatus));

    const lsa = (idxLSA != null) ? normalizeInt_(row[idxLSA]) : null;
    const ala = (idxALA != null) ? normalizeInt_(row[idxALA]) : null;
    const tat = (idxTAT != null) ? normalizeInt_(row[idxTAT]) : null;
    set('Last Status Aging', (lsa != null) ? lsa : '');
    set('LSA', (lsa != null) ? lsa : '');
    set('Activity Log Aging', (ala != null) ? ala : '');
    set('ALA', (ala != null) ? ala : '');
    set('TAT', (tat != null) ? tat : '');

    set('Q-L (Months)', (idxQL != null) ? (normalizeInt_(row[idxQL]) ?? '') : '');
    set('Product', productRaw);

    const sumIns = (idxSumInsuredRaw != null) ? normalizeNumber_(row[idxSumInsuredRaw]) : null;
    const orAmt = (headerIndexRaw[h.ownRiskAmount] != null) ? normalizeNumber_(row[headerIndexRaw[h.ownRiskAmount]]) : null;
    const nett = (headerIndexRaw[h.nettClaimAmount] != null) ? normalizeNumber_(row[headerIndexRaw[h.nettClaimAmount]]) : null;

    set('Sum Insured Amount', (sumIns != null) ? sumIns : '');
    set('Sum Insured', (sumIns != null) ? sumIns : '');

    // Special Case: OR checkbox (if present) + OR Amount (money)
    if (idxH['OR'] != null) set('OR', getOrCheckboxForOutput_(row, headerIndexRaw));
    if (idxH['Claim Own Risk Amount'] != null) set('Claim Own Risk Amount', (orAmt != null) ? orAmt : '');
    if (idxH['OR Amount'] != null) set('OR Amount', (orAmt != null) ? orAmt : '');

    const claimAmt = (nett != null) ? nett : null;
    const claimAmtOut = (claimAmt != null) ? claimAmt : '';
    set('Claim Amount', claimAmtOut);
    set('Repair/Replace Amount', claimAmtOut);
    set('Nett Claim Amount', claimAmtOut);

    const selisih = (sumIns != null && claimAmt != null) ? (sumIns - claimAmt) : '';
    set('Selisih', selisih);

    // Reason labels (must be multi)
    const reasons = [];
    if (isFlex) { reasons.push('Flex'); matchedFlex++; }
    if (over12) { reasons.push('Second-Year (Market Value)'); matchedOver12++; }
    if (qlLe1) { reasons.push('First-Month Policy'); matchedQL01++; }
    if (mqFlag) { reasons.push('Policy Remaining ≤ 1 Month'); matchedMQ++; }
    set('Reason', reasons.join(' | '));
    // Start/End Date columns (Raw Data policy dates)
    const ps0 = (idxPolicyStart != null) ? coerceDateOnly_(row[idxPolicyStart]) : null;
    const pe0 = (idxPolicyEnd != null) ? coerceDateOnly_(row[idxPolicyEnd]) : null;
    if (ps0) set('Start Date', ps0);
    if (pe0) set('End Date', pe0);

    // Details (only for First-Month / Policy Remaining <= 1 Month / Second-Year)
    try {
      const fmt2y = (typeof formatShortDate2y_ === 'function')
        ? formatShortDate2y_
        : function(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'd MMM yy'); };

      const parts = [];

      if (!subDate) parts.push('Submission Date: (missing)');
      const dayMs = 24 * 60 * 60 * 1000;

      if (qlLe1 && ps0 && subDate) {
        const daysAfter = Math.round((subDate.getTime() - ps0.getTime()) / dayMs);
        parts.push('Start Date Policy : ' + fmt2y(ps0) + '\n' +
                   'The claim was submitted ' + daysAfter + ' days after the policy started.');
      }

      if (mqFlag && pe0 && subDate) {
        const daysBefore = Math.round((pe0.getTime() - subDate.getTime()) / dayMs);
        parts.push('End Date Policy : ' + fmt2y(pe0) + '\n' +
                   'The claim was submitted ' + daysBefore + ' days before the policy ended.');
      }
      if (over12) {
        const mpaText = (monthPolicyAging != null) ? String(monthPolicyAging) : '(missing)';
        const lines = [];
        lines.push('month_policy_aging : ' + mpaText + ' months');
        if (ps0) lines.push('Start Date Policy : ' + fmt2y(ps0));
        if (pe0) lines.push('End Date Policy : ' + fmt2y(pe0));
        parts.push(lines.join('\n'));
      }

      if (parts.length) set('Details', parts.join('\n\n'));
    } catch (e) {}

    // Color/Note in Special Case:
// Requirement: "First-Month Policy", "Second-Year (Market Value)", and "Policy Remaining" note+highlight
// must be applied in Operational Sheets (not in Special Case). Keep Flex highlighting for convenience.
let color = null;
if (isFlex) color = '#f4c7c3';

// Note (for Claim Number cell) — keep Flex only; other policy rules are handled in operational sheets.
try {
  const parts2 = [];
  if (isFlex) parts2.push('Flex');
  claimNotes.push(parts2.length ? ('Special Case: ' + parts2.join(' | ')) : '');
} catch (e) { claimNotes.push(''); }
rowsOut.push(out);
    dbUrls.push(dbUrl);
    claimColors.push(color);
  }

  if (!rowsOut.length) {
    const metrics0 = {
      mode: SPECIAL_FLAGS.MODE,
      candidates: candidates,
      written: 0,
      skippedAssoc: skippedAssoc,
      skippedExisting: skippedExisting,
      skippedDoneStatus: skippedDoneStatus,
      minYear: minYear,
      skippedYearGate: (minYear > 0 ? skippedYearGate : 0),
      skippedMissingSubmissionDate: skippedMissingSubmissionDate,
      matched: {
        flex: matchedFlex,
        qlOver12: matchedOver12,
        qlLe1: matchedQL01,
        mqLe1: matchedMQ
      }
    };
    return { count: 0, metrics: metrics0 };
  }

  const metrics = {
    deletedExcluded: 0,
    mode: SPECIAL_FLAGS.MODE,
    candidates: candidates,
    written: rowsOut.length,
    skippedAssoc: skippedAssoc,
    skippedExisting: skippedExisting,
    skippedDoneStatus: skippedDoneStatus,
    minYear: minYear,
    skippedYearGate: (minYear > 0 ? skippedYearGate : 0),
      skippedMissingSubmissionDate: skippedMissingSubmissionDate,
    matched: {
      flex: matchedFlex,
      qlOver12: matchedOver12,
      qlLe1: matchedQL01,
      mqLe1: matchedMQ
    }
  };

  // UPSERT: update existing rows by claim number, but only for controlled columns
  if (SPECIAL_FLAGS.MODE === 'UPSERT' && claimCol0 > -1) {
    const rowMap = {};
    const urlMap = {};
    const colorMap = {};
    const noteMap = {};
    const updatedRowNums = [];
    const appendRows = [];
    const appendUrls = [];
    const appendColors = [];
    const appendNotes = [];

    for (let r = 0; r < rowsOut.length; r++) {
      const out = rowsOut[r];
      const claimKey = String(out[claimCol0] || '').trim().toUpperCase();
      if (!claimKey) continue;

      const existingRowNum = existingMap[claimKey];
      if (existingRowNum != null) {
        const idx0 = existingRowNum - 2;
        const base = existingValsAll[idx0] ? existingValsAll[idx0].slice() : new Array(header.length).fill('');
        for (let ci = 0; ci < controlledIdx.length; ci++) base[controlledIdx[ci]] = out[controlledIdx[ci]];
        rowMap[existingRowNum] = base;
        urlMap[existingRowNum] = dbUrls[r];
        colorMap[existingRowNum] = claimColors[r];
        noteMap[existingRowNum] = claimNotes[r];
        updatedRowNums.push(existingRowNum);
      } else {
        appendRows.push(out);
        appendUrls.push(dbUrls[r]);
        appendColors.push(claimColors[r]);
        appendNotes.push(claimNotes[r]);
      }
    }

    if (updatedRowNums.length) {
      __writeRowSegments_(sh, updatedRowNums, rowMap, header.length);
      if (dbLinkCol0 > -1) __setDbLinkRichTextSegments_(sh, dbLinkCol0, updatedRowNums, urlMap);
      if (SPECIAL_FLAGS.COLORIZE_CLAIM_CELL && claimCol0 > -1) __setBgSegments_(sh, claimCol0, updatedRowNums, colorMap);
      if (SPECIAL_FLAGS.COLORIZE_CLAIM_CELL && claimCol0 > -1) __setNotesSegments_(sh, claimCol0, updatedRowNums, noteMap);

      // Apply formats only to updated blocks (cheap)
      const segs = __groupConsecutive_(updatedRowNums);
      for (let s = 0; s < segs.length; s++) {
        const seg = segs[s];
        applyOperationalColumnSchema_(sh, header, seg[0], seg.length, { orIsMoney: false });
        applyDbLinkFormatting_(sh, header, seg.length, seg[0]);
      }
    }

    if (appendRows.length) {
      const startRow = Math.max(sh.getLastRow() + 1, 2);
      safeSetValues_(sh.getRange(startRow, 1, appendRows.length, header.length), appendRows);
      if (dbLinkCol0 > -1) __setDbLinkRichTextRange_(sh, dbLinkCol0, startRow, appendUrls);
      if (SPECIAL_FLAGS.COLORIZE_CLAIM_CELL && claimCol0 > -1 && !__isDryRun05c__()) {
        sh.getRange(startRow, claimCol0 + 1, appendRows.length, 1)
          .setBackgrounds(appendColors.map(c => [c || '']));
        try { sh.getRange(startRow, claimCol0 + 1, appendRows.length, 1).setNotes(appendNotes.map(n => [n || ''])); } catch (e) {}
      }
      applyOperationalColumnSchema_(sh, header, startRow, appendRows.length, { orIsMoney: false });
      applyDbLinkFormatting_(sh, header, appendRows.length, startRow);
    }

    metrics.upsert = { updated: updatedRowNums.length, appended: appendRows.length };
    metrics.deletedExcluded = deletedExcluded;
    return { count: rowsOut.length, metrics: metrics };
  }

  // REBUILD / APPEND_NEW_ONLY: write block
  const startRow = (SPECIAL_FLAGS.MODE === 'REBUILD') ? 2 : Math.max(sh.getLastRow() + 1, 2);
  safeSetValues_(sh.getRange(startRow, 1, rowsOut.length, header.length), rowsOut);

  if (dbLinkCol0 > -1) __setDbLinkRichTextRange_(sh, dbLinkCol0, startRow, dbUrls);

  if (SPECIAL_FLAGS.COLORIZE_CLAIM_CELL && !__isDryRun05c__() && claimCol0 > -1) {
    sh.getRange(startRow, claimCol0 + 1, rowsOut.length, 1)
      .setBackgrounds(claimColors.map(c => [c || '']));
    try { sh.getRange(startRow, claimCol0 + 1, rowsOut.length, 1).setNotes(claimNotes.map(n => [n || ''])); } catch (e) {}
  }

  applyOperationalColumnSchema_(sh, header, startRow, rowsOut.length, { orIsMoney: false });
  applyDbLinkFormatting_(sh, header, rowsOut.length, startRow);
  metrics.deletedExcluded = (typeof deletedExcluded !== 'undefined') ? deletedExcluded : 0;
  return { count: rowsOut.length, metrics };
}

/** Optional: EV-Bike — enabled by default. Set RUNTIME.enableEvBike = false to disable. */
function processEVBike_(ss, rawValues, headerIndexRaw, pic) { // `pic` kept for backward compatibility
  // Default ON. Only skip when explicitly disabled.
  if (typeof RUNTIME !== 'undefined' && RUNTIME && RUNTIME.enableEvBike === false) return 0;

  const h = CONFIG.headers;
  const sh = ss.getSheetByName('EV-Bike');
  if (!sh) return 0;

  const idxClaim = headerIndexRaw[h.claimNumber];
  if (idxClaim == null) return 0;


  const OPTIONAL_FLAGS = (typeof __OPTIONAL_FLAGS !== 'undefined' && __OPTIONAL_FLAGS) ? __OPTIONAL_FLAGS : {};
  const EXCLUDED_LAST_STATUSES = getSpecialCaseExcludedStatuses_();

  const idxBP = headerIndexRaw[h.businessPartner];
  const idxOwner = headerIndexRaw['customer_name'];
  const idxPolicyNum = headerIndexRaw['qoala_policy_number'];
  const idxSumInsured = (headerIndexRaw[h.sumInsured] != null) ? headerIndexRaw[h.sumInsured] : headerIndexRaw['sum_insured_amount'];
  const excludedPolicySet = (typeof getEvBikeExcludedPolicyNumberSet_ === 'function') ? getEvBikeExcludedPolicyNumberSet_() : new Set(['GODA-20250729-4SHFZ','GODA-20250729-X97UC','GODA-20250729-KESB4']);

  const idxLastStatus = headerIndexRaw[h.lastStatus];
  const idxAssociate = headerIndexRaw[h.associate];
  const idxDashboard = headerIndexRaw[h.dashboardLink];
  const idxClaimLastUpdated = (headerIndexRaw['claim_last_updated_datetime'] != null) ? headerIndexRaw['claim_last_updated_datetime'] : null;
  const idxIns = headerIndexRaw[h.insuranceCode];
  const idxInsPartner = (h.insurancePartnerName && headerIndexRaw[h.insurancePartnerName] != null)
    ? headerIndexRaw[h.insurancePartnerName]
    : headerIndexRaw['insurance_partner_name'];

  let header = __getHeaderRow05c_(sh);
  let idxH = buildHeaderIndex_(header);

  // Remove deprecated EV-Bike columns when still present.
  try {
    const dropSet = new Set(['Start Date', 'End Date', 'Details'].map(__normalizeHeaderText05c_));
    const toDelete = [];
    for (let i = 0; i < header.length; i++) {
      if (dropSet.has(__normalizeHeaderText05c_(header[i]))) toDelete.push(i + 1);
    }
    for (let i = toDelete.length - 1; i >= 0; i--) {
      sh.deleteColumn(toDelete[i]);
    }
    if (toDelete.length) {
      header = __getHeaderRow05c_(sh);
      idxH = buildHeaderIndex_(header);
    }
  } catch (eDrop) {}

  // Ensure mandatory Status Type only when this sheet schema includes Last Status.
  if (idxH['Last Status'] != null) {
    header = __ensureAppendColumnIfMissing05c_(sh, header, 'Status Type');
    idxH = buildHeaderIndex_(header);
  }
  // Keep schema flexible: do not force legacy EV-Bike columns (Start Date / End Date / Details).

  const patterns = (CONFIG.patterns.evBikePartners || []).map(s => String(s || '').toLowerCase());
  const computeTatFromSubmission_ = (v) => {
    const d = coerceDate_(v);
    if (!d) return '';
    const now = new Date();
    const start = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    const end = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    const diff = Math.floor((end - start) / 86400000);
    return diff >= 0 ? diff : '';
  };
  // Additional source: Submission sheet (EV-Bike claims should be included even if not present in Raw Data yet).
  const submissionCandidates = {};
  try {
    const subSh = ss.getSheetByName('Submission');
    if (subSh && subSh.getLastRow() > 1 && subSh.getLastColumn() > 1) {
      const subHeader = __getHeaderRow05c_(subSh);
      const subIdx = buildHeaderIndex_(subHeader);
      const subVals = subSh.getRange(2, 1, subSh.getLastRow() - 1, subHeader.length).getValues();

      const sClaim = subIdx['Claim Number'];
      const sPartner = (subIdx['Partner Name'] != null) ? subIdx['Partner Name'] : subIdx['Partner'];
      const sPolicy = (subIdx['Policy Number'] != null) ? subIdx['Policy Number'] : subIdx['Policy Num'];
      const sOwner = (subIdx['Owner Name'] != null) ? subIdx['Owner Name'] : subIdx['Customer Name'];
	      const sInsurance = subIdx['Insurance'];
	      const sSum = subIdx['Sum Insured'];
	      const sSubDate = subIdx['Submission Date'];
	      const sDbLink = subIdx['DB Link'];
          const sLastStatus = subIdx['Last Status'];

      if (sClaim != null) {
        for (let i = 0; i < subVals.length; i++) {
          const r = subVals[i];
          const c = String(r[sClaim] || '').trim();
          if (!c) continue;

          const partner = String((sPartner != null) ? r[sPartner] : '').trim().toLowerCase();
          if (patterns.length && !patterns.some(p => p && partner.indexOf(p) > -1)) continue;

          const pol = String((sPolicy != null) ? r[sPolicy] : '').trim();
          if (pol && excludedPolicySet.has(pol)) continue;

          const key = c.toUpperCase();
	          submissionCandidates[key] = {
	            submissionDate: (sSubDate != null) ? r[sSubDate] : '',
	            ownerName: (sOwner != null) ? r[sOwner] : '',
	            policyNumber: pol,
	            partnerName: (sPartner != null) ? r[sPartner] : '',
	            insurance: (sInsurance != null) ? r[sInsurance] : '',
	            sumInsured: (sSum != null) ? r[sSum] : '',
	            dbUrl: (sDbLink != null) ? r[sDbLink] : '',
              lastStatus: (sLastStatus != null) ? r[sLastStatus] : ''
	          };
        }
      }
    }
  } catch (e) {}


  const claimColIdx = idxH['Claim Number'];
  const dbLinkCol0 = idxH['DB Link'];
  const existing = (sh.getLastRow() > 1) ? sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getValues() : [];
  const existingMap = {};

  if (claimColIdx != null) {
    for (let i = 0; i < existing.length; i++) {
      const c = String(existing[i][claimColIdx] || '').trim();
      if (c) existingMap[c.toUpperCase()] = i;
    }
  }

  const values = existing.slice();
  const touchedRowNums = [];
  const urlMap = {};
  let evRawMatchedCount = 0;
  let evSubmissionOverlayCount = 0;
  let evSkippedExcludedCount = 0;

  const ensureRow = i => (values[i] ? values[i] : (values[i] = new Array(header.length).fill('')));

  // Avoid duplicate processing when Submission has repeated Claim Number rows.
  const seenClaims = new Set();

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];

    const partner = String((idxBP != null) ? row[idxBP] : '' || '').toLowerCase();
    if (!patterns.some(p => p && partner.indexOf(p) > -1)) continue;

    const claimUp = String(row[idxClaim] || '').toUpperCase();

    if (!claimUp) continue;
    if (seenClaims.has(claimUp)) continue;
    seenClaims.add(claimUp);
    const lastStatus = String((idxLastStatus != null) ? row[idxLastStatus] : '' || '').trim();
    const lastStatusKey = lastStatus.toUpperCase();

    if (OPTIONAL_FLAGS.EVBIKE_SKIP_EXCLUDED_LAST_STATUSES && EXCLUDED_LAST_STATUSES.has(lastStatusKey)) {
      evSkippedExcludedCount++;
      continue;
    }

    let pos = existingMap[claimUp];
    if (pos == null) { pos = values.length; existingMap[claimUp] = pos; }
    const out = ensureRow(pos);

    const set = (k, v) => { const j = idxH[k]; if (j != null) out[j] = v; };

    set('Submission Date', buildSubmissionDateCell_(row, headerIndexRaw));
    set('Claim Number', row[idxClaim]);

    const dbUrl = (idxDashboard != null) ? row[idxDashboard] : '';
    set('DB Link', dbUrl ? 'LINK' : '');
    if (pos != null) {
      const rn = 2 + pos;
      touchedRowNums.push(rn);
      urlMap[rn] = dbUrl;
    }

    set('Partner Name', (idxBP != null) ? row[idxBP] : '');
    set('Insurance', mapInsuranceShort_((idxInsPartner != null) ? row[idxInsPartner] : ''));
    set('Owner Name', (idxOwner != null) ? row[idxOwner] : '');
    const polNum0 = String((idxPolicyNum != null) ? row[idxPolicyNum] : '').trim();
    if (polNum0 && excludedPolicySet.has(polNum0)) continue;
    set('Policy Number', polNum0);
	    // Sum Insured
	    set('Sum Insured', (idxSumInsured != null) ? row[idxSumInsured] : '');
        if (idxH['TAT'] != null) {
          const tatRaw = normalizeInt_((headerIndexRaw[h.daysAgingFromSubmission] != null) ? row[headerIndexRaw[h.daysAgingFromSubmission]] : '');
          set('TAT', (tatRaw != null) ? tatRaw : computeTatFromSubmission_(buildSubmissionDateCell_(row, headerIndexRaw)));
        }

    // Optional columns if present in EV-Bike sheet schema
    set('Last Status', lastStatus);
    if (idxH['Status Type'] != null) set('Status Type', __getStatusType05c_(lastStatus));
    evRawMatchedCount++;
  }


  // Overlay / include Submission-sourced EV-Bike claims.
  try {
    const managed = (typeof EVBIKE_POLICY !== 'undefined' && EVBIKE_POLICY && Array.isArray(EVBIKE_POLICY.MANAGED_HEADERS))
      ? EVBIKE_POLICY.MANAGED_HEADERS
      : ['Submission Date','Owner Name','Policy Number','Partner Name','Insurance','Sum Insured','DB Link','Last Status','TAT'];

    for (const key in submissionCandidates) {
      const info = submissionCandidates[key];
      const claim = key;
      if (!claim) continue;
      if (seenClaims.has(claim)) continue;
      seenClaims.add(claim);

      // Ensure in-sheet row exists (append if missing)
      let pos = (claimColIdx != null) ? existingMap[claim] : null;
      let out = null;
      if (pos != null) {
        out = values[pos];
      } else {
        out = new Array(header.length).fill('');
        const newPos = values.length;
        values.push(out);
        existingMap[claim] = newPos;
        touchedRowNums.push(2 + newPos);
      }

      function setH(name, val) {
        const j = idxH[name];
        if (j != null) out[j] = val;
      }

      // Always set Claim Number if possible (safe)
      setH('Claim Number', claim);

      // Only overwrite managed columns (Status is intentionally untouched).
	      setH('Submission Date', info.submissionDate || '');
	      setH('Owner Name', info.ownerName || '');
	      setH('Policy Number', info.policyNumber || '');
	      setH('Partner Name', info.partnerName || '');
	      setH('Insurance', info.insurance || '');
	      setH('Sum Insured', info.sumInsured || '');
          setH('Last Status', info.lastStatus || '');
          if (idxH['Status Type'] != null) setH('Status Type', __getStatusType05c_(info.lastStatus || ''));
          if (idxH['TAT'] != null) setH('TAT', computeTatFromSubmission_(info.submissionDate));

      // DB Link handling: expect URL; display text stays 'LINK'
      const url = String(info.dbUrl || '').trim();
      if (url) {
        setH('DB Link', 'LINK');
        const rn = 2 + existingMap[claim];
        urlMap[rn] = url;
        if (dbLinkCol0 != null) touchedRowNums.push(rn);
      }
      evSubmissionOverlayCount++;
    }
  } catch (e) {}

  if (!values.length) return 0;
  // EV-Bike Last Status is user-managed free text (no enforced dropdown).
  // Clear DV BEFORE write to prevent setValues rejection:
  // "violates data validation rules ... Please enter one of ..."
  try {
    const idxLastStatusEv = idxH['Last Status'];
    if (idxLastStatusEv != null && !__isDryRun05c__()) {
      const rowsToClear = Math.max(values.length, (sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0));
      if (rowsToClear > 0) sh.getRange(2, idxLastStatusEv + 1, rowsToClear, 1).clearDataValidations();
    }
  } catch (eDvEv) {}

  // Never overwrite manual Status dropdown column on EV-Bike.
  // Write left/right segments around "Status" when column exists.
  const idxStatus = (idxH['Status'] != null) ? idxH['Status'] : -1;
  if (idxStatus === -1) {
    safeSetValues_(sh.getRange(2, 1, values.length, header.length), values);
  } else {
    if (idxStatus > 0) {
      const left = values.map(r => r.slice(0, idxStatus));
      safeSetValues_(sh.getRange(2, 1, values.length, idxStatus), left);
    }
    if (idxStatus < header.length - 1) {
      const right = values.map(r => r.slice(idxStatus + 1));
      safeSetValues_(sh.getRange(2, idxStatus + 2, values.length, header.length - idxStatus - 1), right);
    }
  }

  // Apply RichText hyperlink only for touched rows
  if (dbLinkCol0 != null && touchedRowNums.length) __setDbLinkRichTextSegments_(sh, dbLinkCol0, touchedRowNums, urlMap);
  try {
    if (typeof logLine_ === 'function') {
      logLine_(
        'INFO',
        'EVBIKE_METRICS',
        'rows=' + values.length
          + ' raw=' + evRawMatchedCount
          + ' submission_overlay=' + evSubmissionOverlayCount
          + ' skip_excluded=' + evSkippedExcludedCount,
        '',
        'INFO'
      );
    }
  } catch (eM) {}
  return values.length;
}
