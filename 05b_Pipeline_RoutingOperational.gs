/***************************************
 * 05b_Pipeline_RoutingOperational.gs
 * Split from: 05_Pipeline_Routing_Optional.gs
 * Scope:
 *  - Routing map compilation + sheet writer builders
 *  - Clear operational sheets
 *  - Route Raw -> operational sheets (in-memory writers + flush)
 *  - Sorting operational + optional sheets
 ***************************************/
'use strict';

/**
 * Returns TRUE when DRY_RUN is enabled.
 * Supports both global DRY_RUN constant and helper isDryRun_() (load-order safe).
 */
function __isDryRun05b__() {
  try {
    if (typeof isDryRun_ === 'function') return !!isDryRun_();
  } catch (e) {}
  try {
    return (typeof DRY_RUN !== 'undefined') ? !!DRY_RUN : false;
  } catch (e2) {}
  return false;
}


/**
 * Runtime flow resolver (load-order safe).
 * Expected values: main | sub | form
 */
function __getRuntimeFlow05b_() {
  try {
    if (typeof RUNTIME === 'object' && RUNTIME) {
      const v = RUNTIME.flowName || RUNTIME.flow || RUNTIME.FLOW;
      if (v) return String(v).trim().toLowerCase();
    }
  } catch (e) {}
  return 'main';
}

/**
 * Ensure a header exists by appending it at the right-most end (first row).
 * - If there is an empty header cell after the last non-empty header, it will be used.
 * - Otherwise inserts a new column at the end.
 */
function __ensureAppendHeader05b_(sh, headerName) {
  try {
    if (!sh) return;
    const name = String(headerName || '').trim();
    if (!name) return;

    const lastCol = Math.max(sh.getLastColumn() || 1, 1);
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    if (header.indexOf(name) > -1) return;

    // Find last non-empty header cell.
    let lastNonEmpty = -1;
    for (let i = header.length - 1; i >= 0; i--) {
      if (header[i]) { lastNonEmpty = i; break; }
    }
    const targetCol = lastNonEmpty + 2; // 1-based

    if (__isDryRun05b__()) return;

    if (targetCol <= lastCol) {
      sh.getRange(1, targetCol).setValue(name)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    } else {
      sh.insertColumnAfter(lastCol);
      sh.getRange(1, lastCol + 1).setValue(name)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
  } catch (e) {
    // best-effort only (do not fail routing)
  }
}

/** Status Type resolver (prefers 06c mapping when available). */
function __getStatusType05b_(lastStatus) {
  const s = String(lastStatus || '').trim();
  if (!s) return '';
  try {
    if (typeof getStatusType06c_ === 'function') return String(getStatusType06c_(s) || '');
  } catch (e) {}

  try {
    const map = (CONFIG && (CONFIG.STATUS_TYPE_MAP || CONFIG.statusTypeMap || CONFIG.statusTypeByLastStatus)) || null;
    if (map && map[s] != null) return String(map[s] || '');
  } catch (e2) {}

  return '';
}

/** Pick Activity Log value from raw row based on flow. */
function __pickActivityLogValue05b_(rawRow, headerIndexRaw, flowName) {
  const flow = String(flowName || '').trim().toLowerCase();
  const preferSub = (flow === 'sub');
  const keys = preferSub
    ? ['activity_log', 'last_activity_log']
    : ['last_activity_log', 'activity_log'];
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const ix = headerIndexRaw[k];
    if (ix == null) continue;
    const v = rawRow[ix];
    if (v === '' || v == null) continue;
    return v;
  }
  return '';
}

/** Pick Activity Log datetime from raw row based on flow (best-effort). */
function __pickActivityLogDatetimeValue05b_(rawRow, headerIndexRaw, flowName) {
  const flow = String(flowName || '').trim().toLowerCase();
  const preferSub = (flow === 'sub');
  const keys = preferSub
    ? ['activity_log_datetime', 'activity_log_date_time', 'activity_log_updated_datetime', 'last_activity_log_datetime', 'last_activity_log_date_time']
    : ['last_activity_log_datetime', 'last_activity_log_date_time', 'activity_log_datetime', 'activity_log_date_time', 'activity_log_updated_datetime'];
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const ix = headerIndexRaw[k];
    if (ix == null) continue;
    const v = rawRow[ix];
    if (v === '' || v == null) continue;
    return v;
  }
  return '';
}


/**
 * Finish-status detector (SUB + SC routing patch).
 * NOTE: status may contain trailing spaces (e.g., DONE_REPAIR ).
 */
function isFinishStatus05a_(status) {
  const s = String(status || '').trim();
  if (!s) return false;

  // Prefer shared constant from 00.gs
  try {
    if (typeof FINISH_STATUSES !== 'undefined' && Array.isArray(FINISH_STATUSES)) {
      return FINISH_STATUSES.indexOf(s) !== -1;
    }
  } catch (e) {}

  // Fallback (legacy minimal list)
  return (
    s === 'DONE_REPAIR' ||
    s === 'WAITING_WALKIN_FINISH' ||
    s === 'COURIER_PICKED_UP' ||
    s === 'WAITING_COURIER_FINISH' ||
    s === 'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH'
  );
}

function uniq05a_(arr) {
  const out = [];
  const seen = {};
  (arr || []).forEach(v => {
    const k = String(v || '');
    if (!k) return;
    if (seen[k]) return;
    seen[k] = true;
    out.push(v);
  });
  return out;
}



/** =========================
 *  Routing core
 *  ========================= */

function compileRoutingIndex_(routingMap) {
  const idx = {};
  Object.keys(routingMap).forEach(sheetName => {
    (routingMap[sheetName] || []).forEach(status => {
      if (!idx[status]) idx[status] = [];
      idx[status].push(sheetName);
    });
  });
  return idx;
}


/** =========================
 * SC sheet split + Type dropdown helpers (single master workflow)
 * ========================= */

function containsAnyKeyword05b_(text, keywords) {
  const s = String(text || '').toLowerCase();
  if (!s) return false;
  const list = keywords || [];
  for (let i = 0; i < list.length; i++) {
    const k = String(list[i] || '').toLowerCase().trim();
    if (k && s.indexOf(k) > -1) return true;
  }
  return false;
}

function scoreKeywords05b_(text, keywords) {
  const s = String(text || '').toLowerCase();
  if (!s) return 0;
  const list = keywords || [];
  let score = 0;
  for (let i = 0; i < list.length; i++) {
    const k = String(list[i] || '').toLowerCase().trim();
    if (k && s.indexOf(k) > -1) score++;
  }
  return score;
}


function filterScTargets05b_(targets, scNameVal, scFarhan, scMeilani, scIvan, kwFarhan, kwMeilani, kwIvan, scFallback) {
  if (!targets || !targets.length) return targets || [];

  const hasFarhan = targets.indexOf(scFarhan) > -1;
  const hasMeilani = targets.indexOf(scMeilani) > -1;
  const hasIvan = targets.indexOf(scIvan) > -1;

  if (!hasFarhan && !hasMeilani && !hasIvan) return targets;

  const nonSc = targets.filter(x => x !== scFarhan && x !== scMeilani && x !== scIvan);

  const candidates = [];
  if (hasFarhan) candidates.push({ name: scFarhan, score: scoreKeywords05b_(scNameVal, kwFarhan) });
  if (hasIvan) candidates.push({ name: scIvan, score: scoreKeywords05b_(scNameVal, kwIvan) });
  if (hasMeilani) candidates.push({ name: scMeilani, score: scoreKeywords05b_(scNameVal, kwMeilani) });

  // Pick the best match. Tie-break is deterministic by insertion order above.
  let chosen = null;
  let best = 0;
  for (let i = 0; i < candidates.length; i++) {
    const c = candidates[i];
    if (!chosen || c.score > best) { chosen = c.name; best = c.score; }
  }

  // Fail-closed: if no keyword match, route to quarantine sheet to prevent mis-mapping.
  if (!chosen || best <= 0) {
    const fb = String(scFallback || '').trim();
    if (fb) chosen = fb;
    else if (hasIvan) chosen = scIvan;
    else if (hasMeilani) chosen = scMeilani;
    else chosen = scFarhan;
  }

  const out = nonSc.slice();
  if (chosen) out.push(chosen);
  return out;
}


/**
 * Ensure the SC fallback/quarantine sheet exists.
 * - Creates the sheet if missing (best-effort).
 * - Copies header row (values + formats) from a template SC sheet if available.
 */
function ensureScFallbackSheet05b_(ss, fallbackSheetName, templateSheetName) {
  try {
    const fb = String(fallbackSheetName || '').trim();
    if (!fb) return;
    if (!ss) return;
    if (ss.getSheetByName(fb)) return;
    if (__isDryRun05b__()) return;

    const sh = ss.insertSheet(fb);

    const tmplName = String(templateSheetName || '').trim();
    const tmpl = tmplName ? ss.getSheetByName(tmplName) : null;
    if (tmpl) {
      const lc = Math.max(tmpl.getLastColumn() || 1, 1);
      tmpl.getRange(1, 1, 1, lc).copyTo(sh.getRange(1, 1, 1, lc), { contentsOnly: false });
      try { sh.setFrozenRows(Math.max(tmpl.getFrozenRows() || 0, 1)); } catch (e1) { try { sh.setFrozenRows(1); } catch (e2) {} }
    } else {
      sh.getRange(1, 1).setValue('Claim Number');
      try { sh.setFrozenRows(1); } catch (e3) {}
    }

    try {
      const lc2 = Math.max(sh.getLastColumn() || 1, 1);
      sh.getRange(1, 1, 1, lc2).setHorizontalAlignment('center').setVerticalAlignment('middle');
    } catch (e4) {}
  } catch (e) {
    // best-effort only
  }
}



/**
 * Apply formats & alignments for common columns with minimal calls
 * opts:
 *  - orIsMoney: boolean (true => format OR as money; false => format OR as checkbox)
 */
function applyOperationalColumnSchema_(sh, header, startRow, nRows, opts) {
  if (!sh || !header || nRows <= 0) return;

  opts = opts || {};
  const orIsMoney = !!opts.orIsMoney;

  const idx = buildHeaderIndex_(header);
  const fmt = (colName, numberFormat, align) => {
    const c0 = idx[colName];
    if (c0 == null) return;
    const r = sh.getRange(startRow, c0 + 1, nRows, 1);
    if (numberFormat) safeSetNumberFormat_(r, numberFormat);
    if (align && !DRY_RUN) r.setHorizontalAlignment(align);
  };

  // Dates
  fmt('Submission Date', __FORMATS.DATE, 'center');
  // Defensive: Submission Date must never be checkbox-validated.
  try {
    const cSub = idx['Submission Date'];
    if (cSub != null && !DRY_RUN) {
      sh.getRange(startRow, cSub + 1, nRows, 1).clearDataValidations();
    }
  } catch (eSubDv) {}
  fmt('Last Status Date', __FORMATS.DATE, 'center');
  fmt('Last Status Datetime', __FORMATS.DATETIME, 'center');
  fmt('Timestamp', __FORMATS.TIMESTAMP, 'center');

  // Numbers
  // Legacy abbreviations were renamed:
  //  - LSA => Last Status Aging
  //  - ALA => Activity Log Aging
  // Keep both for backward compatibility.
  fmt('LSA', __FORMATS.INT, 'right');
  fmt('Last Status Aging', __FORMATS.INT, 'right');
  fmt('ALA', __FORMATS.INT, 'right');
  fmt('Activity Log Aging', __FORMATS.INT, 'right');
  fmt('TAT', __FORMATS.INT, 'right');
  fmt('Q-L (Months)', __FORMATS.INT, 'right');
  fmt('M-L (Months)', __FORMATS.INT, 'right');
  fmt('M-Q (Months)', __FORMATS.INT, 'right');

  // Money-like
  fmt('Sum Insured', __FORMATS.MONEY0, 'right');
  fmt('Sum Insured Amount', __FORMATS.MONEY0, 'right');
  // Legacy Special Case header (renamed to Claim Amount)
  fmt('Repair/Replace Amount', __FORMATS.MONEY0, 'right');
  fmt('Claim Amount', __FORMATS.MONEY0, 'right');
  fmt('Claim Own Risk Amount', __FORMATS.MONEY0, 'right');
  fmt('Nett Claim Amount', __FORMATS.MONEY0, 'right');
  fmt('OR Amount', __FORMATS.MONEY0, 'right');

  // Percent
  fmt('% Approval', '0%', 'right');

  // OR checkbox vs money (Special Case)
  if (orIsMoney) {
    fmt('OR', __FORMATS.MONEY0, 'right');
  } else {
    // checkbox: no number format; keep center
    fmt('OR', null, 'center');
  }
}

/**
 * Apply RichText hyperlinks to the 'DB Link' column (post-flush).
 * - Writes plain display text in setValues pass, then upgrades to RichText hyperlinks here.
 * - Avoids locale-sensitive HYPERLINK() formula errors (#ERROR!).
 */
function applyDbLinkRichTextFromWriter_(w, startRow) {
  if (!w || !w.sheet || !w.header || !w.rows || !w.rows.length) return;
  startRow = startRow || 2;
  const idx = w.header.indexOf('DB Link');
  if (idx === -1) return;
  const n = w.rows.length;
  const urls = (w.dbLinkUrls || []).slice(0, n).map(u => [u]);
  const rich = makeRichTextHyperlinks2d_(urls, 'LINK');
  if (!DRY_RUN) safeSetRichTextValues_(w.sheet.getRange(startRow, idx + 1, n, 1), rich);
}

/** Build claim highlight sets from the full Raw Data (independent of routing). */
function buildOperationalClaimHighlightSetsFromRaw_(rawValues, headerIndexRaw) {
  const h = CONFIG.headers;
  const ixClaim = headerIndexRaw[h.claimNumber];
  const ixPartner = headerIndexRaw[h.businessPartner];
  const ixProduct = headerIndexRaw[h.productName];
  const ixDaysToEnd = headerIndexRaw[h.daysToEndPolicy];

  const expired = new Set();
  const flex = new Set();
  const b2b = new Set();
  const duplicate = new Map(); // claimKey -> dynamic note

  // New highlight sets
  const secondYear = new Set();
  const firstMonthPolicy = new Set();
  const remaining1Month = new Set();

  if (ixClaim == null) return { expired, flex, b2b, duplicate, secondYear, firstMonthPolicy, remaining1Month };

  const specialPartners = ((CONFIG.patterns && CONFIG.patterns.specialPartners) || [])
    .map(s => String(s || '').trim())
    .filter(Boolean);

  const b2bPartners = ((CONFIG.patterns && CONFIG.patterns.b2bPartners) || [])
    .map(s => String(s || '').trim())
    .filter(Boolean);

  // Inputs for DUPLICATE detection (load-order safe; uses raw header names when needed).
  const ixPolicyNo = headerIndexRaw['qoala_policy_number'];
  const ixSource = (headerIndexRaw[h.sourceSystem] != null) ? headerIndexRaw[h.sourceSystem] : headerIndexRaw['source_system_name'];
  const ixSub = headerIndexRaw[h.claimSubmissionDate];
  const ixLastStatus = headerIndexRaw[h.lastStatus];

  const hasDupInputs = (ixPolicyNo != null && ixSource != null && ixSub != null && ixLastStatus != null);
  const oldByPolicy = new Map();

  const dupPolicy = (typeof DUPLICATE_DETECTION_POLICY === 'object' && DUPLICATE_DETECTION_POLICY) ? DUPLICATE_DETECTION_POLICY : {};
  const tokens = (Array.isArray(dupPolicy.policyTokens) && dupPolicy.policyTokens.length) ? dupPolicy.policyTokens : ['SFP', 'SFX', 'SMR'];
  const newServiceName = String(dupPolicy.newServiceName || 'NEW SERVICE').toUpperCase();
  const oldServiceName = String(dupPolicy.oldServiceName || 'OLD SERVICE').toUpperCase();
  const maxDays = (dupPolicy.maxDaysDiff != null) ? Number(dupPolicy.maxDaysDiff) : 62;

  const formatOldDate_ = (d) => {
    const dd = (typeof coerceDateOnly_ === 'function') ? coerceDateOnly_(d) : (d instanceof Date ? d : null);
    if (!dd) return '';
    if (typeof formatDowShortDate_ === 'function') return formatDowShortDate_(dd);
    try {
      return Utilities.formatDate(dd, Session.getScriptTimeZone(), 'EEE, dd MMM yyyy');
    } catch (e) {
      return String(dd);
    }
  };

  // Helper: find raw index by multiple candidate header names.
  const findRawIndex_ = (candidates) => {
    const a = candidates || [];
    for (let i = 0; i < a.length; i++) {
      const k = a[i];
      if (!k) continue;
      if (headerIndexRaw[k] != null) return headerIndexRaw[k];
    }
    return null;
  };

  const ixQLMonths = findRawIndex_([
    h.qLMonths,
    'Q-L (Months)',
    'q_l_months',
    'ql_months',
    'q_l_(months)'
  ]);

  const ixMonthPolicyAging = findRawIndex_([
    h.monthPolicyAging,
    'month_policy_aging',
    'Month Policy Aging',
    'month policy aging',
    'month_policy_aging_(months)',
    'month_policy_aging_months'
  ]);


  const ixPolicyStart = findRawIndex_([
    h.policyStartDate,
    'Policy Start Date',
    'policy_start_date'
  ]);

  const ixPolicyEnd = findRawIndex_([
    h.policyEndDate,
    'Policy End Date',
    'policy_end_date'
  ]);

  const toDateOnly_ = (v) => {
    return (typeof coerceDateOnly_ === 'function') ? coerceDateOnly_(v) : (v instanceof Date ? v : null);
  };

  const diffDaysLocal_ = (d1, d2) => {
    const a = toDateOnly_(d1);
    const b = toDateOnly_(d2);
    if (!a || !b) return null;
    const t1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    const t2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
    return Math.floor((t2 - t1) / 86400000);
  };

  // Pass 1: expired/flex/b2b + collect OLD SERVICE candidates by policy number + new highlight sets.
  for (let i = 0; i < (rawValues || []).length; i++) {
    const r = rawValues[i] || [];
    const claim = String(r[ixClaim] || '').trim();
    if (!claim) continue;

    const claimKey = claim.toUpperCase();

    const partner = (ixPartner != null) ? String(r[ixPartner] || '').trim() : '';
    const product = (ixProduct != null) ? String(r[ixProduct] || '').trim() : '';

    // IMPORTANT: avoid treating blank/invalid values as 0 (which would falsely mark everything as EXPIRED).
    const daysToEnd = (ixDaysToEnd != null) ? parseOptionalIntStrict_(r[ixDaysToEnd]) : null;
    if (daysToEnd != null && daysToEnd <= 0) expired.add(claimKey);

    const isFlex =
      matchesAnyCi_(partner, specialPartners) ||
      includesCi_(claim, 'SFX') ||
      includesCi_(product, 'flex');
    if (isFlex) flex.add(claimKey);

    const isB2B =
      matchesAnyCi_(partner, b2bPartners) ||
      includesCi_(claim, 'SMR');
    if (isB2B) b2b.add(claimKey);

    // New rules:
    // 1) Second-Year (Market Value) STRICT: month_policy_aging > 12 (Raw Data)
    if (ixMonthPolicyAging != null) {
      const m = parseOptionalIntStrict_(r[ixMonthPolicyAging]);
      if (m != null && m > 12) secondYear.add(claimKey);
    }
    // 2) First-Month Policy if (claim_submission_date - policy_start_date) < 30 days
    if (ixSub != null && ixPolicyStart != null) {
      const dSub = toDateOnly_(r[ixSub]);
      const dStart = toDateOnly_(r[ixPolicyStart]);
      const diff = (typeof diffDays_ === 'function') ? diffDays_(dStart, dSub) : diffDaysLocal_(dStart, dSub);
      if (diff != null && diff >= 0 && diff < 30) firstMonthPolicy.add(claimKey);
    }

    // 3) Policy Remaining <= 1 Month if (policy_end_date - claim_submission_date) < 30 days
    if (ixSub != null && ixPolicyEnd != null) {
      const dSub = toDateOnly_(r[ixSub]);
      const dEnd = toDateOnly_(r[ixPolicyEnd]);
      const diff = (typeof diffDays_ === 'function') ? diffDays_(dSub, dEnd) : diffDaysLocal_(dSub, dEnd);
      if (diff != null && diff >= 0 && diff < 30) remaining1Month.add(claimKey);
    }

    if (hasDupInputs) {
      const policyNo = String(r[ixPolicyNo] || '').trim();
      const src = String(r[ixSource] || '').trim().toUpperCase();
      if (policyNo && src === oldServiceName) {
        const sub = r[ixSub];
        const subDate = toDateOnly_(sub);
        const lastStatus = String(r[ixLastStatus] || '').trim();
        const arr = oldByPolicy.get(policyNo) || [];
        arr.push({ claim: claim, subDate: subDate, lastStatus: lastStatus });
        oldByPolicy.set(policyNo, arr);
      }
    }
  }

  // Pass 2: DUPLICATE notes for NEW SERVICE rows with SFP/SFX/SMR policy numbers.
  if (hasDupInputs && oldByPolicy.size) {
    const tokenLower = tokens.map(t => String(t || '').toLowerCase()).filter(Boolean);

    for (let i = 0; i < (rawValues || []).length; i++) {
      const r = rawValues[i] || [];
      const claim = String(r[ixClaim] || '').trim();
      if (!claim) continue;

      const claimKey = claim.toUpperCase();

      const policyNo = String(r[ixPolicyNo] || '').trim();
      if (!policyNo) continue;

      const src = String(r[ixSource] || '').trim().toUpperCase();
      if (src !== newServiceName) continue;

      const policyLower = policyNo.toLowerCase();
      let hasToken = false;
      for (let k = 0; k < tokenLower.length; k++) {
        if (policyLower.indexOf(tokenLower[k]) > -1) { hasToken = true; break; }
      }
      if (!hasToken) continue;

      const oldList = oldByPolicy.get(policyNo);
      if (!oldList || !oldList.length) continue;

      const newSub = r[ixSub];
      const newSubDate = toDateOnly_(newSub);
      if (!newSubDate) continue;

      let best = null;
      let bestDiff = null;

      for (let j = 0; j < oldList.length; j++) {
        const o = oldList[j];
        if (!o || !o.subDate) continue;
        const diff = (typeof diffDays_ === 'function') ? diffDays_(o.subDate, newSubDate) : diffDaysLocal_(o.subDate, newSubDate);
        if (diff == null) continue;
        if (diff < 0) continue; // only consider OLD before NEW
        if (diff > maxDays) continue;
        if (bestDiff == null || diff < bestDiff) { best = o; bestDiff = diff; }
      }

      if (!best) continue;

      const oldDateStr = formatOldDate_(best.subDate);
      const oldClaim = String(best.claim || '').trim();
      const oldLastStatus = String(best.lastStatus || '').trim();

      if (!oldClaim) continue;

      const note =
        `Duplicate Claim - Refer to Claim Number ${oldClaim}
` +
        `Submitted on ${oldDateStr || 'Unknown date'} - ${oldLastStatus}.`;

      duplicate.set(claimKey, note);
    }
  }

  return { expired, flex, b2b, duplicate, secondYear, firstMonthPolicy, remaining1Month };
}





function getOperationalClaimHighlightPolicy_() {
  const p = (CONFIG && (CONFIG.CLAIM_HIGHLIGHT_POLICY || CONFIG.claimHighlightPolicy))
    || ((typeof CLAIM_HIGHLIGHT_POLICY !== 'undefined' && CLAIM_HIGHLIGHT_POLICY) ? CLAIM_HIGHLIGHT_POLICY : {})
    || {};
  const colors = p.COLORS || p.colors || {};
  const notes = p.NOTES_CANONICAL || p.notesCanonical || {};

  const pick = (flatKey, colorKey, noteKey, defBg, defNote) => {
    const flat = p[flatKey] || {};
    return {
      bg: flat.bg || colors[colorKey] || defBg,
      note: flat.note || notes[noteKey] || defNote
    };
  };

  const dup = p.duplicate || {};

  return {
    expired: pick('expired', 'EXPIRED', 'EXPIRED', '#fff2cc', 'Policy already expired.'),
    flex: pick('flex', 'FLEX', 'FLEX', '#f4c7c3', 'Flex claim.'),
    b2b: pick('b2b', 'B2B', 'B2B', '#c9daf8', 'B2B claim.'),

    // New rules (colors per requirement)
    secondYear: pick('secondYear', 'SECOND_YEAR', 'SECOND_YEAR', '#d9ead3', 'Second-Year (Market Value).'),
    firstMonthPolicy: pick('firstMonthPolicy', 'FIRST_MONTH_POLICY', 'FIRST_MONTH_POLICY', '#d9d2e9', 'First-Month Policy.'),
    remaining1Month: pick('remaining1Month', 'REMAINING_1_MONTH', 'REMAINING_1_MONTH', '#fce5cd', 'Policy Remaining <= 1 Month.'),

    duplicate: {
      bg: dup.bg || colors.DUPLICATE || '#dd7e6b',
      notePrefix: dup.notePrefix || notes.DUPLICATE_PREFIX || 'Duplicate Claim - Refer to Claim Number'
    }
  };
}



function getClaimNumberHeaderAliases_() {
  const aliases = (CONFIG && (CONFIG.CLAIM_NUMBER_HEADER_ALIASES || CONFIG.claimNumberHeaderAliases))
    || ((typeof CLAIM_HIGHLIGHT_POLICY !== 'undefined' && CLAIM_HIGHLIGHT_POLICY && CLAIM_HIGHLIGHT_POLICY.CLAIM_NUMBER_HEADER_ALIASES)
      ? CLAIM_HIGHLIGHT_POLICY.CLAIM_NUMBER_HEADER_ALIASES
      : [])
    || [];
  const out = (aliases && aliases.length) ? aliases : ['Claim Number'];
  return out.map(s => String(s || '').trim()).filter(Boolean);
}

/**
 * Parse an optional integer in a strict way.
 * - Returns null for blanks/undefined/invalid.
 * - Never returns 0 for blank cells (prevents mass false-positives).
 */
function parseOptionalIntStrict_(v) {
  if (v === '' || v == null) return null;
  // Treat Date objects as invalid for integer parsing
  if (Object.prototype.toString.call(v) === '[object Date]') return null;
  if (typeof v === 'number') return isFinite(v) ? Math.trunc(v) : null;

  const s = String(v || '').trim();
  if (!s) return null;

  // Accept plain ints and numeric strings. Avoid parsing arbitrary text to 0.
  if (!/^-?\d+(?:\.\d+)?$/.test(s)) return null;

  const n = parseInt(s, 10);
  return isNaN(n) ? null : n;
}


function resolveHeaderIndexByAliases_(headerArr, aliases) {
  if (!headerArr || !headerArr.length) return null;
  const a = (aliases && aliases.length) ? aliases : ['Claim Number'];
  if (typeof findHeaderIndexAny_ === 'function') {
    const idx0 = findHeaderIndexAny_(headerArr, a, { enableSnakeCase: true });
    return (idx0 == null) ? null : idx0;
  }
  if (typeof idxAny_ === 'function') {
    const idx1 = idxAny_(headerArr, a, { enableSnakeCase: true });
    return (idx1 == null) ? null : idx1;
  }

  // Conservative fallback when shared utils are unavailable.
  const idx = buildHeaderIndex_(headerArr);
  for (let i = 0; i < a.length; i++) {
    const key = String(a[i] || '').trim();
    if (!key) continue;
    if (idx[key] != null) return idx[key];
  }

  const lower = headerArr.map(h => String(h || '').trim().toLowerCase());
  for (let i = 0; i < a.length; i++) {
    const key = String(a[i] || '').trim().toLowerCase();
    if (!key) continue;
    const j = lower.indexOf(key);
    if (j > -1) return j;
  }
  return null;
}

function normalizeColor_(c) {
  const s = String(c || '').trim().toLowerCase();
  return s;
}

function formatLogDate05b_(v) {
  const d = (typeof coerceDateOnly_ === 'function') ? coerceDateOnly_(v) : (v instanceof Date ? v : null);
  if (d) {
    try {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {
      return d.toDateString ? d.toDateString() : String(d);
    }
  }
  const s = String(v || '').trim();
  return s;
}


/** =========================
 * DB classification (Operational)
 * =========================
 * Requirement:
 *  - DB = OLD if Claim Number contains: SFP, SFX, SMR
 *  - DB = NEW if Claim Number contains: VVMAR, GADLD
 */
function computeDbValueFromClaimNumber05b_(claimNumber) {
  try {
    if (typeof computeDbValueFromClaimNumber_ === 'function') return computeDbValueFromClaimNumber_(claimNumber);
  } catch (e) {}
  const s = String(claimNumber == null ? '' : claimNumber).trim().toUpperCase();
  if (!s) return '';
  if (s.indexOf('SFP') !== -1 || s.indexOf('SFX') !== -1 || s.indexOf('SMR') !== -1) return 'OLD';
  if (s.indexOf('VVMAR') !== -1 || s.indexOf('GADLD') !== -1) return 'NEW';
  return '';
}


/**
 * Insurance shortening (Operational).
 * Requirement mapping (case-insensitive substring):
 * - great eastern general insurance -> GEGI
 * - tokio marine -> TMI
 * - msig indonesia -> MSIG
 * - seainsure -> MIGI
 * - sompo insurance -> Sompo
 * - axa mandiri insurance -> AXA
 * - simas insurtech -> Simas
 */
function normalizeInsuranceShort05b_(insuranceName) {
  const s = String(insuranceName == null ? '' : insuranceName).trim();
  if (!s) return '';
  try {
    if (typeof mapInsuranceShort_ === 'function') {
      const mapped = String(mapInsuranceShort_(s) || '').trim();
      if (mapped) return mapped;
    }
  } catch (e) {}
  return s;
}


/** Strict finite number coercion (blank-safe; avoids treating blank as 0). */
function __toFiniteNumber05b_(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v || '').trim();
  if (!s) return null;
  // Remove common thousand separators; keep minus and decimal.
  const n = Number(s.replace(/,/g, ''));
  return isFinite(n) ? n : null;
}



/** Apply highlight + note on Claim Number cells for OPERATIONAL sheets only. */
function applyOperationalClaimHighlightsByRaw_(ss, rawValues, headerIndexRaw, pic) {
  if (__isDryRun05b__()) return;

  // Notes are written based on RAW flags (as before). Background fill is derived from the final note,
  // so only rows that actually have EXPIRED/FLEX/B2B notes get colored. This also cleans up any
  // historical "mass fill" mistakes by clearing marker colors on non-flagged rows.
  const sets = buildOperationalClaimHighlightSetsFromRaw_(rawValues, headerIndexRaw);
  const policy = getOperationalClaimHighlightPolicy_();

  const isAdmin = (pic === 'Admin');
  const sheets = isAdmin ? CONFIG.sheetsByPic.adminOperational : CONFIG.sheetsByPic.picOperational;

  const cExpired = normalizeColor_(policy.expired.bg);
  const cFlex = normalizeColor_(policy.flex.bg);
  const cB2b = normalizeColor_(policy.b2b.bg);
  const cDup = normalizeColor_(policy.duplicate.bg);
  const cSecondYear = normalizeColor_(policy.secondYear.bg);
  const cFirstMonth = normalizeColor_(policy.firstMonthPolicy.bg);
  const cRemaining1Month = normalizeColor_(policy.remaining1Month.bg);

  const expiredNote = String(policy.expired.note || 'Policy already expired.');
  const flexNote = String(policy.flex.note || 'Flex claim.');
  const b2bNote = String(policy.b2b.note || 'B2B claim.');
  const secondYearNote = String(policy.secondYear.note || 'Second-Year (Market Value).');
  const firstMonthPolicyNote = String(policy.firstMonthPolicy.note || 'First-Month Policy.');
  const remaining1MonthNote = String(policy.remaining1Month.note || 'Policy Remaining <= 1 Month.');

  // Accept legacy Flex note variants.
  const flexNoteAlt = new Set([flexNote, 'FLEX claim.', 'Flex claim.']);
  const dupPrefix = String((policy.duplicate && policy.duplicate.notePrefix) || 'Duplicate Claim - Refer to Claim Number');


  const markerFromNote_ = (note) => {
    const n = String(note || '').trim();
    if (!n) return null;
    if (n.indexOf(dupPrefix) === 0) return 'duplicate';
    if (n === expiredNote) return 'expired';
    if (n === b2bNote) return 'b2b';
    if (n === secondYearNote) return 'secondYear';
    if (n === firstMonthPolicyNote) return 'firstMonthPolicy';
    if (n === remaining1MonthNote) return 'remaining1Month';
    if (flexNoteAlt.has(n)) return 'flex';
    return null;
  };

  const desiredBgFromMarker_ = (marker) => {
    if (marker === 'expired') return policy.expired.bg;
    if (marker === 'flex') return policy.flex.bg;
    if (marker === 'b2b') return policy.b2b.bg;
    if (marker === 'duplicate') return policy.duplicate.bg;
    if (marker === 'secondYear') return policy.secondYear.bg;
    if (marker === 'firstMonthPolicy') return policy.firstMonthPolicy.bg;
    if (marker === 'remaining1Month') return policy.remaining1Month.bg;
    return '';
  };

  const isMarkerBg_ = (bg) => {
    const c = normalizeColor_(bg);
    return c && (c === cExpired || c === cFlex || c === cB2b || c === cDup || c === cSecondYear || c === cFirstMonth || c === cRemaining1Month);
  };

  
const __setBgs05b__ = (range, matrix, sheetName) => {
  try {
    if (typeof safeSetBackgrounds_ === 'function') return safeSetBackgrounds_(range, matrix);
    range.setBackgrounds(matrix);
  } catch (e) {
    const a1 = (range && range.getA1Notation) ? range.getA1Notation() : '?';
    throw new Error(`[05b] Failed to set backgrounds on ${sheetName || '?'} ${a1}: ${e && e.message ? e.message : e}`);
  }
};

const __setNotes05b__ = (range, matrix, sheetName) => {
  try {
    if (typeof safeSetNotes_ === 'function') return safeSetNotes_(range, matrix);
    range.setNotes(matrix);
  } catch (e) {
    const a1 = (range && range.getA1Notation) ? range.getA1Notation() : '?';
    throw new Error(`[05b] Failed to set notes on ${sheetName || '?'} ${a1}: ${e && e.message ? e.message : e}`);
  }
};

// Priority (single fill): FLEX > B2B > EXPIRED
  sheets.forEach(name => {
    try {
      const sh = ss.getSheetByName(name);
      if (!sh) return;

      const lastRow = sh.getLastRow();
      if (lastRow < 2) return;

      const lastCol = Math.max(sh.getLastColumn(), 1);
      const header = sh.getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map(v => String(v || '').trim());

      const idxClaim = resolveHeaderIndexByAliases_(header, getClaimNumberHeaderAliases_());
      if (idxClaim == null) return;

      const n = lastRow - 1;
      const rng = sh.getRange(2, idxClaim + 1, n, 1);

      const vals = rng.getValues();
      const bgs = rng.getBackgrounds();
      const notes = rng.getNotes();

      let bgChanged = false;
      let noteChanged = false;

      for (let i = 0; i < n; i++) {
      const claim = String((vals[i] && vals[i][0]) || '').trim();
      const claimKey = claim ? claim.toUpperCase() : '';

      // 1) Compute desired note from RAW (keep existing behavior).
      let desiredNote = null;
      if (claimKey) {
        const dupNote = (sets.duplicate && typeof sets.duplicate.get === 'function') ? sets.duplicate.get(claimKey) : null;
        if (dupNote) desiredNote = dupNote;
        else if (sets.expired.has(claimKey)) desiredNote = expiredNote;
        else if (sets.flex.has(claimKey)) desiredNote = flexNote;
        else if (sets.b2b.has(claimKey)) desiredNote = b2bNote;
        else if (sets.secondYear && sets.secondYear.has(claimKey)) desiredNote = secondYearNote;
        else if (sets.firstMonthPolicy && sets.firstMonthPolicy.has(claimKey)) desiredNote = firstMonthPolicyNote;
        else if (sets.remaining1Month && sets.remaining1Month.has(claimKey)) desiredNote = remaining1MonthNote;
      }

      // Apply/clear our marker note only.
      if (desiredNote != null) {
        if (notes[i][0] !== desiredNote) { notes[i][0] = desiredNote; noteChanged = true; }
      } else {
        const curMarker = markerFromNote_(notes[i][0]);
        if (curMarker) {
          if (notes[i][0] !== '') { notes[i][0] = ''; noteChanged = true; }
        }
      }

      // 2) Background is derived from the final note state (robust + cleanup).
      const effNote = (desiredNote != null) ? desiredNote : String(notes[i][0] || '');
      const marker = markerFromNote_(effNote);
      const desiredBg = desiredBgFromMarker_(marker);

      if (desiredBg) {
        if (normalizeColor_(bgs[i][0]) !== normalizeColor_(desiredBg)) { bgs[i][0] = desiredBg; bgChanged = true; }
      } else {
        // Clear only our marker colors to avoid wiping user formatting.
        if (isMarkerBg_(bgs[i][0])) { bgs[i][0] = ''; bgChanged = true; }
      }
    }

      if (bgChanged) __setBgs05b__(rng, bgs, name);
      if (noteChanged) __setNotes05b__(rng, notes, name);
    } catch (sheetErr) {
      try { if (typeof logLine_ === 'function') logLine_('WARN', 'HIGHLIGHT_SKIP', String(name || ''), String(sheetErr), 'WARN'); } catch (eLog) {}
    }
  });
}


/** Build per-sheet writer (fast, typed) */
function buildSheetWriters_(ss, routingMap, headerIndexRaw, pic) {
  const h = CONFIG.headers;
  const writers = {};

  const flowName = __getRuntimeFlow05b_();


  // Single-master default: treat missing pic as Admin semantics.
  const isAdmin = (pic === 'Admin' || pic == null);

  const opsPolicy =
    (CONFIG && (CONFIG.opsRouting || CONFIG.opsRoutingPolicy || CONFIG.OPS_ROUTING_POLICY)) || null;

  const scFarhanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_FARHAN) ? opsPolicy.SHEETS.SC_FARHAN : 'SC - Farhan';
  const scMeilaniName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_MEILANI) ? opsPolicy.SHEETS.SC_MEILANI : 'SC - Meilani';

  const scIvanName = (opsPolicy && opsPolicy.SHEETS && (opsPolicy.SHEETS.SC_IVAN || opsPolicy.SHEETS.SC_IVAN_NAME)) ? (opsPolicy.SHEETS.SC_IVAN || opsPolicy.SHEETS.SC_IVAN_NAME) : 'SC - Meindar';

  // Precompute Type lookup for SC sheets (write only if header has "Type").
// Mapping lives in OPS_ROUTING_POLICY.TYPE_BY_LAST_STATUS (00.gs) and is matched by Last Status.
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

  // Precedence (specific > general):
  // - SC - On Rep / SC - Wait Rep override Insurance (overlap in status universe).
  const resolveType = (statusVal) => {
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

  Object.keys(routingMap).forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    // Mandatory column for operational sheets (append if missing)
    __ensureAppendHeader05b_(sh, 'Status Type');

    const lc = sh.getLastColumn() || 20;
    const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v || '').trim());
    const idxH = buildHeaderIndex_(header);
    const dbLinkUrls = [];

    const isScSheet = (sheetName === scFarhanName || sheetName === scMeilaniName || sheetName === scIvanName || sheetName === 'SC - Unmapped');

    const getRaw = (rawRow, rawKey) => {
      const ix = headerIndexRaw[rawKey];
      return (ix != null) ? rawRow[ix] : '';
    };

    // Read first NON-EMPTY raw value across multiple possible raw keys.
    const getRawAny = (rawRow, rawKeys) => {
      const list = rawKeys || [];
      for (let i = 0; i < list.length; i++) {
        const k = list[i];
        if (!k) continue;
        const ix = headerIndexRaw[k];
        if (ix == null) continue;
        const v = rawRow[ix];
        if (v === '' || v == null) continue;
        return v;
      }
      return '';
    };
    // Read from one logical raw source only, but tolerate header formatting variants
    // of that exact source key (e.g. "claim submitted datetime" vs "claim_submitted_datetime").
    const getRawByCanonicalKey = (rawRow, canonicalKey) => {
      const target = String(canonicalKey || '').trim().toLowerCase();
      if (!target) return '';
      const rawKeys = Object.keys(headerIndexRaw || {});
      for (let i = 0; i < rawKeys.length; i++) {
        const k = rawKeys[i];
        if (!k) continue;
        if (canonicalizeHeaderSnake_(k) !== target) continue;
        const ix = headerIndexRaw[k];
        if (ix == null) continue;
        return rawRow[ix];
      }
      return '';
    };

    writers[sheetName] = {
      sheet: sh,
      header,
      rows: [],
      dbLinkUrls,

      build: (rawRow) => {
        const out = header.map(() => '');

        const set = (colName, value) => {
          const c0 = idxH[colName];
          if (c0 == null) return;
          out[c0] = (value != null) ? value : '';
        };

        // Core columns (best-effort; only sets if headers exist)
        const claimNumberVal = String(getRaw(rawRow, h.claimNumber) || '').trim();
        const lastStatusVal = getRaw(rawRow, h.lastStatus);
        set('Claim Number', claimNumberVal);
        set('Last Status', lastStatusVal);
        set('Partner', getRaw(rawRow, h.businessPartner));

        // Activity Log (optional column)
        if (idxH['Activity Log'] != null) {
          set('Activity Log', __pickActivityLogValue05b_(rawRow, headerIndexRaw, flowName));
        }
        if (idxH['Activity Log Datetime'] != null) {
          const v = __pickActivityLogDatetimeValue05b_(rawRow, headerIndexRaw, flowName);
          const dt = coerceDateTime_(v);
          set('Activity Log Datetime', dt ? dt : '');
        }

        // Status Type (mandatory)
        if (idxH['Status Type'] != null) {
          set('Status Type', __getStatusType05b_(lastStatusVal));
        }
        set('Product', getRaw(rawRow, h.productName));

        // Extended operational columns (best-effort; only applies when those headers exist)
        // - Partner Name / Insurance / Device Type / Service Center
        set('Partner Name', getRaw(rawRow, h.businessPartner));
        set('Insurance', normalizeInsuranceShort05b_(getRawAny(rawRow, [h.insuranceName, h.insurance, 'insurance_name', 'insurance'])));
        set('Device Type', getRawAny(rawRow, [h.deviceType, 'device_type', 'deviceType']));
        set('Store Name', getRawAny(rawRow, ['3. All Transaction - qoala_policy_number → outlet_name', 'outlet_name', 'store_name', 'Store Name']));
        set('PA Name', getRawAny(rawRow, ['3. All Transaction - qoala_policy_number → pa_name', 'pa_name', 'PA Name']));
        set('SPA Name', getRawAny(rawRow, ['3. All Transaction - qoala_policy_number → spa_name', 'spa_name', 'SPA Name']));
        set('Service Center', getRawAny(rawRow, [h.serviceCenter, h.serviceCenterName, h.scName, 'service_center', 'service_center_name', 'sc_name']));
        set('Service Center Name', getRawAny(rawRow, [h.serviceCenterName, h.serviceCenter, h.scName, 'service_center_name', 'service_center', 'sc_name']));
        // - Device Brand / IMEI
        set('Device Brand', getRawAny(rawRow, [h.deviceBrand, 'device_brand', 'brand']));
        set('IMEI/SN', getRawAny(rawRow, [h.imeiNumber, h.imei, 'imei_number', 'imei', 'serial_number', 'sn']));

        // - DB OLD/NEW (computed from Claim Number codes; fallback to raw when not matched)
        const dbComputed = computeDbValueFromClaimNumber05b_(claimNumberVal);
        const dbRaw = getRawAny(rawRow, [h.dbClass, 'DB', 'db', 'db_class', 'old_new']);
        set('DB', dbComputed || dbRaw);

        // - Activity Log Aging (ALA) & TAT (best-effort)
        const activityLogAgingVal = normalizeInt_(getRawAny(rawRow, [
          h.activityLogAging,
          h.ala,
          'Activity Log Aging',
          'ALA',
          'activity_log_aging',
          'ala'
        ]));
        set('Activity Log Aging', activityLogAgingVal);
        set('ALA', activityLogAgingVal);
        set('TAT', normalizeInt_(getRawAny(rawRow, [h.tat, 'TAT', 'tat'])));

        // Dates (submission)
        if (idxH['Submission Date'] != null) {
          // Source of truth remains ONE field: claim_submitted_datetime.
          // Only tolerate formatting variants of this exact header name.
          const rawSubmissionVal = getRawByCanonicalKey(rawRow, 'claim_submitted_datetime');
          const d = coerceDate_(rawSubmissionVal);
          // IMPORTANT: never blank-out when parser misses a valid source representation.
          // Keep raw value as fallback so Submission Date is still populated.
          set('Submission Date', d ? d : (rawSubmissionVal != null ? rawSubmissionVal : ''));
        }
        if (idxH['Submission Datetime'] != null) {
          const rawSubmissionDatetimeVal = getRawByCanonicalKey(rawRow, 'claim_submitted_datetime');
          const dt = coerceDateTime_(rawSubmissionDatetimeVal);
          set('Submission Datetime', dt ? dt : '');
        }

        // Policy start/end
        if (idxH['Policy Start Date'] != null) {
          const d = coerceDate_(getRaw(rawRow, h.policyStartDate));
          set('Policy Start Date', d ? d : '');
        }
        if (idxH['Policy End Date'] != null) {
          const d = coerceDate_(getRaw(rawRow, h.policyEndDate));
          set('Policy End Date', d ? d : '');
        }

        // Aging
        if (idxH['Days to End Policy'] != null) {
          const v = normalizeInt_(getRaw(rawRow, h.daysToEndPolicy));
          set('Days to End Policy', (v != null) ? v : '');
        }
        if (idxH['Days Aging'] != null) {
          const v = normalizeInt_(getRaw(rawRow, h.daysAgingFromSubmission));
          set('Days Aging', (v != null) ? v : '');
        }
        // Last Status Aging (LSA) - keep both headers for backward compatibility
        const lastStatusAgingVal = normalizeInt_(getRawAny(rawRow, [
          h.lastStatusAging,
          'Last Status Aging',
          'LSA',
          'last_status_aging',
          'lsa'
        ]));
        set('Last Status Aging', lastStatusAgingVal);
        set('LSA', lastStatusAgingVal);

        // Money-like (best-effort)
        const sumInsured = normalizeMoney_(getRawAny(rawRow, [
          h.sumInsuredAmount,
          'sum_insured_amount',
          'sum_insured',
          'sum_insured_amount_idr'
        ]));

        const nettAmt = normalizeMoney_(getRawAny(rawRow, [
          h.nettClaimAmount,
          'nett_claim_amount',
          'nettClaimAmount',
          'nett_claim_amount_idr'
        ]));

        const claimAmtRaw = normalizeMoney_(getRawAny(rawRow, [
          h.claimAmount,
          'claim_amount'
        ]));
        const claimAmt = (claimAmtRaw != null && claimAmtRaw !== '') ? claimAmtRaw : nettAmt;

        const orAmtRaw = normalizeMoney_(getRawAny(rawRow, [
          h.orAmount,
          'or_amount',
          'orAmount'
        ]));

        const ownRiskRaw = normalizeMoney_(getRawAny(rawRow, [
          h.claimOwnRiskAmount,
          'claim_own_risk_amount',
          'claim_or_amount'
        ]));
        const ownRiskAmt = (ownRiskRaw != null && ownRiskRaw !== '') ? ownRiskRaw : orAmtRaw;

        set('Sum Insured', sumInsured);
        set('Sum Insured Amount', sumInsured);
        set('Claim Amount', claimAmt);
        // Legacy Special Case header (renamed to Claim Amount)
        set('Repair/Replace Amount', claimAmt);
        // OR Amount renamed to Claim Own Risk Amount
        set('Claim Own Risk Amount', ownRiskAmt);
        set('OR Amount', ownRiskAmt);
        set('Nett Claim Amount', nettAmt);

        // Selisih (Special Case) = Sum Insured Amount - Claim Amount
        if (idxH['Selisih'] != null) {
          const s = __toFiniteNumber05b_(sumInsured);
          const a = __toFiniteNumber05b_(claimAmt);
          set('Selisih', (s != null && a != null) ? (s - a) : '');
        }

        // % Approval: store as ratio (e.g. 0.75), format applied by schema formatter.
        if (idxH['% Approval'] != null) {
          const a = __toFiniteNumber05b_(claimAmt);
          const s = __toFiniteNumber05b_(sumInsured);
          set('% Approval', (a != null && s != null && s !== 0) ? (a / s) : '');
        }

        // OR checkbox (keep text if sheet expects checkbox; formatting handled elsewhere)
        set('OR', getRaw(rawRow, h.orFlag));

        // Links
        if (idxH['DB Link'] != null) {
          const url = String(getRaw(rawRow, h.dashboardLink) || '').trim();
          set('DB Link', url ? url : '');
          if (url) dbLinkUrls.push(url);
        }

        // Timestamp if exists
        if (idxH['Timestamp'] != null) {
          const ts = coerceTimestamp_(getRaw(rawRow, h.timestamp));
          set('Timestamp', ts ? ts : '');
        }

        // Status column exists, but we WILL NOT WRITE it on flush (avoid validation traps)
        set('Status', getRaw(rawRow, h.status));

        // Associate (single-master default: allow Admin semantics)
        if (isAdmin && idxH['Associate'] != null) {
          set('Associate', sanitizeAssociateForWrite_(getRaw(rawRow, h.associate)));
        }

        // SC sheet "Type" dropdown auto-fill
        if (isScSheet && idxH['Type'] != null) {
          const typeLabel = resolveType(lastStatusVal);
          if (typeLabel) set('Type', typeLabel);
        }

        // Last Status Date/Datetime:
        // Single-master: prefer Raw last_update; fallback to last_activity_log_date if needed.
        if (idxH['Last Status Date'] != null || idxH['Last Status Datetime'] != null) {
          const rawVal = getRawAny(rawRow, [h.claimLastUpdatedDatetime, 'claim_last_updated_datetime', h.lastUpdate, 'last_update_datetime', 'last_update', h.lastActivityLogDate, 'last_activity_log_date', 'last_activity_log_datetime']);
          const dt = coerceDateTime_(rawVal);
          if (idxH['Last Status Date'] != null) set('Last Status Date', dt ? dt : '');
          if (idxH['Last Status Datetime'] != null) set('Last Status Datetime', dt ? dt : '');
        }

        return out;
      }
    };
  });

  return writers;
}

function clearOperationalSheets_(ss, pic) {
  // Single-master default: treat missing pic as Admin semantics.
  const isAdmin = (pic === 'Admin' || pic == null);
  const sheetsBase = isAdmin ? CONFIG.sheetsByPic.adminOperational : CONFIG.sheetsByPic.picOperational;
  const sheets = sheetsBase.slice();
  try {
    if (ss.getSheetByName('SC - Unmapped') && sheets.indexOf('SC - Unmapped') === -1) sheets.push('SC - Unmapped');
  } catch (e) {}

  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const buffer = (name === 'Ask Detail') ? 1500 : 400;
    // clearFormat() does not remove data validations (dropdowns stay intact)
    clearSheetDataHard_(sh, { bufferRows: buffer, clearFormats: true, preserveTemplateRow: true });

// FIX: clearSheetDataHard_ with preserveTemplateRow=true preserves row 2's format AND DV
// as a template (intentional for Status dropdown chip style). However, if Submission Date
// column had stale checkbox DV from a previous bad run, that DV remains on row 2 and will
// bleed into all new rows when the routing writer writes Date values—causing cells to render
// as checkboxes instead of dates. Defensively clear Submission Date DV across ALL rows
// (including row 2) before any new data is written.
if (!__isDryRun05b__()) {
  try {
    const lc = Math.max(sh.getLastColumn(), 1);
    const mr = sh.getMaxRows();
    const hdrVals = sh.getRange(1, 1, 1, lc).getValues()[0];
    const subDateIdx = hdrVals.findIndex(function(h) {
      return String(h == null ? '' : h).trim().toLowerCase() === 'submission date';
    });
    if (subDateIdx !== -1 && mr > 1) {
      sh.getRange(2, subDateIdx + 1, mr - 1, 1).clearDataValidations();
    }
  } catch (_eClearSubDate) {}
}

    // Requirement: Operational header (row 1) must be centered (horizontal) and middle (vertical).
    if (!__isDryRun05b__()) {
      try {
        const lastCol = Math.max(sh.getLastColumn(), 1);
        sh.getRange(1, 1, 1, lastCol)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      } catch (e) {}
    }
  });
}

function routeRawToOperationalSheetsInMemory_(ss, rawValues, headerIndexRaw, pic) {
  const h = CONFIG.headers;
  if (!rawValues || !rawValues.length) return { total: 0, perSheet: {}, unknown: [], missingStatus: [] };

  // Single-master default: treat missing pic as Admin semantics.
  const isAdmin = (pic === 'Admin' || pic == null);

  // Prefer new routing policy; fallback to legacy maps if present.
  const routingMapOrig = (CONFIG && (CONFIG.statusRoutingAdmin || CONFIG.statusRoutingPIC)) || {};
  const routingMap = Object.assign({}, routingMapOrig);


  // Always include SC quarantine sheet to prevent mis-mapping (fail-closed).
  if (routingMap['SC - Unmapped'] == null) routingMap['SC - Unmapped'] = [];

  // Ensure the quarantine sheet exists before building writers.
  ensureScFallbackSheet05b_(ss, 'SC - Unmapped', 'SC - Farhan');

  const routingIndex = compileRoutingIndex_(routingMap);
  const writers = buildSheetWriters_(ss, routingMap, headerIndexRaw, pic);

  const opsPolicy =
    (CONFIG && (CONFIG.opsRouting || CONFIG.opsRoutingPolicy || CONFIG.OPS_ROUTING_POLICY)) || null;

  const scFarhanName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_FARHAN) ? opsPolicy.SHEETS.SC_FARHAN : 'SC - Farhan';
  const scMeilaniName = (opsPolicy && opsPolicy.SHEETS && opsPolicy.SHEETS.SC_MEILANI) ? opsPolicy.SHEETS.SC_MEILANI : 'SC - Meilani';

  const scIvanName = (opsPolicy && opsPolicy.SHEETS && (opsPolicy.SHEETS.SC_IVAN || opsPolicy.SHEETS.SC_IVAN_NAME)) ? (opsPolicy.SHEETS.SC_IVAN || opsPolicy.SHEETS.SC_IVAN_NAME) : 'SC - Meindar';

  const scFallbackName = 'SC - Unmapped';

  const scKeywords = (opsPolicy && opsPolicy.SC_NAME_KEYWORDS) ? opsPolicy.SC_NAME_KEYWORDS : {};
  const kwFarhan = scKeywords[scFarhanName] || scKeywords['SC - Farhan'] || [];
  const kwMeilani = scKeywords[scMeilaniName] || scKeywords['SC - Meilani'] || [];
  const kwIvan = scKeywords[scIvanName] || scKeywords['SC - Meindar'] || [];

  const idxClaim = headerIndexRaw[h.claimNumber];
  const idxLastStatus = headerIndexRaw[h.lastStatus];
  const idxDaysAging = headerIndexRaw[h.daysAgingFromSubmission];
  const idxPartnerName = headerIndexRaw[h.businessPartner];
  const idxScName = headerIndexRaw[h.scName];

  const unknownStatuses = [];
  const missingStatus = [];
  const routeCount = {};
  let total = 0;

  for (let r = 0; r < rawValues.length; r++) {
    const rawRow = rawValues[r];
    if (!rawRow) continue;
    const rowNumber = r + 2;

    const claimVal = (idxClaim != null) ? String(rawRow[idxClaim] || '').trim() : '';
    const statusVal = (idxLastStatus != null) ? String(rawRow[idxLastStatus] || '').trim() : '';
    const partnerVal = (idxPartnerName != null) ? String(rawRow[idxPartnerName] || '').trim() : '';
    const daysAging = (idxDaysAging != null) ? normalizeInt_(rawRow[idxDaysAging]) : null;

    if (!statusVal) {
      missingStatus.push({ rowNumber, claim: claimVal, partner: partnerVal });
      continue;
    }

    let targets = (routingIndex[statusVal] || []).slice();

    // Patch B1: force Finish statuses into SC routing.
    if (isFinishStatus05a_(statusVal)) {
      targets = uniq05a_(targets.concat([scFarhanName, scMeilaniName, scIvanName, scFallbackName]));
    }

    // SC sheet split by sc_name keywords (only if targets include SC sheets)
    const scNameVal = (idxScName != null) ? String(rawRow[idxScName] || '') : '';
    if (targets.length && (targets.indexOf(scFarhanName) > -1 || targets.indexOf(scMeilaniName) > -1 || targets.indexOf(scIvanName) > -1)) {
      targets = filterScTargets05b_(targets, scNameVal, scFarhanName, scMeilaniName, scIvanName, kwFarhan, kwMeilani, kwIvan, scFallbackName);
    }

    // NOTE: previously there were EV-Bike / PIC suppressions. Those are intentionally removed.
    // We always route all incoming data in the single master workflow.

    if (!targets.length) {
      unknownStatuses.push({
      rowNumber,
      claim: claimVal,
      partner: partnerVal,
      status: statusVal,
      sc: scNameVal,
        daysAging,
      });
      continue;
    }

    for (let t = 0; t < targets.length; t++) {
      const sheetName = targets[t];
      const w = writers[sheetName];
      if (!w) continue;
      w.rows.push(w.build(rawRow));
      routeCount[sheetName] = (routeCount[sheetName] || 0) + 1;
      total++;
    }
  }

  // Flush writers: NEVER write "Status" column to avoid validation trap
  Object.keys(writers).forEach(sheetName => {
    const w = writers[sheetName];
    if (!w.rows.length) return;

    const sh = w.sheet;
    const header = w.header;
    const n = w.rows.length;
    const startRow = 2;

    const idxStatus = header.indexOf('Status');
    if (idxStatus === -1) {
      safeSetValues_(sh.getRange(startRow, 1, n, header.length), w.rows);
    } else {
      if (idxStatus > 0) {
        const left = w.rows.map(r => r.slice(0, idxStatus));
        safeSetValues_(sh.getRange(startRow, 1, n, idxStatus), left);
      }
      if (idxStatus < header.length - 1) {
        const right = w.rows.map(r => r.slice(idxStatus + 1));
        safeSetValues_(sh.getRange(startRow, idxStatus + 2, n, header.length - idxStatus - 1), right);
      }
    }

    // Ensure header alignment (center + middle) on the destination sheet.
    if (!__isDryRun05b__()) {
      try {
        sh.getRange(1, 1, 1, Math.max(header.length, 1))
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      } catch (e) {}
    }

    // Column formatting minimal (does not touch data validations)
    applyOperationalColumnSchema_(sh, header, startRow, n, { orIsMoney: false });

    // DB Link RichText if supported
    if (typeof applyDbLinkRichTextFromWriter_ === 'function') {
      applyDbLinkRichTextFromWriter_(w, startRow);
    }
  });

  // Highlight + note by claim: operational sheets only (derived from full Raw, regardless of routing)
  applyOperationalClaimHighlightsByRaw_(ss, rawValues, headerIndexRaw, pic);

  // Unknown status logging (single-master: log only, do not spam Details by default)
  if (unknownStatuses.length && typeof logLine_ === 'function') {
    const maxN = (typeof MAPPING_ERROR_LOG_POLICY === 'object' && MAPPING_ERROR_LOG_POLICY && MAPPING_ERROR_LOG_POLICY.maxPerRun != null)
      ? Number(MAPPING_ERROR_LOG_POLICY.maxPerRun)
      : 200;

    for (let i = 0; i < unknownStatuses.length && i < maxN; i++) {
      const x = unknownStatuses[i] || {};
      const dateStr = formatLogDate05b_(x.submissionDateVal);
      const notes = `Last Status: ${x.status || ''} | Claim Number: ${x.claim || ''} | SC: ${x.sc || ''} | Date: ${dateStr || ''}`;
      logLine_('MAP', 'Last status not routed', '', notes, 'WARN');
    }
  }

  // Keep old return shape mostly, but remove EV-Bike artifacts.
  return { total, perSheet: routeCount, unknown: unknownStatuses, missingStatus };
}



/** =========================
 *  Sorting
 *  ========================= */

function sortOperationalSheets_(ss, pic) {
  // Requirement: preserve filters & sort by:
  //  1) Last Status Date (A->Z)
  //  2) Last Status (A->Z)
  const isAdmin = (pic === 'Admin' || pic == null);
  const ops = isAdmin ? CONFIG.sheetsByPic.adminOperational : CONFIG.sheetsByPic.picOperational;
  const optional = (!isAdmin) ? (CONFIG.sheetsByPic.optional || []) : []; // Admin skips optional
  const names = ops.concat(optional);

  // Include SC quarantine sheet if present
  try {
    if (ss.getSheetByName('SC - Unmapped') && names.indexOf('SC - Unmapped') === -1) names.push('SC - Unmapped');
  } catch (e) {}

  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 2 || lastCol < 1) return;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    const idx = buildHeaderIndex_(header);

    const idxLsd = (idx['Last Status Date'] != null) ? idx['Last Status Date'] : -1;
    const idxLs = (idx['Last Status'] != null) ? idx['Last Status'] : -1;

    const criteriaAbs = [];
    if (idxLsd > -1) criteriaAbs.push({ colAbs: idxLsd + 1, ascending: true });
    if (idxLs > -1) criteriaAbs.push({ colAbs: idxLs + 1, ascending: true });

    if (!criteriaAbs.length || __isDryRun05b__()) return;

    try {
      const f = (typeof sh.getFilter === 'function') ? sh.getFilter() : null;
      if (f && typeof f.getRange === 'function') {
        const fr = f.getRange();
        const frColStart = fr.getColumn();
        const frColEnd = fr.getColumn() + fr.getNumColumns() - 1;

        // Safety: only sort inside filter range if it covers required columns.
        let ok = (fr.getNumRows() > 1);
        for (let i = 0; i < criteriaAbs.length; i++) {
          const cAbs = criteriaAbs[i].colAbs;
          if (cAbs < frColStart || cAbs > frColEnd) { ok = false; break; }
        }
        if (!ok) return;

        const sortRange = fr.offset(1, 0, fr.getNumRows() - 1, fr.getNumColumns());
        const criteriaRel = criteriaAbs.map(x => ({ column: (x.colAbs - frColStart + 1), ascending: x.ascending }));
        sortRange.sort(criteriaRel);
        return;
      }

      // No filter: sort full data range.
      const criteria = criteriaAbs.map(x => ({ column: x.colAbs, ascending: x.ascending }));
      sh.getRange(2, 1, lastRow - 1, lastCol).sort(criteria);
    } catch (e) {
      // best-effort only
    }
  });
}
