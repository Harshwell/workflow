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

