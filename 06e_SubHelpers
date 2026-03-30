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
