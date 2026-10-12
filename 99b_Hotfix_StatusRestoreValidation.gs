/***************************************************************
 * 99b_Hotfix_StatusRestoreValidation.gs
 *
 * MAIN Stage 2 restore hardening.
 *
 * Root issue:
 * restoreOpsManualFromBackupSheet06c_ can receive a legacy/manual Status
 * from _OPS_MANUAL_BACKUP that is not present in the canonical Status
 * dropdown. Writing that value into a validated Status cell throws and
 * leaves MAIN_STAGE2 checkpointed.
 *
 * Policy:
 * - Overview!C945 remains the canonical native Status template.
 * - Never disable/bypass Status validation.
 * - Only restore values that exist in STATUS_DROPDOWN_OPTIONS.
 * - Invalid legacy values are skipped, logged with exact sheet/cell/claim/value,
 *   and the existing target value remains untouched.
 ***************************************************************/
'use strict';

(function installStatusRestoreValidationHotfix06g_() {
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();
  if (g.__STATUS_RESTORE_VALIDATION_HOTFIX_06G__) return;

  const original = g.__restoreStatusValuesWithCanonicalValidation06c_;
  if (typeof original !== 'function') return;

  function getCanonicalStatusSet_() {
    let options = [];
    try {
      if (typeof STATUS_DROPDOWN_OPTIONS !== 'undefined' && Array.isArray(STATUS_DROPDOWN_OPTIONS)) {
        options = STATUS_DROPDOWN_OPTIONS.slice();
      }
    } catch (e) {}

    if (!options.length) {
      options = [
        'Pending Logistic', 'Pending Front', 'DONE', 'Pending Insurance',
        'Pending TO', 'Pending Finance', 'Pending PIC', 'Pending Cust',
        'Pending Buss. Team', 'Pending SC (TA)', 'Pending SC (Repair)',
        'Pending SC (Estimation)', 'Delivering', 'Delivered',
        'Waiting Courier', 'Re-pickup'
      ];
    }

    return new Set(options.map(function(v) { return String(v == null ? '' : v).trim(); }).filter(Boolean));
  }

  function logInvalid_(sourceLabel, sh, row, col1, claim, value) {
    try {
      if (typeof logLine_ === 'function') {
        logLine_(
          'MAIN_STAGE2_STATUS_SKIP',
          'Skipped invalid Status during backup restore',
          'source=' + String(sourceLabel || '') +
            ' sheet=' + String(sh && sh.getName ? sh.getName() : '') +
            ' cell=' + String(sh && sh.getRange ? sh.getRange(row, col1).getA1Notation() : '') +
            ' claim=' + String(claim || ''),
          'value=' + String(value == null ? '' : value) +
            ' canonicalTemplate=Overview!C945',
          'WARN'
        );
      }
    } catch (e) {}
  }

  g.__restoreStatusValuesWithCanonicalValidation06c_ = function(sh, startRow, col1, values, claims, sourceLabel) {
    if (!sh || !values || !values.length || col1 < 1) return { written: 0, failed: 0, skippedInvalid: 0 };

    const allowed = getCanonicalStatusSet_();
    const filtered = [];
    let skippedInvalid = 0;

    for (let i = 0; i < values.length; i++) {
      const row = startRow + i;
      const raw = values[i] && values[i][0] != null ? values[i][0] : '';
      const value = String(raw).trim();
      const claim = claims && claims[i] ? claims[i][0] : '';

      // Blank Status is allowed to remain blank; do not force a dropdown value.
      if (!value || allowed.has(value)) {
        filtered.push([raw]);
      } else {
        skippedInvalid++;
        filtered.push(['']);
        logInvalid_(sourceLabel, sh, row, col1, claim, raw);
      }
    }

    // Delegate the actual canonical-validation/template handling to the
    // existing hardened implementation, but never pass it an invalid value.
    let result = original(sh, startRow, col1, filtered, claims, sourceLabel);
    result = result || {};
    result.skippedInvalid = skippedInvalid;
    return result;
  };

  g.__STATUS_RESTORE_VALIDATION_HOTFIX_06G__ = true;
})();
