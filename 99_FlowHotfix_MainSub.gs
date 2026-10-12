/***************************************************************
 * 99_FlowHotfix_MainSub.gs
 * MAIN/SUB trigger coordination hotfix.
 *
 * Goals:
 * - Do not hard-skip SUB at MAIN_HOUR; let the existing shared lock/pending
 *   mechanism serialize MAIN and SUB safely.
 * - Recover a stale MAIN_PIPELINE_STAGE2 checkpoint before SUB proceeds.
 * - Keep the existing runSubEmailIngest business logic unchanged.
 ***************************************************************/
'use strict';

(function installMainSubFlowHotfix06g_() {
  const g = (typeof globalThis !== 'undefined') ? globalThis : Function('return this')();
  if (g.__MAIN_SUB_FLOW_HOTFIX_06G__) return;

  const original = g.runSubEmailIngest;
  if (typeof original !== 'function') return;

  const STALE_MS = 15 * 60 * 1000;

  g.runSubEmailIngest = function(maxThreads, options) {
    const opt = Object.assign({}, options || {});
    const props = PropertiesService.getScriptProperties();
    const checkpoint = props.getProperty('MAIN_PIPELINE_STAGE2');

    // A checkpoint that survives well beyond one normal continuation window is
    // no longer safe to leave blocking SUB indefinitely. Use the existing
    // checkpoint-aware retry entry point; it preserves the durable state and
    // either completes it or re-arms the continuation trigger.
    if (checkpoint && typeof g.retryMainPipelineStage2_ === 'function') {
      let state = null;
      try { state = JSON.parse(checkpoint); } catch (eParse) { state = null; }

      const createdAt = Number(state && (state.createdAt || state.updatedAt) || 0);
      const ageMs = createdAt > 0 ? Date.now() - createdAt : 0;
      const stale = createdAt > 0 && ageMs >= STALE_MS;

      if (stale) {
        try {
          if (typeof g.logLine_ === 'function') {
            g.logLine_(
              'MAIN_RECOVERY',
              'Auto-retry stale MAIN continuation before SUB',
              'runId=' + String(state && state.token || '') + ' step=' + String(state && state.nextStep || 0),
              'ageMs=' + ageMs,
              'WARN'
            );
          }
          g.retryMainPipelineStage2_();
        } catch (err) {
          try {
            if (typeof g.logLine_ === 'function') {
              g.logLine_(
                'MAIN_RECOVERY_ERR',
                'Stale MAIN continuation retry failed; SUB remains gated',
                'runId=' + String(state && state.token || '') + ' step=' + String(state && state.nextStep || 0),
                String(err && err.stack ? err.stack : err),
                'ERROR'
              );
            }
          } catch (eLog) {}
        }
      }
    }

    // The original SUB entrypoint hard-skips MAIN_HOUR before its shared lock.
    // Bypass only that time-window skip. If MAIN is actually running, the
    // existing __withTryLockSub06a_ path still queues SUB instead of colliding.
    const tz = (typeof g.getTzSafe_ === 'function')
      ? g.getTzSafe_()
      : ((typeof Session !== 'undefined' && Session.getScriptTimeZone) ? Session.getScriptTimeZone() : 'Asia/Jakarta');
    const hour = Number(Utilities.formatDate(new Date(), tz, 'H'));
    const mainHour = Number(g.FLOW_SCHEDULE_POLICY && g.FLOW_SCHEDULE_POLICY.MAIN_HOUR);

    if (hour === mainHour && !opt.forceAfterMain) {
      opt.forceAfterMain = true;
      try {
        if (typeof g.logLine_ === 'function') {
          g.logLine_('SUB_MAIN_WINDOW', 'SUB allowed during MAIN hour; shared lock handles serialization', 'hour=' + hour, '', 'INFO');
        }
      } catch (eLog2) {}
    }

    return original(maxThreads, opt);
  };

  g.__MAIN_SUB_FLOW_HOTFIX_06G__ = true;
})();
