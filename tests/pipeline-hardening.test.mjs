import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { loadFunctions } from '../scripts/lib/source-contracts.mjs';

const optionalSource = fs.readFileSync('05c_Pipeline_OptionalSheets.gs', 'utf8');
const routingSource = fs.readFileSync('05b_Pipeline_RoutingOperational.gs', 'utf8');
const pipelineSource = fs.readFileSync('06b_PipelineAndEnrichment.gs', 'utf8');

function sheet(name, rows) {
  const writes = [];
  const notes = rows.map(() => ['']);
  const colors = rows.map(() => ['#ffffff']);
  return {
    rows, writes, notes, colors, getName: () => name,
    getLastRow: () => rows.length, getLastColumn: () => rows[0].length,
    getRange(row, col, count, width) {
      return {
        getValues: () => rows.slice(row - 1, row - 1 + count).map(r => r.slice(col - 1, col - 1 + width).map(v => String(v).startsWith('=') ? 'calculated' : v)),
        clearDataValidations() {},
        setValues(values) {
          writes.push({ row, col, count, width });
          values.forEach((r, i) => {
            if (!rows[row - 1 + i]) rows[row - 1 + i] = new Array(rows[0].length).fill('');
            r.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
          });
        },
        getNotes: () => notes.slice(row - 1, row - 1 + count).map(r => r.slice()),
        getBackgrounds: () => colors.slice(row - 1, row - 1 + count).map(r => r.slice()),
        setNotes(values) { if (this.failNotes) throw new Error('note write failed'); values.forEach((r, i) => { notes[row - 1 + i] = r.slice(); }); },
        setBackgrounds(values) { values.forEach((r, i) => { colors[row - 1 + i] = r.slice(); }); }
      };
    }
  };
}

test('SUB EV-Bike and Doss append new claims and write only two existing columns', () => {
  for (const [name, token] of [['EV-Bike', 'VVMAR'], ['Doss', 'DOSS']]) {
    const header = ['Claim Number', 'Last Status', 'Last Status Aging', 'Status', 'Remarks', 'DB Link', 'TAT', 'Owner Name'];
    const sh = sheet(name, [header, [token + '-001', 'OLD', 8, 'Manual', 'Keep', 'rich-link', '=1+1', 'Original']]);
    const before = sh.rows[1].slice();
    const api = loadFunctions(optionalSource, ['updateTokenOptionalSheetFromSubRaw05c_', 'processEVBike_', 'processDoss_'], {
      RUNTIME: { flowName: 'sub' }, CONFIG: { headers: { claimNumber: 'claim_number', lastStatus: 'claim_last_status_name', businessPartner: 'partner_name' }, patterns: { evBikePartners: ['Ofero'] } },
      __getHeaderRow05c_: () => header, buildHeaderIndex_: h => Object.fromEntries(h.map((v, i) => [v, i])),
      __isDryRun05c__: () => false, safeSetValues_: (range, values) => range.setValues(values),
      normalizeInt_: value => Number(value), mapInsuranceShort_: value => value,
      buildSubmissionDateCell_: () => '', __setDbLinkRichTextSegments_: () => {},
      __sortOptionalSheetBySubmissionDate05c_: () => {}
    });
    const index = { claim_number: 0, claim_last_status_name: 1, days_aging_from_last_activity: 2, holder_name: 3, partner_name: 4 };
    const raw = [[token + '-001', 'NEW', 0, 'Changed owner', ''], [token + '-002', 'NEW', 2, 'New owner', ''], [token + '-002', 'DUPLICATE', 9, '', '']];
    const process = () => name === 'Doss' ? api.processDoss_({ getSheetByName: () => sh }, raw, index, 'SUB') : api.processEVBike_({ getSheetByName: () => sh }, raw, index, 'SUB');
    process();
    assert.equal(sh.rows.length, 3);
    assert.equal(sh.rows[1][1], 'NEW');
    assert.equal(sh.rows[1][2], 0);
    assert.deepEqual(sh.rows[1].slice(3), before.slice(3));
    assert.equal(sh.rows[2][0], token + '-002');
    assert.equal(sh.rows[2][7], 'New owner');
    assert.ok(sh.writes.filter(w => w.row === 2).every(w => w.width === 1 && [2, 3].includes(w.col)));
    process();
    assert.equal(sh.rows.length, 3, 'rerun must not append duplicates');
    raw[0][1] = ''; raw[0][2] = '';
    process();
    assert.equal(sh.rows[1][1], 'NEW', 'blank source preserves prior status');
    assert.equal(sh.rows[1][2], 0);
  }
});

function flags(sh, extras = {}) {
  const empty = () => new Set();
  const policy = Object.fromEntries(['expired', 'flex', 'b2b', 'duplicate', 'secondYear', 'firstMonthPolicy', 'remaining1Month', 'migrationPolicy'].map(k => [k, { bg: '#d9ead3', note: k === 'flex' ? 'Flex claim.' : k === 'b2b' ? 'B2B claim.' : k }]));
  policy.priority = ['secondYear', 'flex', 'b2b'];
  return loadFunctions(routingSource, ['applyOperationalClaimHighlightsByRaw_', '__writeFlagMatrixBatched05b_'], {
    CONFIG: { sheetsByPic: { picOperational: ['Submission'] } }, RUNTIME: { flowName: 'main' },
    __isDryRun05b__: () => false, resolveHeaderIndexByAliases_: () => 0, getClaimNumberHeaderAliases_: () => ['Claim Number'],
    getOperationalClaimHighlightPolicy_: () => policy, normalizeColor_: c => String(c || '').toLowerCase(),
    __sanitizeSheetFillColor05b_: c => c, __loadClaimedActivePolicyMap05b_: () => new Map(),
    buildOperationalClaimHighlightSetsFromRaw_: () => ({ expired: empty(), flex: empty(), b2b: empty(), secondYear: new Set(['A', 'B', 'C']), policyByClaim: new Map() }),
    ...extras
  });
}

test('flagging checkpoints bounded batches and resumes without skipping rows', () => {
  const sh = sheet('Submission', [['Claim Number'], ['A'], ['B'], ['C']]);
  const api = flags(sh);
  let cursor;
  const opts = { batchSize: 2, deadline: Date.now() + 10000, strict: true,
    onCheckpoint(value) { cursor = value; opts.deadline = 1; } };
  const first = api.applyOperationalClaimHighlightsByRaw_({ getSheetByName: () => sh }, [], {}, 'Master', opts);
  assert.equal(first.complete, false);
  assert.equal(cursor.row, 4);
  assert.equal(sh.notes[3][0], '');
  const second = api.applyOperationalClaimHighlightsByRaw_({ getSheetByName: () => sh }, [], {}, 'Master', { cursor, batchSize: 2, strict: true });
  assert.equal(second.complete, true);
  assert.ok(sh.notes.slice(1).every(r => r[0] === 'secondYear'));
  assert.ok(sh.colors.slice(1).every(r => r[0] === '#d9ead3'));
});

test('a failed note write still applies colors and retains the failed batch for retry', () => {
  const sh = sheet('Submission', [['Claim Number'], ['A']]);
  const getRange = sh.getRange.bind(sh);
  sh.getRange = (...args) => { const range = getRange(...args); range.failNotes = true; return range; };
  const api = flags(sh);
  let checkpointed = false;
  assert.throws(() => api.applyOperationalClaimHighlightsByRaw_({ getSheetByName: () => sh }, [], {}, 'Master', { strict: true, onCheckpoint() { checkpointed = true; } }), /Flag batch/);
  assert.equal(sh.colors[1][0], '#d9ead3');
  assert.equal(checkpointed, false);
});

test('failed colors do not suppress notes and fallback calls remain bounded', () => {
  const sh = sheet('Submission', [['Claim Number'], ...Array.from({ length: 20 }, () => ['A'])]);
  const getRange = sh.getRange.bind(sh);
  let attempts = 0;
  sh.getRange = (...args) => {
    const range = getRange(...args);
    range.setBackgrounds = () => { attempts++; throw new Error('background write failed'); };
    return range;
  };
  assert.throws(() => flags(sh).applyOperationalClaimHighlightsByRaw_({ getSheetByName: () => sh }, [], {}, 'Master', { strict: true }), /Flag batch/);
  assert.ok(sh.notes.slice(1).every(r => r[0] === 'secondYear'));
  assert.ok(attempts <= 8);
});

test('stale combined Flex/B2B notes and marker fills are removed, manual notes survive', () => {
  const sh = sheet('Submission', [['Claim Number'], ['A'], ['B']]);
  sh.notes[1] = ['Flex claim.\n\nB2B claim.']; sh.colors[1] = ['#d9ead3'];
  sh.notes[2] = ['Manual note']; sh.colors[2] = ['#123456'];
  const api = flags(sh, { buildOperationalClaimHighlightSetsFromRaw_: () => ({ expired: new Set(), flex: new Set(), b2b: new Set(), policyByClaim: new Map() }) });
  api.applyOperationalClaimHighlightsByRaw_({ getSheetByName: () => sh }, [], {}, 'Master');
  assert.equal(sh.notes[1][0], '');
  assert.equal(sh.colors[1][0], null);
  assert.equal(sh.notes[2][0], 'Manual note');
  assert.equal(sh.colors[2][0], '#123456');
});

function continuation(clock) {
  let stored;
  let deleted = false;
  const triggers = [];
  const props = { setProperty: (_key, value) => { stored = JSON.parse(value); }, deleteProperty: () => { deleted = true; } };
  const api = loadFunctions(pipelineSource, ['runMainStage2Steps06b_'], {
    Date: { now: () => clock.time }, armMainPipelineStage2Trigger06b_: delay => triggers.push(delay),
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {} }, setProgress_() {}, logLine_() {}
  });
  return { api, props, triggers, state: () => stored, deleted: () => deleted };
}

test('MAIN continuation resumes after completed steps without clearing/routing again', () => {
  const clock = { time: 0 };
  const env = continuation(clock);
  const calls = [];
  const steps = [{ name: 'ROUTE', run() { calls.push('ROUTE'); clock.time = 100; } }, { name: 'HIGHLIGHT', run() { calls.push('HIGHLIGHT'); } }];
  const first = env.api.runMainStage2Steps06b_({}, steps, env.props, 50);
  assert.equal(first.staged, true);
  assert.equal(env.state().nextStep, 1);
  assert.equal(env.deleted(), false);
  assert.equal(env.triggers[0], 7 * 60 * 1000, 'watchdog must precede work');
  env.api.runMainStage2Steps06b_(env.state(), steps, env.props, 1000);
  assert.deepEqual(calls, ['ROUTE', 'HIGHLIGHT']);
  assert.equal(env.deleted(), true);
});

test('MAIN mandatory flag failure preserves checkpoint and suppresses cleanup', () => {
  const env = continuation({ time: 1 });
  let cleaned = false;
  const steps = [{ name: 'HIGHLIGHT', run() { throw new Error('write unavailable'); } }, { name: 'TRASH', run() { cleaned = true; } }];
  assert.throws(() => env.api.runMainStage2Steps06b_({}, steps, env.props, 100), /write unavailable/);
  assert.equal(env.deleted(), false);
  assert.equal(cleaned, false);
  assert.match(env.state().lastError, /HIGHLIGHT/);
  assert.throws(() => env.api.runMainStage2Steps06b_({ stepAttempts: 3 }, steps, env.props, 100), /Retry limit/);
});

test('SUB generic enrichment excludes optional sheets and waits for pending MAIN', () => {
  const source = fs.readFileSync('06a_EntryPoints.gs', 'utf8');
  let ran = false;
  let marked = false;
  const api = loadFunctions(source, ['__updateOperationalSheetsFromRaw06a_', '__withTryLockSub06a_'], {
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'pending' }) },
    withTryScriptLock_: (_timeout, fn) => ({ acquired: true, result: fn() }),
    __markSubPendingAfterBusyLock06a_: () => { marked = true; }
  });
  api.__updateOperationalSheetsFromRaw06a_({ getSheetByName() { throw new Error('Optional sheet must not enter generic updater'); } }, ['EV-Bike', 'Doss'], new Map(), {});
  const result = api.__withTryLockSub06a_(() => { ran = true; });
  assert.equal(result.pending, true);
  assert.equal(marked, true);
  assert.equal(ran, false);
});

test('SUB check-only idempotency does not consume failed refreshes', () => {
  let cache = {};
  const api = loadFunctions(fs.readFileSync('01_Utils.gs', 'utf8'), ['checkAndMarkTransaction_'], {
    __loadTxnCache_: () => ({ ...cache }), __saveTxnCache_: value => { cache = value; }
  });
  assert.equal(api.checkAndMarkTransaction_('SUB-test', 60000, { checkOnly: true }).duplicate, false);
  assert.deepEqual(cache, {});
  assert.equal(api.checkAndMarkTransaction_('SUB-test', 60000).duplicate, false);
  assert.equal(api.checkAndMarkTransaction_('SUB-test', 60000, { checkOnly: true }).duplicate, true);
});

test('SUB other operational sheets retain their broader update contract', () => {
  const source = fs.readFileSync('06a_EntryPoints.gs', 'utf8');
  const api = loadFunctions(source, ['__updateOperationalSheetsFromRaw06a_', '__normalizeClaimKeySub06a_'], {
    __getScFallbackSheet06a_: () => 'SC - Unmapped', isDryRun_: () => false,
    safeSetValues_: (range, values) => range.setValues(values)
  });
  for (const name of ['Submission', 'Start', 'SC - Farhan']) {
    const sh = sheet(name, [['Claim Number', 'Last Status', 'Last Status Aging', 'Service Center', 'Activity Log'], ['C-001', 'OLD', 8, 'Old SC', 'Old activity']]);
    const rawMap = new Map([['C-001', { last_status: 'NEW', last_status_aging: 0, sc_name: 'New SC', activity_log: 'New activity' }]]);
    api.__updateOperationalSheetsFromRaw06a_({ getSheetByName: () => sh }, [name], rawMap, {});
    assert.deepEqual(sh.rows[1], ['C-001', 'NEW', 0, 'New SC', 'New activity'], name);
  }
});

test('stage 2 resumes flagging and delays cleanup and SUB until complete', () => {
  let stored = { token: 'run', profile: 'Master', rawSheetId: 42, rawRows: 1, createdAt: Date.now() };
  let payload = JSON.stringify(stored);
  const props = { getProperty: () => payload, setProperty: (_key, value) => { payload = value; stored = JSON.parse(value); }, deleteProperty: () => { payload = null; } };
  const calls = [];
  const raw = { getSheetId: () => 42, getLastRow: () => 2, getLastColumn: () => 1, getRange: () => ({ getValues: () => [['C-001']] }) };
  const context = {
    PropertiesService: { getScriptProperties: () => props }, withLock_: fn => fn(),
    RUNTIME: {}, CONFIG: { spreadsheets: { Master: 'test' } },
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => raw }) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ after: () => ({ create() {} }) }) }) },
    PIPELINE_FLAGS: { TRASH_UPLOADED_FILES: true },
    resolveSpreadsheetKey_: () => 'Master', resetRunState_() {}, setLogRunContext_() {}, logLine_() {}, setProgress_() {},
    getRawHeader_: () => ['claim_number'], buildHeaderIndex_: () => ({ claim_number: 0 }), preflightRoutableCount_: () => ({ total: 1 }),
    clearOperationalSheets_: () => calls.push('CLEAR'), routeRawToOperationalSheetsInMemory_: () => { calls.push('ROUTE'); return { total: 1 }; },
    shouldRunWeeklyReportBaseNow06b_: () => false, getOperationalSheetNames06b_: () => [],
    flushTrashQueueBestEffort_: () => calls.push('TRASH'), __drainPendingSubAfterMain06a_: () => calls.push('SUB')
  };
  for (const name of ['applyTemplateRowToOperationalSheets_', 'restoreOpsFieldsFromRawBackup_', 'restoreNamedOpsFieldsFromRaw06c_', 'applyUpdateStatusRichTextToOperational_', 'applyRemarksRichTextToOperational_', 'restoreOpsManualFromMainTempForSub06c_', 'restoreOpsManualFromBackupSheet06c_', 'enrichOperationalSheetsFromRaw06_', 'applyStrictSubmissionDateAndMonth06b_', 'autofillBranchInScSheets06_', 'applyFinishTypeInScSheets06_', 'processB2B_', 'processSpecialCase_', 'processEVBike_', 'processDoss_', 'sanitizeProblematicDataValidations06_', 'recomputeExclusionTat_', 'reorderRawDataColumns06_', 'sortOperationalSheetsPreserveFilter06b_', 'refreshReportBaseFromOperational06_']) context[name] = () => calls.push(name);
  let highlighted = false;
  context.applyOperationalClaimHighlightsByRaw_ = (_ss, _rows, _index, _profile, opts) => {
    opts.onCheckpoint({ sheet: 0, row: 502 });
    if (!highlighted) { highlighted = true; return { complete: false }; }
    assert.equal(opts.cursor.row, 502);
    return { complete: true };
  };
  const api = loadFunctions(pipelineSource, ['armMainPipelineStage2Trigger06b_', 'runMainStage2Steps06b_', 'runMainPipelineStage2_'], context);
  assert.equal(api.runMainPipelineStage2_().staged, true);
  assert.equal(calls.includes('TRASH'), false);
  assert.equal(calls.includes('SUB'), false);
  assert.equal(payload !== null, true);
  api.runMainPipelineStage2_();
  assert.equal(calls.filter(c => c === 'CLEAR').length, 1);
  assert.equal(calls.filter(c => c === 'ROUTE').length, 1);
  assert.deepEqual(calls.slice(-2), ['TRASH', 'SUB']);
  assert.equal(payload, null);
});
