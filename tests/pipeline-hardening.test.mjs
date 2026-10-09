import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { evaluateInitializer, loadFunctions } from '../scripts/lib/source-contracts.mjs';

const optionalSource = fs.readFileSync('05c_Pipeline_OptionalSheets.gs', 'utf8');
const routingSource = fs.readFileSync('05b_Pipeline_RoutingOperational.gs', 'utf8');
const pipelineSource = fs.readFileSync('06b_PipelineAndEnrichment.gs', 'utf8');

function sheet(name, rows) {
  const writes = [];
  const notes = rows.map(() => ['']);
  const colors = rows.map(() => ['#ffffff']);
  return {
    rows, writes, notes, colors, getName: () => name,
    getMaxRows: () => 1000, getLastRow: () => rows.length, getLastColumn: () => rows[0].length,
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
  for (const [name, token] of [['EV-Bike', 'VVMAR'], ['Doss', 'DOSS'], ['TPL', 'CAH8'], ['Drone', 'HDSH']]) {
    const header = ['Claim Number', 'Last Status', 'Last Status Aging', 'Status', 'Remarks', 'DB Link', 'TAT', 'Owner Name'];
    const sh = sheet(name, [header, [token + '-001', 'OLD', 8, 'Manual', 'Keep', 'rich-link', '=1+1', 'Original']]);
    const before = sh.rows[1].slice();
    const api = loadFunctions(optionalSource, ['updateTokenOptionalSheetFromSubRaw05c_', 'processEVBike_', 'processDoss_', 'processTPL_', 'processDrone_', 'matchesVerifiedOptionalRow05c_'], {
      VERIFIED_OPTIONAL_SHEET_POLICY: evaluateInitializer(fs.readFileSync('00_Config.gs', 'utf8'), 'VERIFIED_OPTIONAL_SHEET_POLICY'),
      RUNTIME: { flowName: 'sub' }, CONFIG: { headers: { claimNumber: 'claim_number', lastStatus: 'claim_last_status_name', businessPartner: 'partner_name' }, patterns: { evBikePartners: ['Ofero'] } },
      __getHeaderRow05c_: () => header, buildHeaderIndex_: h => Object.fromEntries(h.map((v, i) => [v, i])),
      __isDryRun05c__: () => false, safeSetValues_: (range, values) => range.setValues(values),
      normalizeInt_: value => Number(value), mapInsuranceShort_: value => value,
      buildSubmissionDateCell_: () => '', __setDbLinkRichTextSegments_: () => {},
      __sortOptionalSheetBySubmissionDate05c_: () => {}
    });
    const index = { claim_number: 0, claim_last_status_name: 1, days_aging_from_last_activity: 2, holder_name: 3, partner_name: 4, partner_code: 5 };
    const raw = [[token + '-001', 'NEW', 0, 'Changed owner', '', token], [token + '-002', 'NEW', 2, 'New owner', '', token], [token + '-002', 'DUPLICATE', 9, '', '', token]];
    const process = () => name === 'TPL' ? api.processTPL_({ getSheetByName: () => sh }, raw, index, 'SUB') : name === 'Drone' ? api.processDrone_({ getSheetByName: () => sh }, raw, index, 'SUB') : name === 'Doss' ? api.processDoss_({ getSheetByName: () => sh }, raw, index, 'SUB') : api.processEVBike_({ getSheetByName: () => sh }, raw, index, 'SUB');
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

test('TPL and Drone accept each verification independently and reject unrelated claims', () => {
  const api = loadFunctions(optionalSource, ['matchesVerifiedOptionalRow05c_'], {
    VERIFIED_OPTIONAL_SHEET_POLICY: evaluateInitializer(fs.readFileSync('00_Config.gs', 'utf8'), 'VERIFIED_OPTIONAL_SHEET_POLICY')
  });
  const cases = [
    ['TPL', 'business_partner_name', 'Cahaya ID'], ['TPL', 'partner_name', 'Cahaya ID'], ['TPL', 'partner_code', 'CAH8'],
    ['Drone', 'product_name', 'Qoala Drone 12 Month'], ['Drone', 'device_brand', 'DJI'],
    ['Drone', 'device_type', 'DJI FLIP (DJI RC 2 GL)'], ['Drone', 'device_type', 'CAMERA DRONE'],
    ['Drone', 'device_type', 'DJI LITO X1 - CAMERA DRONE'], ['Drone', 'partner_code', 'HDSH'],
    ['Drone', 'repairer_location_store_name', 'Skylensindo Service Center'], ['Drone', 'sc_name', 'Skylensindo Service Center']
  ];
  for (const [name, header, value] of cases) {
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, [value.toLowerCase()], { [header]: 0 }), true, header + ':' + value);
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, [''], { [header]: 0 }), false);
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, ['unrelated'], { [header]: 0 }), false);
  }
  for (const name of ['TPL', 'Drone']) {
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, ['CAH8-HDSH'], { claim_number: 0 }), false);
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, ['CAH8'], { insurance_partner_code: 0 }), false);
    assert.equal(api.matchesVerifiedOptionalRow05c_(name, [], {}), false);
  }
  assert.equal(api.matchesVerifiedOptionalRow05c_('Drone', ['HDSHX'], { partner_code: 0 }), false);
  assert.equal(api.matchesVerifiedOptionalRow05c_('TPL', ['CAH80'], { partner_code: 0 }), false);
  assert.equal(api.matchesVerifiedOptionalRow05c_('Drone', ['Drone', 'DJI', 'Skylensindo'], { product_name: 0, device_type: 1, sc_name: 2 }), true);
});

test('MAIN TPL and Drone upsert matching Raw claims with aging and preserve manual columns', () => {
  for (const [name, code, processor] of [['TPL', 'CAH8', 'processTPL_'], ['Drone', 'HDSH', 'processDrone_']]) {
    const header = ['Claim Number', 'Last Status', 'Last Status Aging', 'Status', 'Remarks', 'Owner Name'];
    const sh = sheet(name, [header, ['C-001', 'OLD', 8, 'Manual', 'Keep', 'Original']]);
    const api = loadFunctions(optionalSource, ['processEVBike_', 'processTPL_', 'processDrone_', 'matchesVerifiedOptionalRow05c_'], {
      VERIFIED_OPTIONAL_SHEET_POLICY: evaluateInitializer(fs.readFileSync('00_Config.gs', 'utf8'), 'VERIFIED_OPTIONAL_SHEET_POLICY'),
      RUNTIME: { flowName: 'main', enableEvBike: false },
      CONFIG: { headers: { claimNumber: 'claim_number', lastStatus: 'claim_last_status_name', businessPartner: 'business_partner_name' }, patterns: {} },
      __getHeaderRow05c_: () => header, buildHeaderIndex_: h => Object.fromEntries(h.map((v, i) => [v, i])),
      __normalizeHeaderText05c_: value => value, getSpecialCaseExcludedStatuses_: () => new Set(),
      getEvBikeExcludedPolicyNumberSet_: () => new Set(['EV-only-excluded']),
      __isDryRun05c__: () => false, safeSetValues_: (range, values) => range.setValues(values),
      normalizeInt_: value => Number(value), mapInsuranceShort_: value => value,
      buildSubmissionDateCell_: () => '', __sortOptionalSheetBySubmissionDate05c_: () => {}
    });
    const index = { claim_number: 0, claim_last_status_name: 1, days_aging_from_last_activity: 2, holder_name: 3, partner_code: 4, policy_number: 5 };
    const raw = [['C-001', 'MAIN', 0, 'Changed owner', code, 'EV-only-excluded'], ['C-002', 'NEW', 2, 'New owner', code], ['C-002', 'DUPLICATE', 9, '', code], ['VVMAR-REJECT', 'NEW', 4, 'Unrelated', 'OTHER']];
    const ss = { getSheetByName: sheetName => sheetName === name ? sh : null };
    api[processor](ss, raw, index, 'Master');
    assert.equal(sh.rows.length, 3);
    assert.deepEqual(sh.rows[1], ['C-001', 'MAIN', 0, 'Manual', 'Keep', 'Changed owner']);
    assert.equal(sh.rows[2][0], 'C-002');
    api[processor](ss, raw, index, 'Master');
    assert.equal(sh.rows.length, 3);
    assert.throws(() => api[processor]({ getSheetByName: () => null }, raw, index, 'Master'), /Optional sheet missing/);
  }
});

test('SUB optional refresh calls TPL and Drone after each OLD/NEW source and surfaces failures', () => {
  const source = fs.readFileSync('06a_EntryPoints.gs', 'utf8');
  const calls = [];
  const context = { buildHeaderIndex_: () => ({ claim_number: 0 }) };
  for (const name of ['processEVBike_', 'processDoss_', 'processTPL_', 'processDrone_']) {
    context[name] = (_ss, rows, _index, pic) => { calls.push([name, rows[0][0], pic]); return 1; };
  }
  const ss = { getSheetByName: name => ({ getLastRow: () => 2, getLastColumn: () => 1, getRange: () => ({ getValues: () => [['claim_number'], [name]] }) }) };
  const api = loadFunctions(source, ['__refreshTokenOptionalSheetsFromSubRaw06a_', '__getSubRelocationSheetNames06a_'], context);
  const result = api.__refreshTokenOptionalSheetsFromSubRaw06a_(ss, ['Raw OLD', 'Raw NEW']);
  assert.equal(result.tpl, 2); assert.equal(result.drone, 2);
  assert.deepEqual(calls.filter(c => c[0] === 'processDrone_'), [['processDrone_', 'Raw OLD', 'SUB'], ['processDrone_', 'Raw NEW', 'SUB']]);
  assert.deepEqual(Array.from(api.__getSubRelocationSheetNames06a_(['TPL', 'Drone', 'Doss', 'EV-Bike', 'Start', 'Exclusion'])), ['Start', 'Exclusion']);
  context.processDrone_ = () => { throw new Error('Drone write failed'); };
  const failed = loadFunctions(source, ['__refreshTokenOptionalSheetsFromSubRaw06a_'], context);
  assert.throws(() => failed.__refreshTokenOptionalSheetsFromSubRaw06a_(ss, ['Raw NEW']), /Drone write failed/);
});

test('SUB submission date sync cannot write into the four selective optional sheets', () => {
  const api = loadFunctions(pipelineSource, ['applyStrictSubmissionDateAndMonth06b_'], {
    RUNTIME: { flowName: 'sub' }, resolveRawIdx06_: (index, aliases) => aliases.map(name => index[name]).find(value => value != null)
  });
  api.applyStrictSubmissionDateAndMonth06b_({ getSheetByName() { throw new Error('Optional sheet must not be touched'); } },
    [['C-001', '2026-10-09']], { claim_number: 0, claim_submitted_datetime: 1 }, { sheets: ['EV-Bike', 'Doss', 'TPL', 'Drone'] });
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
  api.__updateOperationalSheetsFromRaw06a_({ getSheetByName() { throw new Error('Optional sheet must not enter generic updater'); } }, ['EV-Bike', 'Doss', 'TPL', 'Drone'], new Map(), {});
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
    sv03_getCanonicalStatusTemplateCell_() {},
    getRawHeader_: () => ['claim_number'], buildHeaderIndex_: () => ({ claim_number: 0 }), preflightRoutableCount_: () => ({ total: 1 }),
    clearOperationalSheets_: () => calls.push('CLEAR'), routeRawToOperationalSheetsInMemory_: () => { calls.push('ROUTE'); return { total: 1 }; },
    shouldRunWeeklyReportBaseNow06b_: () => false, getOperationalSheetNames06b_: () => [],
    flushTrashQueueBestEffort_: () => calls.push('TRASH'), __drainPendingSubAfterMain06a_: () => calls.push('SUB')
  };
  for (const name of ['applyTemplateRowToOperationalSheets_', 'restoreOpsFieldsFromRawBackup_', 'restoreNamedOpsFieldsFromRaw06c_', 'applyUpdateStatusRichTextToOperational_', 'applyRemarksRichTextToOperational_', 'restoreOpsManualFromMainTempForSub06c_', 'restoreOpsManualFromBackupSheet06c_', 'enrichOperationalSheetsFromRaw06_', 'applyStrictSubmissionDateAndMonth06b_', 'autofillBranchInScSheets06_', 'applyFinishTypeInScSheets06_', 'processB2B_', 'processSpecialCase_', 'processEVBike_', 'processDoss_', 'processTPL_', 'processDrone_', 'sanitizeProblematicDataValidations06_', 'recomputeExclusionTat_', 'reorderRawDataColumns06_', 'sortOperationalSheetsPreserveFilter06b_', 'refreshReportBaseFromOperational06_']) context[name] = () => calls.push(name);
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
  assert.equal(calls.filter(c => c === 'processTPL_').length, 1);
  assert.equal(calls.filter(c => c === 'processDrone_').length, 1);
  assert.deepEqual(calls.slice(-2), ['TRASH', 'SUB']);
  assert.equal(payload, null);
});

function statusTemplateFixture() {
  const source = fs.readFileSync('03_SheetsAndValidation.gs', 'utf8') + '\n' + fs.readFileSync('06c_PostProcessAndUtils.gs', 'utf8');
  const templateConfig = evaluateInitializer(fs.readFileSync('00_Config.gs', 'utf8'), 'STATUS_DROPDOWN_TEMPLATE');
  const metadata = { displayStyle: 'CHIP', optionColors: { Delivered: '#00ff00', DONE: '#123456' }, helpText: 'Status template' };
  let rule = { getCriteriaType: () => 'VALUE_IN_LIST' };
  let failCopy = false;
  const range = {
    values: [['Delivered'], [''], ['Legacy manual status']], metadata: { displayStyle: 'ARROW' },
    getA1Notation: () => 'D2:D4', getSheet: () => sh,
    clearDataValidations() { this.metadata = null; },
    setValues(values) { this.values = values.map(r => Array.from(r)); }
  };
  const sh = { getName: () => 'Finish', getParent: () => ss, getRange: () => range };
  const cell = { getDataValidation: () => rule, copyTo(dst, mode, transposed) {
    assert.equal(mode, 'DATA_VALIDATION'); assert.equal(transposed, false);
    if (failCopy) throw new Error('native copy failed');
    dst.metadata = metadata;
  }, getValue() { throw new Error('Template selection must not be read/copied'); } };
  const ss = { getSheetByName: name => name === 'Overview' ? { getRange(address) { assert.equal(address, 'C945'); return cell; } } : sh };
  const api = loadFunctions(source, ['sv03_getCanonicalStatusTemplateCell_', 'sv03_applyGeneralStatusValidationToRange_', '__restoreStatusValuesWithCanonicalValidation06c_', 'sv03_syncDropdownForWorkbook_', 'restoreOpsManualFromMainTempForSub06c_'], {
    STATUS_DROPDOWN_TEMPLATE: templateConfig,
    DRY_RUN: false, getOperationalSheetsForBackup_: () => ['Finish'],
    __normalizeHeaderText06_: value => String(value).trim(), __claimKey06_: value => String(value).trim().toUpperCase(),
    __findHeaderIndexFlexible06_: (header, name) => header.indexOf(name),
    SpreadsheetApp: { CopyPasteType: { PASTE_DATA_VALIDATION: 'DATA_VALIDATION' } },
    getWorkbookProfile_: () => 'PIC', sv03_getProfileSpec_: () => ({ operational: ['Finish'], optional: ['EV-Bike', 'Doss'] }),
    sv03_findHeaderCol1_: () => 4, SV03_DROPDOWN_SYNC: { BUFFER_ROWS: 5 }
  });
  return { api, range, sh, ss, metadata, invalidate() { rule = null; }, fail() { failCopy = true; } };
}

test('Status restore uses Overview C945 native metadata and preserves filled, blank, and legacy values', () => {
  const fixture = statusTemplateFixture();
  const values = fixture.range.values.map(r => r.slice());
  fixture.api.__restoreStatusValuesWithCanonicalValidation06c_(fixture.sh, 2, 4, values, [['C-1'], ['C-2'], ['C-3']], 'TEST');
  assert.deepEqual(fixture.range.values, values);
  assert.equal(fixture.range.metadata, fixture.metadata);
  fixture.api.sv03_applyGeneralStatusValidationToRange_(fixture.range);
  assert.deepEqual(fixture.range.values, values);
});

test('invalid Status template is rejected before restore mutates values or validation', () => {
  const fixture = statusTemplateFixture();
  fixture.invalidate();
  const before = fixture.range.values.map(r => r.slice());
  assert.throws(() => fixture.api.__restoreStatusValuesWithCanonicalValidation06c_(fixture.sh, 2, 4, [['DONE']], [['C-1']], 'TEST'), /Overview!C945/);
  assert.deepEqual(fixture.range.values, before);
  assert.deepEqual(fixture.range.metadata, { displayStyle: 'ARROW' });
});

test('native Status copy failure propagates without falling back to rebuilt dropdown', () => {
  const fixture = statusTemplateFixture();
  fixture.fail();
  assert.throws(() => fixture.api.sv03_applyGeneralStatusValidationToRange_(fixture.range), /native copy failed/);
  assert.deepEqual(fixture.range.metadata, { displayStyle: 'ARROW' });
});

test('workbook Status sync ignores row-2 templates and copies C945 to Raw, ops, and optional sheets', () => {
  const fixture = statusTemplateFixture();
  const targets = [];
  const getSheet = fixture.ss.getSheetByName;
  fixture.ss.getSheetByName = name => name === 'Overview' ? getSheet(name) : {
    getMaxRows: () => 1000, getLastRow: () => 4,
    getRange(row, col, count, width) {
      assert.equal(row, 2); assert.equal(col, 4); assert.equal(width, 1); assert.equal(count, 8);
      targets.push(name); return fixture.range;
    }
  };
  fixture.api.sv03_syncDropdownForWorkbook_(fixture.ss, 'Master', 'Status', ['old fallback'], { exactOptions: ['old fallback'] });
  assert.deepEqual(targets, ['Raw Data', 'Finish', 'EV-Bike', 'Doss']);
  assert.equal(fixture.range.metadata, fixture.metadata);
});

test('MAIN temp restore cannot replace C945 dropdown with an old arrow rule', () => {
  const fixture = statusTemplateFixture();
  fixture.range.values = [['']];
  fixture.range.getValues = () => fixture.range.values;
  Object.assign(fixture.sh, {
    getLastRow: () => 2, getLastColumn: () => 3,
    getRange(row, col) {
      if (row === 1) return { getValues: () => [['Claim Number', 'Service Center', 'Status']] };
      if (col === 1) return { getValues: () => [['C-1']] };
      if (col === 2) return { getValues: () => [['SC-1']] };
      return fixture.range;
    }
  });
  const backup = {
    getLastRow: () => 2, getLastColumn: () => 9,
    getRange(row) {
      if (row === 1) return { getValues: () => [['Claim Number', 'Service Center', 'Sheet', 'Row', 'At', 'Update Status', 'Timestamp', 'Status', 'Remarks']] };
      return {
        getValues: () => [['C-1', 'SC-1', 'Finish', 2, '', '', '', 'DONE', '']],
        copyTo(dst) { dst.values = [['DONE']]; dst.metadata = { displayStyle: 'ARROW' }; }
      };
    }
  };
  const getSheet = fixture.ss.getSheetByName;
  fixture.ss.getSheetByName = name => name === '_OPS_MAIN_SUB_TEMP' ? backup : getSheet(name);
  fixture.api.restoreOpsManualFromMainTempForSub06c_(fixture.ss, 'Master', { deleteAfterRestore: false });
  assert.deepEqual(fixture.range.values, [['DONE']]);
  assert.equal(fixture.range.metadata, fixture.metadata);
});
