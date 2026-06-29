#!/usr/bin/env node
'use strict';

/**
 * Local static smoke-check for Apps Script sources (no external deps).
 * - Loads project source files in deterministic order.
 * - Provides minimal GAS stubs so top-level eval can run in Node.
 * - Executes runSelfCheck_() when available.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = process.cwd();
const SOURCE_FILES = [
  '00_Config.gs',
  '01_Utils.gs',
  '02_LogAndDetails.gs',
  '03_SheetsAndValidation.gs',
  '04_ParseAndAging.gs',
  '05a_Pipeline_RawMutate_Backup.gs',
  '05b_Pipeline_RoutingOperational.gs',
  '05c_Pipeline_OptionalSheets.gs',
  '06a_EntryPoints.gs',
  '06b_PipelineAndEnrichment.gs',
  '06c_PostProcessAndUtils.gs'
];

function createContext() {
  const labelStub = { getName: () => '', getThreads: () => [] };
  const threadStub = {
    getId: () => 'thread',
    getMessages: () => [],
    removeLabel: () => {},
    moveToTrash: () => {},
    markRead: () => {}
  };

  const context = {
    console,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    Date,
    Math,
    JSON,
    RegExp,
    String,
    Number,
    Boolean,
    Array,
    Object,
    Map,
    Set,
    WeakMap,
    WeakSet,

    Utilities: {
      formatDate: (d) => (d instanceof Date ? d.toISOString() : String(d || '')),
      sleep: () => {},
      getUuid: () => 'uuid-smoke-check'
    },
    Logger: { log: () => {} },
    Session: { getScriptTimeZone: () => 'Asia/Jakarta' },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }),
      getDocumentProperties: () => ({ getProperty: () => null, setProperty: () => {} })
    },
    CacheService: {
      getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} })
    },
    LockService: {
      getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {}, tryLock: () => true })
    },
    MailApp: { sendEmail: () => {} },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger: () => {},
      newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ nearMinute: () => ({ create: () => {} }) }) }) })
    },
    DriveApp: {
      getFileById: () => ({ setTrashed: () => {}, getBlob: () => ({}) }),
      createFile: () => ({ getId: () => 'file-id', setTrashed: () => {} })
    },
    GmailApp: {
      search: () => [],
      getUserLabelByName: () => labelStub,
      createLabel: () => labelStub,
      getThreadById: () => threadStub
    },
    SpreadsheetApp: {
      openById: () => {
        const sheetStub = {
          getLastRow: () => 1,
          getLastColumn: () => 1,
          getRange: () => ({ getValues: () => [['']], setValues: () => {}, setNumberFormat: () => {} })
        };
        const sheets = {
          'Raw Data': sheetStub,
          'Log': sheetStub,
          'Daily': sheetStub,
          'Past': sheetStub
        };
        return {
          getId: () => 'ss-id',
          getSpreadsheetTimeZone: () => 'Asia/Jakarta',
          setSpreadsheetTimeZone: () => {},
          getSheetByName: (name) => sheets[String(name || '').trim()] || null,
          insertSheet: () => sheetStub,
          getSheets: () => Object.keys(sheets).map((n) => ({ getName: () => n }))
        };
      },
      create: () => ({ getId: () => 'ss-id', getSheets: () => [] })
    }
  };

  context.global = context;
  context.globalThis = context;
  return vm.createContext(context);
}

function loadSourceFiles(ctx) {
  for (const rel of SOURCE_FILES) {
    const file = path.join(ROOT, rel);
    if (!fs.existsSync(file)) {
      throw new Error(`Missing source file: ${rel}`);
    }
    const code = fs.readFileSync(file, 'utf8');
    vm.runInContext(code, ctx, { filename: rel, displayErrors: true });
  }
}

function runSmoke() {
  const ctx = createContext();
  loadSourceFiles(ctx);

  const hasRunSelfCheck = vm.runInContext("typeof runSelfCheck_ === 'function'", ctx);
  if (!hasRunSelfCheck) {
    throw new Error('runSelfCheck_ is not defined after source load.');
  }

  const report = vm.runInContext('runSelfCheck_()', ctx);
  if (!report || report.ok !== true) {
    const payload = JSON.stringify(report || {}, null, 2);
    throw new Error(`runSelfCheck_ reported failure:\n${payload}`);
  }

  // Regression guard:
  // 05a helper must normalize missing header to -1
  // because downstream code relies on -1 sentinel checks.
  const guard = vm.runInContext(`(function () {
    if (typeof __findHeaderIndex05a_ !== 'function') return { ok: false, reason: '__findHeaderIndex05a_ missing' };
    const miss = __findHeaderIndex05a_(['A', 'B'], 'C');
    const hit = __findHeaderIndex05a_(['A', 'B'], 'B');
    return { ok: miss === -1 && hit === 1, miss: miss, hit: hit };
  })()`, ctx);
  if (!guard || guard.ok !== true) {
    throw new Error('05a header-index regression guard failed: ' + JSON.stringify(guard || {}));
  }

  const workflowGuard = vm.runInContext(`(function () {
    function makeGridSheet(name, data) {
      return {
        name: name,
        data: data,
        getName: function () { return name; },
        getLastRow: function () { return this.data.length; },
        getLastColumn: function () { return this.data[0] ? this.data[0].length : 0; },
        getRange: function (row, col, numRows, numCols) {
          const sh = this;
          return {
            getValues: function () {
              const out = [];
              for (let r = 0; r < numRows; r++) {
                const line = [];
                for (let c = 0; c < numCols; c++) line.push((sh.data[row - 1 + r] || [])[col - 1 + c]);
                out.push(line);
              }
              return out;
            },
            setValues: function (values) {
              for (let r = 0; r < values.length; r++) {
                while (sh.data.length <= row - 1 + r) sh.data.push(new Array(sh.getLastColumn()).fill(''));
                for (let c = 0; c < values[r].length; c++) sh.data[row - 1 + r][col - 1 + c] = values[r][c];
              }
            }
          };
        }
      };
    }

    const b2b = makeGridSheet('B2B', [
      ['Claim Number', 'Last Status', 'Service Center', 'Last Status Aging', 'Stage Aging'],
      ['XY', 'INSURANCE_CLAIM_REVIEW', 'Old SC', 37, 37]
    ]);
    const b2bChanged = updateB2BStatusServiceCenterFromRaw05c_(b2b, [
      ['XY', 'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH', 'New SC', 5]
    ], {
      claim_number: 0,
      claim_last_status_name: 1,
      repairer_location_store_name: 2,
      days_aging_from_last_activity: 3
    });
    const b2bRow = b2b.data[1];
    const b2bOk = b2bChanged === 1
      && b2bRow[1] === 'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH'
      && b2bRow[2] === 'New SC'
      && b2bRow[3] === 5
      && b2bRow[4] === 0;

    const policy = getOperationalClaimHighlightPolicy_();
    const markerBg = policy.remaining1Month.bg;
    const highlightSheet = {
      getName: function () { return 'SC - Farhan'; },
      getLastRow: function () { return 2; },
      getLastColumn: function () { return 1; },
      values: [['Claim Number'], ['ABC']],
      bgs: [[markerBg]],
      notes: [['Flagging retained from MAIN\\n\\nSubmission Date : 25 May 26']],
      getRange: function (row, col, numRows, numCols) {
        const sh = this;
        return {
          getValues: function () {
            if (row === 1) return [sh.values[0].slice(0, numCols)];
            return [[sh.values[row - 1][col - 1]]];
          },
          getBackgrounds: function () { return sh.bgs.map(function (r) { return r.slice(); }); },
          setBackgrounds: function (v) { sh.bgs = v.map(function (r) { return r.slice(); }); },
          getNotes: function () { return sh.notes.map(function (r) { return r.slice(); }); },
          setNotes: function (v) { sh.notes = v.map(function (r) { return r.slice(); }); }
        };
      }
    };
    const ss = { getSheetByName: function () { return highlightSheet; } };
    applyOperationalClaimHighlightsByRaw_(ss, [['ABC']], { claim_number: 0 }, 'SUB');
    const highlightOk = typeof __shouldPreserveSubHighlight05b_ === 'function'
      && __shouldPreserveSubHighlight05b_('SUB', markerBg, 'Flagging retained from MAIN') === true
      && highlightSheet.notes[0][0].indexOf('Flagging retained from MAIN') === 0
      && normalizeColor_(highlightSheet.bgs[0][0]) === normalizeColor_(markerBg);

    const finishCloneOk = typeof __shouldKeepScRowAndCloneFinishSub06a_ === 'function'
      && __shouldKeepScRowAndCloneFinishSub06a_('SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH') === true;

    const parsedSubmissionDate = normalizeDate_('25 May 26');
    const submissionDateOk = parsedSubmissionDate
      && parsedSubmissionDate.getFullYear() === 2026
      && parsedSubmissionDate.getMonth() === 4
      && parsedSubmissionDate.getDate() === 25;

    return { ok: b2bOk && highlightOk && finishCloneOk && submissionDateOk, b2bOk, highlightOk, finishCloneOk, submissionDateOk, b2bRow: b2bRow, bg: highlightSheet.bgs[0][0], note: highlightSheet.notes[0][0] };
  })()`, ctx);
  if (!workflowGuard || workflowGuard.ok !== true) {
    throw new Error('MAIN/SUB workflow regression guard failed: ' + JSON.stringify(workflowGuard || {}, null, 2));
  }

  console.log('✅ static_smoke_check: PASS');
  console.log(JSON.stringify({ ok: report.ok, warnings: report.warnings || [], summary: report.summary || {}, guard: guard, workflowGuard: workflowGuard }, null, 2));
}

try {
  runSmoke();
} catch (err) {
  console.error('❌ static_smoke_check: FAIL');
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
}
