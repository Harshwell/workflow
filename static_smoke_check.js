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
  '00_Config',
  '01_Utils',
  '02_LogAndDetails',
  '03_SheetsAndValidation',
  '04_ParseAndAging',
  '05a_Pipeline_RawMutate_Backup',
  '05b_Pipeline_RoutingOperational',
  '05c_Pipeline_OptionalSheets',
  '06a_EntryPoints.gs',
  '06a_EntryPoints',
  '06b_PipelineAndEnrichment',
  '06c_PostProcessAndUtils',
  '06d_IntegratedMaintenance',
  '06e_SubHelpers.gs',
  '06f_RuntimeAssertions.gs'
  '06e_SubHelpers',
  '06f_RuntimeAssertions'
  '06e_SubHelpers'
  '06d_IntegratedMaintenance'
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

  console.log('✅ static_smoke_check: PASS');
  console.log(JSON.stringify({ ok: report.ok, warnings: report.warnings || [], summary: report.summary || {} }, null, 2));
}

try {
  runSmoke();
} catch (err) {
  console.error('❌ static_smoke_check: FAIL');
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
}
