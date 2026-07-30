/**
 * SC-Meilani
 *
 * Manual Google Sheets menu for mirroring filtered Salvage and Salvage Repair data
 * from source workbook into the active destination workbook.
 *
 * Menu:
 * - SC Meilani > Salvage
 * - SC Meilani > Salvage Repair
 */

const SC_MEILANI_CONFIG = Object.freeze({
  sourceSpreadsheetId: '1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc',
  menuName: 'SC Meilani',
  logSheetName: 'Log SC-Meilani',
  headerRow: 1,
  allowedBranches: Object.freeze([
    'Xiaomi Authorized',
    'Unicom',
    'Samsung Exclusive',
  ]),
  salvage: Object.freeze({
    sourceSheetName: 'Salvage 25-26',
    yosHeader: 'YoS',
    remarksHeader: 'Remarks',
    requiredRemarksValue: 'Unit belum ada',
    targets: Object.freeze([
      Object.freeze({ year: 2025, sheetName: 'Salvage TLO/BER 2025' }),
      Object.freeze({ year: 2026, sheetName: 'Salvage TLO/BER 2026' }),
    ]),
  }),
  repair: Object.freeze({
    sourceSheetName: 'Raw NEW',
    targetSheetName: 'Pickup Sparepart Repair (Salvage)',
    allowedLastStatuses: Object.freeze([
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_WALKIN',
      'SERVICE_CENTER_CLAIM_WAITING_WALKIN_FINISH',
      'SERVICE_CENTER_CLAIM_DONE',
      'SERVICE_CENTER_CLAIM_DONE_REPAIR_PICKUP',
      'SERVICE_CENTER_CLAIM_WAITING_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH',
      'COURIER_CLAIM_PICKUP_FINISH_DONE',
      'INSURANCE_CLAIM_WAITING_PAID_REPAIR',
      'INSURANCE_CLAIM_PAID_REPAIR',
      'INSURANCE_CLAIM_WAITING_PAID',
    ]),
  }),
  outputHeaders: Object.freeze([
    'Submission Date',
    'Submission Month',
    'Claim Number',
    'DB Link',
    'Approval Date',
    'Branch',
    'Service Center',
    'Insurance',
    'Sum Insured',
    'Device Brand',
    'Device Type',
    'IMEI/SN',
    'Last Status',
  ]),
  headerAliases: Object.freeze({
    'Submission Date': Object.freeze(['Submission Date', 'Submitted Datetime', 'Submitted Date', 'claim_submitted_at', 'created_at']),
    'Submission Month': Object.freeze(['Submission Month', 'Submitted Month', 'Month']),
    'Claim Number': Object.freeze(['Claim Number', 'Claim No', 'Claim', 'claim_number']),
    'DB Link': Object.freeze(['DB Link', 'Dashboard Link', 'Link']),
    'Approval Date': Object.freeze(['Approval Date', 'Approved Date', 'Insurance Approval Date', 'approval_date']),
    'Branch': Object.freeze(['Branch', 'Service Center Branch', 'SC Branch', 'branch']),
    'Service Center': Object.freeze(['Service Center', 'Service Center Name', 'SC Name', 'SC', 'service_center_name']),
    'Insurance': Object.freeze(['Insurance', 'Insurance Name', 'insurance_name']),
    'Sum Insured': Object.freeze(['Sum Insured', 'SI', 'sum_insured']),
    'Device Brand': Object.freeze(['Device Brand', 'Brand', 'device_brand']),
    'Device Type': Object.freeze(['Device Type', 'Type', 'device_type']),
    'IMEI/SN': Object.freeze(['IMEI/SN', 'IMEI', 'SN', 'Serial Number', 'IMEI SN', 'imei', 'serial_number']),
    'Last Status': Object.freeze(['Last Status', 'Status Terakhir', 'claim_last_status_name', 'last_status']),
    'YoS': Object.freeze(['YoS', 'YOS', 'Year of Salvage', 'Year']),
    'Remarks': Object.freeze(['Remarks', 'Remark', 'Notes', 'Note']),
  }),
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(SC_MEILANI_CONFIG.menuName)
    .addItem('Salvage', 'runSCMeilaniSalvage')
    .addItem('Salvage Repair', 'runSCMeilaniSalvageRepair')
    .addSeparator()
    .addItem('Run All', 'runSCMeilaniAll')
    .addToUi();
}

function runSCMeilaniAll() {
  runSCMeilaniSalvage();
  runSCMeilaniSalvageRepair();
}

function runSCMeilaniSalvage() {
  return scMeilaniWithLock_('SALVAGE', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    const sourceSheet = scMeilaniRequireSheet_(ctx.sourceSpreadsheet, cfg.salvage.sourceSheetName);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet);
    const branchCol = scMeilaniRequireHeader_(sourceMeta.headerMap, 'Branch');
    const yosCol = scMeilaniRequireHeader_(sourceMeta.headerMap, cfg.salvage.yosHeader);
    const remarksCol = scMeilaniRequireHeader_(sourceMeta.headerMap, cfg.salvage.remarksHeader);

    const results = [];
    cfg.salvage.targets.forEach(function (target) {
      const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, target.sheetName);
      const targetMeta = scMeilaniReadTargetMeta_(targetSheet);
      const rowNumbers = scMeilaniCollectRowNumbers_(sourceMeta.values, function (row) {
        return scMeilaniIsAllowedBranch_(row[branchCol])
          && scMeilaniMatchesYear_(row[yosCol], target.year)
          && scMeilaniEqualsText_(row[remarksCol], cfg.salvage.requiredRemarksValue);
      });

      const written = scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, rowNumbers, cfg.outputHeaders);
      results.push(target.sheetName + ': ' + written + ' row');
    });

    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage selesai. ' + results.join(' | '));
    return results;
  });
}

function runSCMeilaniSalvageRepair() {
  return scMeilaniWithLock_('SALVAGE_REPAIR', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    const sourceSheet = scMeilaniRequireSheet_(ctx.sourceSpreadsheet, cfg.repair.sourceSheetName);
    const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, cfg.repair.targetSheetName);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet);
    const targetMeta = scMeilaniReadTargetMeta_(targetSheet);
    const branchCol = scMeilaniRequireHeader_(sourceMeta.headerMap, 'Branch');
    const lastStatusCol = scMeilaniRequireHeader_(sourceMeta.headerMap, 'Last Status');
    const allowedStatuses = scMeilaniToKeySet_(cfg.repair.allowedLastStatuses);

    const rowNumbers = scMeilaniCollectRowNumbers_(sourceMeta.values, function (row) {
      return scMeilaniIsAllowedBranch_(row[branchCol])
        && allowedStatuses[scMeilaniStatusKey_(row[lastStatusCol])] === true;
    });

    const written = scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, rowNumbers, cfg.outputHeaders);
    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage Repair selesai. ' + written + ' row ditulis.');
    return written;
  });
}

function scMeilaniWithLock_(flowName, runner) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Run dibatalkan: proses lain masih berjalan.');
  }

  const startedAt = new Date();
  const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSpreadsheet = SpreadsheetApp.openById(SC_MEILANI_CONFIG.sourceSpreadsheetId);
  const ctx = {
    flowName: flowName,
    destinationSpreadsheet: destinationSpreadsheet,
    sourceSpreadsheet: sourceSpreadsheet,
    startedAt: startedAt,
  };

  try {
    scMeilaniToast_(destinationSpreadsheet, flowName + ' berjalan...');
    const result = runner(ctx);
    scMeilaniLog_(destinationSpreadsheet, flowName, 'SUCCESS', startedAt, 'OK', JSON.stringify(result));
    return result;
  } catch (err) {
    scMeilaniLog_(destinationSpreadsheet, flowName, 'ERROR', startedAt, scMeilaniErrorMessage_(err), '');
    scMeilaniToast_(destinationSpreadsheet, flowName + ' gagal: ' + scMeilaniErrorMessage_(err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function scMeilaniReadSourceMeta_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Source sheet kosong: ' + sheet.getName());

  const headerValues = values[SC_MEILANI_CONFIG.headerRow - 1] || [];
  return {
    sheet: sheet,
    values: values,
    headerValues: headerValues,
    headerMap: scMeilaniBuildHeaderMap_(headerValues),
  };
}

function scMeilaniReadTargetMeta_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), SC_MEILANI_CONFIG.outputHeaders.length, 1);
  const headerValues = sheet.getRange(SC_MEILANI_CONFIG.headerRow, 1, 1, lastColumn).getValues()[0];
  const headerMap = scMeilaniBuildHeaderMap_(headerValues);

  SC_MEILANI_CONFIG.outputHeaders.forEach(function (header, index) {
    if (scMeilaniFindHeaderIndex_(headerMap, header) == null) {
      sheet.getRange(SC_MEILANI_CONFIG.headerRow, index + 1).setValue(header);
    }
  });

  const refreshedHeaders = sheet.getRange(SC_MEILANI_CONFIG.headerRow, 1, 1, Math.max(sheet.getLastColumn(), SC_MEILANI_CONFIG.outputHeaders.length)).getValues()[0];
  return {
    sheet: sheet,
    headerValues: refreshedHeaders,
    headerMap: scMeilaniBuildHeaderMap_(refreshedHeaders),
  };
}

function scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, sourceRowNumbers, outputHeaders) {
  scMeilaniValidateMirrorHeaders_(sourceMeta, targetMeta, outputHeaders);
  scMeilaniClearBody_(targetSheet);
  if (!sourceRowNumbers.length) return 0;

  const targetStartRow = SC_MEILANI_CONFIG.headerRow + 1;
  const sourceCols = outputHeaders.map(function (header) {
    return scMeilaniRequireHeader_(sourceMeta.headerMap, header) + 1;
  });
  const targetCols = outputHeaders.map(function (header) {
    return scMeilaniRequireHeader_(targetMeta.headerMap, header) + 1;
  });

  for (let i = 0; i < sourceRowNumbers.length; i++) {
    for (let h = 0; h < outputHeaders.length; h++) {
      const sourceCell = sourceSheet.getRange(sourceRowNumbers[i], sourceCols[h]);
      const targetCell = targetSheet.getRange(targetStartRow + i, targetCols[h]);
      sourceCell.copyTo(targetCell, { contentsOnly: false });
    }
  }

  return sourceRowNumbers.length;
}

function scMeilaniValidateMirrorHeaders_(sourceMeta, targetMeta, outputHeaders) {
  const missingSource = [];
  const missingTarget = [];

  outputHeaders.forEach(function (header) {
    if (scMeilaniFindHeaderIndex_(sourceMeta.headerMap, header) == null) missingSource.push(header);
    if (scMeilaniFindHeaderIndex_(targetMeta.headerMap, header) == null) missingTarget.push(header);
  });

  if (missingSource.length) {
    throw new Error('Header source tidak ditemukan di "' + sourceMeta.sheet.getName() + '": ' + missingSource.join(', '));
  }
  if (missingTarget.length) {
    throw new Error('Header target tidak ditemukan di "' + targetMeta.sheet.getName() + '": ' + missingTarget.join(', '));
  }
}

function scMeilaniClearBody_(sheet) {
  const headerRow = SC_MEILANI_CONFIG.headerRow;
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow <= headerRow) return;
  sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastColumn).clear();
}

function scMeilaniCollectRowNumbers_(values, predicate) {
  const out = [];
  for (let r = SC_MEILANI_CONFIG.headerRow; r < values.length; r++) {
    const row = values[r] || [];
    if (predicate(row)) out.push(r + 1);
  }
  return out;
}

function scMeilaniBuildHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach(function (header, index) {
    const key = scMeilaniHeaderKey_(header);
    if (!key || map[key] != null) return;
    map[key] = index;
  });
  return map;
}

function scMeilaniRequireHeader_(headerMap, canonicalHeader) {
  const idx = scMeilaniFindHeaderIndex_(headerMap, canonicalHeader);
  if (idx == null) throw new Error('Header wajib tidak ditemukan: ' + canonicalHeader);
  return idx;
}

function scMeilaniFindHeaderIndex_(headerMap, canonicalHeader) {
  const aliases = SC_MEILANI_CONFIG.headerAliases[canonicalHeader] || [canonicalHeader];
  for (let i = 0; i < aliases.length; i++) {
    const idx = headerMap[scMeilaniHeaderKey_(aliases[i])];
    if (idx != null) return idx;
  }
  return null;
}

function scMeilaniRequireSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  return sheet;
}

function scMeilaniIsAllowedBranch_(value) {
  const branch = scMeilaniNormalizeText_(value);
  const allowed = SC_MEILANI_CONFIG.allowedBranches;
  for (let i = 0; i < allowed.length; i++) {
    if (branch === scMeilaniNormalizeText_(allowed[i])) return true;
  }
  return false;
}

function scMeilaniMatchesYear_(value, year) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getFullYear() === Number(year);
  const text = String(value == null ? '' : value).trim();
  return text === String(year) || new RegExp('(^|\\D)' + year + '(\\D|$)').test(text);
}

function scMeilaniEqualsText_(value, expected) {
  return scMeilaniNormalizeText_(value) === scMeilaniNormalizeText_(expected);
}

function scMeilaniToKeySet_(values) {
  const out = {};
  (values || []).forEach(function (value) {
    out[scMeilaniStatusKey_(value)] = true;
  });
  return out;
}

function scMeilaniStatusKey_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function scMeilaniHeaderKey_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function scMeilaniNormalizeText_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function scMeilaniToast_(spreadsheet, message) {
  try {
    spreadsheet.toast(String(message || ''), SC_MEILANI_CONFIG.menuName, 8);
  } catch (err) {
    Logger.log(message);
  }
}

function scMeilaniLog_(spreadsheet, flowName, status, startedAt, message, details) {
  const sheet = scMeilaniGetOrCreateLogSheet_(spreadsheet);
  const endedAt = new Date();
  sheet.appendRow([
    endedAt,
    flowName,
    status,
    startedAt,
    endedAt,
    Math.round((endedAt.getTime() - startedAt.getTime()) / 1000),
    message,
    details || '',
  ]);
}

function scMeilaniGetOrCreateLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SC_MEILANI_CONFIG.logSheetName);
  if (sheet) return sheet;

  sheet = spreadsheet.insertSheet(SC_MEILANI_CONFIG.logSheetName);
  sheet.getRange(1, 1, 1, 8).setValues([[
    'Timestamp',
    'Flow',
    'Status',
    'Started At',
    'Ended At',
    'Duration Seconds',
    'Message',
    'Details',
  ]]);
  sheet.setFrozenRows(1);
  return sheet;
}

function scMeilaniErrorMessage_(err) {
  return err && err.message ? err.message : String(err);
}
