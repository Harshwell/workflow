/**
 * SC-Meilani
 *
 * Manual Google Sheets menu for mirroring Salvage data and upserting Salvage Repair data
 * from source workbook into the active destination workbook.
 *
 * Menu:
 * - SC Meilani > Salvage
 * - SC Meilani > Salvage Repair
 */

const SC_MEILANI_CONFIG = Object.freeze({
  sourceSpreadsheetId: '1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc',
  destinationSpreadsheetId: '1dU9dt01Ld_ykMJWQxupvIHyLXV6ArnGeXCBV72q31lU',
  scriptVersion: '2026-08-03-salvage-batch-log-5',
  menuName: 'SC Meilani',
  logSheetName: 'Log SC-Meilani',
  headerRow: 1,
  headerScanRows: 20,
  fallbackColumns: Object.freeze({
    Branch: 6,
    YoS: 8,
    Remarks: 20,
  }),
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
    identifierHeader: 'Claim Number',
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
    'Submission Date': Object.freeze(['Submission Date', 'Submitted Datetime', 'Submitted Date', 'Claim Submitted Datetime', 'claim_submitted_datetime', 'claim_submission_date', 'claim_submitted_at', 'created_at']),
    'Submission Month': Object.freeze(['Submission Month', 'Submitted Month', 'Submission by Month', 'Month', 'claim_submission_month', 'submitted_month']),
    'Claim Number': Object.freeze(['Claim Number', 'Claim No', 'Claim', 'claim_number']),
    'DB Link': Object.freeze(['DB Link', 'Dashboard Link', 'Link', 'dashboard_link', 'db_link', 'dblink', 'database_link']),
    'Approval Date': Object.freeze(['Approval Date', 'Approved Date', 'Insurance Approval Date', 'approval_date', 'claim_approved_at', 'claim_approval_date', 'claim_last_updated_datetime']),
    'Branch': Object.freeze(['Branch', 'Service Center Branch', 'SC Branch', 'branch', 'service_center_branch', 'sc_branch']),
    'Service Center': Object.freeze(['Service Center', 'Service Center Name', 'SC Name', 'SC', 'service_center_name', 'service_center', 'sc_name']),
    'Insurance': Object.freeze(['Insurance', 'Insurance Name', 'insurance_name', 'insurer_name', 'insurance_code']),
    'Sum Insured': Object.freeze(['Sum Insured', 'SI', 'sum_insured', 'insured_value', 'sum_insured_amount']),
    'Device Brand': Object.freeze(['Device Brand', 'Brand', 'device_brand', 'brand_name', 'device_brand_name']),
    'Device Type': Object.freeze(['Device Type', 'Type', 'device_type', 'device_model']),
    'IMEI/SN': Object.freeze(['IMEI/SN', 'IMEI', 'SN', 'Serial Number', 'IMEI SN', 'imei', 'imei_number', 'serial_number', 'device_imei', 'device_imei2']),
    'Last Status': Object.freeze(['Last Status', 'Status Terakhir', 'claim_last_status_name', 'last_status', 'last_status_name']),
    'YoS': Object.freeze(['YoS', 'YOS', 'Year of Salvage', 'Year']),
    'Remarks': Object.freeze(['Remarks', 'Remark', 'Notes', 'Note']),
  }),
});

let SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_ = false;

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
  SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_ = false;
  try {
    runSCMeilaniSalvage();
    SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_ = true;
    runSCMeilaniSalvageRepair();
  } finally {
    SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_ = false;
  }
}

function runSCMeilaniSalvage() {
  return scMeilaniWithLock_('SALVAGE', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure salvage source sheet', 'START', 0, 'Checking Salvage source sheet.', 'source="' + cfg.salvage.sourceSheetName + '"', ctx.startedAt);
    const sourceSheet = scMeilaniRequireSheet_(ctx.sourceSpreadsheet, cfg.salvage.sourceSheetName);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure salvage source sheet', 'SUCCESS', 1, 'Source sheet found.', sourceSheet.getName(), ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read salvage source headers', 'START', 0, 'Reading source data and validating filter headers.', '', ctx.startedAt);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet, ['Claim Number', 'Branch', cfg.salvage.yosHeader, cfg.salvage.remarksHeader]);
    const branchCol = scMeilaniRequireSourceColumn_(sourceMeta, 'Branch', cfg.fallbackColumns.Branch);
    const yosCol = scMeilaniRequireSourceColumn_(sourceMeta, cfg.salvage.yosHeader, cfg.fallbackColumns.YoS);
    const remarksCol = scMeilaniRequireSourceColumn_(sourceMeta, cfg.salvage.remarksHeader, cfg.fallbackColumns.Remarks);
    const sourceRows = Math.max(sourceMeta.values.length - sourceMeta.headerRowNumber, 0);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read salvage source headers', 'SUCCESS', sourceRows, 'Source data loaded.', 'headerRow=' + sourceMeta.headerRowNumber + ', rows=' + sourceRows, ctx.startedAt);

    const results = [];
    cfg.salvage.targets.forEach(function (target) {
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Process target ' + target.sheetName, 'START', 0, 'Preparing target year ' + target.year + '.', '', ctx.startedAt);
      const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, target.sheetName);
      const targetMeta = scMeilaniReadTargetMeta_(targetSheet);
      const rowNumbers = scMeilaniCollectRowNumbers_(sourceMeta, function (row) {
        return scMeilaniIsAllowedBranch_(row[branchCol])
          && scMeilaniMatchesYear_(row[yosCol], target.year)
          && scMeilaniEqualsText_(row[remarksCol], cfg.salvage.requiredRemarksValue);
      });
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter target ' + target.sheetName, rowNumbers.length ? 'SUCCESS' : 'SKIPPED', rowNumbers.length, 'Filter completed for target year ' + target.year + '.', 'sourceRows=' + sourceRows + ', matched=' + rowNumbers.length, ctx.startedAt);

      const written = scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, rowNumbers, cfg.outputHeaders, ctx);
      results.push(target.sheetName + ': ' + written + ' row');
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Process target ' + target.sheetName, 'SUCCESS', written, 'Target completed.', '', ctx.startedAt);
    });

    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage selesai. ' + results.join(' | '));
    return results;
  });
}
function runSCMeilaniSalvageRepair() {
  return scMeilaniWithLock_('SALVAGE_REPAIR', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure source/target sheets', 'START', 0, 'Checking required sheets.', 'source="' + cfg.repair.sourceSheetName + '", target="' + cfg.repair.targetSheetName + '"', ctx.startedAt);
    const sourceSheet = scMeilaniRequireSheet_(ctx.sourceSpreadsheet, cfg.repair.sourceSheetName);
    const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, cfg.repair.targetSheetName);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure source/target sheets', 'SUCCESS', 2, 'Required sheets found.', 'sourceId=' + ctx.sourceSpreadsheet.getId() + ', destinationId=' + ctx.destinationSpreadsheet.getId(), ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read and validate headers', 'START', 0, 'Reading source and target headers.', '', ctx.startedAt);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet, [cfg.repair.identifierHeader, 'Branch', 'Last Status']);
    const targetMeta = scMeilaniReadTargetMeta_(targetSheet);
    const columnMap = scMeilaniResolveRepairColumnMap_(sourceMeta, targetMeta);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read and validate headers', 'SUCCESS', cfg.outputHeaders.length, 'Header validation passed.', 'sourceHeaderRow=' + sourceMeta.headerRowNumber + ', targetHeaderRow=' + targetMeta.headerRowNumber + ', optionalMissingSource=' + columnMap.optionalMissingSource.join(', '), ctx.startedAt);

    const allowedStatuses = scMeilaniToKeySet_(cfg.repair.allowedLastStatuses);
    const targetIndex = scMeilaniBuildTargetIdentifierIndex_(targetSheet, targetMeta, cfg.repair.identifierHeader);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Scan existing target', 'SUCCESS', targetIndex.uniqueCount, 'Existing target identifiers indexed.', 'duplicateTarget=' + targetIndex.duplicateCount + ', duplicateSamples=' + targetIndex.duplicateSamples.join(' | '), ctx.startedAt);

    const sourceRows = Math.max(sourceMeta.values.length - sourceMeta.headerRowNumber, 0);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter Salvage Repair source', 'START', sourceRows, 'Filtering Raw NEW rows by branch and repair last status.', 'allowedBranches=' + cfg.allowedBranches.join(', ') + ', allowedStatuses=' + cfg.repair.allowedLastStatuses.length, ctx.startedAt);
    const sourceSnapshot = scMeilaniCollectRepairRecords_(sourceMeta, columnMap, allowedStatuses, targetIndex);
    scMeilaniWriteReasonLogs_(ctx, 'Filter Salvage Repair source', 'SKIPPED', sourceSnapshot.skippedByReason);
    scMeilaniWriteReasonLogs_(ctx, 'Filter Salvage Repair source', 'FAILED', sourceSnapshot.failedByReason);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter Salvage Repair source', 'SUCCESS', sourceSnapshot.validRecords.length, 'Filtering completed.', 'sourceRows=' + sourceRows + ', valid=' + sourceSnapshot.validRecords.length + ', skipped=' + sourceSnapshot.skippedCount + ', failed=' + sourceSnapshot.failedCount, ctx.startedAt);

    const writeResult = scMeilaniUpsertRepairRecords_(targetSheet, targetMeta, columnMap, targetIndex, sourceSnapshot.validRecords, ctx);
    const verify = scMeilaniVerifyRepairResult_(targetSheet, targetMeta, cfg.repair.identifierHeader, sourceSnapshot.validRecords);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Post-write verification', verify.missingClaims.length || verify.duplicateCount ? 'FAILED' : 'SUCCESS', verify.matchedSourceClaims, 'Target verification completed.', 'targetClaims=' + verify.targetClaims + ', duplicateTargetAfter=' + verify.duplicateCount + ', missingAfter=' + verify.missingClaims.join(' | '), ctx.startedAt);

    const summary = {
      sourceRows: sourceRows,
      valid: sourceSnapshot.validRecords.length,
      skipped: sourceSnapshot.skippedCount,
      updated: writeResult.updated,
      appended: writeResult.appended,
      failed: sourceSnapshot.failedCount + writeResult.failed + verify.missingClaims.length,
      duplicateTargetSkipped: sourceSnapshot.duplicateTargetSkipped,
      targetSheetName: cfg.repair.targetSheetName,
    };

    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage Repair selesai. Valid ' + summary.valid + ', update ' + summary.updated + ', append ' + summary.appended + ', skip ' + summary.skipped + ', failed ' + summary.failed + '.');
    return summary;
  });
}
function scMeilaniWithLock_(flowName, runner) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Run dibatalkan: proses lain masih berjalan.');
  }

  const startedAt = new Date();
  let destinationSpreadsheet = null;

  try {
    scMeilaniValidateRuntime_();
    destinationSpreadsheet = SpreadsheetApp.openById(SC_MEILANI_CONFIG.destinationSpreadsheetId);
    const preservedLog = SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_;
    if (preservedLog) {
      SC_MEILANI_PRESERVE_LOG_ON_NEXT_RUN_ = false;
    } else {
      scMeilaniResetLog_(destinationSpreadsheet);
    }
    scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime bootstrap', 'START', 0, 'Run started.', 'scriptVersion=' + SC_MEILANI_CONFIG.scriptVersion + ', logMode=' + (preservedLog ? 'PRESERVED_RUN_ALL' : 'RESET'), startedAt);

    const sourceSpreadsheet = SpreadsheetApp.openById(SC_MEILANI_CONFIG.sourceSpreadsheetId);
    const ctx = {
      flowName: flowName,
      destinationSpreadsheet: destinationSpreadsheet,
      sourceSpreadsheet: sourceSpreadsheet,
      startedAt: startedAt,
    };

    scMeilaniToast_(destinationSpreadsheet, flowName + ' berjalan...');
    const result = runner(ctx);
    scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime completion', 'SUCCESS', scMeilaniResultCount_(result), 'Run completed.', JSON.stringify(result), startedAt);
    return result;
  } catch (err) {
    if (destinationSpreadsheet) {
      scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime failure', 'FAILED', 0, 'Run failed.', scMeilaniErrorMessage_(err), startedAt);
      scMeilaniToast_(destinationSpreadsheet, flowName + ' gagal: ' + scMeilaniErrorMessage_(err));
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}
function scMeilaniValidateRuntime_() {
  if (!String(SC_MEILANI_CONFIG.sourceSpreadsheetId || '').trim()) {
    throw new Error('Source spreadsheet ID kosong.');
  }
  if (!String(SC_MEILANI_CONFIG.destinationSpreadsheetId || '').trim()) {
    throw new Error('Destination spreadsheet ID kosong.');
  }
  if (SC_MEILANI_CONFIG.sourceSpreadsheetId === SC_MEILANI_CONFIG.destinationSpreadsheetId) {
    throw new Error('Source dan destination spreadsheet ID tidak boleh sama.');
  }
}
function scMeilaniReadSourceMeta_(sheet, requiredHeaders) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Source sheet kosong: ' + sheet.getName());

  const headerRowIndex = scMeilaniFindBestHeaderRowIndex_(values, requiredHeaders);
  const headerValues = values[headerRowIndex] || [];
  return {
    sheet: sheet,
    values: values,
    headerRowIndex: headerRowIndex,
    headerRowNumber: headerRowIndex + 1,
    headerValues: headerValues,
    headerMap: scMeilaniBuildHeaderMap_(headerValues),
  };
}

function scMeilaniReadTargetMeta_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), SC_MEILANI_CONFIG.outputHeaders.length, 1);
  const scanRowCount = Math.max(Math.min(sheet.getMaxRows(), SC_MEILANI_CONFIG.headerScanRows), 1);
  const scanValues = sheet.getRange(1, 1, scanRowCount, lastColumn).getValues();
  const headerRowIndex = scMeilaniFindBestHeaderRowIndex_(scanValues, ['Claim Number']);
  const headerRowNumber = headerRowIndex + 1;
  const headerValues = sheet.getRange(headerRowNumber, 1, 1, lastColumn).getValues()[0];
  const headerMap = scMeilaniBuildHeaderMap_(headerValues);

  SC_MEILANI_CONFIG.outputHeaders.forEach(function (header, index) {
    if (scMeilaniFindHeaderIndex_(headerMap, header) == null) {
      sheet.getRange(headerRowNumber, index + 1).setValue(header);
    }
  });

  const refreshedHeaders = sheet.getRange(headerRowNumber, 1, 1, Math.max(sheet.getLastColumn(), SC_MEILANI_CONFIG.outputHeaders.length)).getValues()[0];
  return {
    sheet: sheet,
    headerRowIndex: headerRowIndex,
    headerRowNumber: headerRowNumber,
    headerValues: refreshedHeaders,
    headerMap: scMeilaniBuildHeaderMap_(refreshedHeaders),
  };
}

function scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, sourceRowNumbers, outputHeaders, ctx) {
  scMeilaniValidateMirrorHeaders_(sourceMeta, targetMeta, outputHeaders);
  scMeilaniLogStepSafe_(ctx, 'Clear target body ' + targetSheet.getName(), 'START', Math.max(targetSheet.getLastRow() - targetMeta.headerRowNumber, 0), 'Clearing existing target body before mirror.', 'target=' + targetSheet.getName());
  scMeilaniClearBody_(targetSheet, targetMeta.headerRowNumber);
  scMeilaniLogStepSafe_(ctx, 'Clear target body ' + targetSheet.getName(), 'SUCCESS', 0, 'Target body cleared.', 'target=' + targetSheet.getName());

  if (!sourceRowNumbers.length) {
    scMeilaniLogStepSafe_(ctx, 'Batch mirror ' + targetSheet.getName(), 'SKIPPED', 0, 'No matching source rows for this target.', 'target=' + targetSheet.getName());
    return 0;
  }

  const sourceCols = outputHeaders.map(function (header) {
    return scMeilaniRequireHeader_(sourceMeta.headerMap, header);
  });
  const targetCols = outputHeaders.map(function (header) {
    return scMeilaniRequireHeader_(targetMeta.headerMap, header);
  });
  const targetStartRow = targetMeta.headerRowNumber + 1;
  const lastTargetCol = Math.max.apply(null, targetCols) + 1;
  const matrix = sourceRowNumbers.map(function (rowNumber) {
    const row = sourceMeta.values[rowNumber - 1] || [];
    const out = scMeilaniBlankArray_(lastTargetCol);
    outputHeaders.forEach(function (header, index) {
      out[targetCols[index]] = row[sourceCols[index]];
    });
    return out;
  });

  scMeilaniEnsureRows_(targetSheet, targetMeta.headerRowNumber + matrix.length);
  scMeilaniLogStepSafe_(ctx, 'Batch mirror ' + targetSheet.getName(), 'START', matrix.length, 'Writing matched rows in one batch.', 'columns=' + lastTargetCol);
  targetSheet.getRange(targetStartRow, 1, matrix.length, lastTargetCol).setValues(matrix);
  scMeilaniLogStepSafe_(ctx, 'Batch mirror ' + targetSheet.getName(), 'SUCCESS', matrix.length, 'Rows written.', 'target=' + targetSheet.getName());
  return sourceRowNumbers.length;
}

function scMeilaniLogStepSafe_(ctx, processName, status, count, message, details) {
  if (!ctx || !ctx.destinationSpreadsheet) return;
  scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName || '', processName, status, count, message, details || '', ctx.startedAt);
}
function scMeilaniResolveRepairColumnMap_(sourceMeta, targetMeta) {
  const cfg = SC_MEILANI_CONFIG;
  const source = {};
  const target = {};
  const missingTarget = [];
  const optionalMissingSource = [];

  cfg.outputHeaders.forEach(function (header) {
    const targetIdx = scMeilaniFindHeaderIndex_(targetMeta.headerMap, header);
    if (targetIdx == null) missingTarget.push(header);
    target[header] = targetIdx;
    source[header] = scMeilaniFindHeaderIndex_(sourceMeta.headerMap, header);
  });

  const identifierSource = scMeilaniRequireSourceColumn_(sourceMeta, cfg.repair.identifierHeader);
  const lastStatusSource = scMeilaniRequireSourceColumn_(sourceMeta, 'Last Status');
  let branchSource = scMeilaniFindHeaderIndex_(sourceMeta.headerMap, 'Branch');
  const serviceCenterSource = scMeilaniFindHeaderIndex_(sourceMeta.headerMap, 'Service Center');
  if (branchSource == null) branchSource = serviceCenterSource;
  if (branchSource == null && cfg.fallbackColumns.Branch) branchSource = cfg.fallbackColumns.Branch - 1;
  if (branchSource == null) {
    throw new Error('Header source Salvage Repair tidak punya Branch atau Service Center di "' + sourceMeta.sheet.getName() + '" row ' + sourceMeta.headerRowNumber + '. Header terbaca: ' + scMeilaniPreviewHeaders_(sourceMeta.headerValues));
  }

  const identifierTarget = scMeilaniFindHeaderIndex_(targetMeta.headerMap, cfg.repair.identifierHeader);
  if (identifierTarget == null) missingTarget.push(cfg.repair.identifierHeader);

  cfg.outputHeaders.forEach(function (header) {
    if (source[header] == null && header !== 'Submission Month' && header !== 'Branch') optionalMissingSource.push(header);
  });
  if (source.Branch == null && serviceCenterSource != null) optionalMissingSource.push('Branch derived from Service Center');

  if (missingTarget.length) {
    throw new Error('Header target Salvage Repair tidak lengkap di "' + targetMeta.sheet.getName() + '" row ' + targetMeta.headerRowNumber + ': ' + scMeilaniUnique_(missingTarget).join(', ') + '. Header terbaca: ' + scMeilaniPreviewHeaders_(targetMeta.headerValues));
  }

  return {
    source: source,
    target: target,
    identifierSource: identifierSource,
    branchSource: branchSource,
    lastStatusSource: lastStatusSource,
    serviceCenterSource: serviceCenterSource,
    identifierTarget: identifierTarget,
    optionalMissingSource: scMeilaniUnique_(optionalMissingSource),
    sourceMeta: sourceMeta,
    targetLastColumn: Math.max(targetMeta.sheet.getLastColumn(), cfg.outputHeaders.length),
  };
}
function scMeilaniGetRepairOutputValue_(row, columnMap, header, branchValue, serviceCenterValue) {
  if (header === 'Submission Month') {
    const monthIdx = columnMap.source[header];
    if (monthIdx != null) return row[monthIdx];
    return scMeilaniDeriveSubmissionMonth_(scMeilaniGetRepairSourceValue_(row, columnMap, 'Submission Date'));
  }
  if (header === 'Branch') return branchValue || '';
  if (header === 'Service Center') return serviceCenterValue || scMeilaniGetRepairSourceValue_(row, columnMap, header);
  if (header === 'IMEI/SN') {
    const imei = scMeilaniGetRepairSourceValue_(row, columnMap, header);
    if (imei) return imei;
    return scMeilaniFirstExistingHeaderValue_(row, columnMap.sourceMeta, ['device_imei', 'device_imei2']);
  }
  return scMeilaniGetRepairSourceValue_(row, columnMap, header);
}

function scMeilaniGetRepairSourceValue_(row, columnMap, header) {
  const idx = columnMap.source[header];
  if (idx == null) return '';
  return row[idx];
}

function scMeilaniFirstExistingHeaderValue_(row, sourceMeta, aliases) {
  if (!sourceMeta) return '';
  for (let i = 0; i < aliases.length; i++) {
    const idx = sourceMeta.headerMap[scMeilaniHeaderKey_(aliases[i])];
    if (idx == null) continue;
    const value = row[idx];
    if (String(value == null ? '' : value).trim()) return value;
  }
  return '';
}

function scMeilaniResolveRepairBranch_(branchValue, serviceCenterValue) {
  if (scMeilaniIsAllowedBranch_(branchValue)) return branchValue;
  const text = scMeilaniNormalizeText_(String(branchValue || '') + ' ' + String(serviceCenterValue || ''));
  if (!text) return branchValue || '';
  if (text.indexOf('xiaomi') >= 0) return 'Xiaomi Authorized';
  if (text.indexOf('unicom') >= 0) return 'Unicom';
  if (text.indexOf('samsung exclusive') >= 0 || (text.indexOf('samsung') >= 0 && text.indexOf('exclusive') >= 0)) return 'Samsung Exclusive';
  return branchValue || serviceCenterValue || '';
}
function scMeilaniBuildTargetIdentifierIndex_(targetSheet, targetMeta, identifierHeader) {
  const identifierCol = scMeilaniRequireHeader_(targetMeta.headerMap, identifierHeader) + 1;
  const startRow = targetMeta.headerRowNumber + 1;
  const rowCount = Math.max(targetSheet.getLastRow() - targetMeta.headerRowNumber, 0);
  const rowByKey = {};
  const duplicateKeys = {};
  const duplicateSamples = [];
  let uniqueCount = 0;
  let duplicateCount = 0;

  if (rowCount > 0) {
    const values = targetSheet.getRange(startRow, identifierCol, rowCount, 1).getValues();
    values.forEach(function (item, index) {
      const key = scMeilaniIdentifierKey_(item[0]);
      if (!key) return;
      const rowNumber = startRow + index;
      if (rowByKey[key]) {
        duplicateKeys[key] = true;
        duplicateCount += 1;
        if (duplicateSamples.length < 10) duplicateSamples.push(String(item[0]) + ' @ row ' + rowNumber + ' (first row ' + rowByKey[key] + ')');
        return;
      }
      rowByKey[key] = rowNumber;
      uniqueCount += 1;
    });
  }

  return {
    rowByKey: rowByKey,
    duplicateKeys: duplicateKeys,
    uniqueCount: uniqueCount,
    duplicateCount: duplicateCount,
    duplicateSamples: duplicateSamples,
  };
}

function scMeilaniCollectRepairRecords_(sourceMeta, columnMap, allowedStatuses, targetIndex) {
  const cfg = SC_MEILANI_CONFIG;
  const records = [];
  const seenSource = {};
  const skippedByReason = {};
  const failedByReason = {};
  let skippedCount = 0;
  let failedCount = 0;
  let duplicateTargetSkipped = 0;

  for (let r = sourceMeta.headerRowIndex + 1; r < sourceMeta.values.length; r++) {
    const row = sourceMeta.values[r] || [];
    const rowNumber = r + 1;
    try {
      const claimValue = row[columnMap.identifierSource];
      const claimKey = scMeilaniIdentifierKey_(claimValue);
      const serviceCenterValue = columnMap.serviceCenterSource == null ? '' : row[columnMap.serviceCenterSource];
      const branchValue = scMeilaniResolveRepairBranch_(row[columnMap.branchSource], serviceCenterValue);
      const statusValue = row[columnMap.lastStatusSource];
      const statusKey = scMeilaniStatusKey_(statusValue);

      if (!claimKey) {
        skippedCount += 1;
        scMeilaniAddReason_(skippedByReason, 'Claim Number kosong', rowNumber, claimValue, 'Identifier wajib kosong.');
        continue;
      }
      if (targetIndex.duplicateKeys[claimKey]) {
        skippedCount += 1;
        duplicateTargetSkipped += 1;
        scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di target', rowNumber, claimValue, 'Target punya duplicate identifier; update dilewati supaya tidak menulis ke row ambigu.');
        continue;
      }
      if (seenSource[claimKey]) {
        skippedCount += 1;
        scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di source', rowNumber, claimValue, 'Duplicate source; first row ' + seenSource[claimKey] + ' sudah dipakai.');
        continue;
      }
      if (!scMeilaniIsAllowedBranch_(branchValue)) {
        skippedCount += 1;
        scMeilaniAddReason_(skippedByReason, 'Branch bukan scope Salvage Repair', rowNumber, claimValue, 'Branch="' + String(branchValue || '') + '".');
        continue;
      }
      if (allowedStatuses[statusKey] !== true) {
        skippedCount += 1;
        scMeilaniAddReason_(skippedByReason, 'Last Status bukan Salvage Repair', rowNumber, claimValue, 'Last Status="' + String(statusValue || '') + '".');
        continue;
      }

      seenSource[claimKey] = rowNumber;
      records.push({
        claimKey: claimKey,
        claimValue: claimValue,
        sourceRowNumber: rowNumber,
        values: cfg.outputHeaders.map(function (header) {
          return scMeilaniGetRepairOutputValue_(row, columnMap, header, branchValue, serviceCenterValue);
        }),
      });
    } catch (err) {
      failedCount += 1;
      scMeilaniAddReason_(failedByReason, 'Row processing error', rowNumber, row[columnMap.identifierSource], scMeilaniErrorMessage_(err));
    }
  }

  return {
    validRecords: records,
    skippedByReason: skippedByReason,
    failedByReason: failedByReason,
    skippedCount: skippedCount,
    failedCount: failedCount,
    duplicateTargetSkipped: duplicateTargetSkipped,
  };
}

function scMeilaniUpsertRepairRecords_(targetSheet, targetMeta, columnMap, targetIndex, records, ctx) {
  scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Upsert target rows', 'START', records.length, 'Preparing batch upsert for valid Salvage Repair rows.', 'identifier=' + SC_MEILANI_CONFIG.repair.identifierHeader, ctx.startedAt);
  if (!records.length) {
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Upsert target rows', 'SKIPPED', 0, 'No valid records to write.', '', ctx.startedAt);
    return { updated: 0, appended: 0, failed: 0 };
  }

  const lastColumn = columnMap.targetLastColumn;
  const dataStartRow = targetMeta.headerRowNumber + 1;
  const existingRowCount = Math.max(targetSheet.getLastRow() - targetMeta.headerRowNumber, 0);
  let existingValues = [];
  if (existingRowCount > 0) {
    existingValues = targetSheet.getRange(dataStartRow, 1, existingRowCount, lastColumn).getValues();
  }

  const appendRows = [];
  let updated = 0;
  let failed = 0;

  records.forEach(function (record) {
    try {
      const targetRow = targetIndex.rowByKey[record.claimKey];
      if (!targetRow) {
        const newRow = scMeilaniBlankArray_(lastColumn);
        scMeilaniApplyRepairValuesToRow_(newRow, columnMap, record.values);
        appendRows.push({ row: newRow, record: record });
        return;
      }

      const rowIndex = targetRow - dataStartRow;
      if (rowIndex < 0 || rowIndex >= existingValues.length) {
        failed += 1;
        scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Upsert target rows', 'FAILED', 1, 'Indexed target row is outside loaded range.', 'sourceRow=' + record.sourceRowNumber + ', claim=' + record.claimValue + ', targetRow=' + targetRow + ', loadedRows=' + existingValues.length, ctx.startedAt);
        return;
      }

      scMeilaniApplyRepairValuesToRow_(existingValues[rowIndex], columnMap, record.values);
      updated += 1;
    } catch (err) {
      failed += 1;
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Upsert target rows', 'FAILED', 1, 'Failed preparing row for batch write.', 'sourceRow=' + record.sourceRowNumber + ', claim=' + record.claimValue + ', error=' + scMeilaniErrorMessage_(err), ctx.startedAt);
    }
  });

  if (updated > 0) {
    try {
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch update existing rows', 'START', updated, 'Writing existing target rows in one batch.', 'loadedRows=' + existingValues.length + ', columns=' + lastColumn, ctx.startedAt);
      scMeilaniWriteExistingRepairColumns_(targetSheet, dataStartRow, existingValues, columnMap);
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch update existing rows', 'SUCCESS', updated, 'Existing rows updated.', '', ctx.startedAt);
    } catch (err) {
      failed += updated;
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch update existing rows', 'FAILED', updated, 'Batch update failed.', scMeilaniErrorMessage_(err), ctx.startedAt);
      updated = 0;
    }
  } else {
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch update existing rows', 'SKIPPED', 0, 'No existing target rows needed update.', '', ctx.startedAt);
  }

  let appended = 0;
  if (appendRows.length) {
    const startRow = Math.max(targetSheet.getLastRow() + 1, targetMeta.headerRowNumber + 1);
    scMeilaniEnsureRows_(targetSheet, startRow + appendRows.length - 1);
    try {
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch append new rows', 'START', appendRows.length, 'Appending new target rows in one batch.', 'startRow=' + startRow + ', columns=' + lastColumn, ctx.startedAt);
      targetSheet.getRange(startRow, 1, appendRows.length, lastColumn).setValues(appendRows.map(function (item) { return item.row; }));
      appended = appendRows.length;
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch append new rows', 'SUCCESS', appended, 'New rows appended.', '', ctx.startedAt);
    } catch (err) {
      failed += appendRows.length;
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch append new rows', 'FAILED', appendRows.length, 'Batch append failed.', scMeilaniErrorMessage_(err), ctx.startedAt);
    }
  } else {
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Batch append new rows', 'SKIPPED', 0, 'No new rows to append.', '', ctx.startedAt);
  }

  scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Upsert target rows', failed ? 'FAILED' : 'SUCCESS', updated + appended, 'Upsert completed.', 'updated=' + updated + ', appended=' + appended + ', failed=' + failed, ctx.startedAt);
  return { updated: updated, appended: appended, failed: failed };
}

function scMeilaniWriteExistingRepairColumns_(targetSheet, dataStartRow, existingValues, columnMap) {
  if (!existingValues.length) return;
  SC_MEILANI_CONFIG.outputHeaders.forEach(function (header) {
    const targetIdx = columnMap.target[header];
    if (targetIdx == null) return;
    const columnValues = existingValues.map(function (row) {
      return [row[targetIdx]];
    });
    targetSheet.getRange(dataStartRow, targetIdx + 1, existingValues.length, 1).setValues(columnValues);
  });
}
function scMeilaniVerifyRepairResult_(targetSheet, targetMeta, identifierHeader, records) {
  const identifierCol = scMeilaniRequireHeader_(targetMeta.headerMap, identifierHeader) + 1;
  const rowCount = Math.max(targetSheet.getLastRow() - targetMeta.headerRowNumber, 0);
  const claimSet = {};
  const duplicateSet = {};
  let duplicateCount = 0;

  if (rowCount > 0) {
    const values = targetSheet.getRange(targetMeta.headerRowNumber + 1, identifierCol, rowCount, 1).getValues();
    values.forEach(function (item) {
      const key = scMeilaniIdentifierKey_(item[0]);
      if (!key) return;
      if (claimSet[key]) {
        duplicateSet[key] = true;
        duplicateCount += 1;
      }
      claimSet[key] = true;
    });
  }

  const missingClaims = [];
  records.forEach(function (record) {
    if (!claimSet[record.claimKey] && missingClaims.length < 20) missingClaims.push(String(record.claimValue));
  });

  return {
    targetClaims: Object.keys(claimSet).length,
    duplicateCount: duplicateCount,
    matchedSourceClaims: records.length - missingClaims.length,
    missingClaims: missingClaims,
  };
}

function scMeilaniApplyRepairValuesToRow_(targetRow, columnMap, values) {
  SC_MEILANI_CONFIG.outputHeaders.forEach(function (header, index) {
    const targetIdx = columnMap.target[header];
    if (targetIdx == null) return;
    targetRow[targetIdx] = values[index];
  });
}

function scMeilaniDeriveSubmissionMonth_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), 1);
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) return new Date(parsed.getFullYear(), parsed.getMonth(), 1);
  return '';
}

function scMeilaniEnsureRows_(sheet, requiredLastRow) {
  const maxRows = sheet.getMaxRows();
  if (maxRows >= requiredLastRow) return;
  sheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
}

function scMeilaniBlankArray_(length) {
  const out = [];
  for (let i = 0; i < length; i++) out.push('');
  return out;
}

function scMeilaniAddReason_(bucket, reason, rowNumber, claimValue, detail) {
  if (!bucket[reason]) bucket[reason] = { count: 0, samples: [] };
  bucket[reason].count += 1;
  if (bucket[reason].samples.length < 10) {
    bucket[reason].samples.push('row ' + rowNumber + ', claim=' + String(claimValue || '') + ', ' + detail);
  }
}

function scMeilaniWriteReasonLogs_(ctx, processName, status, bucket) {
  Object.keys(bucket || {}).forEach(function (reason) {
    const item = bucket[reason];
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, processName + ' - ' + reason, status, item.count, reason, item.samples.join(' | '), ctx.startedAt);
  });
}

function scMeilaniIdentifierKey_(value) {
  return String(value == null ? '' : value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase();
}

function scMeilaniUnique_(values) {
  const seen = {};
  const out = [];
  (values || []).forEach(function (value) {
    if (seen[value]) return;
    seen[value] = true;
    out.push(value);
  });
  return out;
}
function scMeilaniWriteMirroredColumn_(sourceSheet, targetSheet, sourceRowNumbers, sourceColumn, targetStartRow, targetColumn) {
  const snapshots = [];

  for (let i = 0; i < sourceRowNumbers.length; i++) {
    const sourceCell = sourceSheet.getRange(sourceRowNumbers[i], sourceColumn);
    snapshots.push(scMeilaniReadCellSnapshot_(sourceCell));
  }

  const targetRange = targetSheet.getRange(targetStartRow, targetColumn, sourceRowNumbers.length, 1);
  const values = snapshots.map(function (item) { return [item.value]; });
  targetRange.setValues(values);

  scMeilaniTryWriteRange_(targetRange, 'setNumberFormats', snapshots.map(function (item) { return [item.numberFormat]; }));
  scMeilaniTryWriteRange_(targetRange, 'setBackgrounds', snapshots.map(function (item) { return [item.background]; }));
  scMeilaniTryWriteRange_(targetRange, 'setFontColors', snapshots.map(function (item) { return [item.fontColor]; }));
  scMeilaniTryWriteRange_(targetRange, 'setFontFamilies', snapshots.map(function (item) { return [item.fontFamily]; }));
  scMeilaniTryWriteRange_(targetRange, 'setFontSizes', snapshots.map(function (item) { return [item.fontSize]; }));
  scMeilaniTryWriteRange_(targetRange, 'setFontWeights', snapshots.map(function (item) { return [item.fontWeight]; }));
  scMeilaniTryWriteRange_(targetRange, 'setFontStyles', snapshots.map(function (item) { return [item.fontStyle]; }));
  scMeilaniTryWriteRange_(targetRange, 'setHorizontalAlignments', snapshots.map(function (item) { return [item.horizontalAlignment]; }));
  scMeilaniTryWriteRange_(targetRange, 'setVerticalAlignments', snapshots.map(function (item) { return [item.verticalAlignment]; }));
  scMeilaniTryWriteRange_(targetRange, 'setWraps', snapshots.map(function (item) { return [item.wrap]; }));

  try {
    const richTextValues = snapshots.map(function (item) { return [item.richText]; });
    targetRange.setRichTextValues(richTextValues);
  } catch (err) {
    // Rich text is optional; keep the data if the source cell style cannot be serialized.
  }
}

function scMeilaniReadCellSnapshot_(cell) {
  return {
    value: scMeilaniSafeCall_(function () { return cell.getValue(); }, ''),
    richText: scMeilaniSafeCall_(function () { return cell.getRichTextValue(); }, null),
    numberFormat: scMeilaniSafeCall_(function () { return cell.getNumberFormat(); }, '@'),
    background: scMeilaniSafeCall_(function () { return cell.getBackground(); }, '#ffffff'),
    fontColor: scMeilaniSafeCall_(function () { return cell.getFontColor(); }, '#000000'),
    fontFamily: scMeilaniSafeCall_(function () { return cell.getFontFamily(); }, 'Arial'),
    fontSize: scMeilaniSafeCall_(function () { return cell.getFontSize(); }, 10),
    fontWeight: scMeilaniSafeCall_(function () { return cell.getFontWeight(); }, 'normal'),
    fontStyle: scMeilaniSafeCall_(function () { return cell.getFontStyle(); }, 'normal'),
    horizontalAlignment: scMeilaniSafeCall_(function () { return cell.getHorizontalAlignment(); }, 'left'),
    verticalAlignment: scMeilaniSafeCall_(function () { return cell.getVerticalAlignment(); }, 'middle'),
    wrap: scMeilaniSafeCall_(function () { return cell.getWrap(); }, false),
  };
}

function scMeilaniTryWriteRange_(range, setterName, values) {
  try {
    range[setterName](values);
  } catch (err) {
    // Ignore style write failure and keep going. Data already exists.
  }
}

function scMeilaniSafeCall_(fn, fallback) {
  try {
    const value = fn();
    return value === undefined ? fallback : value;
  } catch (err) {
    return fallback;
  }
}

function scMeilaniValidateMirrorHeaders_(sourceMeta, targetMeta, outputHeaders) {
  const missingSource = [];
  const missingTarget = [];

  outputHeaders.forEach(function (header) {
    if (scMeilaniFindHeaderIndex_(sourceMeta.headerMap, header) == null) missingSource.push(header);
    if (scMeilaniFindHeaderIndex_(targetMeta.headerMap, header) == null) missingTarget.push(header);
  });

  if (missingSource.length) {
    throw new Error('Header source tidak ditemukan di "' + sourceMeta.sheet.getName() + '" row ' + sourceMeta.headerRowNumber + ': ' + missingSource.join(', '));
  }
  if (missingTarget.length) {
    throw new Error('Header target tidak ditemukan di "' + targetMeta.sheet.getName() + '" row ' + targetMeta.headerRowNumber + ': ' + missingTarget.join(', '));
  }
}

function scMeilaniClearBody_(sheet, headerRow) {
  const effectiveHeaderRow = headerRow || SC_MEILANI_CONFIG.headerRow;
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow <= effectiveHeaderRow) return;
  sheet.getRange(effectiveHeaderRow + 1, 1, lastRow - effectiveHeaderRow, lastColumn).clear();
}

function scMeilaniCollectRowNumbers_(sourceMeta, predicate) {
  const values = sourceMeta.values || [];
  const out = [];
  for (let r = sourceMeta.headerRowIndex + 1; r < values.length; r++) {
    const row = values[r] || [];
    if (predicate(row)) out.push(r + 1);
  }
  return out;
}

function scMeilaniFindBestHeaderRowIndex_(values, requiredHeaders) {
  const scanLimit = Math.min(values.length, SC_MEILANI_CONFIG.headerScanRows || 20);
  let bestIndex = SC_MEILANI_CONFIG.headerRow - 1;
  let bestScore = -1;
  const requiredCount = (requiredHeaders || []).length;
  const minScore = Math.max(2, requiredCount);

  for (let r = 0; r < scanLimit; r++) {
    const map = scMeilaniBuildHeaderMap_(values[r] || []);
    let score = 0;
    SC_MEILANI_CONFIG.outputHeaders.forEach(function (header) {
      if (scMeilaniFindHeaderIndex_(map, header) != null) score += 1;
    });
    (requiredHeaders || []).forEach(function (header) {
      if (scMeilaniFindHeaderIndex_(map, header) != null) score += 5;
    });
    if (score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }

  if (bestScore < minScore) return SC_MEILANI_CONFIG.headerRow - 1;
  return bestIndex;
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

function scMeilaniRequireSourceColumn_(sourceMeta, canonicalHeader, fallbackColumnNumber) {
  const idx = scMeilaniFindHeaderIndex_(sourceMeta.headerMap, canonicalHeader);
  if (idx != null) return idx;

  if (fallbackColumnNumber && fallbackColumnNumber > 0) {
    return fallbackColumnNumber - 1;
  }

  throw new Error(
    'Header source tidak ditemukan di "' + sourceMeta.sheet.getName() +
    '" row ' + sourceMeta.headerRowNumber + ': ' + canonicalHeader +
    '. Header terbaca: ' + scMeilaniPreviewHeaders_(sourceMeta.headerValues)
  );
}

function scMeilaniPreviewHeaders_(headers) {
  return (headers || [])
    .map(function (header) { return String(header == null ? '' : header).trim(); })
    .filter(function (header) { return !!header; })
    .slice(0, 30)
    .join(' | ');
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
  if (!sheet) throw new Error('Sheet tidak ditemukan di spreadsheet ' + spreadsheet.getId() + ': ' + sheetName);
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
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[_\s]+/g, ' ');
}

function scMeilaniNormalizeText_(value) {
  return String(value == null ? '' : value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
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

function scMeilaniResetLog_(spreadsheet) {
  const sheet = scMeilaniGetOrCreateLogSheet_(spreadsheet);
  sheet.clear();
  scMeilaniWriteLogHeader_(sheet);
}

function scMeilaniLog_(spreadsheet, flowName, status, startedAt, message, details) {
  scMeilaniLogStep_(spreadsheet, flowName, 'Runtime completion', status === 'ERROR' ? 'FAILED' : status, 0, message, details, startedAt);
}

function scMeilaniLogStep_(spreadsheet, flowName, processName, status, count, message, details, startedAt) {
  const sheet = scMeilaniGetOrCreateLogSheet_(spreadsheet);
  scMeilaniWriteLogHeader_(sheet);
  const endedAt = new Date();
  const normalizedStatus = scMeilaniNormalizeLogStatus_(status);
  sheet.appendRow([
    endedAt,
    flowName,
    processName,
    normalizedStatus,
    Number(count || 0),
    message || '',
    details || '',
    startedAt || '',
    endedAt,
    startedAt ? Math.round((endedAt.getTime() - startedAt.getTime()) / 1000) : '',
  ]);
  scMeilaniFlush_();
}

function scMeilaniWriteLogHeader_(sheet) {
  const headers = [
    'Timestamp',
    'Flow',
    'Process',
    'Status',
    'Count',
    'Message',
    'Details',
    'Started At',
    'Ended At',
    'Duration Seconds',
  ];
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('|');
  if (existing !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    scMeilaniFormatLogSheet_(sheet, headers.length);
  }
  sheet.setFrozenRows(1);
}

function scMeilaniFormatLogSheet_(sheet, headerCount) {
  try {
    sheet.getRange(1, 1, Math.min(Math.max(sheet.getMaxRows(), 1000), 5000), headerCount).setNumberFormat('@');
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange('E:E').setNumberFormat('0');
    sheet.getRange('H:I').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange('J:J').setNumberFormat('0');
    sheet.getRange(1, 1, 1, headerCount).setFontWeight('bold');
  } catch (err) {
    // Formatting should never block the flow.
  }
}


function scMeilaniFlush_() {
  try {
    SpreadsheetApp.flush();
  } catch (err) {
    // Best-effort only; logging must not fail the main flow.
  }
}
function scMeilaniNormalizeLogStatus_(status) {
  const value = String(status || '').trim().toUpperCase();
  if (value === 'START' || value === 'SUCCESS' || value === 'SKIPPED' || value === 'FAILED') return value;
  if (value === 'ERROR') return 'FAILED';
  return 'FAILED';
}

function scMeilaniResultCount_(result) {
  if (result == null) return 0;
  if (typeof result === 'number') return result;
  if (Array.isArray(result)) return result.length;
  if (typeof result === 'object') {
    if (result.valid != null) return Number(result.valid) || 0;
    if (result.appended != null || result.updated != null) return (Number(result.appended) || 0) + (Number(result.updated) || 0);
  }
  return 0;
}

function scMeilaniGetOrCreateLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SC_MEILANI_CONFIG.logSheetName);
  if (sheet) return sheet;

  sheet = spreadsheet.insertSheet(SC_MEILANI_CONFIG.logSheetName);
  scMeilaniWriteLogHeader_(sheet);
  return sheet;
}

function scMeilaniErrorMessage_(err) {
  if (!err) return '';
  const base = err && err.message ? err.message : String(err);
  const stack = err && err.stack ? String(err.stack).split('\n').slice(0, 4).join(' | ') : '';
  return stack ? base + ' | stack: ' + stack : base;
}
