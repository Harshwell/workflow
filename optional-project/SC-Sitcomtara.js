/**
 * SC-Sitcomtara
 *
 * Manual Google Sheets menu for mirroring Salvage data and upserting Salvage Repair data
 * from source workbook into the active destination workbook.
 *
 * Menu:
 * - SC - Sitcomtara > Salvage
 * - SC - Sitcomtara > Salvage Repair
 */

const SC_MEILANI_CONFIG = Object.freeze({
  sourceSpreadsheetId: '1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc',
  destinationSpreadsheetId: '1yXQ4FwvPbBX22vfwDnsnrefGkq3cOoR8vc_tJILgZFg',
  scriptVersion: '2026-08-05-repair-start-approval-cutoff-1',
  menuName: 'SC - Sitcomtara',
  logSheetName: 'Log SC-Sitcomtara',
  headerRow: 1,
  headerScanRows: 20,
  fallbackColumns: Object.freeze({
    Branch: 6,
    YoS: 8,
    Remarks: 20,
  }),
  allowedBranches: Object.freeze([
    'Sitcomtara',
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
  repairRemarksBackup: Object.freeze({
    sourceSheetName: 'Repair',
    targetSheetName: 'SC - Sitcomtara',
    identifierHeader: 'Claim Number',
    remarksHeader: 'Remarks',
  }),
  repairUpdateSync: Object.freeze({
    sourceSheetName: 'Repair',
    targetSheetNames: Object.freeze(['SC - Farhan']),
    identifierHeader: 'Claim Number',
    sourceUpdateHeader: 'Update from Service Center',
    targetRemarksHeader: 'Remarks',
    skipBlankUpdates: true,
  }),
  repairFlow: Object.freeze({
    overviewSheetName: 'Overview',
    branchRangeA1: 'G2:G4',
    sourceSheetNames: Object.freeze(['SC - Farhan', 'SC - Meilani', 'SC - Meindar']),
    targetSheetName: 'Repair',
    updateHeader: 'Update from Service Center',
    statusTypeHeader: 'Status Type',
    statusTypeOptions: Object.freeze([
      'Start',
      'On Repair',
      'Waiting Repair',
      'Waiting Estimate',
      'Waiting Receive Unit',
      'Waiting Insurance Approval',
      'Waiting Payment Cust',
      'Finish',
    ]),
    outputHeaders: Object.freeze([
      'Claim Number',
      'Device Brand',
      'Device Type',
      'IMEI/SN',
      'Service Center Name',
      'Dashboard Link',
      'Last Status',
      'Status Type',
      'Last Status Date',
      'Last Status Aging',
      'Update from Service Center',
    ]),
  }),
  repair: Object.freeze({
    sourceSheetName: 'Raw Data',
    targetSheetName: 'Pickup Sparepart Repair (Salvage)',
    identifierHeader: 'Claim Number',
    sourceHeaders: Object.freeze({
      'Submission Date': 'claim_submitted_datetime',
      'Claim Number': 'claim_number',
      'Approval Date': 'ins_approve_datetime',
      'Service Center': 'repairer_location_store_name',
      'Insurance': 'insurance_partner_code',
      'Sum Insured': 'sum_insured_amount',
      'Device Brand': 'device_brand',
      'Device Type': 'device_type',
      'IMEI/SN': 'imei_number',
      'Last Status': 'claim_last_status_name',
    }),
    minApprovalDate: Object.freeze({ year: 2026, month: 8, day: 1 }),
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
    'Submission Date': Object.freeze(['claim_submitted_datetime', 'Submission Date', 'Submitted Datetime', 'Submitted Date', 'Claim Submitted Datetime', 'claim_submission_date', 'claim_submitted_at', 'created_at']),
    'Submission Month': Object.freeze(['Submission Month', 'Submitted Month', 'Submission by Month', 'Month', 'claim_submission_month', 'submitted_month']),
    'Claim Number': Object.freeze(['claim_number', 'Claim Number', 'Claim No', 'Claim']),
    'DB Link': Object.freeze(['DB Link', 'Dashboard Link', 'Link', 'dashboard_link', 'db_link', 'dblink', 'database_link']),
    'Approval Date': Object.freeze(['ins_approve_datetime', 'Approval Date', 'Approved Date', 'Insurance Approval Date', 'approval_date', 'claim_approved_at', 'claim_approval_date', 'claim_last_updated_datetime']),
    'Branch': Object.freeze(['Branch', 'Service Center Branch', 'SC Branch', 'branch', 'service_center_branch', 'sc_branch']),
    'Service Center': Object.freeze(['repairer_location_store_name', 'Service Center', 'Service Center Name', 'SC Name', 'SC', 'service_center_name', 'service_center', 'sc_name']),
    'Insurance': Object.freeze(['insurance_partner_code', 'Insurance', 'Insurance Name', 'insurance_name', 'insurer_name', 'insurance_code']),
    'Sum Insured': Object.freeze(['sum_insured_amount', 'Sum Insured', 'SI', 'sum_insured', 'insured_value']),
    'Device Brand': Object.freeze(['device_brand', 'Device Brand', 'Brand', 'brand_name', 'device_brand_name']),
    'Device Type': Object.freeze(['device_type', 'Device Type', 'Type', 'device_model']),
    'IMEI/SN': Object.freeze(['imei_number', 'IMEI/SN', 'IMEI', 'SN', 'Serial Number', 'IMEI SN', 'imei', 'serial_number', 'device_imei', 'device_imei2']),
    'Last Status': Object.freeze(['claim_last_status_name', 'Last Status', 'Status Terakhir', 'last_status', 'last_status_name']),
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
    .addItem('Start Repair', 'runSCMeilaniRepairStart')
    .addItem('Backup Repair Remarks', 'runSCMeilaniBackupRepairRemarks')
    .addItem('Update SC Universe Remarks', 'runSCMeilaniRepairUpdateToScUniverse')
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
      }, ctx);
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter target ' + target.sheetName, rowNumbers.length ? 'SUCCESS' : 'SKIPPED', rowNumbers.length, 'Filter completed for target year ' + target.year + '.', 'sourceRows=' + sourceRows + ', matched=' + rowNumbers.length, ctx.startedAt);

      const written = scMeilaniMirrorRows_(sourceSheet, sourceMeta, targetSheet, targetMeta, rowNumbers, cfg.outputHeaders, ctx);
      results.push(target.sheetName + ': ' + written + ' row');
      scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Process target ' + target.sheetName, 'SUCCESS', written, 'Target completed.', '', ctx.startedAt);
    });

    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage selesai. ' + results.join(' | '));
    return results;
  });
}


function runSCMeilaniRepairStart() {
  return scMeilaniWithLock_('REPAIR_START', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    const repairCfg = cfg.repairFlow;
    const overviewSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, repairCfg.overviewSheetName);
    const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, repairCfg.targetSheetName);
    const selectedBranches = scMeilaniReadSelectedBranches_(overviewSheet, repairCfg.branchRangeA1);
    if (!selectedBranches.length) throw new Error('Overview!' + repairCfg.branchRangeA1 + ' belum berisi branch yang dipilih.');

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read selected Repair branches', 'SUCCESS', selectedBranches.length, 'Selected branch filters loaded from Overview.', selectedBranches.join(', '), ctx.startedAt);
    const existingUpdates = scMeilaniBackupRepairUpdateValues_(targetSheet, repairCfg);
    const records = scMeilaniCollectRepairStartRecords_(ctx.destinationSpreadsheet, repairCfg, selectedBranches, ctx);
    const written = scMeilaniRewriteRepairSheet_(targetSheet, repairCfg, records, existingUpdates, ctx);
    scMeilaniToast_(ctx.destinationSpreadsheet, 'Start Repair selesai. Branch ' + selectedBranches.join(', ') + ', row ' + written + '.');
    return { selectedBranches: selectedBranches, records: records.length, written: written };
  });
}

function runSCMeilaniBackupRepairRemarks() {
  return scMeilaniWithLock_('BACKUP_REPAIR_REMARKS', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    const backupCfg = cfg.repairRemarksBackup;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure Repair and SC sheets', 'START', 0, 'Checking required sheets for Repair Remarks backup.', 'source="' + backupCfg.sourceSheetName + '", target="' + backupCfg.targetSheetName + '"', ctx.startedAt);
    const sourceSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, backupCfg.sourceSheetName);
    const targetSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, backupCfg.targetSheetName);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure Repair and SC sheets', 'SUCCESS', 2, 'Required sheets found.', 'source="' + sourceSheet.getName() + '", target="' + targetSheet.getName() + '"', ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read Repair and SC headers', 'START', 0, 'Reading Claim Number and Remarks columns.', '', ctx.startedAt);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet, [backupCfg.identifierHeader, backupCfg.remarksHeader]);
    const targetMeta = scMeilaniReadSourceMeta_(targetSheet, [backupCfg.identifierHeader, backupCfg.remarksHeader]);
    const sourceClaimCol = scMeilaniRequireSourceColumn_(sourceMeta, backupCfg.identifierHeader);
    const sourceRemarksCol = scMeilaniRequireSourceColumn_(sourceMeta, backupCfg.remarksHeader);
    const targetClaimCol = scMeilaniRequireSourceColumn_(targetMeta, backupCfg.identifierHeader);
    const targetRemarksCol = scMeilaniRequireSourceColumn_(targetMeta, backupCfg.remarksHeader);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read Repair and SC headers', 'SUCCESS', 4, 'Header validation passed.', 'sourceHeaderRow=' + sourceMeta.headerRowNumber + ', targetHeaderRow=' + targetMeta.headerRowNumber, ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Match Repair claims', 'START', Math.max(sourceMeta.values.length - sourceMeta.headerRowNumber, 0), 'Building Claim Number index from Repair sheet.', 'identifier=' + backupCfg.identifierHeader, ctx.startedAt);
    const sourceIndex = scMeilaniBuildSingleRowIndex_(sourceMeta, sourceClaimCol);
    const targetIndex = scMeilaniBuildSingleRowIndex_(targetMeta, targetClaimCol);
    const duplicateSourceCount = Object.keys(sourceIndex.duplicateKeys).length;
    const duplicateTargetCount = Object.keys(targetIndex.duplicateKeys).length;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Match Repair claims', 'SUCCESS', sourceIndex.rowByKeyCount, 'Claim Number index built.', 'sourceDuplicates=' + duplicateSourceCount + ', targetDuplicates=' + duplicateTargetCount + ', targetClaims=' + targetIndex.rowByKeyCount, ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Copy Remarks 1:1', 'START', targetIndex.rowByKeyCount, 'Copying Repair Remarks cells into SC - Meilani Remarks by Claim Number.', 'copyMode=cellCopyTo', ctx.startedAt);
    const result = scMeilaniCopyRemarksByClaim_(sourceSheet, targetSheet, sourceIndex, targetIndex, sourceRemarksCol + 1, targetRemarksCol + 1, ctx);
    scMeilaniWriteReasonLogs_(ctx, 'Copy Remarks 1:1', 'SKIPPED', result.skippedByReason);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Copy Remarks 1:1', 'SUCCESS', result.copied, 'Repair Remarks backup completed.', 'copied=' + result.copied + ', skipped=' + result.skipped + ', sourceDuplicates=' + duplicateSourceCount + ', targetDuplicates=' + duplicateTargetCount, ctx.startedAt);
    scMeilaniToast_(ctx.destinationSpreadsheet, 'Backup Repair Remarks selesai. Copy ' + result.copied + ', skip ' + result.skipped + '.');
    return result;
  });
}


function runSCMeilaniRepairUpdateToScUniverse() {
  return scMeilaniWithLock_('REPAIR_UPDATE_TO_SC_UNIVERSE', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    const syncCfg = cfg.repairUpdateSync;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure Repair and SC Universe sheets', 'START', 0, 'Checking required sheets for Repair update sync.', 'source="' + syncCfg.sourceSheetName + '", targets="' + syncCfg.targetSheetNames.join(', ') + '"', ctx.startedAt);
    const sourceSheet = scMeilaniRequireSheet_(ctx.destinationSpreadsheet, syncCfg.sourceSheetName);
    const targetSheets = syncCfg.targetSheetNames.map(function (sheetName) { return scMeilaniRequireSheet_(ctx.destinationSpreadsheet, sheetName); });
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure Repair and SC Universe sheets', 'SUCCESS', targetSheets.length + 1, 'Required sheets found.', 'source="' + sourceSheet.getName() + '", targets="' + syncCfg.targetSheetNames.join(', ') + '"', ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read Repair and SC Universe headers', 'START', 0, 'Reading Claim Number, Update from Service Center, and Remarks columns.', '', ctx.startedAt);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet, [syncCfg.identifierHeader, syncCfg.sourceUpdateHeader]);
    const targetMetas = targetSheets.map(function (sheet) { return scMeilaniReadSourceMeta_(sheet, [syncCfg.identifierHeader, syncCfg.targetRemarksHeader]); });
    const sourceClaimCol = scMeilaniRequireSourceColumn_(sourceMeta, syncCfg.identifierHeader);
    const sourceUpdateCol = scMeilaniRequireSourceColumn_(sourceMeta, syncCfg.sourceUpdateHeader);
    const targetMaps = targetMetas.map(function (targetMeta, index) {
      return {
        sheet: targetSheets[index],
        meta: targetMeta,
        claimCol: scMeilaniRequireSourceColumn_(targetMeta, syncCfg.identifierHeader),
        remarksCol: scMeilaniRequireSourceColumn_(targetMeta, syncCfg.targetRemarksHeader),
      };
    });
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read Repair and SC Universe headers', 'SUCCESS', 2 + targetMaps.length * 2, 'Header validation passed.', 'sourceHeaderRow=' + sourceMeta.headerRowNumber + ', targetSheets=' + syncCfg.targetSheetNames.join(', '), ctx.startedAt);

    const result = scMeilaniCopyRepairUpdatesToScUniverse_(sourceSheet, sourceMeta, sourceClaimCol, sourceUpdateCol, targetMaps, syncCfg, ctx);
    scMeilaniWriteReasonLogs_(ctx, 'Copy Update from Service Center to SC Universe Remarks', 'SKIPPED', result.skippedByReason);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Copy Update from Service Center to SC Universe Remarks', result.failed ? 'FAILED' : 'SUCCESS', result.updated, 'Repair update sync completed.', 'updated=' + result.updated + ', skipped=' + result.skipped + ', failed=' + result.failed, ctx.startedAt);
    scMeilaniToast_(ctx.destinationSpreadsheet, 'Update SC Universe Remarks selesai. Update ' + result.updated + ', skip ' + result.skipped + ', failed ' + result.failed + '.');
    return result;
  });
}

function runSCMeilaniSalvageRepair() {
  return scMeilaniWithLock_('SALVAGE_REPAIR', function (ctx) {
    const cfg = SC_MEILANI_CONFIG;
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure source/target sheets', 'START', 0, 'Checking required sheets.', 'source="' + cfg.repair.sourceSheetName + '", target="' + cfg.repair.targetSheetName + '"', ctx.startedAt);
    const sourceSheet = scMeilaniRequireSheet_(ctx.sourceSpreadsheet, cfg.repair.sourceSheetName);
    const targetSheet = scMeilaniResolveOptionalTargetSheet_(ctx.destinationSpreadsheet, cfg.repair.targetSheetName);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Ensure source/target sheets', 'SUCCESS', 2, 'Required sheets found.', 'sourceId=' + ctx.sourceSpreadsheet.getId() + ', destinationId=' + ctx.destinationSpreadsheet.getId(), ctx.startedAt);

    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read and validate headers', 'START', 0, 'Reading source and target headers.', '', ctx.startedAt);
    const sourceMeta = scMeilaniReadSourceMeta_(sourceSheet, [cfg.repair.identifierHeader, 'Branch', 'Last Status', 'Submission Date', 'Approval Date']);
    const targetMeta = scMeilaniReadSourceMeta_(targetSheet, [cfg.repair.identifierHeader]);
    const columnMap = scMeilaniResolveRepairColumnMap_(sourceMeta, targetMeta);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Read and validate headers', 'SUCCESS', cfg.outputHeaders.length, 'Header validation passed.', 'sourceHeaderRow=' + sourceMeta.headerRowNumber + ', targetHeaderRow=' + targetMeta.headerRowNumber + ', optionalMissingSource=' + columnMap.optionalMissingSource.join(', '), ctx.startedAt);

    const allowedStatuses = scMeilaniToKeySet_(cfg.repair.allowedLastStatuses);
    const targetIndex = scMeilaniBuildTargetIdentifierIndex_(targetSheet, targetMeta, cfg.repair.identifierHeader);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Scan existing target', 'SUCCESS', targetIndex.uniqueCount, 'Existing target identifiers indexed.', 'duplicateTarget=' + targetIndex.duplicateCount + ', duplicateSamples=' + targetIndex.duplicateSamples.join(' | '), ctx.startedAt);

    const sourceRows = Math.max(sourceMeta.values.length - sourceMeta.headerRowNumber, 0);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter Salvage Repair source', 'START', sourceRows, 'Filtering ' + cfg.repair.sourceSheetName + ' rows by branch, repair last status, and Approval Date cutoff.', 'allowedBranches=' + cfg.allowedBranches.join(', ') + ', allowedStatuses=' + cfg.repair.allowedLastStatuses.length + ', minApprovalDate=' + scMeilaniFormatDateConfig_(cfg.repair.minApprovalDate), ctx.startedAt);
    const sourceSnapshot = scMeilaniCollectRepairRecords_(sourceMeta, columnMap, allowedStatuses, targetIndex, ctx);
    scMeilaniWriteReasonLogs_(ctx, 'Filter Salvage Repair source', 'SKIPPED', sourceSnapshot.skippedByReason);
    scMeilaniWriteReasonLogs_(ctx, 'Filter Salvage Repair source', 'FAILED', sourceSnapshot.failedByReason);
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Filter Salvage Repair source', 'SUCCESS', sourceSnapshot.validRecords.length, 'Filtering completed.', 'sourceRows=' + sourceRows + ', valid=' + sourceSnapshot.validRecords.length + ', skipped=' + sourceSnapshot.skippedCount + ', failed=' + sourceSnapshot.failedCount, ctx.startedAt);

    scMeilaniCheckStop_(ctx, 'SALVAGE_REPAIR before upsert');
    const writeResult = scMeilaniUpsertRepairRecords_(targetSheet, targetMeta, columnMap, targetIndex, sourceSnapshot.validRecords, ctx);
    const sortResult = scMeilaniSortRepairTarget_(targetSheet, targetMeta, columnMap, ctx);
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
      targetSheetName: targetSheet.getName(),
      sortedRows: sortResult.sortedRows,
    };

    scMeilaniToast_(ctx.destinationSpreadsheet, 'Salvage Repair selesai. Valid ' + summary.valid + ', update ' + summary.updated + ', append ' + summary.appended + ', skip ' + summary.skipped + ', failed ' + summary.failed + '.');
    return summary;
  });
}
function scMeilaniWithLock_(flowName, runner) {
  const stopRequestedAtMs = scMeilaniRequestPreviousRunsStop_(flowName);
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const waitStartedAt = new Date();
  try {
    lock.waitLock(300000);
    lockAcquired = true;
  } catch (lockErr) {
    throw new Error('Run dibatalkan: proses sebelumnya belum berhenti setelah 5 menit. Stop signal sudah dikirim pada ' + new Date(stopRequestedAtMs).toISOString() + '. Detail: ' + scMeilaniErrorMessage_(lockErr));
  }

  const startedAt = new Date();
  const runId = scMeilaniCreateRunId_(flowName, startedAt);
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
    scMeilaniRegisterActiveRun_(runId, flowName, startedAt);
    scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime bootstrap', 'START', 0, 'Run started; previous run stop signal sent before acquiring lock.', 'scriptVersion=' + SC_MEILANI_CONFIG.scriptVersion + ', runId=' + runId + ', stopRequestedAt=' + new Date(stopRequestedAtMs).toISOString() + ', lockWaitSeconds=' + Math.round((startedAt.getTime() - waitStartedAt.getTime()) / 1000) + ', logMode=' + (preservedLog ? 'PRESERVED_RUN_ALL' : 'RESET'), startedAt);

    const sourceSpreadsheet = SpreadsheetApp.openById(SC_MEILANI_CONFIG.sourceSpreadsheetId);
    const ctx = {
      flowName: flowName,
      destinationSpreadsheet: destinationSpreadsheet,
      sourceSpreadsheet: sourceSpreadsheet,
      startedAt: startedAt,
      startedAtMs: startedAt.getTime(),
      runId: runId,
    };

    scMeilaniCheckStop_(ctx, 'Runtime before runner');
    scMeilaniToast_(destinationSpreadsheet, flowName + ' berjalan...');
    const result = runner(ctx);
    scMeilaniCheckStop_(ctx, 'Runtime before completion');
    scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime completion', 'SUCCESS', scMeilaniResultCount_(result), 'Run completed.', JSON.stringify(result), startedAt);
    return result;
  } catch (err) {
    if (destinationSpreadsheet) {
      const isStop = scMeilaniIsStopError_(err);
      scMeilaniLogStep_(destinationSpreadsheet, flowName, 'Runtime failure', isStop ? 'SKIPPED' : 'FAILED', 0, isStop ? 'Run stopped by newer execution.' : 'Run failed.', scMeilaniErrorMessage_(err), startedAt);
      scMeilaniToast_(destinationSpreadsheet, flowName + (isStop ? ' dihentikan oleh run baru.' : ' gagal: ' + scMeilaniErrorMessage_(err)));
    }
    throw err;
  } finally {
    scMeilaniClearActiveRun_(runId);
    if (lockAcquired) lock.releaseLock();
  }
}
function scMeilaniRequestPreviousRunsStop_(flowName) {
  const nowMs = Date.now();
  try {
    PropertiesService.getScriptProperties().setProperties({
      SC_MEILANI_STOP_REQUESTED_AT_MS: String(nowMs),
      SC_MEILANI_STOP_REQUESTED_BY_FLOW: String(flowName || ''),
    }, false);
  } catch (err) {}
  return nowMs;
}

function scMeilaniCreateRunId_(flowName, startedAt) {
  try {
    return String(flowName || 'RUN') + '-' + Utilities.getUuid();
  } catch (err) {
    return String(flowName || 'RUN') + '-' + String(startedAt.getTime()) + '-' + String(Math.floor(Math.random() * 1e9));
  }
}

function scMeilaniRegisterActiveRun_(runId, flowName, startedAt) {
  try {
    PropertiesService.getScriptProperties().setProperties({
      SC_MEILANI_ACTIVE_RUN_ID: String(runId || ''),
      SC_MEILANI_ACTIVE_FLOW: String(flowName || ''),
      SC_MEILANI_ACTIVE_STARTED_AT_MS: String(startedAt ? startedAt.getTime() : Date.now()),
    }, false);
  } catch (err) {}
}

function scMeilaniClearActiveRun_(runId) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('SC_MEILANI_ACTIVE_RUN_ID') === String(runId || '')) {
      props.deleteProperty('SC_MEILANI_ACTIVE_RUN_ID');
      props.deleteProperty('SC_MEILANI_ACTIVE_FLOW');
      props.deleteProperty('SC_MEILANI_ACTIVE_STARTED_AT_MS');
    }
  } catch (err) {}
}

function scMeilaniCheckStop_(ctx, checkpoint) {
  if (!ctx || !ctx.startedAtMs) return;
  let stopMs = 0;
  let stopBy = '';
  try {
    const props = PropertiesService.getScriptProperties();
    stopMs = Number(props.getProperty('SC_MEILANI_STOP_REQUESTED_AT_MS') || 0);
    stopBy = props.getProperty('SC_MEILANI_STOP_REQUESTED_BY_FLOW') || '';
  } catch (err) {
    return;
  }
  if (stopMs > ctx.startedAtMs) {
    const err = new Error('SC_MEILANI_STOP_REQUESTED: dihentikan oleh run baru. checkpoint=' + String(checkpoint || '') + ', requestedBy=' + stopBy + ', requestedAt=' + new Date(stopMs).toISOString() + ', currentRunId=' + String(ctx.runId || ''));
    err.scMeilaniStopped = true;
    throw err;
  }
}

function scMeilaniIsStopError_(err) {
  return !!(err && (err.scMeilaniStopped || String(err.message || '').indexOf('SC_MEILANI_STOP_REQUESTED') >= 0));
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

  scMeilaniCheckStop_(ctx, 'Mirror before write ' + targetSheet.getName());
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

function scMeilaniResolveOptionalTargetSheet_(spreadsheet, sheetName) {
  const configuredName = String(sheetName == null ? '' : sheetName).trim();
  if (configuredName) return scMeilaniRequireSheet_(spreadsheet, configuredName);
  const activeSheet = spreadsheet.getActiveSheet();
  if (!activeSheet) throw new Error('Target sheet Salvage Repair tidak dikonfigurasi dan active sheet tidak tersedia.');
  return activeSheet;
}

function scMeilaniFindRepairSourceHeaderIndex_(sourceMeta, canonicalHeader) {
  const configuredHeader = SC_MEILANI_CONFIG.repair.sourceHeaders[canonicalHeader];
  if (configuredHeader) {
    const configuredIdx = sourceMeta.headerMap[scMeilaniHeaderKey_(configuredHeader)];
    if (configuredIdx != null) return configuredIdx;
  }
  return scMeilaniFindHeaderIndex_(sourceMeta.headerMap, canonicalHeader);
}

function scMeilaniRequireRepairSourceColumn_(sourceMeta, canonicalHeader) {
  const idx = scMeilaniFindRepairSourceHeaderIndex_(sourceMeta, canonicalHeader);
  if (idx != null) return idx;
  const configuredHeader = SC_MEILANI_CONFIG.repair.sourceHeaders[canonicalHeader];
  throw new Error(
    'Header source Salvage Repair tidak ditemukan di "' + sourceMeta.sheet.getName() +
    '" row ' + sourceMeta.headerRowNumber + ': ' + (configuredHeader || canonicalHeader) +
    '. Header terbaca: ' + scMeilaniPreviewHeaders_(sourceMeta.headerValues)
  );
}

function scMeilaniResolveRepairColumnMap_(sourceMeta, targetMeta) {
  const cfg = SC_MEILANI_CONFIG;
  const source = {};
  const target = {};
  const missingTarget = [];
  const optionalMissingSource = [];

  cfg.outputHeaders.forEach(function (header) {
    const targetIdx = scMeilaniFindHeaderIndex_(targetMeta.headerMap, header);
    target[header] = targetIdx;
    source[header] = scMeilaniFindRepairSourceHeaderIndex_(sourceMeta, header);
  });

  const identifierSource = scMeilaniRequireRepairSourceColumn_(sourceMeta, cfg.repair.identifierHeader);
  const lastStatusSource = scMeilaniRequireRepairSourceColumn_(sourceMeta, 'Last Status');
  const submissionDateSource = scMeilaniFindRepairSourceHeaderIndex_(sourceMeta, 'Submission Date');
  const approvalDateSource = scMeilaniRequireRepairSourceColumn_(sourceMeta, 'Approval Date');
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
    if (source[header] == null && header !== 'Submission Month' && header !== 'Branch' && header !== 'DB Link') optionalMissingSource.push(header);
  });
  if (source.Branch == null && serviceCenterSource != null) optionalMissingSource.push('Branch derived from Service Center');

  if (missingTarget.length) {
    throw new Error('Header target Salvage Repair wajib tidak ditemukan di "' + targetMeta.sheet.getName() + '" row ' + targetMeta.headerRowNumber + ': ' + scMeilaniUnique_(missingTarget).join(', ') + '. Kolom target lain bersifat opsional. Header terbaca: ' + scMeilaniPreviewHeaders_(targetMeta.headerValues));
  }

  return {
    source: source,
    target: target,
    identifierSource: identifierSource,
    branchSource: branchSource,
    lastStatusSource: lastStatusSource,
    submissionDateSource: submissionDateSource,
    approvalDateSource: approvalDateSource,
    serviceCenterSource: serviceCenterSource,
    identifierTarget: identifierTarget,
    optionalMissingSource: scMeilaniUnique_(optionalMissingSource),
    sourceMeta: sourceMeta,
    targetLastColumn: Math.max(targetMeta.sheet.getLastColumn(), identifierTarget + 1),
  };
}
function scMeilaniGetRepairOutputValue_(row, columnMap, header, branchValue, serviceCenterValue) {
  if (header === 'Submission Month') {
    const monthIdx = columnMap.source[header];
    if (monthIdx != null) return row[monthIdx];
    return scMeilaniDeriveSubmissionMonth_(scMeilaniGetRepairSourceValue_(row, columnMap, 'Submission Date'));
  }
  if (header === 'Branch') return branchValue || '';
  if (header === 'DB Link') return scMeilaniBuildDashboardLinkFormula_(scMeilaniGetRepairSourceValue_(row, columnMap, 'Claim Number'));
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


function scMeilaniBuildDashboardLinkFormula_(claimValue) {
  const claim = String(claimValue == null ? '' : claimValue).trim();
  const literal = scMeilaniSheetsStringLiteral_(claim);
  const host = 'https:' + ['//internal', 'qoala', 'app'].join('.');
  const gadgetBase = host + '/gadget/claim/';
  const partnershipBase = host + '/partnership/claim/';
  return '=IF(' + literal + '="", "", LET(link, IF(REGEXMATCH(' + literal + ', "SFP|SFX|SMR|SPP"), "' + gadgetBase + '" & ' + literal + ', "' + partnershipBase + '" & ' + literal + '), HYPERLINK(link, "LINK")))';
}

function scMeilaniSheetsStringLiteral_(value) {
  return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
}

function scMeilaniIsOnOrAfterDateConfig_(value, dateConfig) {
  const parsed = scMeilaniParseDateOnly_(value);
  if (!parsed) return false;
  const cutoff = new Date(Number(dateConfig.year), Number(dateConfig.month) - 1, Number(dateConfig.day || 1));
  return parsed.getTime() >= cutoff.getTime();
}

function scMeilaniParseDateOnly_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const text = String(value == null ? '' : value).trim();
  if (!text) return null;
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function scMeilaniFormatDateConfig_(dateConfig) {
  const year = String(dateConfig && dateConfig.year || '');
  const month = String(dateConfig && dateConfig.month || '').padStart(2, '0');
  const day = String(dateConfig && dateConfig.day || 1).padStart(2, '0');
  return year + '-' + month + '-' + day;
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

function scMeilaniCollectRepairRecords_(sourceMeta, columnMap, allowedStatuses, targetIndex, ctx) {
  const cfg = SC_MEILANI_CONFIG;
  const records = [];
  const seenSource = {};
  const skippedByReason = {};
  const failedByReason = {};
  let skippedCount = 0;
  let failedCount = 0;
  let duplicateTargetSkipped = 0;

  for (let r = sourceMeta.headerRowIndex + 1; r < sourceMeta.values.length; r++) {
    if ((r - sourceMeta.headerRowIndex) % 250 === 0) scMeilaniCheckStop_(ctx, 'Collect repair records row ' + (r + 1));
    const row = sourceMeta.values[r] || [];
    const rowNumber = r + 1;
    try {
      const claimValue = row[columnMap.identifierSource];
      const claimKey = scMeilaniIdentifierKey_(claimValue);
      const serviceCenterValue = columnMap.serviceCenterSource == null ? '' : row[columnMap.serviceCenterSource];
      const branchValue = scMeilaniResolveRepairBranch_(row[columnMap.branchSource], serviceCenterValue);
      const statusValue = row[columnMap.lastStatusSource];
      const statusKey = scMeilaniStatusKey_(statusValue);
      const approvalDateValue = row[columnMap.approvalDateSource];

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
      if (!scMeilaniIsOnOrAfterDateConfig_(approvalDateValue, cfg.repair.minApprovalDate)) {
        skippedCount += 1;
        scMeilaniAddReason_(skippedByReason, 'Approval Date sebelum cutoff Salvage Repair', rowNumber, claimValue, 'Approval Date="' + String(approvalDateValue || '') + '", cutoff=' + scMeilaniFormatDateConfig_(cfg.repair.minApprovalDate) + '.');
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

  records.forEach(function (record, index) {
    if (index % 250 === 0) scMeilaniCheckStop_(ctx, 'Prepare upsert record ' + (index + 1));
    try {
      const targetRow = targetIndex.rowByKey[record.claimKey];
      if (!targetItem) {
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
    scMeilaniCheckStop_(ctx, 'Before batch update existing rows');
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
    scMeilaniCheckStop_(ctx, 'Before batch append new rows');
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

function scMeilaniSortRepairTarget_(targetSheet, targetMeta, columnMap, ctx) {
  const branchIdx = columnMap.target.Branch;
  const approvalIdx = columnMap.target['Approval Date'];
  const rowCount = Math.max(targetSheet.getLastRow() - targetMeta.headerRowNumber, 0);
  const lastColumn = Math.max(targetSheet.getLastColumn(), columnMap.targetLastColumn || 1);
  if (!rowCount) {
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Sort target rows', 'SKIPPED', 0, 'No target rows to sort.', '', ctx.startedAt);
    return { sortedRows: 0, skipped: true };
  }
  if (branchIdx == null || approvalIdx == null) {
    scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Sort target rows', 'SKIPPED', 0, 'Sort columns missing; Branch and Approval Date target columns are optional.', 'hasBranch=' + (branchIdx != null) + ', hasApprovalDate=' + (approvalIdx != null), ctx.startedAt);
    return { sortedRows: 0, skipped: true };
  }
  scMeilaniCheckStop_(ctx, 'Before sort target rows');
  targetSheet.getRange(targetMeta.headerRowNumber + 1, 1, rowCount, lastColumn).sort([
    { column: branchIdx + 1, ascending: true },
    { column: approvalIdx + 1, ascending: true },
  ]);
  scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Sort target rows', 'SUCCESS', rowCount, 'Target sorted by Branch A-Z and Approval Date oldest first.', 'branchColumn=' + (branchIdx + 1) + ', approvalDateColumn=' + (approvalIdx + 1), ctx.startedAt);
  return { sortedRows: rowCount, skipped: false };
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




function scMeilaniReadSelectedBranches_(overviewSheet, rangeA1) {
  const values = overviewSheet.getRange(rangeA1).getValues();
  const out = [];
  const seen = {};
  values.forEach(function (row) {
    const value = String((row && row[0]) || '').trim();
    const key = scMeilaniNormalizeText_(value);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(value);
  });
  return out;
}

function scMeilaniBackupRepairUpdateValues_(targetSheet, repairCfg) {
  const meta = scMeilaniReadSourceMeta_(targetSheet, [repairCfg.outputHeaders[0], repairCfg.updateHeader]);
  const claimCol = scMeilaniRequireSourceColumn_(meta, 'Claim Number');
  const updateCol = scMeilaniRequireSourceColumn_(meta, repairCfg.updateHeader);
  const backup = {};
  for (let r = meta.headerRowIndex + 1; r < meta.values.length; r++) {
    const row = meta.values[r] || [];
    const key = scMeilaniIdentifierKey_(row[claimCol]);
    if (!key) continue;
    backup[key] = row[updateCol];
  }
  return backup;
}

function scMeilaniCollectRepairStartRecords_(spreadsheet, repairCfg, selectedBranches, ctx) {
  const selected = selectedBranches.map(scMeilaniNormalizeText_);
  const records = [];
  repairCfg.sourceSheetNames.forEach(function (sheetName) {
    const sheet = scMeilaniRequireSheet_(spreadsheet, sheetName);
    const meta = scMeilaniReadSourceMeta_(sheet, ['Claim Number', 'Service Center']);
    const map = scMeilaniResolveRepairStartSourceColumns_(meta);
    for (let r = meta.headerRowIndex + 1; r < meta.values.length; r++) {
      if ((r - meta.headerRowIndex) % 250 === 0) scMeilaniCheckStop_(ctx, 'Collect Repair Start ' + sheetName + ' row ' + (r + 1));
      const row = meta.values[r] || [];
      const claim = row[map.claimNumber];
      if (!scMeilaniIdentifierKey_(claim)) continue;
      const branchValue = map.branch == null ? row[map.serviceCenter] : row[map.branch];
      const resolvedBranch = scMeilaniResolveRepairStartBranch_(branchValue, row[map.serviceCenter]);
      if (selected.indexOf(scMeilaniNormalizeText_(resolvedBranch)) < 0 && selected.indexOf(scMeilaniNormalizeText_(branchValue)) < 0) continue;
      records.push(scMeilaniBuildRepairStartRecord_(row, map, resolvedBranch));
    }
  });
  return scMeilaniDedupeRepairStartRecords_(records);
}

function scMeilaniResolveRepairStartSourceColumns_(meta) {
  return {
    claimNumber: scMeilaniRequireSourceColumn_(meta, 'Claim Number'),
    deviceBrand: scMeilaniFindHeaderIndex_(meta.headerMap, 'Device Brand'),
    deviceType: scMeilaniFindHeaderIndex_(meta.headerMap, 'Device Type'),
    imeiSn: scMeilaniFindHeaderIndex_(meta.headerMap, 'IMEI/SN'),
    serviceCenter: scMeilaniFindHeaderIndex_(meta.headerMap, 'Service Center'),
    branch: scMeilaniFindHeaderIndex_(meta.headerMap, 'Branch'),
    dashboardLink: scMeilaniFindHeaderIndex_(meta.headerMap, 'DB Link'),
    lastStatus: scMeilaniFindHeaderIndex_(meta.headerMap, 'Last Status'),
    statusType: scMeilaniFindHeaderIndex_(meta.headerMap, 'Status Type'),
    lastStatusDate: scMeilaniFindHeaderIndex_(meta.headerMap, 'Last Status Date'),
    lastStatusAging: scMeilaniFindHeaderIndex_(meta.headerMap, 'Last Status Aging'),
  };
}

function scMeilaniBuildRepairStartRecord_(row, map, branchValue) {
  return {
    claimNumber: row[map.claimNumber],
    deviceBrand: map.deviceBrand == null ? '' : row[map.deviceBrand],
    deviceType: map.deviceType == null ? '' : row[map.deviceType],
    imeiSn: map.imeiSn == null ? '' : row[map.imeiSn],
    serviceCenter: map.serviceCenter == null ? branchValue : row[map.serviceCenter],
    dashboardLink: map.dashboardLink == null ? '' : row[map.dashboardLink],
    lastStatus: map.lastStatus == null ? '' : row[map.lastStatus],
    statusType: map.statusType == null ? '' : row[map.statusType],
    lastStatusDate: map.lastStatusDate == null ? '' : row[map.lastStatusDate],
    lastStatusAging: map.lastStatusAging == null ? '' : row[map.lastStatusAging],
  };
}

function scMeilaniDedupeRepairStartRecords_(records) {
  const seen = {};
  const out = [];
  (records || []).forEach(function (record) {
    const key = scMeilaniIdentifierKey_(record.claimNumber);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(record);
  });
  return out;
}

function scMeilaniRewriteRepairSheet_(targetSheet, repairCfg, records, existingUpdates, ctx) {
  const meta = scMeilaniReadSourceMeta_(targetSheet, ['Claim Number']);
  const headerMap = meta.headerMap;
  const targetCols = repairCfg.outputHeaders.map(function (header, index) {
    let idx = scMeilaniFindHeaderIndex_(headerMap, header);
    if (idx == null) {
      idx = index;
      targetSheet.getRange(meta.headerRowNumber, idx + 1).setValue(header);
    }
    return idx;
  });
  const lastCol = Math.max(targetSheet.getLastColumn(), Math.max.apply(null, targetCols) + 1);
  if (targetSheet.getLastRow() > meta.headerRowNumber) {
    targetSheet.getRange(meta.headerRowNumber + 1, 1, targetSheet.getLastRow() - meta.headerRowNumber, lastCol).clearContent();
  }
  const matrix = records.map(function (record) {
    const row = scMeilaniBlankArray_(lastCol);
    const key = scMeilaniIdentifierKey_(record.claimNumber);
    const values = [record.claimNumber, record.deviceBrand, record.deviceType, record.imeiSn, record.serviceCenter, record.dashboardLink, record.lastStatus, record.statusType, record.lastStatusDate, record.lastStatusAging, existingUpdates[key] || ''];
    targetCols.forEach(function (col, index) { row[col] = values[index]; });
    return row;
  });
  if (matrix.length) targetSheet.getRange(meta.headerRowNumber + 1, 1, matrix.length, lastCol).setValues(matrix);
  const statusTypeCol = targetCols[repairCfg.outputHeaders.indexOf(repairCfg.statusTypeHeader)];
  scMeilaniApplyRepairStatusTypeValidation_(targetSheet, meta.headerRowNumber + 1, statusTypeCol, Math.max(matrix.length, targetSheet.getMaxRows() - meta.headerRowNumber));
  scMeilaniLogStep_(ctx.destinationSpreadsheet, ctx.flowName, 'Rewrite Repair sheet', 'SUCCESS', matrix.length, 'Repair sheet refreshed from selected SC branches.', 'target=' + targetSheet.getName(), ctx.startedAt);
  return matrix.length;
}


function scMeilaniApplyRepairStatusTypeValidation_(sheet, startRow, zeroBasedColumn, rowCount) {
  if (zeroBasedColumn == null || zeroBasedColumn < 0) throw new Error('Kolom Status Type tidak ditemukan di sheet Repair.');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SC_MEILANI_CONFIG.repairFlow.statusTypeOptions.slice(), true)
    .setAllowInvalid(false)
    .setHelpText('Pilih Status Type yang tersedia.')
    .build();
  sheet.getRange(startRow, zeroBasedColumn + 1, Math.max(rowCount, 1), 1).setDataValidation(rule);
}

function scMeilaniResolveRepairStartBranch_(branchValue, serviceCenterValue) {
  const direct = String(branchValue == null ? '' : branchValue).trim();
  if (direct) return direct;
  return String(serviceCenterValue == null ? '' : serviceCenterValue).trim();
}

function scMeilaniCopyRepairUpdatesToScUniverse_(sourceSheet, sourceMeta, sourceClaimCol, sourceUpdateCol, targetMaps, syncCfg, ctx) {
  const sourceIndex = scMeilaniBuildRepairUpdateIndex_(sourceMeta, sourceClaimCol, sourceUpdateCol, syncCfg.skipBlankUpdates);
  const targetIndex = scMeilaniBuildScUniverseIndex_(targetMaps);
  const skippedByReason = {};
  let updated = 0;
  let skipped = sourceIndex.blankSkipped;
  let failed = 0;
  if (sourceIndex.blankSkipped) scMeilaniAddReason_(skippedByReason, 'Update from Service Center kosong', 0, '', 'Blank update dilewati supaya tidak menghapus Remarks SC - Universe. count=' + sourceIndex.blankSkipped);

  Object.keys(sourceIndex.rowByKey).forEach(function (claimKey, index) {
    if (index % 100 === 0) scMeilaniCheckStop_(ctx, 'Copy Repair update claim ' + (index + 1));
    const sourceItem = sourceIndex.rowByKey[claimKey];
    const targetItem = targetIndex.rowByKey[claimKey];

    if (sourceIndex.duplicateKeys[claimKey]) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di Repair', sourceItem.rowNumber, sourceItem.claimValue, 'Duplicate source; copy dilewati supaya update deterministic.');
      return;
    }
    if (targetIndex.duplicateKeys[claimKey]) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di SC Universe', targetItem ? targetItem.rowNumber : 0, sourceItem.claimValue, 'Target punya duplicate identifier; copy dilewati supaya tidak menulis ke row ambigu.');
      return;
    }
    if (!targetItem) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Claim Number tidak ada di SC Universe', sourceItem.rowNumber, sourceItem.claimValue, 'Tidak ada pasangan target untuk update Remarks.');
      return;
    }

    try {
      const sourceCell = sourceSheet.getRange(sourceItem.rowNumber, sourceUpdateCol + 1);
      const targetCell = targetItem.sheet.getRange(targetItem.rowNumber, targetItem.remarksCol + 1);
      sourceCell.copyTo(targetCell);
      updated += 1;
    } catch (err) {
      failed += 1;
      scMeilaniAddReason_(skippedByReason, 'Copy Update from Service Center gagal', sourceItem.rowNumber, sourceItem.claimValue, scMeilaniErrorMessage_(err));
    }
  });

  return {
    updated: updated,
    skipped: skipped,
    failed: failed,
    skippedByReason: skippedByReason,
  };
}


function scMeilaniBuildScUniverseIndex_(targetMaps) {
  const rowByKey = {};
  const duplicateKeys = {};
  (targetMaps || []).forEach(function (targetMap) {
    const values = targetMap.meta.values || [];
    for (let r = targetMap.meta.headerRowIndex + 1; r < values.length; r++) {
      const row = values[r] || [];
      const key = scMeilaniIdentifierKey_(row[targetMap.claimCol]);
      if (!key) continue;
      if (rowByKey[key]) {
        duplicateKeys[key] = true;
        continue;
      }
      rowByKey[key] = {
        sheet: targetMap.sheet,
        rowNumber: r + 1,
        remarksCol: targetMap.remarksCol,
      };
    }
  });
  return {
    rowByKey: rowByKey,
    duplicateKeys: duplicateKeys,
  };
}

function scMeilaniBuildRepairUpdateIndex_(meta, claimColumnIndex, updateColumnIndex, skipBlankUpdates) {
  const rowByKey = {};
  const duplicateKeys = {};
  let blankSkipped = 0;

  for (let r = meta.headerRowIndex + 1; r < meta.values.length; r++) {
    const row = meta.values[r] || [];
    const claimValue = row[claimColumnIndex];
    const key = scMeilaniIdentifierKey_(claimValue);
    if (!key) continue;
    const updateValue = row[updateColumnIndex];
    if (skipBlankUpdates && !String(updateValue == null ? '' : updateValue).trim()) {
      blankSkipped += 1;
      continue;
    }
    if (rowByKey[key]) {
      duplicateKeys[key] = true;
      continue;
    }
    rowByKey[key] = {
      rowNumber: r + 1,
      claimValue: claimValue,
    };
  }

  return {
    rowByKey: rowByKey,
    duplicateKeys: duplicateKeys,
    blankSkipped: blankSkipped,
  };
}

function scMeilaniBuildSingleRowIndex_(meta, claimColumnIndex) {
  const rowByKey = {};
  const duplicateKeys = {};
  let rowByKeyCount = 0;

  for (let r = meta.headerRowIndex + 1; r < meta.values.length; r++) {
    const row = meta.values[r] || [];
    const key = scMeilaniIdentifierKey_(row[claimColumnIndex]);
    if (!key) continue;
    if (rowByKey[key]) {
      duplicateKeys[key] = true;
      continue;
    }
    rowByKey[key] = r + 1;
    rowByKeyCount += 1;
  }

  return {
    rowByKey: rowByKey,
    duplicateKeys: duplicateKeys,
    rowByKeyCount: rowByKeyCount,
  };
}

function scMeilaniCopyRemarksByClaim_(sourceSheet, targetSheet, sourceIndex, targetIndex, sourceRemarksColumn, targetRemarksColumn, ctx) {
  const skippedByReason = {};
  let copied = 0;
  let skipped = 0;
  const targetKeys = Object.keys(targetIndex.rowByKey || {});

  for (let i = 0; i < targetKeys.length; i++) {
    if (i % 100 === 0) scMeilaniCheckStop_(ctx, 'Copy Repair Remarks claim ' + (i + 1));
    const key = targetKeys[i];
    const targetRow = targetIndex.rowByKey[key];
    const sourceRow = sourceIndex.rowByKey[key];

    if (targetIndex.duplicateKeys[key]) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di target', targetRow, key, 'Target punya duplicate identifier; copy dilewati supaya tidak menulis ke row ambigu.');
      continue;
    }
    if (sourceIndex.duplicateKeys[key]) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Duplicate Claim Number di source', sourceRow || 0, key, 'Repair punya duplicate identifier; copy dilewati supaya backup tetap deterministic.');
      continue;
    }
    if (!sourceRow) {
      skipped += 1;
      scMeilaniAddReason_(skippedByReason, 'Claim Number tidak ada di Repair', targetRow, key, 'Tidak ada pasangan Repair untuk target SC - Meilani.');
      continue;
    }

    const sourceCell = sourceSheet.getRange(sourceRow, sourceRemarksColumn);
    const targetCell = targetSheet.getRange(targetRow, targetRemarksColumn);
    sourceCell.copyTo(targetCell);
    copied += 1;
  }

  return {
    copied: copied,
    skipped: skipped,
    skippedByReason: skippedByReason,
  };
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

function scMeilaniCollectRowNumbers_(sourceMeta, predicate, ctx) {
  const values = sourceMeta.values || [];
  const out = [];
  for (let r = sourceMeta.headerRowIndex + 1; r < values.length; r++) {
    if ((r - sourceMeta.headerRowIndex) % 250 === 0) scMeilaniCheckStop_(ctx, 'Collect row numbers row ' + (r + 1));
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
