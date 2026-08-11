/**
 * APPLE CLAIM SYNC
 *
 * Source : Raw Data
 * Filter :
 * - device_brand contains "Apple"
 * - claim_submitted_datetime >= 2026-07-01
 * - all statuses except EXCLUDED_STATUSES
 *
 * Output : unique Claim Number only
 * Sync   : installable OnEdit + manual recheck
 */

const CONFIG = {
  SOURCE_SPREADSHEET_ID: '1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc',
  SOURCE_SHEET_NAME: 'Raw Data',

  TARGET_SPREADSHEET_ID: '18_JazMtrwsj7loSfPhtduhvSvR_SuZDrWxtf5rJkjJM',
  TARGET_SHEET_GID: 0,

  HEADER_ROW: 1,
  TARGET_HEADER: 'Claim Number',

  SOURCE_HEADERS: {
    CLAIM_NUMBER: 'claim_number',
    DEVICE_BRAND: 'device_brand',
    STATUS: 'claim_last_status_name',
    SUBMITTED_AT: 'claim_submitted_datetime'
  },

  MIN_SUBMITTED_DATE: new Date(2026, 6, 1) // 1 Jul 2026
};

const EXCLUDED_STATUSES = new Set([
  'CLAIM_INITIATE',
  'QOALA_ASK_DETAIL',
  'CUSTOMER_RESUBMIT_DOCUMENT',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_REQ_QOALA',
  'CLAIM_EXPIRE',
  'QOALA_CLAIM_REOPEN',

  'QOALA_CLAIM_APPROVE_WALKIN',
  'WAITING_WALKIN_START',
  'CLAIM_EXPIRE_WALKIN',
  'QOALA_CLAIM_REOPEN_WALKIN',
  'QOALA_CLAIM_APPROVE_PICKUP',
  'WAITING_PICKUP_START',
  'COURIER_PICKUP_START',
  'COURIER_PICKUP_START_DONE',

  'QOALA_CLAIM_REJECT',
  'QOALA_CLAIM_REJECT_PICKUP',
  'QOALA_CLAIM_REJECT_WALKIN',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_WALKIN',
  'CUSTOMER_REJECT_PAYMENT_DEDUCTIBLE_EXCESS_FEE_PICKUP',
  'INSURANCE_CLAIM_REJECT_WALKIN',
  'INSURANCE_CLAIM_REJECT_PICKUP',
  'SERVICE_CENTER_CLAIM_WAITING_WALKIN_REJECT',
  'SERVICE_CENTER_CLAIM_DONE_REJECT',
  'SERVICE_CENTER_CLAIM_WAITING_PICKUP_REJECT',
  'COURIER_CLAIM_PICKUP_REJECT',
  'COURIER_CLAIM_PICKUP_REJECT_DONE',
  'CLAIM_CANCELLED'
]);

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Apple Claim')
    .addItem('Recheck Claim Number', 'manualRecheckAppleClaims')
    .addSeparator()
    .addItem('Setup / Reset OnEdit Trigger', 'createOnEditTrigger')
    .addToUi();
}

function setupAppleClaimSync() {
  createOnEditTrigger_(false);
  const result = syncAppleClaims_('INITIAL_SETUP');

  showAlertSafe_(
    'Apple Claim Sync Berhasil',
    buildResultMessage_(result) + '\n\nInstallable OnEdit trigger sudah aktif.'
  );
}

function onEditAppleClaims(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const ss = sheet.getParent();

  if (ss.getId() !== CONFIG.SOURCE_SPREADSHEET_ID) return;
  if (sheet.getName() !== CONFIG.SOURCE_SHEET_NAME) return;
  if (e.range.getLastRow() <= CONFIG.HEADER_ROW) return;

  syncAppleClaims_('ON_EDIT');
}

function manualRecheckAppleClaims() {
  try {
    showToastSafe_('Sedang melakukan full recheck...', 'Apple Claim', 5);

    const result = syncAppleClaims_('MANUAL_RECHECK');

    showToastSafe_(
      `${result.uniqueClaims} Claim Number berhasil disinkronkan.`,
      'Apple Claim',
      5
    );

    showAlertSafe_(
      'Apple Claim Recheck Selesai',
      buildResultMessage_(result)
    );
  } catch (error) {
    console.error(`[MANUAL_RECHECK] ERROR: ${error.message}`);
    showAlertSafe_('Apple Claim Recheck Gagal', error.message);
    throw error;
  }
}

function syncAppleClaims_(mode) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    console.log(`[${mode}] Mulai Apple Claim Sync`);

    const sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSS.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      throw new Error(`Sheet "${CONFIG.SOURCE_SHEET_NAME}" tidak ditemukan.`);
    }

    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();

    if (lastRow <= CONFIG.HEADER_ROW || lastColumn === 0) {
      writeTargetClaims_([]);
      return emptyResult_();
    }

    const values = sourceSheet
      .getRange(
        CONFIG.HEADER_ROW,
        1,
        lastRow - CONFIG.HEADER_ROW + 1,
        lastColumn
      )
      .getValues();

    const headers = values[0].map(normalizeHeader_);

    const claimIndex = findHeaderIndex_(headers, CONFIG.SOURCE_HEADERS.CLAIM_NUMBER);
    const brandIndex = findHeaderIndex_(headers, CONFIG.SOURCE_HEADERS.DEVICE_BRAND);
    const statusIndex = findHeaderIndex_(headers, CONFIG.SOURCE_HEADERS.STATUS);
    const submittedIndex = findHeaderIndex_(headers, CONFIG.SOURCE_HEADERS.SUBMITTED_AT);

    const uniqueClaims = new Set();

    let scannedRows = 0;
    let appleRows = 0;
    let dateEligibleRows = 0;
    let eligibleRows = 0;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      scannedRows++;

      const claimNumber = String(row[claimIndex] ?? '').trim();
      if (!claimNumber) continue;

      const deviceBrand = String(row[brandIndex] ?? '').trim().toLowerCase();
      if (!deviceBrand.includes('apple')) continue;
      appleRows++;

      const submittedAt = parseSheetDate_(row[submittedIndex]);
      if (!submittedAt || submittedAt < CONFIG.MIN_SUBMITTED_DATE) continue;
      dateEligibleRows++;

      const status = String(row[statusIndex] ?? '').trim().toUpperCase();
      if (!status || EXCLUDED_STATUSES.has(status)) continue;

      eligibleRows++;
      uniqueClaims.add(claimNumber);
    }

    const claims = Array.from(uniqueClaims);
    writeTargetClaims_(claims);

    const result = {
      scannedRows,
      appleRows,
      dateEligibleRows,
      eligibleRows,
      uniqueClaims: claims.length
    };

    console.log(`[${mode}] ${JSON.stringify(result)}`);
    return result;

  } catch (error) {
    console.error(`[${mode}] ERROR: ${error.message}`);
    console.error(error.stack || '');
    throw error;
  } finally {
    if (locked) {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  }
}

function writeTargetClaims_(claimNumbers) {
  const targetSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  const targetSheet = targetSS
    .getSheets()
    .find(sheet => sheet.getSheetId() === CONFIG.TARGET_SHEET_GID);

  if (!targetSheet) {
    throw new Error(`Sheet tujuan GID ${CONFIG.TARGET_SHEET_GID} tidak ditemukan.`);
  }

  const lastColumn = Math.max(targetSheet.getLastColumn(), 1);
  const headers = targetSheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn)
    .getValues()[0]
    .map(normalizeHeader_);

  const headerIndex = headers.indexOf(normalizeHeader_(CONFIG.TARGET_HEADER));

  if (headerIndex === -1) {
    throw new Error(`Kolom "${CONFIG.TARGET_HEADER}" tidak ditemukan di target.`);
  }

  const targetColumn = headerIndex + 1;
  const dataStartRow = CONFIG.HEADER_ROW + 1;
  const targetLastRow = targetSheet.getLastRow();

  if (targetLastRow >= dataStartRow) {
    targetSheet
      .getRange(
        dataStartRow,
        targetColumn,
        targetLastRow - dataStartRow + 1,
        1
      )
      .clearContent();
  }

  if (!claimNumbers.length) {
    console.log('[TARGET] Tidak ada eligible Claim Number.');
    return;
  }

  const requiredLastRow = dataStartRow + claimNumbers.length - 1;
  const currentMaxRows = targetSheet.getMaxRows();

  if (requiredLastRow > currentMaxRows) {
    targetSheet.insertRowsAfter(
      currentMaxRows,
      requiredLastRow - currentMaxRows
    );
  }

  targetSheet
    .getRange(dataStartRow, targetColumn, claimNumbers.length, 1)
    .setValues(claimNumbers.map(claimNumber => [claimNumber]));

  console.log(`[TARGET] ${claimNumbers.length} Claim Number ditulis.`);
}

function createOnEditTrigger() {
  createOnEditTrigger_(true);
}

function createOnEditTrigger_(showAlert) {
  const handlerName = 'onEditAppleClaims';

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handlerName)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(CONFIG.SOURCE_SPREADSHEET_ID)
    .onEdit()
    .create();

  console.log('[TRIGGER] Installable OnEdit trigger berhasil dibuat/reset.');

  if (showAlert) {
    showAlertSafe_(
      'Apple Claim',
      'Installable OnEdit trigger berhasil dibuat / di-reset.'
    );
  }
}

function deleteOnEditTrigger() {
  const handlerName = 'onEditAppleClaims';
  let deleted = 0;

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handlerName)
    .forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    });

  showAlertSafe_('Apple Claim', `${deleted} OnEdit trigger berhasil dihapus.`);
}

function normalizeHeader_(value) {
  return String(value ?? '').trim().toLowerCase();
}

function findHeaderIndex_(headers, headerName) {
  const index = headers.indexOf(normalizeHeader_(headerName));

  if (index === -1) {
    throw new Error(
      `Header "${headerName}" tidak ditemukan di sheet "${CONFIG.SOURCE_SHEET_NAME}".`
    );
  }

  return index;
}

function parseSheetDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return normalizeDate_(value);
  }

  const text = String(value ?? '').trim();
  if (!text) return null;

  let match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (match) {
    return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : normalizeDate_(parsed);
}

function normalizeDate_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function emptyResult_() {
  return {
    scannedRows: 0,
    appleRows: 0,
    dateEligibleRows: 0,
    eligibleRows: 0,
    uniqueClaims: 0
  };
}

function buildResultMessage_(result) {
  return [
    `Rows scanned: ${result.scannedRows}`,
    `Apple rows: ${result.appleRows}`,
    `Date >= Jul 2026: ${result.dateEligibleRows}`,
    `Eligible rows: ${result.eligibleRows}`,
    `Unique Claim Number: ${result.uniqueClaims}`
  ].join('\n');
}

function showToastSafe_(message, title, seconds) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) return;
    spreadsheet.toast(message, title || 'Apple Claim', seconds || 5);
  } catch (error) {
    console.log(`[UI] Toast skipped: ${error.message}`);
  }
}

function showAlertSafe_(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (error) {
    console.log(`[UI] ${title}: ${message}`);
    console.log(`[UI] Alert skipped: ${error.message}`);
  }
}
