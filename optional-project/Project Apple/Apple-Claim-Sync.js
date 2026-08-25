/**
 * APPLE CLAIM SYNC
 *
 * Source : Raw Data
 * Filter :
 * - device_brand contains "Apple"
 * - claim_submitted_datetime >= 2026-07-01
 * - all statuses except EXCLUDED_STATUSES
 *
 * Output : append-only unique Claim Number; existing target rows are preserved
 * Sync   : daily 09:00 trigger + manual recheck
 */

const APPLE_CLAIM_SYNC_CONFIG = {
  SOURCE_SPREADSHEET_PROPERTY: 'APPLE_CLAIM_SOURCE_SPREADSHEET_ID',
  SOURCE_SHEET_NAME: 'Raw Data',
  TARGET_SHEET_NAME: 'General',

  TARGET_SPREADSHEET_PROPERTY: 'APPLE_CLAIM_TARGET_SPREADSHEET_ID',

  SOURCE_HEADER_ROW: 1,
  TARGET_HEADER_SCAN_ROWS: 20,
  TARGET_HEADER: 'Claim Number',
  RUN_TIMESTAMP_CELL: 'G2',
  RUN_STATUS_CELL: 'H2',
  RUN_TIMESTAMP_FORMAT: 'yyyy-mm-dd hh:mm:ss',

  SOURCE_HEADERS: {
    CLAIM_NUMBER: 'claim_number',
    DEVICE_BRAND: 'device_brand',
    STATUS: 'claim_last_status_name',
    SUBMITTED_AT: 'claim_submitted_datetime'
  },

  MIN_SUBMITTED_DATE: new Date(2026, 6, 1), // 1 Jul 2026
  DAILY_TRIGGER_HOUR: 9
};

const APPLE_CLAIM_EXCLUDED_STATUSES = new Set([
  'CLAIM_INITIATE',
  'QOALA_ASK_DETAIL',
  'CUSTOMER_RESUBMIT_DOCUMENT',
  'QOALA_CLAIM_RESUBMIT_DOCUMENT_REQ_QOALA',
  'CLAIM_EXPIRE',
  'QOALA_CLAIM_REOPEN',
  'CLAIM_EXPIRE_WALKIN',

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
    .addItem('Setup / Reset Automation', 'setupAppleClaimSync')
    .addToUi();
}

function setupAppleClaimSync() {
  createAutomationTriggers_();
  const result = syncAppleClaims_('INITIAL_SETUP');

  showAlertSafe_(
    'Apple Claim Sync Berhasil',
    buildResultMessage_(result) +
      '\n\nTrigger OnEdit REQ FU dan sync harian pukul 09:00 sudah aktif.'
  );
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

function dailyAppleClaimSync() {
  syncAppleClaims_('DAILY_09');
}

function syncAppleClaims_(mode) {
  const lock = LockService.getScriptLock();
  const startedAt = new Date();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    console.log(`[${mode}] Mulai Apple Claim Sync`);
    writeAppleClaimRunStatus_(startedAt, 'ON PROGRESS');

    const sourceSS = SpreadsheetApp.openById(
      getRequiredAppleClaimSpreadsheetId_(
        APPLE_CLAIM_SYNC_CONFIG.SOURCE_SPREADSHEET_PROPERTY
      )
    );
    const sourceSheet = sourceSS.getSheetByName(
      APPLE_CLAIM_SYNC_CONFIG.SOURCE_SHEET_NAME
    );

    if (!sourceSheet) {
      throw new Error(`Sheet "${APPLE_CLAIM_SYNC_CONFIG.SOURCE_SHEET_NAME}" tidak ditemukan.`);
    }

    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();

    if (lastRow === 0 || lastColumn === 0) {
      const targetResult = writeTargetClaims_([]);
      const result = Object.assign(emptyResult_(), targetResult);
      writeAppleClaimRunStatus_(startedAt, 'UPDATED');
      return result;
    }

    const values = sourceSheet
      .getRange(
        APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADER_ROW,
        1,
        lastRow - APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADER_ROW + 1,
        lastColumn
      )
      .getValues();

    const headers = values[0].map(normalizeSourceHeader_);
    const claimIndex = findSourceHeaderIndex_(
      headers,
      APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADERS.CLAIM_NUMBER
    );
    const brandIndex = findSourceHeaderIndex_(
      headers,
      APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADERS.DEVICE_BRAND
    );
    const statusIndex = findSourceHeaderIndex_(
      headers,
      APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADERS.STATUS
    );
    const submittedIndex = findSourceHeaderIndex_(
      headers,
      APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADERS.SUBMITTED_AT
    );

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
      if (!submittedAt || submittedAt < APPLE_CLAIM_SYNC_CONFIG.MIN_SUBMITTED_DATE) continue;
      dateEligibleRows++;

      const status = String(row[statusIndex] ?? '').trim().toUpperCase();
      if (!status || APPLE_CLAIM_EXCLUDED_STATUSES.has(status)) continue;

      eligibleRows++;
      uniqueClaims.add(claimNumber);
    }

    const claims = Array.from(uniqueClaims);
    const targetResult = writeTargetClaims_(claims);

    const result = {
      scannedRows,
      appleRows,
      dateEligibleRows,
      eligibleRows,
      uniqueClaims: claims.length,
      existingClaims: targetResult.existingClaims,
      appendedClaims: targetResult.appendedClaims
    };

    writeAppleClaimRunStatus_(startedAt, 'UPDATED');
    console.log(`[${mode}] ${JSON.stringify(result)}`);
    return result;

  } catch (error) {
    if (locked) writeAppleClaimRunStatusSafe_(startedAt, 'FAILED', mode);
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
  const targetSS = SpreadsheetApp.openById(
    getRequiredAppleClaimSpreadsheetId_(
      APPLE_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_PROPERTY
    )
  );
  const targetSheet = targetSS.getSheetByName(APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME);

  if (!targetSheet) {
    throw new Error(
      `Sheet tujuan "${APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME}" tidak ditemukan di target.`
    );
  }

  const targetHeader = findTargetClaimHeader_(targetSheet);
  const targetColumn = targetHeader.column;
  const dataStartRow = targetHeader.row + 1;
  const targetLastRow = targetSheet.getLastRow();
  const existingClaimRows = targetLastRow >= dataStartRow
    ? targetSheet
      .getRange(dataStartRow, targetColumn, targetLastRow - dataStartRow + 1, 1)
      .getDisplayValues()
    : [];
  const existingClaims = new Set();

  existingClaimRows.forEach(([claimNumber], rowOffset) => {
    const normalizedClaim = String(claimNumber ?? '').trim();
    if (!normalizedClaim) return;
    if (existingClaims.has(normalizedClaim)) {
      throw new Error(
        `Duplicate Claim Number "${normalizedClaim}" ditemukan di target ` +
          `sheet "${APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME}" sekitar row ` +
          `${dataStartRow + rowOffset}. Hapus duplicate sebelum sync diulang.`
      );
    }
    existingClaims.add(normalizedClaim);
  });

  const newClaims = claimNumbers.filter(claimNumber => {
    return !existingClaims.has(String(claimNumber ?? '').trim());
  });

  if (!newClaims.length) {
    console.log(
      `[TARGET] Tidak ada Claim Number baru; ${existingClaims.size} row existing dipertahankan.`
    );
    return {
      existingClaims: existingClaims.size,
      appendedClaims: 0
    };
  }

  const appendStartRow = Math.max(targetLastRow + 1, dataStartRow);
  const requiredLastRow = appendStartRow + newClaims.length - 1;
  const currentMaxRows = targetSheet.getMaxRows();

  if (requiredLastRow > currentMaxRows) {
    targetSheet.insertRowsAfter(
      currentMaxRows,
      requiredLastRow - currentMaxRows
    );
  }

  targetSheet
    .getRange(appendStartRow, targetColumn, newClaims.length, 1)
    .setNumberFormat('@')
    .setValues(newClaims.map(claimNumber => [claimNumber]));

  console.log(
    `[TARGET] ${newClaims.length} Claim Number baru ditambahkan; ` +
      `${existingClaims.size} row existing dipertahankan.`
  );
  return {
    existingClaims: existingClaims.size,
    appendedClaims: newClaims.length
  };
}

function createAutomationTriggers_() {
  createAppleClaimOnEditTrigger_(false);
  createDailyTrigger_();
}

function createDailyTrigger_() {
  const handlerName = 'dailyAppleClaimSync';

  deleteTriggersByHandler_(handlerName);

  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .atHour(APPLE_CLAIM_SYNC_CONFIG.DAILY_TRIGGER_HOUR)
    .everyDays(1)
    .create();

  console.log(
    `[TRIGGER] Daily trigger pukul ${String(APPLE_CLAIM_SYNC_CONFIG.DAILY_TRIGGER_HOUR).padStart(2, '0')}:00 berhasil dibuat/reset.`
  );
}

function deleteTriggersByHandler_(handlerName) {
  let deleted = 0;

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handlerName)
    .forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    });

  return deleted;
}

function getRequiredAppleClaimSpreadsheetId_(propertyName) {
  const value = String(
    PropertiesService.getScriptProperties().getProperty(propertyName) || ''
  ).trim();
  if (!value) {
    throw new Error(
      `Script Property "${propertyName}" wajib diisi dengan Spreadsheet ID.`
    );
  }
  const urlMatch = value.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  const spreadsheetId = urlMatch ? urlMatch[1] : value;
  if (!/^[A-Za-z0-9_-]{20,}$/.test(spreadsheetId)) {
    throw new Error(
      `Script Property "${propertyName}" harus berisi Spreadsheet ID atau URL Google Sheets yang valid.`
    );
  }
  return spreadsheetId;
}

function normalizeSourceHeader_(value) {
  return String(value ?? '').trim().toLowerCase();
}

function findSourceHeaderIndex_(headers, headerName) {
  const index = headers.indexOf(normalizeSourceHeader_(headerName));
  if (index === -1) {
    throw new Error(
      `Header source "${headerName}" tidak ditemukan pada row ` +
        `${APPLE_CLAIM_SYNC_CONFIG.SOURCE_HEADER_ROW} sheet ` +
        `"${APPLE_CLAIM_SYNC_CONFIG.SOURCE_SHEET_NAME}".`
    );
  }
  return index;
}

function normalizeTargetHeader_(value) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function findTargetClaimHeader_(targetSheet) {
  const lastRow = targetSheet.getLastRow();
  const lastColumn = targetSheet.getLastColumn();
  const scanRows = Math.min(
    Math.max(lastRow, 1),
    APPLE_CLAIM_SYNC_CONFIG.TARGET_HEADER_SCAN_ROWS
  );
  if (lastColumn === 0) {
    throw new Error(
      `Sheet target "${APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME}" kosong; ` +
        `header "${APPLE_CLAIM_SYNC_CONFIG.TARGET_HEADER}" tidak ditemukan.`
    );
  }

  const values = targetSheet.getRange(1, 1, scanRows, lastColumn).getValues();
  const expected = normalizeTargetHeader_(APPLE_CLAIM_SYNC_CONFIG.TARGET_HEADER);
  for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
    for (let columnOffset = 0; columnOffset < values[rowOffset].length; columnOffset++) {
      if (normalizeTargetHeader_(values[rowOffset][columnOffset]) === expected) {
        console.log(
          `[TARGET] Header Claim Number terdeteksi pada row ${rowOffset + 1}, ` +
            `column ${columnOffset + 1}.`
        );
        return { row: rowOffset + 1, column: columnOffset + 1 };
      }
    }
  }

  throw new Error(
    `Header target "${APPLE_CLAIM_SYNC_CONFIG.TARGET_HEADER}" tidak ditemukan ` +
      `pada row 1-${scanRows} sheet "${APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME}".`
  );
}

function writeAppleClaimRunStatus_(startedAt, status) {
  const targetSS = SpreadsheetApp.openById(
    getRequiredAppleClaimSpreadsheetId_(
      APPLE_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_PROPERTY
    )
  );
  const targetSheet = targetSS.getSheetByName(
    APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME
  );
  if (!targetSheet) {
    throw new Error(
      `Sheet tujuan "${APPLE_CLAIM_SYNC_CONFIG.TARGET_SHEET_NAME}" tidak ditemukan di target.`
    );
  }

  const timestampRange = targetSheet.getRange(
    APPLE_CLAIM_SYNC_CONFIG.RUN_TIMESTAMP_CELL
  );
  timestampRange
    .setValue(startedAt)
    .setNumberFormat(APPLE_CLAIM_SYNC_CONFIG.RUN_TIMESTAMP_FORMAT);
  targetSheet
    .getRange(APPLE_CLAIM_SYNC_CONFIG.RUN_STATUS_CELL)
    .setValue(status);
}

function writeAppleClaimRunStatusSafe_(startedAt, status, mode) {
  try {
    writeAppleClaimRunStatus_(startedAt, status);
  } catch (statusError) {
    console.error(
      `[${mode}] Gagal menulis status ${status}: ${statusError.message}`
    );
  }
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
    uniqueClaims: 0,
    existingClaims: 0,
    appendedClaims: 0
  };
}

function buildResultMessage_(result) {
  return [
    `Rows scanned: ${result.scannedRows}`,
    `Apple rows: ${result.appleRows}`,
    `Date >= Jul 2026: ${result.dateEligibleRows}`,
    `Eligible rows: ${result.eligibleRows}`,
    `Unique Claim Number eligible: ${result.uniqueClaims}`,
    `Existing target claims preserved: ${result.existingClaims}`,
    `New claims appended: ${result.appendedClaims}`
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
