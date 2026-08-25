/**
 * APPLE CLAIM REQ FU ONEDIT
 *
 * Installable OnEdit khusus target sheet REQ FU. Script ini tidak menjalankan
 * Apple Claim Sync dan tidak bereaksi terhadap edit pada workbook/source lain.
 */

const APPLE_CLAIM_ON_EDIT_CONFIG = {
  SHEET_NAME: 'REQ FU',
  HEADER_ROW: 1,
  STATUS_COLUMN: 14,
  CLOSED_BACKGROUND: '#b7b7b7'
};

function onEditAppleClaims(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (
    sheet.getParent().getId() !==
    getRequiredAppleClaimSpreadsheetId_(
      APPLE_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_PROPERTY
    )
  ) return;
  if (sheet.getName() !== APPLE_CLAIM_ON_EDIT_CONFIG.SHEET_NAME) return;

  applyReqFuClosedFormatting_(sheet, e.range);
}

function applyReqFuClosedFormatting_(sheet, editedRange) {
  const statusColumn = APPLE_CLAIM_ON_EDIT_CONFIG.STATUS_COLUMN;
  const editsStatus =
    editedRange.getColumn() <= statusColumn &&
    editedRange.getLastColumn() >= statusColumn;

  if (
    !editsStatus ||
    editedRange.getLastRow() <= APPLE_CLAIM_ON_EDIT_CONFIG.HEADER_ROW
  ) return;

  const firstRow = Math.max(
    editedRange.getRow(),
    APPLE_CLAIM_ON_EDIT_CONFIG.HEADER_ROW + 1
  );
  const rowCount = editedRange.getLastRow() - firstRow + 1;
  const lastColumn = Math.max(sheet.getLastColumn(), statusColumn);
  const statuses = sheet.getRange(firstRow, statusColumn, rowCount, 1).getValues();
  const fontLines = statuses.map(([status]) => {
    const fontLine = String(status ?? '').trim().toUpperCase() === 'CLOSED'
      ? 'line-through'
      : 'none';
    return Array(lastColumn).fill(fontLine);
  });
  const backgrounds = statuses.map(([status]) => {
    const background = String(status ?? '').trim().toUpperCase() === 'CLOSED'
      ? APPLE_CLAIM_ON_EDIT_CONFIG.CLOSED_BACKGROUND
      : null;
    return Array(lastColumn).fill(background);
  });

  const rowsRange = sheet.getRange(firstRow, 1, rowCount, lastColumn);
  rowsRange.setFontLines(fontLines);
  rowsRange.setBackgrounds(backgrounds);
}

function createOnEditTrigger() {
  createAppleClaimOnEditTrigger_(true);
}

function createAppleClaimOnEditTrigger_(showAlert) {
  const handlerName = 'onEditAppleClaims';
  deleteAppleClaimOnEditTriggers_();

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(
      getRequiredAppleClaimSpreadsheetId_(
        APPLE_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_PROPERTY
      )
    )
    .onEdit()
    .create();

  console.log('[TRIGGER] Installable OnEdit khusus target REQ FU berhasil dibuat/reset.');

  if (showAlert) {
    showAlertSafe_(
      'Apple Claim',
      'Installable OnEdit khusus target REQ FU berhasil dibuat / di-reset.'
    );
  }
}

function deleteOnEditTrigger() {
  const deleted = deleteAppleClaimOnEditTriggers_();
  showAlertSafe_('Apple Claim', `${deleted} OnEdit trigger berhasil dihapus.`);
}

function deleteAppleClaimOnEditTriggers_() {
  return deleteTriggersByHandler_('onEditAppleClaims');
}
