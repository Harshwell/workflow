/**
 * SAMSUNG CLAIM REQ FU ONEDIT
 *
 * Installable OnEdit khusus target sheet REQ FU. Script ini tidak menjalankan
 * Samsung Claim Sync dan tidak bereaksi terhadap edit pada workbook/source lain.
 */

const SAMSUNG_CLAIM_ON_EDIT_CONFIG = {
  SHEET_NAME: 'REQ FU',
  HEADER_ROW: 1,
  STATUS_COLUMN: 14,
  CLOSED_BACKGROUND: '#b7b7b7'
};

function onEditSamsungClaims(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (
    sheet.getParent().getId() !==
    SAMSUNG_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_ID
  ) return;
  if (sheet.getName() !== SAMSUNG_CLAIM_ON_EDIT_CONFIG.SHEET_NAME) return;

  applyReqFuClosedFormatting_(sheet, e.range);
}

function applyReqFuClosedFormatting_(sheet, editedRange) {
  const statusColumn = SAMSUNG_CLAIM_ON_EDIT_CONFIG.STATUS_COLUMN;
  const editsStatus =
    editedRange.getColumn() <= statusColumn &&
    editedRange.getLastColumn() >= statusColumn;

  if (
    !editsStatus ||
    editedRange.getLastRow() <= SAMSUNG_CLAIM_ON_EDIT_CONFIG.HEADER_ROW
  ) return;

  const firstRow = Math.max(
    editedRange.getRow(),
    SAMSUNG_CLAIM_ON_EDIT_CONFIG.HEADER_ROW + 1
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
      ? SAMSUNG_CLAIM_ON_EDIT_CONFIG.CLOSED_BACKGROUND
      : null;
    return Array(lastColumn).fill(background);
  });

  const rowsRange = sheet.getRange(firstRow, 1, rowCount, lastColumn);
  rowsRange.setFontLines(fontLines);
  rowsRange.setBackgrounds(backgrounds);
}

function createOnEditTrigger() {
  createSamsungClaimOnEditTrigger_(true);
}

function createSamsungClaimOnEditTrigger_(showAlert) {
  const handlerName = 'onEditSamsungClaims';
  deleteSamsungClaimOnEditTriggers_();

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(
      SAMSUNG_CLAIM_SYNC_CONFIG.TARGET_SPREADSHEET_ID
    )
    .onEdit()
    .create();

  console.log('[TRIGGER] Installable OnEdit khusus target REQ FU berhasil dibuat/reset.');

  if (showAlert) {
    showAlertSafe_(
      'Samsung Claim',
      'Installable OnEdit khusus target REQ FU berhasil dibuat / di-reset.'
    );
  }
}

function deleteOnEditTrigger() {
  const deleted = deleteSamsungClaimOnEditTriggers_();
  showAlertSafe_('Samsung Claim', `${deleted} OnEdit trigger berhasil dihapus.`);
}

function deleteSamsungClaimOnEditTriggers_() {
  return deleteTriggersByHandler_('onEditSamsungClaims');
}
