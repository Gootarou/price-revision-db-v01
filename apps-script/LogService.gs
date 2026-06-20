/**
 * Appends process logs to the system log sheet.
 */
function writeProcessLog(processName, targetCaseId, result, message) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateSheet_(spreadsheet, SHEETS.SYSTEM_LOG);
  ensureHeaderRow_(logSheet, getHeadersForSheet_(SHEETS.SYSTEM_LOG));

  logSheet.appendRow([
    new Date(),
    processName || '',
    targetCaseId || '',
    result || '',
    message || '',
    getActiveUserEmail_(),
  ]);
}

function getActiveUserEmail_() {
  const user = Session.getActiveUser();
  return user ? user.getEmail() : '';
}
