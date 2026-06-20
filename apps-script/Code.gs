/**
 * Adds the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_CONFIG.MENU_NAME)
    .addItem(APP_CONFIG.MENU_INITIAL_SETUP, 'runInitialSetup')
    .addToUi();
}

/**
 * Menu entry point for Phase 1 initial setup.
 */
function runInitialSetup() {
  try {
    setupInitialSheets();
    writeProcessLog('初期セットアップ', '', LOG_RESULT.OK, '必要シート、ヘッダー、シート順を確認しました。');
    SpreadsheetApp.getUi().alert('初期セットアップが完了しました。');
  } catch (error) {
    writeProcessLog('初期セットアップ', '', LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('初期セットアップでエラーが発生しました: ' + error.message);
    throw error;
  }
}
