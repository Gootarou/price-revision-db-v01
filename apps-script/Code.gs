/**
 * Adds the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_CONFIG.MENU_NAME)
    .addItem(APP_CONFIG.MENU_INITIAL_SETUP, 'runInitialSetup')
    .addItem(APP_CONFIG.MENU_INPUT_SHEET_INITIALIZE, 'runInputSheetInitialize')
    .addItem(APP_CONFIG.MENU_SAVE, 'runSave')
    .addItem(APP_CONFIG.MENU_LOAD_CASE, 'runLoadCase')
    .addItem(APP_CONFIG.MENU_SAVE_AND_CALCULATE, 'runSaveAndCalculate')
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


/**
 * Menu entry point for Phase 2 input/edit screen initialization.
 */
function runInputSheetInitialize() {
  try {
    initializeInputCaseEditSheet();
    writeProcessLog('入力画面初期化', '', LOG_RESULT.OK, SHEETS.INPUT_CASE_EDIT + ' の入力・編集画面を初期化しました。');
    SpreadsheetApp.getUi().alert('入力画面初期化が完了しました。');
  } catch (error) {
    writeProcessLog('入力画面初期化', '', LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('入力画面初期化でエラーが発生しました: ' + error.message);
    throw error;
  }
}


/**
 * Menu entry point for Phase 3 canonical save.
 */
function runSave() {
  saveInputCaseEdit();
}

/**
 * Menu entry point for Phase 3 case load.
 */
function runLoadCase() {
  loadCaseToInputCaseEdit();
}


/**
 * Menu entry point for Phase 4 save and monthly calculation.
 */
function runSaveAndCalculate() {
  saveAndCalculateInputCaseEdit();
}
