/**
 * Issues canonical IDs using 管理_設定 and LockService.
 */
const ID_SETTING_KEYS = {
  NEXT_SERVICE_SEQ: 'NEXT_SERVICE_SEQ',
  NEXT_CASE_SEQ_PREFIX: 'NEXT_CASE_SEQ_',
};

function issueServiceId() {
  return issueIdWithLock_(ID_SETTING_KEYS.NEXT_SERVICE_SEQ, 'SVC', 6, '採番成功: サービスID');
}

function issueCaseId(fiscalYear) {
  const year = String(fiscalYear || '').trim();
  if (!year) {
    throw new Error('案件ID採番には年度が必要です。');
  }
  return issueIdWithLock_(ID_SETTING_KEYS.NEXT_CASE_SEQ_PREFIX + year, 'REV' + year + '-', 6, '採番成功: 案件ID');
}

function issueIdWithLock_(settingKey, prefix, digitCount, logProcessName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const nextSeq = getNextSequence_(settingKey);
    const issuedId = prefix + padNumber_(nextSeq, digitCount);
    setSettingValue_(settingKey, nextSeq + 1, '次回採番値');
    writeProcessLog(logProcessName, '', LOG_RESULT.OK, settingKey + ' から ' + issuedId + ' を採番しました。');
    return issuedId;
  } catch (error) {
    writeProcessLog('採番失敗', '', LOG_RESULT.ERROR, settingKey + ': ' + error.message);
    throw error;
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      // Lock may not have been acquired. No action needed.
    }
  }
}

function getNextSequence_(settingKey) {
  const value = getSettingValue_(settingKey);
  if (value === '' || value === null || typeof value === 'undefined') {
    setSettingValue_(settingKey, 1, '次回採番値');
    return 1;
  }
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue) || numberValue < 1) {
    throw new Error('管理_設定 の ' + settingKey + ' が正の数値ではありません。');
  }
  return Math.floor(numberValue);
}

function getSettingValue_(settingKey) {
  const sheet = getSettingsSheet_();
  const row = findRowByHeaderValue_(sheet, 'キー', settingKey);
  if (!row) {
    return '';
  }
  const headerMap = getHeaderMap(sheet);
  return sheet.getRange(row, headerMap['値']).getValue();
}

function setSettingValue_(settingKey, value, note) {
  const sheet = getSettingsSheet_();
  const headerMap = getHeaderMap(sheet);
  let row = findRowByHeaderValue_(sheet, 'キー', settingKey);
  if (!row) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, headerMap['キー']).setValue(settingKey);
  }
  sheet.getRange(row, headerMap['値']).setValue(value);
  if (headerMap['備考']) {
    sheet.getRange(row, headerMap['備考']).setValue(note || '');
  }
}

function getSettingsSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, SHEETS.SYSTEM_SETTINGS);
  ensureHeaderRow_(sheet, getHeadersForSheet_(SHEETS.SYSTEM_SETTINGS));
  return sheet;
}

function padNumber_(value, digitCount) {
  return String(value).padStart(digitCount, '0');
}
