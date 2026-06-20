/**
 * Saves and loads canonical price revision cases.
 */
function saveInputCaseEdit() {
  const values = getInputCaseEditValues();
  const caseIdForLog = values.operation['案件ID'] || '';
  try {
    validateRequiredInput_(values);
    const ids = resolveInputIds_(values);
    const now = new Date();

    upsertMasterService_(values, ids.serviceId, now);
    upsertPriceRevisionCase_(values, ids, now);
    replaceWorkTimeBasisDetails_(values, ids, now);
    setInputCaseEditIds(ids.caseId, ids.serviceId);

    writeProcessLog('保存成功', ids.caseId, LOG_RESULT.OK, '入力画面の内容を正本シートへ保存しました。');
    SpreadsheetApp.getUi().alert('保存が完了しました。\n案件ID: ' + ids.caseId + '\nサービスID: ' + ids.serviceId);
  } catch (error) {
    writeProcessLog('保存失敗', caseIdForLog, LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('保存でエラーが発生しました: ' + error.message);
    throw error;
  }
}

function loadCaseToInputCaseEdit() {
  const inputValues = getInputCaseEditValues();
  const caseId = String(inputValues.operation['案件ID'] || '').trim();
  try {
    if (!caseId) {
      writeProcessLog('対象案件ID未入力', '', LOG_RESULT.ERROR, '案件IDを入力してください。');
      throw new Error('案件IDを入力してください。');
    }

    const caseRecord = findCaseById_(caseId);
    if (!caseRecord) {
      writeProcessLog('対象案件ID未存在', caseId, LOG_RESULT.ERROR, '対象案件IDが正本シートに存在しません。');
      throw new Error('対象案件IDが存在しません: ' + caseId);
    }

    const workTimeDetails = findWorkTimeDetailsByCaseId_(caseId);
    setInputCaseEditValues(caseRecord, workTimeDetails);
    writeProcessLog('読込成功', caseId, LOG_RESULT.OK, '正本シートから入力画面へ読み込みました。');
    SpreadsheetApp.getUi().alert('案件読込が完了しました。\n案件ID: ' + caseId);
  } catch (error) {
    writeProcessLog('読込失敗', caseId, LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('案件読込でエラーが発生しました: ' + error.message);
    throw error;
  }
}

function validateRequiredInput_(values) {
  const missing = [];
  [['operation', '年度'], ['basic', '契約先名'], ['basic', '現場名'], ['basic', '業務名'], ['conditions', '現行料金'], ['conditions', '現行単位']].forEach(function(item) {
    const value = values[item[0]][item[1]];
    if (value === '' || value === null || typeof value === 'undefined') {
      missing.push(item[1]);
    }
  });
  if (missing.length > 0) {
    const message = '必須項目不足: ' + missing.join(', ');
    writeProcessLog('必須項目不足', values.operation['案件ID'] || '', LOG_RESULT.ERROR, message);
    throw new Error(message);
  }
}

function resolveInputIds_(values) {
  return {
    serviceId: String(values.operation['サービスID'] || '').trim() || issueServiceId(),
    caseId: String(values.operation['案件ID'] || '').trim() || issueCaseId(values.operation['年度']),
  };
}

function upsertMasterService_(values, serviceId, now) {
  const sheet = getCanonicalSheet_(SHEETS.MASTER_SERVICE);
  const row = findRowByHeaderValue_(sheet, 'サービスID', serviceId) || sheet.getLastRow() + 1;
  const isNew = row > sheet.getLastRow();
  setRowFields_(sheet, row, {
    'サービスID': serviceId,
    '契約先ID': values.basic['契約先ID'],
    '現場ID': values.basic['現場ID'],
    'テナント名': values.basic['テナント名'],
    '業務種別': values.basic['業務種別'],
    '業務名': values.basic['業務名'],
    '計算方式': values.conditions['現行単位'],
    '使用有無': ACTIVE_STATUS.ACTIVE,
    '備考': '',
    '更新日時': now,
  });
  if (isNew) {
    setRowFields_(sheet, row, {'作成日時': now});
  }
}

function upsertPriceRevisionCase_(values, ids, now) {
  const sheet = getCanonicalSheet_(SHEETS.DATA_PRICE_REVISION_CASE);
  const row = findRowByHeaderValue_(sheet, '案件ID', ids.caseId) || sheet.getLastRow() + 1;
  const isNew = row > sheet.getLastRow();
  setRowFields_(sheet, row, {
    '年度': values.operation['年度'],
    '案件ID': ids.caseId,
    'サービスID': ids.serviceId,
    '契約先ID': values.basic['契約先ID'],
    '現場ID': values.basic['現場ID'],
    '契約先名': values.basic['契約先名'],
    '現場名': values.basic['現場名'],
    'テナント名': values.basic['テナント名'],
    '業務種別': values.basic['業務種別'],
    '業務名': values.basic['業務名'],
    '前回改定開始日': values.conditions['前回改定開始日'],
    '今回改定依頼日': values.conditions['今回改定依頼日'],
    '今回改定開始日': values.conditions['今回改定開始日'],
    '現行料金': values.conditions['現行料金'],
    '現行単位': values.conditions['現行単位'],
    '端数処理方式': values.conditions['端数処理方式'],
    '手動調整額': values.conditions['手動調整額'],
    '調整理由': values.conditions['調整理由'],
    '交渉状況': NEGOTIATION_STATUS.NOT_SUBMITTED,
    '計算結果状態': CALCULATION_STATUS.NOT_CALCULATED,
    '確認状態': CONFIRMATION_STATUS.NOT_CONFIRMED,
    '文書転記OK': TRANSFER_STATUS.NG,
    'エラー内容': '',
    '更新日時': now,
  });
  if (isNew) {
    setRowFields_(sheet, row, {'作成日時': now});
  }
}

function findCaseById_(caseId) {
  const sheet = getCanonicalSheet_(SHEETS.DATA_PRICE_REVISION_CASE);
  const row = findRowByHeaderValue_(sheet, '案件ID', caseId);
  return row ? getRowObject_(sheet, row) : null;
}

function getCanonicalSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, sheetName);
  ensureHeaderRow_(sheet, getHeadersForSheet_(sheetName));
  return sheet;
}
