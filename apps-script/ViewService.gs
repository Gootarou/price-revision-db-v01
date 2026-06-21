/**
 * Regenerates the read-only document transfer check view from canonical case data.
 */
function updateDocumentTransferCheckView() {
  try {
    const result = updateDocumentTransferCheckViewCore_();
    if (result.outputCount === 0) {
      writeProcessLog('出力対象データなし', '', LOG_RESULT.WARN, SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK + ' の出力対象データがありません。');
      SpreadsheetApp.getUi().alert('文書転記前確認を更新しました。\n出力対象データはありません。');
      return result;
    }

    writeProcessLog('文書転記前確認更新成功', '', LOG_RESULT.OK, SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK + ' を更新しました。出力件数: ' + result.outputCount + '件');
    SpreadsheetApp.getUi().alert('文書転記前確認を更新しました。\n出力件数: ' + result.outputCount + '件');
    return result;
  } catch (error) {
    const processName = error.processName || '文書転記前確認更新失敗';
    writeProcessLog(processName, '', LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('文書転記前確認の更新でエラーが発生しました: ' + error.message);
    throw error;
  }
}

function updateDocumentTransferCheckViewCore_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(SHEETS.DATA_PRICE_REVISION_CASE);
  const targetSheet = spreadsheet.getSheetByName(SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK);

  if (!sourceSheet || !targetSheet) {
    throw new DocumentTransferViewError_('必要シート不足', '必要なシートが不足しています。');
  }

  const targetHeaders = getHeadersForSheet_(SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK);
  if (!targetHeaders || targetHeaders.length === 0) {
    throw new DocumentTransferViewError_('ヘッダー不足', SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK + ' のヘッダー定義がありません。');
  }

  ensureHeaderRow_(targetSheet, targetHeaders);
  const sourceHeaderMap = getHeaderMap(sourceSheet);
  if (!sourceHeaderMap['案件ID']) {
    throw new DocumentTransferViewError_('ヘッダー不足', SHEETS.DATA_PRICE_REVISION_CASE + ' に 案件ID ヘッダーがありません。');
  }

  const records = collectDocumentTransferCheckRecords_(sourceSheet, sourceHeaderMap, targetHeaders);
  records.sort(compareDocumentTransferCheckRecords_);
  replaceDocumentTransferCheckRows_(targetSheet, targetHeaders, records);
  formatDocumentTransferCheckSheet_(targetSheet, targetHeaders);
  applyDocumentTransferCheckConditionalFormatting_(targetSheet);

  return {
    outputCount: records.length,
  };
}

function collectDocumentTransferCheckRecords_(sourceSheet, sourceHeaderMap, targetHeaders) {
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();
  if (lastRow < 2 || lastColumn === 0) {
    return [];
  }

  const values = sourceSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  return values.reduce(function(records, row) {
    const record = rowValuesToObject_(sourceHeaderMap, row);
    if (String(record['案件ID'] || '').trim() === '') {
      return records;
    }
    records.push(pickFields_(record, targetHeaders));
    return records;
  }, []);
}

function replaceDocumentTransferCheckRows_(targetSheet, targetHeaders, records) {
  const lastRow = targetSheet.getLastRow();
  const lastColumn = Math.max(targetSheet.getLastColumn(), targetHeaders.length);
  if (lastRow > 1 && lastColumn > 0) {
    targetSheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }

  targetSheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
  if (records.length === 0) {
    return;
  }

  const outputValues = records.map(function(record) {
    return targetHeaders.map(function(header) {
      return record[header] === null || typeof record[header] === 'undefined' ? '' : record[header];
    });
  });
  targetSheet.getRange(2, 1, outputValues.length, targetHeaders.length).setValues(outputValues);
}

function formatDocumentTransferCheckSheet_(sheet, headers) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#d9ead3')
    .setFontWeight('bold');

  const maxRows = Math.max(sheet.getLastRow() - 1, 1);
  headers.forEach(function(header, index) {
    const column = index + 1;
    if (isDocumentTransferCheckDateColumn_(header)) {
      sheet.getRange(2, column, maxRows, 1).setNumberFormat('yyyy/mm/dd');
    } else if (isDocumentTransferCheckDecimalColumn_(header)) {
      sheet.getRange(2, column, maxRows, 1).setNumberFormat('#,##0.00');
    } else if (isDocumentTransferCheckIntegerColumn_(header)) {
      sheet.getRange(2, column, maxRows, 1).setNumberFormat('#,##0');
    }
  });

  sheet.autoResizeColumns(1, headers.length);
}

function compareDocumentTransferCheckRecords_(left, right) {
  return compareDocumentTransferValue_(left['年度'], right['年度'])
    || compareDocumentTransferValue_(left['契約先名'], right['契約先名'])
    || compareDocumentTransferValue_(left['現場名'], right['現場名'])
    || compareDocumentTransferValue_(left['業務名'], right['業務名'])
    || compareDocumentTransferValue_(left['案件ID'], right['案件ID']);
}

function compareDocumentTransferValue_(left, right) {
  const leftValue = left === null || typeof left === 'undefined' ? '' : left;
  const rightValue = right === null || typeof right === 'undefined' ? '' : right;
  if (leftValue instanceof Date && rightValue instanceof Date) {
    return leftValue.getTime() - rightValue.getTime();
  }
  if (typeof leftValue === 'number' && typeof rightValue === 'number') {
    return leftValue - rightValue;
  }
  return String(leftValue).localeCompare(String(rightValue), 'ja');
}

function isDocumentTransferCheckDateColumn_(header) {
  return ['前回改定開始日', '今回改定依頼日', '今回改定開始日', '更新日時'].indexOf(header) !== -1;
}

function isDocumentTransferCheckDecimalColumn_(header) {
  return [
    '法定福利費計算倍率',
    '現行時間単価',
    '最終提案時間単価',
    '時間単価差額',
  ].indexOf(header) !== -1;
}

function isDocumentTransferCheckIntegerColumn_(header) {
  return [
    '現行料金',
    '年間総労働時間_明細合計',
    '前回参照最低賃金',
    '今回参照最低賃金',
    '最低賃金上昇額',
    '月額増加額',
    '年間増加額',
    '最終提案料金',
    '提案年額換算',
  ].indexOf(header) !== -1;
}

function DocumentTransferViewError_(processName, message) {
  this.name = 'DocumentTransferViewError_';
  this.processName = processName;
  this.message = message;
}
DocumentTransferViewError_.prototype = new Error();
