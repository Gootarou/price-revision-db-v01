/**
 * Creates missing sheets, keeps existing sheet contents, and maintains header rows.
 */
function setupInitialSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SHEET_ORDER.forEach(function(sheetName) {
    const sheet = getOrCreateSheet_(spreadsheet, sheetName);
    ensureHeaderRow_(sheet, getHeadersForSheet_(sheetName));
  });

  applyDataPriceRevisionCaseValidations_(spreadsheet.getSheetByName(SHEETS.DATA_PRICE_REVISION_CASE));
  applyManagedConditionalFormatting_();
  reorderManagedSheets_(spreadsheet);
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) {
    return existingSheet;
  }
  return spreadsheet.insertSheet(sheetName);
}

function ensureHeaderRow_(sheet, headers) {
  if (!headers || headers.length === 0) {
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const existingValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const existingHeaders = existingValues.filter(function(value) {
    return value !== '';
  });

  if (existingHeaders.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const missingHeaders = headers.filter(function(header) {
    return existingHeaders.indexOf(header) === -1;
  });

  if (missingHeaders.length > 0) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  sheet.setFrozenRows(1);
}

function reorderManagedSheets_(spreadsheet) {
  SHEET_ORDER.forEach(function(sheetName, index) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(index + 1);
  });
}

function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return {};
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  return headers.reduce(function(map, header, index) {
    if (header !== '') {
      map[header] = index + 1;
    }
    return map;
  }, {});
}

function applyDataPriceRevisionCaseValidations_(sheet) {
  if (!sheet) return;
  const headerMap = getHeaderMap(sheet);
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  if (headerMap['確認状態']) {
    sheet.getRange(2, headerMap['確認状態'], maxRows, 1)
      .setDataValidation(buildListValidation_(CONFIRMATION_STATUS_OPTIONS));
  }
  if (headerMap['交渉状況']) {
    sheet.getRange(2, headerMap['交渉状況'], maxRows, 1)
      .setDataValidation(buildListValidation_(NEGOTIATION_STATUS_OPTIONS));
  }
}

function findRowByHeaderValue_(sheet, headerName, targetValue) {
  const headerMap = getHeaderMap(sheet);
  const column = headerMap[headerName];
  const lastRow = sheet.getLastRow();
  if (!column || lastRow < 2) {
    return null;
  }
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0]) === String(targetValue)) {
      return i + 2;
    }
  }
  return null;
}

function setRowFields_(sheet, row, fieldValues) {
  const headerMap = getHeaderMap(sheet);
  Object.keys(fieldValues).forEach(function(header) {
    if (headerMap[header]) {
      sheet.getRange(row, headerMap[header]).setValue(fieldValues[header]);
    }
  });
}

function getRowObject_(sheet, row) {
  const headerMap = getHeaderMap(sheet);
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return rowValuesToObject_(headerMap, values);
}

function rowValuesToObject_(headerMap, values) {
  return Object.keys(headerMap).reduce(function(result, header) {
    result[header] = values[headerMap[header] - 1];
    return result;
  }, {});
}

function deleteRowsByHeaderValue_(sheet, headerName, targetValue) {
  const headerMap = getHeaderMap(sheet);
  const column = headerMap[headerName];
  if (!column) {
    return;
  }
  for (let row = sheet.getLastRow(); row >= 2; row -= 1) {
    if (String(sheet.getRange(row, column).getValue()) === String(targetValue)) {
      sheet.deleteRow(row);
    }
  }
}


function applyManagedConditionalFormatting_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  applyDataPriceRevisionCaseConditionalFormatting_(spreadsheet.getSheetByName(SHEETS.DATA_PRICE_REVISION_CASE));
  applyDocumentTransferCheckConditionalFormatting_(spreadsheet.getSheetByName(SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK));
}

function applyDataPriceRevisionCaseConditionalFormatting_(sheet) {
  applyStatusConditionalFormattingToHeaderSheet_(sheet, [
    '計算結果状態',
    '確認状態',
    '文書転記OK',
    'エラー内容',
    '交渉状況',
  ]);
}

function applyDocumentTransferCheckConditionalFormatting_(sheet) {
  applyStatusConditionalFormattingToHeaderSheet_(sheet, [
    '計算結果状態',
    '確認状態',
    '文書転記OK',
    'エラー内容',
    '交渉状況',
  ]);
}

function applyStatusConditionalFormattingToHeaderSheet_(sheet, targetHeaders) {
  if (!sheet) return;
  const headerMap = getHeaderMap(sheet);
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const managedRanges = targetHeaders
    .filter(function(header) { return !!headerMap[header]; })
    .map(function(header) { return sheet.getRange(2, headerMap[header], maxRows, 1); });
  if (managedRanges.length === 0) return;

  const rules = removeConditionalFormattingRulesForRanges_(sheet.getConditionalFormatRules(), managedRanges);
  addHeaderSheetStatusRules_(rules, sheet, headerMap, maxRows);
  sheet.setConditionalFormatRules(rules);
}

function addHeaderSheetStatusRules_(rules, sheet, headerMap, maxRows) {
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '計算結果状態', CALCULATION_STATUS.ERROR, CONDITIONAL_FORMAT_COLORS.ERROR);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '計算結果状態', CALCULATION_STATUS.OUT_OF_SCOPE, CONDITIONAL_FORMAT_COLORS.MUTED);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '計算結果状態', CALCULATION_STATUS.OK, CONDITIONAL_FORMAT_COLORS.OK);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '確認状態', CONFIRMATION_STATUS.NOT_CONFIRMED, CONDITIONAL_FORMAT_COLORS.WARNING);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '確認状態', CONFIRMATION_STATUS.CONFIRMED, CONDITIONAL_FORMAT_COLORS.OK);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '文書転記OK', TRANSFER_STATUS.NG, CONDITIONAL_FORMAT_COLORS.ERROR);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '文書転記OK', TRANSFER_STATUS.OK, CONDITIONAL_FORMAT_COLORS.OK);
  addNotBlankRuleForHeader_(rules, sheet, headerMap, maxRows, 'エラー内容', CONDITIONAL_FORMAT_COLORS.ERROR);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '交渉状況', NEGOTIATION_STATUS.APPROVED, CONDITIONAL_FORMAT_COLORS.OK);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '交渉状況', NEGOTIATION_STATUS.DECLINED, CONDITIONAL_FORMAT_COLORS.MUTED);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '交渉状況', NEGOTIATION_STATUS.PENDING, CONDITIONAL_FORMAT_COLORS.WARNING);
  addTextRuleForHeader_(rules, sheet, headerMap, maxRows, '交渉状況', NEGOTIATION_STATUS.IN_NEGOTIATION, CONDITIONAL_FORMAT_COLORS.INFO);
}

function addTextRuleForHeader_(rules, sheet, headerMap, maxRows, header, text, background) {
  if (!headerMap[header]) return;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(background)
    .setRanges([sheet.getRange(2, headerMap[header], maxRows, 1)])
    .build());
}

function addNotBlankRuleForHeader_(rules, sheet, headerMap, maxRows, header, background) {
  if (!headerMap[header]) return;
  const column = headerMap[header];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + columnToLetter_(column) + '2<>""')
    .setBackground(background)
    .setRanges([sheet.getRange(2, column, maxRows, 1)])
    .build());
}

function removeConditionalFormattingRulesForRanges_(rules, managedRanges) {
  return rules.filter(function(rule) {
    return !rule.getRanges().some(function(range) {
      return managedRanges.some(function(managedRange) {
        return rangesIntersect_(range, managedRange);
      });
    });
  });
}

function rangesIntersect_(left, right) {
  if (left.getSheet().getName() !== right.getSheet().getName()) return false;
  const leftRowEnd = left.getRow() + left.getNumRows() - 1;
  const rightRowEnd = right.getRow() + right.getNumRows() - 1;
  const leftColumnEnd = left.getColumn() + left.getNumColumns() - 1;
  const rightColumnEnd = right.getColumn() + right.getNumColumns() - 1;
  return left.getRow() <= rightRowEnd
    && right.getRow() <= leftRowEnd
    && left.getColumn() <= rightColumnEnd
    && right.getColumn() <= leftColumnEnd;
}

function columnToLetter_(column) {
  let letter = '';
  let current = column;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    current = Math.floor((current - 1) / 26);
  }
  return letter;
}
