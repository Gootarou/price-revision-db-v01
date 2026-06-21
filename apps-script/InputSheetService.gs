/**
 * Builds and reads the single-case input/edit screen.
 *
 * This sheet is an operation screen only. Canonical data remains in
 * データ_料金改定案件 and 明細_勤務時間根拠.
 */
const INPUT_SHEET_LAYOUT = {
  TITLE_CELL: 'A1',
  OPERATION_START_ROW: 3,
  BASIC_START_ROW: 10,
  CONDITIONS_START_ROW: 20,
  WORK_TIME_START_ROW: 32,
  WORK_TIME_DETAIL_START_ROW: 34,
  WORK_TIME_DETAIL_ROWS: 10,
  RESULT_START_ROW: 47,
  INPUT_COLUMN: 2,
  NOTE_COLUMN: 3,
};

const INPUT_SHEET_STYLE = {
  TITLE_BACKGROUND: '#1f4e78',
  TITLE_FONT: '#ffffff',
  SECTION_BACKGROUND: '#d9eaf7',
  HEADER_BACKGROUND: '#eeeeee',
  INPUT_BACKGROUND: '#fff2cc',
  READONLY_BACKGROUND: '#eaf3f8',
  AUTO_BACKGROUND: '#eeeeee',
  RESULT_BACKGROUND: '#eeeeee',
  NOTE_BACKGROUND: '#f7f7f7',
  BORDER_COLOR: '#b7b7b7',
};

const INPUT_FIELD_BACKGROUND_RULES = {
  OPERATION: {
    '年度': INPUT_SHEET_STYLE.INPUT_BACKGROUND,
    '案件ID': INPUT_SHEET_STYLE.AUTO_BACKGROUND,
    'サービスID': INPUT_SHEET_STYLE.AUTO_BACKGROUND,
    '操作メモ欄': INPUT_SHEET_STYLE.INPUT_BACKGROUND,
  },
  WORK_TIME: {
    '明細行番号': INPUT_SHEET_STYLE.AUTO_BACKGROUND,
    '勤務日数/年': INPUT_SHEET_STYLE.AUTO_BACKGROUND,
    '年間総労働時間': INPUT_SHEET_STYLE.AUTO_BACKGROUND,
  },
};

const INPUT_FIELD_GROUPS = {
  OPERATION: [
    '年度',
    '案件ID',
    'サービスID',
    '操作メモ欄',
  ],
  BASIC: [
    '契約先ID',
    '契約先名',
    '現場ID',
    '現場名',
    'テナント名',
    '業務種別',
    '業務名',
  ],
  CONDITIONS: [
    '前回改定開始日',
    '今回改定依頼日',
    '今回改定開始日',
    '現行料金',
    '現行単位',
    '端数処理方式',
    '手動調整額',
    '調整理由',
  ],
  WORK_TIME: [
    '明細行番号',
    '区分',
    '年間休日',
    '勤務日数/年',
    '労働時間/日',
    '人員/日',
    '年間総労働時間',
    '根拠メモ',
    '使用有無',
    '備考',
  ],
  RESULTS: [
    '年間総労働時間_明細合計',
    '前回参照最低賃金',
    '今回参照最低賃金',
    '最低賃金上昇額',
    '法定福利費計算倍率',
    '労務単価上昇額',
    '暫定年間増加額',
    '暫定月額増加額',
    '端数処理後月額増加額',
    '月額増加額',
    '年間増加額',
    '最終提案料金',
    '提案月額換算',
    '提案年額換算',
    '現行時間単価',
    '最終提案時間単価',
    '時間単価差額',
    '計算結果状態',
    '確認状態',
    '文書転記OK',
    'エラー内容',
  ],
};

/**
 * Initializes the input/edit screen layout without touching canonical data sheets.
 */
function initializeInputCaseEditSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, SHEETS.INPUT_CASE_EDIT);

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(0);
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 1, 190);
  sheet.setColumnWidths(2, 1, 240);
  sheet.setColumnWidths(3, 1, 280);
  sheet.setColumnWidths(4, INPUT_FIELD_GROUPS.WORK_TIME.length - 3, 130);
  sheet.setRowHeights(1, 1, 36);

  sheet.getRange(1, 1, 1, 10)
    .merge()
    .setValue('料金改定案件 入力・編集')
    .setBackground(INPUT_SHEET_STYLE.TITLE_BACKGROUND)
    .setFontColor(INPUT_SHEET_STYLE.TITLE_FONT)
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  buildVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.OPERATION_START_ROW, '操作エリア', INPUT_FIELD_GROUPS.OPERATION);
  buildVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.BASIC_START_ROW, '基本情報エリア', INPUT_FIELD_GROUPS.BASIC);
  buildVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, '日付・料金条件エリア', INPUT_FIELD_GROUPS.CONDITIONS);
  buildWorkTimeInputSection_(sheet);
  buildVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.RESULT_START_ROW, '計算結果表示エリア', INPUT_FIELD_GROUPS.RESULTS);

  applyInputSheetDefaults_(sheet);
  applyWorkTimeHelperFormulas_(sheet);
  applyInputSheetDataValidations_(sheet);
  applyInputSheetNumberFormats_(sheet);
  applyInputSheetLayoutStyle_(sheet);
  applyInputSheetConditionalFormatting_(sheet);
  spreadsheet.setActiveSheet(sheet);
}

/**
 * Reads current values from the input/edit screen. This function does not save data.
 */
function getInputCaseEditValues() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.INPUT_CASE_EDIT);
  if (!sheet) {
    throw new Error('入力画面が見つかりません。入力画面初期化を実行してください。');
  }

  return {
    operation: readVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.OPERATION_START_ROW, INPUT_FIELD_GROUPS.OPERATION),
    basic: readVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.BASIC_START_ROW, INPUT_FIELD_GROUPS.BASIC),
    conditions: readVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, INPUT_FIELD_GROUPS.CONDITIONS),
    workTimeBasis: readWorkTimeInputSection_(sheet),
    results: readVerticalInputSection_(sheet, INPUT_SHEET_LAYOUT.RESULT_START_ROW, INPUT_FIELD_GROUPS.RESULTS),
  };
}

function buildVerticalInputSection_(sheet, startRow, title, fields) {
  sheet.getRange(startRow, 1).setValue(title);
  sheet.getRange(startRow + 1, 1, 1, 3).setValues([['項目', '値', '備考']]);

  const values = fields.map(function(field) {
    return [field, '', ''];
  });
  sheet.getRange(startRow + 2, 1, values.length, 3).setValues(values);
  sheet.getRange(startRow + 1, 1, values.length + 1, 3)
    .setBorder(true, true, true, true, true, true, INPUT_SHEET_STYLE.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
}

function buildWorkTimeInputSection_(sheet) {
  const startRow = INPUT_SHEET_LAYOUT.WORK_TIME_START_ROW;
  const detailStartRow = INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW;
  const fields = INPUT_FIELD_GROUPS.WORK_TIME;

  sheet.getRange(startRow, 1).setValue('勤務時間根拠入力エリア');
  sheet.getRange(detailStartRow - 1, 1, 1, fields.length).setValues([fields]);

  const rows = [];
  for (let i = 0; i < INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS; i += 1) {
    rows.push(['', '', '', '', '', '', '', '', ACTIVE_STATUS.ACTIVE, '']);
  }
  sheet.getRange(detailStartRow, 1, rows.length, fields.length).setValues(rows);
  sheet.getRange(detailStartRow - 1, 1, rows.length + 1, fields.length)
    .setBorder(true, true, true, true, true, true, INPUT_SHEET_STYLE.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
}


function applyInputSheetLayoutStyle_(sheet) {
  const verticalSections = [
    {
      startRow: INPUT_SHEET_LAYOUT.OPERATION_START_ROW,
      fields: INPUT_FIELD_GROUPS.OPERATION,
      inputBackground: INPUT_SHEET_STYLE.INPUT_BACKGROUND,
      fieldBackgrounds: INPUT_FIELD_BACKGROUND_RULES.OPERATION,
    },
    {
      startRow: INPUT_SHEET_LAYOUT.BASIC_START_ROW,
      fields: INPUT_FIELD_GROUPS.BASIC,
      inputBackground: INPUT_SHEET_STYLE.INPUT_BACKGROUND,
    },
    {
      startRow: INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW,
      fields: INPUT_FIELD_GROUPS.CONDITIONS,
      inputBackground: INPUT_SHEET_STYLE.INPUT_BACKGROUND,
    },
    {
      startRow: INPUT_SHEET_LAYOUT.RESULT_START_ROW,
      fields: INPUT_FIELD_GROUPS.RESULTS,
      inputBackground: INPUT_SHEET_STYLE.RESULT_BACKGROUND,
    },
  ];

  verticalSections.forEach(function(section) {
    applyVerticalInputSectionStyle_(sheet, section.startRow, section.fields, section.inputBackground, section.fieldBackgrounds);
  });
  applyWorkTimeInputSectionStyle_(sheet);
}

function applyVerticalInputSectionStyle_(sheet, startRow, fields, inputBackground, fieldBackgrounds) {
  const fieldCount = fields.length;
  sheet.getRange(startRow, 1, 1, 3)
    .merge()
    .setBackground(INPUT_SHEET_STYLE.SECTION_BACKGROUND)
    .setFontWeight('bold')
    .setFontSize(12)
    .setVerticalAlignment('middle');
  sheet.setRowHeights(startRow, 1, 28);

  sheet.getRange(startRow + 1, 1, 1, 3)
    .setBackground(INPUT_SHEET_STYLE.HEADER_BACKGROUND)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(startRow + 2, 1, fieldCount, 1)
    .setBackground(INPUT_SHEET_STYLE.READONLY_BACKGROUND)
    .setFontWeight('bold');
  const inputBackgrounds = fields.map(function(field) {
    return [fieldBackgrounds && fieldBackgrounds[field] ? fieldBackgrounds[field] : inputBackground];
  });
  sheet.getRange(startRow + 2, INPUT_SHEET_LAYOUT.INPUT_COLUMN, fieldCount, 1)
    .setBackgrounds(inputBackgrounds);
  sheet.getRange(startRow + 2, INPUT_SHEET_LAYOUT.NOTE_COLUMN, fieldCount, 1)
    .setBackground(INPUT_SHEET_STYLE.NOTE_BACKGROUND);
  sheet.getRange(startRow + 2, 1, fieldCount, 3)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeights(startRow + 2, fieldCount, 24);
}

function applyWorkTimeInputSectionStyle_(sheet) {
  const startRow = INPUT_SHEET_LAYOUT.WORK_TIME_START_ROW;
  const detailStartRow = INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW;
  const fieldCount = INPUT_FIELD_GROUPS.WORK_TIME.length;
  const dataRows = INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS;

  sheet.getRange(startRow, 1, 1, fieldCount)
    .merge()
    .setBackground(INPUT_SHEET_STYLE.SECTION_BACKGROUND)
    .setFontWeight('bold')
    .setFontSize(12)
    .setVerticalAlignment('middle');
  sheet.setRowHeights(startRow, 1, 28);

  sheet.getRange(detailStartRow - 1, 1, 1, fieldCount)
    .setBackground(INPUT_SHEET_STYLE.HEADER_BACKGROUND)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(detailStartRow, 1, dataRows, fieldCount)
    .setBackground(INPUT_SHEET_STYLE.INPUT_BACKGROUND)
    .setVerticalAlignment('middle')
    .setWrap(true);
  INPUT_FIELD_GROUPS.WORK_TIME.forEach(function(field, index) {
    const background = INPUT_FIELD_BACKGROUND_RULES.WORK_TIME[field] || INPUT_SHEET_STYLE.INPUT_BACKGROUND;
    sheet.getRange(detailStartRow, index + 1, dataRows, 1)
      .setBackground(background);
  });
  sheet.setRowHeights(detailStartRow, dataRows, 26);
}

function applyInputSheetNumberFormats_(sheet) {
  sheet.getRange('B22:B24').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('B25:B25').setNumberFormat('#,##0');
  sheet.getRange('B28:B28').setNumberFormat('#,##0');
  sheet.getRange('C34:G43').setNumberFormat('#,##0.00');
  sheet.getRange('B49:B65').setNumberFormat('#,##0.00');
}

function readVerticalInputSection_(sheet, startRow, fields) {
  const values = sheet.getRange(startRow + 2, INPUT_SHEET_LAYOUT.INPUT_COLUMN, fields.length, 1).getValues();
  return fields.reduce(function(result, field, index) {
    result[field] = values[index][0];
    return result;
  }, {});
}

function readWorkTimeInputSection_(sheet) {
  const fields = INPUT_FIELD_GROUPS.WORK_TIME;
  const rows = sheet
    .getRange(INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW, 1, INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS, fields.length)
    .getValues();

  return rows
    .filter(function(row) {
      return row.some(function(value, index) {
        return fields[index] !== '使用有無' && value !== '';
      });
    })
    .map(function(row) {
      return fields.reduce(function(result, field, index) {
        result[field] = row[index];
        return result;
      }, {});
    });
}

function setInputCaseEditIds(caseId, serviceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.INPUT_CASE_EDIT);
  if (!sheet) {
    throw new Error('入力画面が見つかりません。入力画面初期化を実行してください。');
  }
  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.OPERATION_START_ROW, INPUT_FIELD_GROUPS.OPERATION, {
    '案件ID': caseId,
    'サービスID': serviceId,
  });
}

function setInputCaseEditValues(caseRecord, workTimeDetails) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEETS.INPUT_CASE_EDIT);
  if (!sheet) {
    initializeInputCaseEditSheet();
    sheet = spreadsheet.getSheetByName(SHEETS.INPUT_CASE_EDIT);
  }

  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.OPERATION_START_ROW, INPUT_FIELD_GROUPS.OPERATION, {
    '年度': caseRecord['年度'],
    '案件ID': caseRecord['案件ID'],
    'サービスID': caseRecord['サービスID'],
  });
  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.BASIC_START_ROW, INPUT_FIELD_GROUPS.BASIC, caseRecord);
  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, INPUT_FIELD_GROUPS.CONDITIONS, caseRecord);
  setVerticalInputSectionValuesWithBlank_(sheet, INPUT_SHEET_LAYOUT.RESULT_START_ROW, INPUT_FIELD_GROUPS.RESULTS, caseRecord);
  setWorkTimeInputSectionValues_(sheet, workTimeDetails || []);
  spreadsheet.setActiveSheet(sheet);
}

function applyInputSheetDefaults_(sheet) {
  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.BASIC_START_ROW, INPUT_FIELD_GROUPS.BASIC, {
    '業務種別': BUSINESS_TYPE_OPTIONS[0],
  });
  setVerticalInputSectionValues_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, INPUT_FIELD_GROUPS.CONDITIONS, {
    '現行単位': PRICE_UNIT_OPTIONS[0],
    '端数処理方式': ROUNDING_METHOD.NONE,
  });
}

function applyInputSheetDataValidations_(sheet) {
  applyVerticalInputValidation_(sheet, INPUT_SHEET_LAYOUT.BASIC_START_ROW, INPUT_FIELD_GROUPS.BASIC, '業務種別', BUSINESS_TYPE_OPTIONS);
  applyVerticalInputValidation_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, INPUT_FIELD_GROUPS.CONDITIONS, '現行単位', PRICE_UNIT_OPTIONS);
  applyVerticalInputValidation_(sheet, INPUT_SHEET_LAYOUT.CONDITIONS_START_ROW, INPUT_FIELD_GROUPS.CONDITIONS, '端数処理方式', ROUNDING_METHOD_OPTIONS);

  const activeStatusColumn = INPUT_FIELD_GROUPS.WORK_TIME.indexOf('使用有無') + 1;
  sheet.getRange(INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW, activeStatusColumn, INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS, 1)
    .setDataValidation(buildListValidation_(ACTIVE_STATUS_OPTIONS));
}

function applyVerticalInputValidation_(sheet, startRow, fields, field, options) {
  const fieldIndex = fields.indexOf(field);
  if (fieldIndex === -1) return;
  sheet.getRange(startRow + 2 + fieldIndex, INPUT_SHEET_LAYOUT.INPUT_COLUMN)
    .setDataValidation(buildListValidation_(options));
}

function buildListValidation_(options) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
}

function applyWorkTimeHelperFormulas_(sheet) {
  const detailStartRow = INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW;
  const workDaysColumn = INPUT_FIELD_GROUPS.WORK_TIME.indexOf('勤務日数/年') + 1;
  const annualHoursColumn = INPUT_FIELD_GROUPS.WORK_TIME.indexOf('年間総労働時間') + 1;
  for (let i = 0; i < INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS; i += 1) {
    const row = detailStartRow + i;
    sheet.getRange(row, workDaysColumn).setFormulaR1C1('=IF(RC[-1]="","",365-RC[-1])');
    sheet.getRange(row, annualHoursColumn).setFormulaR1C1('=IF(OR(RC[-3]="",RC[-2]="",RC[-1]=""),"",RC[-3]*RC[-2]*RC[-1])');
  }
}

function setVerticalInputSectionValues_(sheet, startRow, fields, values) {
  fields.forEach(function(field, index) {
    if (Object.prototype.hasOwnProperty.call(values, field)) {
      sheet.getRange(startRow + 2 + index, INPUT_SHEET_LAYOUT.INPUT_COLUMN).setValue(values[field]);
    }
  });
}


function setVerticalInputSectionValuesWithBlank_(sheet, startRow, fields, values) {
  const rows = fields.map(function(field) {
    if (Object.prototype.hasOwnProperty.call(values, field)) {
      return [values[field]];
    }
    return [''];
  });
  sheet.getRange(startRow + 2, INPUT_SHEET_LAYOUT.INPUT_COLUMN, fields.length, 1).setValues(rows);
}

function setWorkTimeInputSectionValues_(sheet, workTimeDetails) {
  const fields = INPUT_FIELD_GROUPS.WORK_TIME;
  const rows = [];
  for (let i = 0; i < INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS; i += 1) {
    const detail = workTimeDetails[i] || {};
    rows.push(fields.map(function(field) {
      if (Object.prototype.hasOwnProperty.call(detail, field)) {
        return detail[field];
      }
      return field === '使用有無' ? ACTIVE_STATUS.ACTIVE : '';
    }));
  }
  sheet.getRange(INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW, 1, INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS, fields.length).setValues(rows);
  applyWorkTimeHelperFormulas_(sheet);
  applyInputSheetDataValidations_(sheet);
  applyInputSheetConditionalFormatting_(sheet);
}

function setInputCaseEditCalculationResults(caseRecord) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.INPUT_CASE_EDIT);
  if (!sheet) {
    throw new Error('入力画面が見つかりません。入力画面初期化を実行してください。');
  }
  setVerticalInputSectionValuesWithBlank_(sheet, INPUT_SHEET_LAYOUT.RESULT_START_ROW, INPUT_FIELD_GROUPS.RESULTS, caseRecord);
}


function applyInputSheetConditionalFormatting_(sheet) {
  if (!sheet) return;
  const managedRanges = getInputSheetConditionalFormattingRanges_(sheet);
  const rules = removeConditionalFormattingRulesForRanges_(sheet.getConditionalFormatRules(), managedRanges);
  addInputResultTextRule_(rules, sheet, '計算結果状態', CALCULATION_STATUS.ERROR, CONDITIONAL_FORMAT_COLORS.ERROR);
  addInputResultTextRule_(rules, sheet, '計算結果状態', CALCULATION_STATUS.OUT_OF_SCOPE, CONDITIONAL_FORMAT_COLORS.MUTED);
  addInputResultTextRule_(rules, sheet, '計算結果状態', CALCULATION_STATUS.OK, CONDITIONAL_FORMAT_COLORS.OK);
  addInputResultTextRule_(rules, sheet, '確認状態', CONFIRMATION_STATUS.NOT_CONFIRMED, CONDITIONAL_FORMAT_COLORS.WARNING);
  addInputResultTextRule_(rules, sheet, '確認状態', CONFIRMATION_STATUS.CONFIRMED, CONDITIONAL_FORMAT_COLORS.OK);
  addInputResultTextRule_(rules, sheet, '文書転記OK', TRANSFER_STATUS.NG, CONDITIONAL_FORMAT_COLORS.ERROR);
  addInputResultTextRule_(rules, sheet, '文書転記OK', TRANSFER_STATUS.OK, CONDITIONAL_FORMAT_COLORS.OK);
  addInputResultNotBlankRule_(rules, sheet, 'エラー内容', CONDITIONAL_FORMAT_COLORS.ERROR);
  addInputWorkTimeInactiveRule_(rules, sheet);
  sheet.setConditionalFormatRules(rules);
}

function getInputSheetConditionalFormattingRanges_(sheet) {
  const ranges = [];
  ['計算結果状態', '確認状態', '文書転記OK', 'エラー内容'].forEach(function(field) {
    const row = getInputResultFieldRow_(field);
    if (row) ranges.push(sheet.getRange(row, INPUT_SHEET_LAYOUT.INPUT_COLUMN));
  });
  ranges.push(sheet.getRange(
    INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW,
    1,
    INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS,
    INPUT_FIELD_GROUPS.WORK_TIME.length
  ));
  return ranges;
}

function addInputResultTextRule_(rules, sheet, field, text, background) {
  const row = getInputResultFieldRow_(field);
  if (!row) return;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(background)
    .setRanges([sheet.getRange(row, INPUT_SHEET_LAYOUT.INPUT_COLUMN)])
    .build());
}

function addInputResultNotBlankRule_(rules, sheet, field, background) {
  const row = getInputResultFieldRow_(field);
  if (!row) return;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B$' + row + '<>""')
    .setBackground(background)
    .setRanges([sheet.getRange(row, INPUT_SHEET_LAYOUT.INPUT_COLUMN)])
    .build());
}

function addInputWorkTimeInactiveRule_(rules, sheet) {
  const activeStatusColumn = INPUT_FIELD_GROUPS.WORK_TIME.indexOf('使用有無') + 1;
  if (activeStatusColumn === 0) return;
  const activeStatusColumnLetter = columnToLetter_(activeStatusColumn);
  const startRow = INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_START_ROW;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + activeStatusColumnLetter + startRow + '="' + ACTIVE_STATUS.INACTIVE + '"')
    .setBackground(CONDITIONAL_FORMAT_COLORS.MUTED)
    .setRanges([sheet.getRange(startRow, 1, INPUT_SHEET_LAYOUT.WORK_TIME_DETAIL_ROWS, INPUT_FIELD_GROUPS.WORK_TIME.length)])
    .build());
}

function getInputResultFieldRow_(field) {
  const fieldIndex = INPUT_FIELD_GROUPS.RESULTS.indexOf(field);
  if (fieldIndex === -1) return null;
  return INPUT_SHEET_LAYOUT.RESULT_START_ROW + 2 + fieldIndex;
}
