/**
 * Creates missing sheets, keeps existing sheet contents, and maintains header rows.
 */
function setupInitialSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SHEET_ORDER.forEach(function(sheetName) {
    const sheet = getOrCreateSheet_(spreadsheet, sheetName);
    ensureHeaderRow_(sheet, getHeadersForSheet_(sheetName));
  });

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
