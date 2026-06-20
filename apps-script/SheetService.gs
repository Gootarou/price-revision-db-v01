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
