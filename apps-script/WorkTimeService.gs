/**
 * Maintains canonical work time basis details.
 */
function replaceWorkTimeBasisDetails_(values, ids, now) {
  const sheet = getCanonicalSheet_(SHEETS.DETAIL_WORK_TIME_BASIS);
  deleteRowsByHeaderValue_(sheet, '案件ID', ids.caseId);
  getPersistableWorkTimeRows_(values.workTimeBasis).forEach(function(detail, index) {
    const row = sheet.getLastRow() + 1;
    setRowFields_(sheet, row, {
      '年度': values.operation['年度'],
      '案件ID': ids.caseId,
      'サービスID': ids.serviceId,
      '契約先名': values.basic['契約先名'],
      '現場名': values.basic['現場名'],
      '業務名': values.basic['業務名'],
      '明細行番号': detail['明細行番号'] || index + 1,
      '区分': detail['区分'],
      '年間休日': detail['年間休日'],
      '勤務日数/年': detail['勤務日数/年'],
      '労働時間/日': detail['労働時間/日'],
      '人員/日': detail['人員/日'],
      '年間総労働時間': detail['年間総労働時間'],
      '根拠メモ': detail['根拠メモ'],
      '使用有無': detail['使用有無'] || ACTIVE_STATUS.ACTIVE,
      '備考': detail['備考'],
      '作成日時': now,
      '更新日時': now,
    });
  });
}

function findWorkTimeDetailsByCaseId_(caseId) {
  const sheet = getCanonicalSheet_(SHEETS.DETAIL_WORK_TIME_BASIS);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !headerMap['案件ID']) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
    .map(function(row, index) {
      return {rowNumber: index + 2, values: row};
    })
    .filter(function(rowData) {
      return rowData.values[headerMap['案件ID'] - 1] === caseId;
    })
    .map(function(rowData) {
      return rowValuesToObject_(headerMap, rowData.values);
    })
    .sort(function(a, b) {
      return Number(a['明細行番号'] || 0) - Number(b['明細行番号'] || 0);
    });
}

function getPersistableWorkTimeRows_(details) {
  return (details || []).filter(function(detail) {
    return INPUT_FIELD_GROUPS.WORK_TIME.some(function(field) {
      return field !== '使用有無' && detail[field] !== '' && detail[field] !== null && typeof detail[field] !== 'undefined';
    });
  });
}
