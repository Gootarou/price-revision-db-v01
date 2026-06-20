/**
 * Read-only sidebar for checking work time basis details.
 */
function showWorkTimeBasisSidebar() {
  let caseIdForLog = '';
  try {
    const payload = buildWorkTimeSidebarPayloadFromSelection_();
    caseIdForLog = payload.caseInfo['案件ID'];
    const template = HtmlService.createTemplateFromFile('WorkTimeSidebar');
    template.payloadJson = JSON.stringify(payload);
    const html = template.evaluate()
      .setTitle('勤務時間根拠確認')
      .setWidth(520);
    SpreadsheetApp.getUi().showSidebar(html);

    if (!payload.summary.isMatched) {
      writeProcessLog('明細合計不一致', caseIdForLog, LOG_RESULT.WARN, '明細側合計とデータ側合計が一致していません。');
    }
    writeProcessLog('勤務時間根拠サイドバー表示成功', caseIdForLog, LOG_RESULT.OK, '勤務時間根拠確認サイドバーを表示しました。');
    return payload;
  } catch (error) {
    writeProcessLog('勤務時間根拠サイドバー表示失敗', caseIdForLog, LOG_RESULT.ERROR, error.message);
    SpreadsheetApp.getUi().alert('勤務時間根拠の確認でエラーが発生しました: ' + error.message);
    throw error;
  }
}

function buildWorkTimeSidebarPayloadFromSelection_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== SHEETS.DATA_PRICE_REVISION_CASE) {
    writeProcessLog('対象シート不正', '', LOG_RESULT.ERROR, SHEETS.DATA_PRICE_REVISION_CASE + ' のデータ行を選択してください。');
    throw new Error(SHEETS.DATA_PRICE_REVISION_CASE + ' のデータ行を選択してください。');
  }

  const activeRange = activeSheet.getActiveRange();
  const selectedRow = activeRange ? activeRange.getRow() : 0;
  if (selectedRow <= 1) {
    writeProcessLog('対象シート不正', '', LOG_RESULT.ERROR, 'ヘッダー行ではなくデータ行を選択してください。');
    throw new Error('ヘッダー行ではなくデータ行を選択してください。');
  }

  const caseRecord = getRowObject_(activeSheet, selectedRow);
  const caseId = String(caseRecord['案件ID'] || '').trim();
  if (!caseId) {
    writeProcessLog('案件ID未取得', '', LOG_RESULT.ERROR, '選択行から案件IDを取得できません。');
    throw new Error('選択行から案件IDを取得できません。');
  }

  const caseSheet = getCanonicalSheet_(SHEETS.DATA_PRICE_REVISION_CASE);
  const canonicalRow = findRowByHeaderValue_(caseSheet, '案件ID', caseId);
  if (!canonicalRow) {
    writeProcessLog('案件ID未存在', caseId, LOG_RESULT.ERROR, '案件IDが正本シートに存在しません。');
    throw new Error('案件IDが正本シートに存在しません: ' + caseId);
  }

  const canonicalCaseRecord = getRowObject_(caseSheet, canonicalRow);
  const details = findWorkTimeDetailsByCaseId_(caseId);
  if (details.length === 0) {
    writeProcessLog('勤務時間根拠未登録', caseId, LOG_RESULT.ERROR, '対象案件IDの勤務時間根拠明細がありません。');
    throw new Error('対象案件IDの勤務時間根拠明細がありません: ' + caseId);
  }

  const activeDetails = details.filter(function(detail) {
    return String(detail['使用有無'] || ACTIVE_STATUS.ACTIVE).trim() === ACTIVE_STATUS.ACTIVE;
  });
  const inactiveDetails = details.filter(function(detail) {
    return String(detail['使用有無'] || ACTIVE_STATUS.ACTIVE).trim() !== ACTIVE_STATUS.ACTIVE;
  });
  const detailTotal = activeDetails.reduce(function(total, detail) {
    return total + toComparableNumber_(detail['年間総労働時間']);
  }, 0);
  const dataTotal = toComparableNumber_(canonicalCaseRecord['年間総労働時間_明細合計']);
  const difference = detailTotal - dataTotal;

  return {
    readOnlyMessage: 'このサイドバーは確認専用です。編集・保存・計算実行はできません。',
    caseInfo: pickFields_(canonicalCaseRecord, [
      '案件ID',
      'サービスID',
      '年度',
      '契約先名',
      '現場名',
      '業務名',
      '現行単位',
      '計算結果状態',
      '確認状態',
      '文書転記OK',
      'エラー内容',
    ]),
    summary: {
      detailCount: details.length,
      activeDetailCount: activeDetails.length,
      inactiveDetailCount: inactiveDetails.length,
      activeAnnualWorkTimeTotal: detailTotal,
      dataAnnualWorkTimeTotal: dataTotal,
      difference: difference,
      isMatched: detailTotal === dataTotal,
      judgment: detailTotal === dataTotal ? '一致' : '不一致',
    },
    details: details.map(function(detail) {
      return pickFields_(detail, [
        '明細行番号',
        '区分',
        '年間休日',
        '勤務日数/年',
        '労働時間/日',
        '人員/日',
        '年間総労働時間',
        '使用有無',
        '根拠メモ',
        '備考',
      ]);
    }),
  };
}

function pickFields_(record, fields) {
  return fields.reduce(function(result, field) {
    result[field] = record[field] === null || typeof record[field] === 'undefined' ? '' : record[field];
    return result;
  }, {});
}

function toComparableNumber_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return 0;
  }
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : 0;
}
