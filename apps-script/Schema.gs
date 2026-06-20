/**
 * Header definitions for all sheets used by Phase 1 setup.
 */
function getSheetHeaders_() {
  const headers = {};

  headers[SHEETS.INPUT_CASE_EDIT] = [
  '項目',
  '値',
  '備考',
];


  headers[SHEETS.DATA_PRICE_REVISION_CASE] = [
  '年度',
  '案件ID',
  'サービスID',
  '契約先ID',
  '現場ID',
  '契約先名',
  '現場名',
  'テナント名',
  '業務種別',
  '業務名',
  '前回改定開始日',
  '今回改定依頼日',
  '今回改定開始日',
  '現行料金',
  '現行単位',
  '現行月額換算',
  '現行年額換算',
  '年間総労働時間_明細合計',
  '前回参照最低賃金',
  '今回参照最低賃金',
  '最低賃金上昇額',
  '法定福利費計算倍率',
  '労務単価上昇額',
  '暫定年間増加額',
  '暫定月額増加額',
  '端数処理方式',
  '端数処理後月額増加額',
  '手動調整額',
  '調整理由',
  '月額増加額',
  '年間増加額',
  '最終提案料金',
  '提案月額換算',
  '提案年額換算',
  '現行時間単価',
  '最終提案時間単価',
  '時間単価差額',
  '承認後料金',
  '交渉状況',
  '計算結果状態',
  '確認状態',
  '文書転記OK',
  'エラー内容',
  '作成日時',
  '更新日時',
];


  headers[SHEETS.DETAIL_WORK_TIME_BASIS] = [
  '年度',
  '案件ID',
  'サービスID',
  '契約先名',
  '現場名',
  '業務名',
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
  '作成日時',
  '更新日時',
];


  headers[SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK] = [
  '年度',
  '案件ID',
  '契約先名',
  '現場名',
  'テナント名',
  '業務名',
  '前回改定開始日',
  '今回改定依頼日',
  '今回改定開始日',
  '現行料金',
  '現行単位',
  '年間総労働時間_明細合計',
  '前回参照最低賃金',
  '今回参照最低賃金',
  '最低賃金上昇額',
  '法定福利費計算倍率',
  '月額増加額',
  '年間増加額',
  '最終提案料金',
  '提案年額換算',
  '計算結果状態',
  '確認状態',
  '交渉状況',
  '文書転記OK',
  'エラー内容',
  '更新日時',
];


  headers[SHEETS.MASTER_SERVICE] = [
  'サービスID',
  '契約先ID',
  '現場ID',
  'テナント名',
  '業務種別',
  '業務名',
  '計算方式',
  '使用有無',
  '備考',
  '作成日時',
  '更新日時',
];


  headers[SHEETS.MASTER_CLIENT] = [
  '契約先ID',
  '契約先名',
  '宛名',
  '敬称',
  '郵便番号',
  '住所',
  '電話番号',
  '備考',
  '作成日時',
  '更新日時',
];


  headers[SHEETS.MASTER_SITE] = [
  '現場ID',
  '契約先ID',
  '現場名',
  '都道府県',
  '住所',
  '備考',
  '作成日時',
  '更新日時',
];


  headers[SHEETS.MASTER_CODE_MAP] = [
  '対応ID',
  '契約先ID',
  '契約先名',
  '現場ID',
  '現場名',
  'テナント名',
  '物件コード',
  '使用有無',
  '備考',
  '作成日時',
  '更新日時',
];


  headers[SHEETS.MASTER_MIN_WAGE] = [
  '都道府県',
  '適用開始日',
  '最低賃金',
  '根拠URL',
  '備考',
];


  headers[SHEETS.MASTER_COEFFICIENT] = [
  '係数ID',
  '適用開始日',
  '法定福利費計算倍率',
  '備考',
];


  headers[SHEETS.SYSTEM_SETTINGS] = [
  'キー',
  '値',
  '備考',
];


  headers[SHEETS.SYSTEM_LOG] = [
  '日時',
  '処理名',
  '対象案件ID',
  '結果',
  'メッセージ',
  '実行ユーザー',
];

  return headers;
}

function getHeadersForSheet_(sheetName) {
  return getSheetHeaders_()[sheetName] || [];
}
