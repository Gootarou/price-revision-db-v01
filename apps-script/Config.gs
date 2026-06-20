/**
 * Application-wide constants for the price revision database.
 * Sheet names must be referenced through this file instead of hard-coded in services.
 */
const APP_CONFIG = {
  MENU_NAME: '料金改定DB',
  MENU_INITIAL_SETUP: '初期セットアップ',
  MENU_INPUT_SHEET_INITIALIZE: '入力画面初期化',
};

const SHEETS = {
  INPUT_CASE_EDIT: '入力_案件編集',
  DATA_PRICE_REVISION_CASE: 'データ_料金改定案件',
  DETAIL_WORK_TIME_BASIS: '明細_勤務時間根拠',
  VIEW_DOCUMENT_TRANSFER_CHECK: '確認_文書転記前',
  MASTER_SERVICE: 'マスタ_契約サービス',
  MASTER_CLIENT: 'マスタ_契約先',
  MASTER_SITE: 'マスタ_現場',
  MASTER_CODE_MAP: 'マスタ_コード対応表',
  MASTER_MIN_WAGE: 'マスタ_最低賃金',
  MASTER_COEFFICIENT: 'マスタ_計算係数',
  SYSTEM_SETTINGS: '管理_設定',
  SYSTEM_LOG: '管理_処理ログ',
};

const SHEET_ORDER = [
  SHEETS.INPUT_CASE_EDIT,
  SHEETS.DATA_PRICE_REVISION_CASE,
  SHEETS.DETAIL_WORK_TIME_BASIS,
  SHEETS.VIEW_DOCUMENT_TRANSFER_CHECK,
  SHEETS.MASTER_SERVICE,
  SHEETS.MASTER_CLIENT,
  SHEETS.MASTER_SITE,
  SHEETS.MASTER_CODE_MAP,
  SHEETS.MASTER_MIN_WAGE,
  SHEETS.MASTER_COEFFICIENT,
  SHEETS.SYSTEM_SETTINGS,
  SHEETS.SYSTEM_LOG,
];

const CALCULATION_STATUS = {
  NOT_CALCULATED: '未計算',
  OK: 'OK',
  OUT_OF_SCOPE: '対象外',
  ERROR: 'エラー',
};

const CONFIRMATION_STATUS = {
  NOT_CONFIRMED: '未確認',
  CONFIRMED: '確認済',
};

const NEGOTIATION_STATUS = {
  NOT_SUBMITTED: '未提出',
  SUBMITTED: '提出済',
  APPROVED: '承認',
  REJECTED: '否認',
  PENDING: '保留',
};

const TRANSFER_STATUS = {
  OK: 'OK',
  NG: 'NG',
};

const ACTIVE_STATUS = {
  ACTIVE: '使用',
  INACTIVE: '不使用',
};

const ROUNDING_METHOD = {
  NONE: 'なし',
  FLOOR_10: '10円未満切り下げ',
  FLOOR_100: '100円未満切り下げ',
  FLOOR_1000: '1000円未満切り下げ',
};

const LOG_RESULT = {
  OK: 'OK',
  ERROR: 'ERROR',
  WARN: 'WARN',
};
