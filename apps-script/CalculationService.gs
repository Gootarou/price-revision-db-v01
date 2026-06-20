/**
 * Calculates saved price revision cases for Phase 4.
 */
function calculatePriceRevisionCase(caseId) {
  const sheet = getCanonicalSheet_(SHEETS.DATA_PRICE_REVISION_CASE);
  const row = findRowByHeaderValue_(sheet, '案件ID', caseId);
  if (!row) {
    throw new Error('対象案件IDが存在しません: ' + caseId);
  }

  const caseRecord = getRowObject_(sheet, row);
  const now = new Date();

  try {
    if (String(caseRecord['現行単位'] || '').trim() !== '月額') {
      const outOfScopeResult = blankCalculationResultFields_(buildCalculationBaseResult_(caseRecord, now));
      outOfScopeResult['計算結果状態'] = CALCULATION_STATUS.OUT_OF_SCOPE;
      outOfScopeResult['文書転記OK'] = TRANSFER_STATUS.NG;
      outOfScopeResult['エラー内容'] = 'v1自動計算対象外';
      setRowFields_(sheet, row, outOfScopeResult);
      writeProcessLog('計算対象外', caseId, LOG_RESULT.WARN, '現行単位が月額ではないため自動計算対象外です。');
      return getRowObject_(sheet, row);
    }

    const previousRevisionStartDate = requireDate_(caseRecord['前回改定開始日'], '前回改定開始日');
    const currentRevisionStartDate = requireDate_(caseRecord['今回改定開始日'], '今回改定開始日');
    const currentPrice = requireNumber_(caseRecord['現行料金'], '現行料金');
    const site = findSiteById_(caseRecord['現場ID']);
    if (!site || !site['都道府県']) {
      throw new CalculationNamedError_('最低賃金未登録', '対象現場の都道府県が見つかりません。現場ID: ' + caseRecord['現場ID']);
    }

    const previousMinWage = findMinWage_(site['都道府県'], previousRevisionStartDate);
    const currentMinWage = findMinWage_(site['都道府県'], currentRevisionStartDate);
    const coefficient = findCalculationCoefficient_(currentRevisionStartDate);
    const workTimeTotal = sumActiveWorkTimeTotal_(caseId);
    const roundingMethod = String(caseRecord['端数処理方式'] || ROUNDING_METHOD.NONE).trim() || ROUNDING_METHOD.NONE;
    const manualAdjustment = caseRecord['手動調整額'] === '' || caseRecord['手動調整額'] === null || typeof caseRecord['手動調整額'] === 'undefined'
      ? 0
      : requireNumber_(caseRecord['手動調整額'], '手動調整額');

    const minWageIncrease = currentMinWage - previousMinWage;
    const laborUnitIncrease = minWageIncrease * coefficient;
    const provisionalAnnualIncrease = workTimeTotal * laborUnitIncrease;
    const provisionalMonthlyIncrease = provisionalAnnualIncrease / 12;
    const roundedMonthlyIncrease = applyRounding_(provisionalMonthlyIncrease, roundingMethod);
    const monthlyIncrease = roundedMonthlyIncrease + manualAdjustment;
    const annualIncrease = monthlyIncrease * 12;
    const currentMonthlyEquivalent = currentPrice;
    const currentAnnualEquivalent = currentMonthlyEquivalent * 12;
    const finalProposalPrice = currentMonthlyEquivalent + monthlyIncrease;
    const proposedMonthlyEquivalent = finalProposalPrice;
    const proposedAnnualEquivalent = proposedMonthlyEquivalent * 12;

    const result = buildCalculationBaseResult_(caseRecord, now);
    Object.assign(result, {
      '現行月額換算': currentMonthlyEquivalent,
      '現行年額換算': currentAnnualEquivalent,
      '年間総労働時間_明細合計': workTimeTotal,
      '前回参照最低賃金': previousMinWage,
      '今回参照最低賃金': currentMinWage,
      '最低賃金上昇額': minWageIncrease,
      '法定福利費計算倍率': coefficient,
      '労務単価上昇額': laborUnitIncrease,
      '暫定年間増加額': provisionalAnnualIncrease,
      '暫定月額増加額': provisionalMonthlyIncrease,
      '端数処理方式': roundingMethod,
      '端数処理後月額増加額': roundedMonthlyIncrease,
      '手動調整額': manualAdjustment,
      '月額増加額': monthlyIncrease,
      '年間増加額': annualIncrease,
      '最終提案料金': finalProposalPrice,
      '提案月額換算': proposedMonthlyEquivalent,
      '提案年額換算': proposedAnnualEquivalent,
      '計算結果状態': CALCULATION_STATUS.OK,
      'エラー内容': '',
    });
    result['文書転記OK'] = determineDocumentTransferStatus_(Object.assign({}, caseRecord, result));

    setRowFields_(sheet, row, result);
    writeProcessLog('計算成功', caseId, LOG_RESULT.OK, '月額契約の料金改定計算が完了しました。');
    return getRowObject_(sheet, row);
  } catch (error) {
    const processName = error.processName || '計算エラー';
    const errorResult = blankCalculationResultFields_({
      '計算結果状態': CALCULATION_STATUS.ERROR,
      '文書転記OK': TRANSFER_STATUS.NG,
      'エラー内容': error.message,
      '更新日時': now,
    });
    setRowFields_(sheet, row, errorResult);
    writeProcessLog(processName, caseId, LOG_RESULT.ERROR, error.message);
    throw error;
  }
}

function buildCalculationBaseResult_(caseRecord, now) {
  return {
    '端数処理方式': String(caseRecord['端数処理方式'] || ROUNDING_METHOD.NONE).trim() || ROUNDING_METHOD.NONE,
    '手動調整額': caseRecord['手動調整額'] === '' || caseRecord['手動調整額'] === null || typeof caseRecord['手動調整額'] === 'undefined' ? 0 : caseRecord['手動調整額'],
    '更新日時': now,
  };
}

function blankCalculationResultFields_(result) {
  [
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
    '端数処理後月額増加額',
    '月額増加額',
    '年間増加額',
    '最終提案料金',
    '提案月額換算',
    '提案年額換算',
  ].forEach(function(field) {
    result[field] = '';
  });
  return result;
}

function findSiteById_(siteId) {
  const sheet = getCanonicalSheet_(SHEETS.MASTER_SITE);
  const row = findRowByHeaderValue_(sheet, '現場ID', siteId);
  return row ? getRowObject_(sheet, row) : null;
}

function findMinWage_(prefecture, targetDate) {
  const record = findEffectiveRecord_(SHEETS.MASTER_MIN_WAGE, '都道府県', prefecture, targetDate, '最低賃金');
  if (!record) {
    throw new CalculationNamedError_('最低賃金未登録', '最低賃金が見つかりません。都道府県: ' + prefecture + '、基準日: ' + formatDateForMessage_(targetDate));
  }
  return requireNumber_(record['最低賃金'], '最低賃金');
}

function findCalculationCoefficient_(targetDate) {
  const record = findEffectiveRecord_(SHEETS.MASTER_COEFFICIENT, null, null, targetDate, '法定福利費計算倍率');
  if (!record) {
    throw new CalculationNamedError_('計算係数未登録', '計算係数が見つかりません。基準日: ' + formatDateForMessage_(targetDate));
  }
  return requireNumber_(record['法定福利費計算倍率'], '法定福利費計算倍率');
}

function findEffectiveRecord_(sheetName, matchHeader, matchValue, targetDate, valueHeader) {
  const sheet = getCanonicalSheet_(sheetName);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !headerMap['適用開始日'] || !headerMap[valueHeader]) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let best = null;
  let bestDate = null;
  values.forEach(function(row) {
    const record = rowValuesToObject_(headerMap, row);
    if (matchHeader && String(record[matchHeader]) !== String(matchValue)) return;
    const effectiveDate = parseDateValue_(record['適用開始日']);
    if (!effectiveDate || effectiveDate.getTime() > targetDate.getTime()) return;
    if (!bestDate || effectiveDate.getTime() > bestDate.getTime()) {
      best = record;
      bestDate = effectiveDate;
    }
  });
  return best;
}

function sumActiveWorkTimeTotal_(caseId) {
  const details = findWorkTimeDetailsByCaseId_(caseId).filter(function(detail) {
    return String(detail['使用有無'] || ACTIVE_STATUS.ACTIVE).trim() === ACTIVE_STATUS.ACTIVE;
  });
  if (details.length === 0) {
    throw new CalculationNamedError_('勤務時間根拠未登録', '使用対象の勤務時間根拠がありません。');
  }
  return details.reduce(function(total, detail) {
    return total + requireNumber_(detail['年間総労働時間'], '年間総労働時間');
  }, 0);
}

function applyRounding_(value, method) {
  if (method === ROUNDING_METHOD.NONE) return value;
  if (method === ROUNDING_METHOD.FLOOR_10) return Math.floor(value / 10) * 10;
  if (method === ROUNDING_METHOD.FLOOR_100) return Math.floor(value / 100) * 100;
  if (method === ROUNDING_METHOD.FLOOR_1000) return Math.floor(value / 1000) * 1000;
  throw new Error('未知の端数処理方式です: ' + method);
}

function determineDocumentTransferStatus_(record) {
  const ok = record['計算結果状態'] === CALCULATION_STATUS.OK
    && record['確認状態'] === CONFIRMATION_STATUS.CONFIRMED
    && !record['エラー内容']
    && record['最終提案料金'] !== '' && record['最終提案料金'] !== null && typeof record['最終提案料金'] !== 'undefined'
    && record['提案月額換算'] !== '' && record['提案月額換算'] !== null && typeof record['提案月額換算'] !== 'undefined'
    && record['提案年額換算'] !== '' && record['提案年額換算'] !== null && typeof record['提案年額換算'] !== 'undefined'
    && record['今回改定開始日'] !== '' && record['今回改定開始日'] !== null && typeof record['今回改定開始日'] !== 'undefined'
    && Number(record['年間総労働時間_明細合計']) > 0;
  return ok ? TRANSFER_STATUS.OK : TRANSFER_STATUS.NG;
}

function requireNumber_(value, label) {
  if (value === '' || value === null || typeof value === 'undefined') throw new Error(label + 'が未入力です。');
  const numberValue = Number(value);
  if (isNaN(numberValue)) throw new Error(label + 'が数値ではありません。');
  return numberValue;
}

function requireDate_(value, label) {
  const date = parseDateValue_(value);
  if (!date) throw new Error(label + 'が日付ではありません。');
  return date;
}

function parseDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (value === '' || value === null || typeof value === 'undefined') return null;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function formatDateForMessage_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function CalculationNamedError_(processName, message) {
  this.name = 'CalculationNamedError';
  this.processName = processName;
  this.message = message;
}
CalculationNamedError_.prototype = Object.create(Error.prototype);
CalculationNamedError_.prototype.constructor = CalculationNamedError_;
