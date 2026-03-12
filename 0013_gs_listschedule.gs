/** =========================================================
 * WBS表示用データ取得
 * type: 'SV' or 'CL'
 * 返却:
 * [
 *   {
 *     key, customer, type, statusText,
 *     sales, primary, manager, support,
 *     summary, workDetail, category,
 *     estimateHours, requestDate, inspectDate, openClose,
 *     caseStart, caseEnd,
 *     inStart, inEnd,
 *     outStart, outEnd
 *   }
 * ]
 * ======================================================= */
function api_getCaseWbsData(type) {
  type = normalizeCaseType_(type);

  const ss = getSS_('Display');
  if (!ss) throw new Error("スプレッドシートが取得できません: getSS_('Display')");

  const config = getCaseWbsSheetConfig_(type);
  const shCase = ss.getSheetByName(config.caseSheetName);
  const shSt   = ss.getSheetByName(config.statusSheetName);

  if (!shCase) throw new Error('シートが見つかりません: ' + config.caseSheetName);
  if (!shSt)   throw new Error('シートが見つかりません: ' + config.statusSheetName);

  const caseValues = shCase.getDataRange().getValues();
  const stValues   = shSt.getDataRange().getValues();

  if (caseValues.length < 2) return [];
  if (stValues.length < 2) return [];

  const caseHeaders = caseValues[0].map(v => String(v || '').trim());
  const stHeaders   = stValues[0].map(v => String(v || '').trim());

  const caseMap = headerMapFromArray_(caseHeaders);
  const stMap   = headerMapFromArray_(stHeaders);

  validateRequiredHeaders_(caseMap, ['見積番号'], config.caseSheetName);
  validateRequiredHeaders_(stMap, ['見積番号'], config.statusSheetName);

  const caseRows = caseValues.slice(1);
  const stRows   = stValues.slice(1);

  const stByKey = buildStatusMapByKey_(stRows, stMap);

  const results = [];

  caseRows.forEach(r => {
    const key = getCellByHeader_(r, caseMap, '見積番号');
    if (!key) return;

    const openCloseRaw = getCellByHeader_(r, caseMap, 'open/close') || 'open';
    const openClose = String(openCloseRaw).trim().toLowerCase();
    if (openClose !== 'open') return;

    const stRow = stByKey[key];
    if (!stRow) return;

    const period = getWbsPeriodSet_(stRow, stMap);

    results.push({
      key: key,
      customer: getCellByHeader_(r, caseMap, '顧客名'),
      type: type,
      statusText: getCellByHeader_(r, caseMap, 'ステータス'),

      sales: getCellByHeader_(r, caseMap, '担当営業'),
      primary: getFirstExistingCell_(r, caseMap, ['担当プリ', '担当SE', '主担当エンジニア']),
      manager: getCellByHeader_(r, caseMap, '管理者'),
      support: getFirstExistingCell_(r, caseMap, ['サポート', 'サポート1', '副担当エンジニア']),

      summary: getFirstExistingCell_(r, caseMap, ['案件概要', '案件名']),
      workDetail: getCellByHeader_(r, caseMap, '作業内容'),
      category: getCellByHeader_(r, caseMap, 'カテゴリタグ'),
      estimateHours: getCellByHeader_(r, caseMap, '見積工数'),
      requestDate: toYmd_(getCellRawByHeader_(r, caseMap, '作業依頼日')),
      inspectDate: toYmd_(getCellRawByHeader_(r, caseMap, '検収予定')),
      openClose: openClose,

      caseStart: period.caseStart,
      caseEnd: period.caseEnd,
      inStart: period.inStart,
      inEnd: period.inEnd,
      outStart: period.outStart,
      outEnd: period.outEnd
    });
  });

  results.sort((a, b) => {
    const ad = a.caseStart || a.inStart || a.outStart || '9999-12-31';
    const bd = b.caseStart || b.inStart || b.outStart || '9999-12-31';

    if (ad < bd) return -1;
    if (ad > bd) return 1;

    return String(a.key).localeCompare(String(b.key), 'ja');
  });

  return results;
}

/** =========================================================
 * 設定
 * ======================================================= */
function getCaseWbsSheetConfig_(type) {
  return {
    caseSheetName:   type === 'CL' ? 'CL案件管理表' : 'SV案件管理表',
    statusSheetName: type === 'CL' ? 'CL案件ステータス' : 'SV案件ステータス'
  };
}

function normalizeCaseType_(type) {
  const v = String(type || 'SV').trim().toUpperCase();
  return v === 'CL' ? 'CL' : 'SV';
}

/** =========================================================
 * マップ / 値取得
 * ======================================================= */
function headerMapFromArray_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (!key) return;
    map[key] = i;
  });
  return map;
}

function validateRequiredHeaders_(map, requiredHeaders, sheetName) {
  const missing = requiredHeaders.filter(h => map[h] == null);
  if (missing.length) {
    throw new Error(
      '必須ヘッダーが見つかりません: ' +
      sheetName +
      ' / ' +
      missing.join(', ')
    );
  }
}

function getCellRawByHeader_(row, map, header) {
  const idx = map[header];
  if (idx == null) return '';
  return row[idx];
}

function getCellByHeader_(row, map, header) {
  const idx = map[header];
  if (idx == null) return '';
  return nz_(row[idx]);
}

function getFirstExistingCell_(row, map, headers) {
  for (let i = 0; i < headers.length; i++) {
    const v = getCellByHeader_(row, map, headers[i]);
    if (v) return v;
  }
  return '';
}

/** =========================================================
 * ステータス行の見積番号マップ
 * ======================================================= */
function buildStatusMapByKey_(stRows, stMap) {
  const out = {};
  stRows.forEach(r => {
    const key = getCellByHeader_(r, stMap, '見積番号');
    if (!key) return;
    out[key] = r;
  });
  return out;
}

/** =========================================================
 * 期間取得
 * ======================================================= */
function getWbsPeriodSet_(stRow, stMap) {
  const caseStart = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '案件管理_作業期間：開始',
    '案件_開始'
  ]));
  const caseEnd = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '案件管理_作業期間：終了',
    '案件_終了'
  ]));

  const inStart = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '社内作業_作業期間：開始',
    '社内_開始'
  ]));
  const inEnd = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '社内作業_作業期間：終了',
    '社内_終了'
  ]));

  let outStart = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '現地作業_作業期間：開始',
    '現地・リモート作業_作業期間：開始',
    '現地_開始'
  ]));
  let outEnd = toYmd_(getFirstExistingRaw_(stRow, stMap, [
    '現地作業_作業期間：終了',
    '現地・リモート作業_作業期間：終了',
    '現地_終了'
  ]));

  if (outStart && !outEnd) outEnd = outStart;
  if (!outStart && outEnd) outStart = outEnd;

  return {
    caseStart: caseStart,
    caseEnd: caseEnd,
    inStart: inStart,
    inEnd: inEnd,
    outStart: outStart,
    outEnd: outEnd
  };
}

function getFirstExistingRaw_(row, map, headers) {
  for (let i = 0; i < headers.length; i++) {
    const idx = map[headers[i]];
    if (idx == null) continue;

    const v = row[idx];
    if (v !== '' && v != null) return v;
  }
  return '';
}

/** =========================================================
 * 文字 / 日付
 * ======================================================= */
function nz_(v) {
  return String(v == null ? '' : v).trim();
}

function toYmd_(v) {
  if (v == null || v === '') return '';

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  if (typeof v === 'number') {
    const d = new Date(v);
    if (!isNaN(d)) {
      return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    }
  }

  const s = String(v).trim();
  if (!s) return '';

  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);
    return y + '-' + mo + '-' + d;
  }

  m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (m) {
    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);
    return y + '-' + mo + '-' + d;
  }

  const d2 = new Date(s);
  if (!isNaN(d2)) {
    return Utilities.formatDate(d2, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  return '';
}