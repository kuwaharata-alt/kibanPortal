/** ======================================================================
 * Modal API
 * - 案件情報
 * - ステータス
 * - 統合取得 / 統合保存
 * - 高速化版
 * ====================================================================== */

const INPUT_CONFIG = {
  SHEET_MEMBER: 'BSOL情報',
  SHEET_MENU: 'PDMenu',
  HEADER_ROW: 1,
  KEY_HEADER: '見積番号',

  MAPPING: {
    '作業難度': '作業難度',
    '作業形態': '作業形態',
    '現地作業場所': '現地作業場所',
    '最寄り': '最寄り',
    '住所': '住所',
    '作業内容': '作業内容',
    '台数': '台数',
    '監督者': '監督者',
    '管理者': '管理者',
    'サポート': 'サポート',
    'BP利用': 'BP利用',
  },
};

const STATUS_APP_CONFIG = {
  MASTER_SHEET: 'ステータスマスタ',
  MASTER_HEADER_ROW: 1,
  CASE_HEADER_ROW: 1,

  KEY_HEADER: '見積番号',
  NOTE_HEADER: '備考',
  UPDATE_AT_HEADER: 'updated_at',
  UPDATE_BY_HEADER: 'updated_by',
  CURRENT_CODE_HEADER: 'current_status_code',
  CASE_STATUS_HEADER: 'ステータス',

  NA_OPTIONS: ['対応有', '対応無'],
};

const STATUS_MODAL_CACHE = {
  MASTER_SEC: 300,     // 5分
  ROW_INDEX_SEC: 180,  // 3分
};

const __STATUS_EXEC_CACHE = {
  headerMap: {},
  rowIndex: {},
};

/** =========================
 * 案件情報マスタ
 * ========================= */

function api_getMembers() {
  try {
    const sh = getMasterSheet_(INPUT_CONFIG.SHEET_MEMBER);
    if (!sh) return { ok: false, error: INPUT_CONFIG.SHEET_MEMBER + 'が見つかりません' };

    const members = listUniqueFromColumn_(sh, 7, {
      startRow: 1,
      skipWords: ['氏名', '名前'],
    });

    return { ok: true, members };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getLevel() {
  try {
    const sh = getMasterSheet_(INPUT_CONFIG.SHEET_MENU);
    if (!sh) return { ok: false, error: INPUT_CONFIG.SHEET_MENU + 'が見つかりません' };

    const levels = listUniqueFromColumn_(sh, 1, {
      startRow: 1,
      skipWords: ['作業難度'],
    });

    return { ok: true, members: levels };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getLocationMaster() {
  try {
    const sh = getMasterSheet_(INPUT_CONFIG.SHEET_MENU);
    if (!sh) return { ok: false, error: INPUT_CONFIG.SHEET_MENU + 'が見つかりません' };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, groups: [] };

    const rows = sh.getRange(2, 2, lastRow - 1, 2).getValues(); // B:C
    return { ok: true, groups: groupByTwoCols_(rows) };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/** =========================
 * 案件情報
 * ========================= */

function api_upsertInput(payload) {
  try {
    const type = normalizeCaseType_(payload?.type);
    const key = norm_(payload?.key);
    const values = payload?.values || {};

    if (!type) return { ok: false, error: 'type が空です' };
    if (!key) return { ok: false, error: 'key が空です' };

    const sh = getDisplayCaseSheet_(type);
    if (!sh) return { ok: false, error: 'sheet not found for type=' + type };

    const updater = getActiveUserNameOrEmail_();
    Logger.log('[api_upsertInput] key=' + key + ' updater=' + updater);

    const res = upsertByMapping_(sh, {
      headerRow: INPUT_CONFIG.HEADER_ROW,
      keyHeader: INPUT_CONFIG.KEY_HEADER,
      key,
      mapping: INPUT_CONFIG.MAPPING,
      values,
      transform: transformInputValue_,
      always: {
        updated_at: nowYmd_(false),
        updated_by: updater,
      },
    });

    if (!res.ok) return res;

    return {
      ok: true,
      sheet: sh.getName(),
      row: res.row,
      inserted: res.inserted,
      missingHeaders: res.missingHeaders,
      updated_by: updater,
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getCaseByKey(arg1, arg2) {
  try {
    const req = normalizeGetCaseRequest_(arg1, arg2);
    const type = normalizeCaseType_(req.type);
    const key = norm_(req.key);

    if (!key) return { ok: false, error: 'key is empty' };

    const sh = getDisplayCaseSheet_(type);
    if (!sh) return { ok: false, error: 'sheet not found for type=' + type };

    const headerRow = INPUT_CONFIG.HEADER_ROW;
    const hm = getHeaderMapCached_(sh, headerRow);
    const keyCol = hm[INPUT_CONFIG.KEY_HEADER];
    if (!keyCol) return { ok: false, error: 'key header not found: 見積番号' };

    const rowNo = findRowByKeyCached_(sh, keyCol, key, headerRow + 1);
    if (!rowNo) return { ok: false, error: 'not found: ' + key + ' (' + type + ')' };

    const obj = readRowAsObjectFast_(sh, rowNo, headerRow);
    const data = normalizeInputRowForClient_(normalizeForClient_(obj));

    return { ok: true, data, from: type };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function normalizeGetCaseRequest_(arg1, arg2) {
  if (arg1 && typeof arg1 === 'object') {
    return { type: arg1.type, key: arg1.key };
  }
  return { type: arg1, key: arg2 };
}

function transformInputValue_(inputKey, value) {
  if (inputKey === 'サポート') return joinMultiValue_(value, '、');
  if (inputKey === '作業形態') return joinMultiValue_(value, '／');
  return norm_(value);
}

function normalizeInputRowForClient_(row) {
  const out = Object.assign({}, row || {});

  if (out['作業開始予定'] != null && out['開始日'] == null) {
    out['開始日'] = ymdSlashToDashSafe_(out['作業開始予定']);
  }
  if (out['作業完了予定'] != null && out['終了日'] == null) {
    out['終了日'] = ymdSlashToDashSafe_(out['作業完了予定']);
  }

  if (out['作業形態'] != null) {
    out['作業形態'] = splitMultiValueSafe_(out['作業形態']);
  }
  if (out['サポート'] != null) {
    out['サポート'] = splitMultiValueSafe_(out['サポート']);
  }

  return out;
}

/** =========================
 * ステータスマスタ
 * ========================= */

function api_getStatusMaster(params) {
  try {
    params = params || {};
    const target = normalizeCaseType_(params.target || 'SV');
    const forceRefresh = !!params.forceRefresh;
    return getStatusMasterCached_(target, forceRefresh);
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function getStatusMasterCached_(target, forceRefresh) {
  target = normalizeCaseType_(target || 'SV');

  const cache = CacheService.getScriptCache();
  const cacheKey = 'status_master_v2_' + target;

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // noop
      }
    }
  }

  const res = loadStatusMasterRaw_(target);
  if (res && res.ok) {
    cache.put(cacheKey, JSON.stringify(res), STATUS_MODAL_CACHE.MASTER_SEC);
  }
  return res;
}

function loadStatusMasterRaw_(target) {
  target = normalizeCaseType_(target || 'SV');

  const sh = getSS_('DB').getSheetByName(STATUS_APP_CONFIG.MASTER_SHEET);
  if (!sh) {
    return { ok: false, error: 'マスタシートが見つかりません: ' + STATUS_APP_CONFIG.MASTER_SHEET };
  }

  const headerRow = STATUS_APP_CONFIG.MASTER_HEADER_ROW;
  const headerMap = getHeaderMapCached_(sh, headerRow);
  const lastRow = sh.getLastRow();

  if (lastRow <= headerRow) {
    return {
      ok: true,
      target,
      rows: [],
      groups: [],
      naOptions: STATUS_APP_CONFIG.NA_OPTIONS
    };
  }

  const values = sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn()).getValues();

  const colTarget = headerMap['対象'];
  const colCode = headerMap['コード'];
  const colBig = headerMap['大分類'];
  const colMid = headerMap['中分類'];
  const colSmall = headerMap['小分類'];
  const colDisp = headerMap['表示名'];
  const colSort = headerMap['sort'];
  const colClosed = headerMap['is_closed'];
  const colColor = headerMap['is_color'];
  const colForce = headerMap['force'];
  const colDefault = headerMap['default'];
  const colNaEnable = headerMap['na_enable'];
  const colRollup = headerMap['rollup_state'];
  const colStHeader = headerMap['Status_Header'];

  if (!colTarget || !colCode || !colBig || !colMid || !colSmall) {
    return { ok: false, error: 'マスタの必須列（対象/コード/大分類/中分類/小分類）が不足しています' };
  }

  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const t = norm_(r[colTarget - 1]).toUpperCase();
    if (t !== target) continue;

    const defRaw = colDefault ? r[colDefault - 1] : '';
    const defStr = norm_(defRaw).toUpperCase();
    const isDefault = defStr === 'TRUE';
    const defaultValue = (!isDefault && defStr !== 'FALSE') ? toYmdSlash_(defRaw, COMMON.TZ) : '';

    rows.push({
      target: t,
      code: norm_(r[colCode - 1]),
      big: norm_(r[colBig - 1]),
      mid: norm_(r[colMid - 1]),
      small: norm_(r[colSmall - 1]),
      display: norm_(colDisp ? r[colDisp - 1] : ''),
      sort: Number(colSort ? r[colSort - 1] : r[colCode - 1]) || 0,
      is_closed: colClosed ? String(r[colClosed - 1]).toUpperCase() === 'TRUE' : false,
      is_color: colColor ? String(r[colColor - 1]).toUpperCase() === 'TRUE' : false,
      force: colForce ? String(r[colForce - 1]).toUpperCase() === 'TRUE' : false,
      is_default: isDefault,
      default_value: defaultValue,
      na_enable: colNaEnable ? String(r[colNaEnable - 1]).toUpperCase() === 'TRUE' : false,
      rollup_state: norm_(colRollup ? r[colRollup - 1] : ''),
      status_header: norm_(colStHeader ? r[colStHeader - 1] : '')
    });
  }

  return {
    ok: true,
    target,
    rows,
    groups: buildGroups_(rows),
    naOptions: STATUS_APP_CONFIG.NA_OPTIONS
  };
}

function buildGroups_(rows) {
  const bigMap = {};

  (rows || []).forEach(r => {
    if (!r.big) return;

    if (!bigMap[r.big]) {
      bigMap[r.big] = {
        big: r.big,
        sort: r.sort || 0,
        na_enable: !!r.na_enable,
        midMap: {}
      };
    } else if (r.na_enable) {
      bigMap[r.big].na_enable = true;
    }

    const bm = bigMap[r.big];
    bm.sort = Math.min(bm.sort || 0, r.sort || bm.sort || 0);

    const midKey = norm_(r.mid);
    if (!bm.midMap[midKey]) {
      bm.midMap[midKey] = {
        mid: r.mid,
        sort: r.sort || 0,
        force: !!r.force,
        is_date: !!r.default_value,
        status_header: r.status_header || statusColName_(r.big, r.mid),
        options: []
      };
    }

    const mm = bm.midMap[midKey];
    mm.sort = Math.min(mm.sort || 0, r.sort || mm.sort || 0);
    if (r.force) mm.force = true;
    if (r.default_value) mm.is_date = true;
    if (!mm.status_header && r.status_header) mm.status_header = r.status_header;

    // is_closed=TRUE は画面表示用 options には含めない
    if (r.is_closed) return;

    mm.options.push({
      code: r.code,
      small: r.small,
      display: r.display,
      sort: r.sort || 0,
      is_closed: !!r.is_closed,
      is_color: !!r.is_color,
      is_default: !!r.is_default,
      rollup_state: r.rollup_state || ''
    });
  });

  return Object.keys(bigMap)
    .map(k => bigMap[k])
    .sort((a, b) => (a.sort || 0) - (b.sort || 0))
    .map(bm => ({
      big: bm.big,
      na_enable: !!bm.na_enable,
      mids: Object.keys(bm.midMap)
        .map(k => bm.midMap[k])
        .sort((a, b) => (a.sort || 0) - (b.sort || 0))
        .map(mm => {
          mm.options.sort((a, b) => (a.sort || 0) - (b.sort || 0));
          return mm;
        })
    }));
}

/** =========================
 * 案件ステータス
 * ========================= */

function api_getCaseStatus(req) {
  try {
    req = req || {};
    const type = normalizeCaseType_(req.type || 'SV');
    const key = norm_(req.key);
    if (!key) return { ok: false, error: '見積番号が空です' };

    const master = req.master && req.master.ok ? req.master : getStatusMasterCached_(type);
    if (!master.ok) return { ok: false, error: 'マスタ取得失敗' };

    return buildCaseStatusResponse_(type, key, master);
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function buildCaseStatusResponse_(type, key, master) {
  const shS = getDisplayStatusSheet_(type);
  if (!shS) return { ok: true, statusMap: {}, naGroups: {}, currentCode: '' };

  const hmS = getHeaderMapCached_(shS, 1);
  const colKey = hmS[STATUS_APP_CONFIG.KEY_HEADER];
  if (!colKey) return { ok: false, error: '見積番号列がありません' };

  const rowNo = findRowByKeyCached_(shS, colKey, key, 2);
  if (!rowNo) return { ok: true, statusMap: {}, naGroups: {}, currentCode: '' };

  const rowVals = readSheetRowValues_(shS, rowNo);

  const statusMap = {};
  const naGroups = {};

  (master.groups || []).forEach(g => {
    let allNa = true;
    let hasValue = false;

    (g.mids || []).forEach(m => {
      const colName = m.status_header || statusColName_(g.big, m.mid);
      const col = hmS[colName];
      if (!col) return;

      const v = rowVals[col - 1];
      if (v == null || v === '') return;

      hasValue = true;

      const stKey = 'ST_' + g.big + '__' + m.mid;
      statusMap[stKey] = m.is_date ? toInputDateValue_(v) : String(v).trim();

      if (String(v).trim() !== 'ー') allNa = false;
    });

    if (g.na_enable && hasValue && allNa) {
      naGroups[g.big] = true;
    }
  });

  getCommentTargetBigs_(type).forEach(big => {
    const col = getCommentCol_(hmS, type, big);
    if (!col) return;

    const v = rowVals[col - 1];
    if (v != null && String(v).trim() !== '') {
      statusMap['CM_' + big] = String(v);
    }
  });

  const currentCode = hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]
    ? norm_(rowVals[hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER] - 1])
    : '';

  return { ok: true, statusMap, naGroups, currentCode };
}

function api_saveCaseStatus(req) {
  try {
    req = req || {};

    const type = normalizeCaseType_(req.type || 'SV');
    const key = norm_(req.key);
    if (!key) return { ok: false, error: '見積番号が空です' };

    const statusMap = req.statusMap || {};
    const naBigs = Array.isArray(req.naBigs) ? req.naBigs : [];
    const currentCode = norm_(req.current_status_code || '');

    const shS = getDisplayStatusSheet_(type);
    const shM = getDisplayCaseSheet_(type);
    if (!shS) return { ok: false, error: '保存先シートがありません' };
    if (!shM) return { ok: false, error: '案件管理表がありません' };

    const hmS = getHeaderMapCached_(shS, 1);
    const hmM = getHeaderMapCached_(shM, STATUS_APP_CONFIG.CASE_HEADER_ROW);

    const colKeyS = hmS[STATUS_APP_CONFIG.KEY_HEADER];
    if (!colKeyS) return { ok: false, error: '見積番号列がありません' };

    const rowNo = findRowByKeyCached_(shS, colKeyS, key, 2);
    if (!rowNo) return { ok: false, error: '見積番号が見つかりません: ' + key };

    const rowNoM = findRowByKeyCached_(shM, hmM[STATUS_APP_CONFIG.KEY_HEADER], key, STATUS_APP_CONFIG.CASE_HEADER_ROW + 1);
    if (!rowNoM) return { ok: false, error: '案件管理表に見積番号が見つかりません: ' + key };

    const master = getStatusMasterCached_(type);
    if (!master.ok) return { ok: false, error: 'マスタ取得失敗' };

    const naSet = buildNaSet_(naBigs);

    clearTargetStatusCols_(shS, rowNo, hmS, master);
    applyNaMarks_(shS, rowNo, hmS, master, naSet);
    applyStatusValues_(shS, rowNo, hmS, statusMap, naSet, master);
    writeCommentAndMeta_(shS, rowNo, hmS, type, statusMap, currentCode);
    writeSummaryColsToStatusSheet_(type, shS, rowNo, hmS, master, statusMap, naSet);

    syncCaseSheetAfterStatusSave_(type, rowNoM, shM, hmM, master, statusMap, naSet, currentCode);

    invalidateRowIndexCache_(shS, hmS[STATUS_APP_CONFIG.KEY_HEADER]);
    invalidateRowIndexCache_(shM, hmM[STATUS_APP_CONFIG.KEY_HEADER]);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function statusColName_(big, mid) {
  return norm_(big) + '_' + norm_(mid);
}

function buildNaSet_(naBigs) {
  const naSet = {};
  (naBigs || []).forEach(b => naSet[norm_(b)] = true);
  return naSet;
}

function clearTargetStatusCols_(shS, rowNo, hmS, master) {
  const uniq = {};
  (master.groups || []).forEach(g => {
    (g.mids || []).forEach(m => {
      const colName = m.status_header || statusColName_(g.big, m.mid);
      const col = hmS[colName];
      if (col) uniq[col] = true;
    });
  });

  const cols = Object.keys(uniq).map(Number).filter(Boolean);
  writeSameRowByCols_(shS, rowNo, cols, () => '');
}

function applyNaMarks_(shS, rowNo, hmS, master, naSet) {
  const uniq = {};
  (master.groups || []).forEach(g => {
    const big = norm_(g.big);
    if (!big || !naSet[big]) return;

    (g.mids || []).forEach(m => {
      const colName = m.status_header || statusColName_(big, m.mid);
      const col = hmS[colName];
      if (col) uniq[col] = true;
    });
  });

  const cols = Object.keys(uniq).map(Number).filter(Boolean);
  writeSameRowByCols_(shS, rowNo, cols, () => 'ー');
}

function applyStatusValues_(shS, rowNo, hmS, statusMap, naSet, master) {
  const setMap = {};
  const uniq = {};

  Object.keys(statusMap || {}).forEach(stKey => {
    if (!String(stKey).startsWith('ST_')) return;

    const raw = String(stKey).slice(3);
    const p = raw.split('__');
    if (p.length < 2) return;

    const big = p[0];
    const mid = p.slice(1).join('__');

    if (naSet[norm_(big)]) return;

    const colName = getStatusHeaderByBigMid_(master, big, mid) || statusColName_(big, mid);
    const col = hmS[colName];
    if (!col) return;

    uniq[col] = true;
    setMap[col] = statusMap[stKey];
  });

  const cols = Object.keys(uniq).map(Number).filter(Boolean);
  writeSameRowByCols_(shS, rowNo, cols, col => setMap[col]);
}

function writeCommentAndMeta_(shS, rowNo, hmS, type, statusMap, currentCode) {
  const updater = getActiveUserNameOrEmail_();
  const updates = {};

  getCommentTargetBigs_(type).forEach(big => {
    const col = getCommentCol_(hmS, type, big);
    if (!col) return;
    updates[col] = statusMap['CM_' + big] == null ? '' : String(statusMap['CM_' + big]);
  });

  if (hmS[STATUS_APP_CONFIG.UPDATE_AT_HEADER]) {
    updates[hmS[STATUS_APP_CONFIG.UPDATE_AT_HEADER]] = nowYmd_(true);
  }

  if (hmS[STATUS_APP_CONFIG.UPDATE_BY_HEADER]) {
    updates[hmS[STATUS_APP_CONFIG.UPDATE_BY_HEADER]] = updater;
  }

  if (hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]) {
    updates[hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]] = currentCode || '';
  }

  if (hmS['open/close']) {
    updates[hmS['open/close']] = openCloseFromCode_(currentCode);
  }

  const cols = Object.keys(updates).map(Number).filter(Boolean);
  if (!cols.length) return;

  writeSameRowByCols_(shS, rowNo, cols, col => updates[col]);
}

/** =========================
 * 保存後同期（ステータスシート / 案件管理表）
 * ========================= */

function writeSummaryColsToStatusSheet_(type, shS, rowNo, hmS, master, statusMap, naSet) {
  const caseSt = resolveCaseStatusByRule_(type, master, statusMap, naSet);
  const inSt = resolveBigRollupState_(master, statusMap, naSet, '社内作業');
  const outSt = resolveBigRollupState_(master, statusMap, naSet, getFieldWorkBigName_(type));
  const docSt = resolveBigRollupState_(master, statusMap, naSet, 'ドキュメント');

  const updates = {};
  if (hmS['案件ST']) updates[hmS['案件ST']] = caseSt || '';
  if (hmS['社内ST']) updates[hmS['社内ST']] = inSt || '';
  if (hmS['現地ST']) updates[hmS['現地ST']] = outSt || '';
  if (hmS['Doc.ST']) updates[hmS['Doc.ST']] = docSt || '';
  if (hmS['DocST']) updates[hmS['DocST']] = docSt || '';

  const cols = Object.keys(updates).map(Number).filter(Boolean);
  if (!cols.length) return;

  writeSameRowByCols_(shS, rowNo, cols, col => updates[col]);
}

function syncCaseSheetAfterStatusSave_(type, rowNoM, shM, hmM, master, statusMap, naSet, currentCode) {
  const updater = getActiveUserNameOrEmail_();

  const caseSt = resolveCaseStatusByRule_(type, master, statusMap, naSet);
  const inSt = resolveBigRollupState_(master, statusMap, naSet, '社内作業');
  const outSt = resolveBigRollupState_(master, statusMap, naSet, getFieldWorkBigName_(type));
  const docSt = resolveBigRollupState_(master, statusMap, naSet, 'ドキュメント');

  const summaryText = [
    `案件ST：${caseSt || ''}`,
    `社内：${inSt || ''}　現地：${outSt || ''}　Doc.：${docSt || ''}`
  ].join('\n');

  const updates = {};

  if (hmM[STATUS_APP_CONFIG.CASE_STATUS_HEADER]) {
    updates[hmM[STATUS_APP_CONFIG.CASE_STATUS_HEADER]] = summaryText;
  }

  if (hmM['案件ST']) updates[hmM['案件ST']] = caseSt || '';
  if (hmM['社内ST']) updates[hmM['社内ST']] = inSt || '';
  if (hmM['現地ST']) updates[hmM['現地ST']] = outSt || '';
  if (hmM['Doc.ST']) updates[hmM['Doc.ST']] = docSt || '';
  if (hmM['DocST']) updates[hmM['DocST']] = docSt || '';

  if (hmM[STATUS_APP_CONFIG.UPDATE_AT_HEADER]) {
    updates[hmM[STATUS_APP_CONFIG.UPDATE_AT_HEADER]] = nowYmd_(true);
  }
  if (hmM[STATUS_APP_CONFIG.UPDATE_BY_HEADER]) {
    updates[hmM[STATUS_APP_CONFIG.UPDATE_BY_HEADER]] = updater;
  }
  if (hmM[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]) {
    updates[hmM[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]] = currentCode || '';
  }
  if (hmM['open/close']) {
    updates[hmM['open/close']] = (caseSt === '案件クローズ') ? 'close' : 'open';
  }

  const cols = Object.keys(updates).map(Number).filter(Boolean);
  if (!cols.length) return;

  writeSameRowByCols_(shM, rowNoM, cols, col => updates[col]);
}

function getMidOptionByValue_(master, big, mid, rawValue) {
  const mm = findMidMaster_(master, big, mid);
  if (!mm) return null;

  const s = String(rawValue ?? '').trim();
  if (!s) return null;

  return (mm.options || []).find(o =>
    norm_(o.code) === norm_(s) ||
    norm_(o.small) === norm_(s) ||
    norm_(o.display) === norm_(s)
  ) || null;
}

function getMidResolvedLabel_(master, statusMap, big, mid) {
  const raw = getMidRawStatus_(statusMap, big, mid);
  if (!raw) return '';

  const opt = getMidOptionByValue_(master, big, mid, raw);
  if (!opt) return String(raw).trim();

  return String(opt.small || opt.display || raw).trim();
}

/** =========================
 * 案件ST / rollup 判定
 * ========================= */

function resolveCaseStatusByRule_(type, master, statusMap, naSet) {
  const assignBig = '案件確認';
  const checkBig = '案件確認';
  const inBig = '社内作業';
  const outBig = getFieldWorkBigName_(type);
  const docBig = 'ドキュメント';
  const caseStBig = '案件ステータス';

  // ① アサイン系
  const assignState = resolveMidsRollupState_(
    master, statusMap, naSet, assignBig, getAssignRelatedMids_(master, type)
  );
  if (assignState === '未着手') return '未アサイン';
  if (assignState === '対応中') return 'アサイン中';

  // ② 詳細確認
  const checkState = resolveMidsRollupState_(
    master, statusMap, naSet, checkBig, getDetailCheckMids_(master, type)
  );
  if (checkState === '未着手' || checkState === '対応中') return '詳細確認中';

  // ③ 作業中
  const inState = resolveBigRollupState_(master, statusMap, naSet, inBig);
  const outState = resolveBigRollupState_(master, statusMap, naSet, outBig);
  if (!(inState === '完了' && outState === '完了')) return '作業中';

  // ④ ドキュメント作成中
  const docState = resolveBigRollupState_(master, statusMap, naSet, docBig);
  if (docState !== '完了') return 'ﾄﾞｷｭﾒﾝﾄ作成中';

  // ⑤ 案件ステータス
  if (!hasBigInMaster_(master, caseStBig)) return '案件クローズ前対応中';

  const caseTaiouRaw   = getMidRawStatus_(statusMap, caseStBig, '案件対応');
  const sagyoRaw       = getMidRawStatus_(statusMap, caseStBig, '作業報告書');
  const fbRaw          = getMidRawStatus_(statusMap, caseStBig, 'FBアンケート');

  const caseTaiouLabel = getMidResolvedLabel_(master, statusMap, caseStBig, '案件対応');
  const sagyoLabel     = getMidResolvedLabel_(master, statusMap, caseStBig, '作業報告書');
  const fbLabel        = getMidResolvedLabel_(master, statusMap, caseStBig, 'FBアンケート');

  const caseTaiouNorm = normalizeSummaryStateStrict_(caseTaiouLabel);
  const sagyoNorm     = normalizeSummaryStateStrict_(sagyoLabel);
  const fbNorm        = normalizeSummaryStateStrict_(fbLabel);

  Logger.log(
    '[resolveCaseStatusByRule_] ' +
    '案件 raw=' + caseTaiouRaw + ' label=' + caseTaiouLabel + ' norm=' + caseTaiouNorm +
    ' / 報告書 raw=' + sagyoRaw + ' label=' + sagyoLabel + ' norm=' + sagyoNorm +
    ' / FB raw=' + fbRaw + ' label=' + fbLabel + ' norm=' + fbNorm
  );

  // 案件対応が SE対応完了 なら専用ステータス
  if (caseTaiouLabel === 'SE対応完了') {
    return 'SE対応完了';
  }

  // 案件対応が未完了ならクローズ前対応中
  if (caseTaiouNorm !== '完了') {
    return '案件クローズ前対応中';
  }

  // 報告書 or FB が残っていればクローズ前対応中
  if (!(sagyoNorm === '完了' && fbNorm === '完了')) {
    return '案件クローズ前対応中';
  }

  // ⑥ CLOSE
  const closeMids = getCloseMids_(master, type);

  // 案件クローズ中分類が無い運用ならここで CLOSE
  if (!closeMids.length) return 'CLOSE';

  const closeState = resolveMidsRollupState_(master, statusMap, naSet, caseStBig, closeMids);
  if (closeState === '完了') return 'CLOSE';

  return '案件クローズ前対応中';
}

function getMidRawStatus_(statusMap, big, mid) {
  const stKey = 'ST_' + norm_(big) + '__' + norm_(mid);
  return String(statusMap?.[stKey] ?? '').trim();
}

function getDetailCheckMids_(master, type) {
  const big = '案件確認';
  const group = (master?.groups || []).find(g => norm_(g.big) === big);
  if (!group) return ['詳細確認'];

  const mids = (group.mids || []).map(m => norm_(m.mid));
  const out = [];

  if (mids.includes('詳細確認')) out.push('詳細確認');

  return out.length ? out : ['詳細確認'];
}

function resolveBigRollupState_(master, statusMap, naSet, big) {
  big = norm_(big);
  if (!big) return '';

  if (naSet && naSet[big]) return 'ー';

  const group = (master?.groups || []).find(g => norm_(g.big) === big);
  if (!group) return '';

  const mids = (group.mids || []).filter(m => !m.is_date);
  if (!mids.length) return '';

  const states = mids.map(m => resolveMidRollupState_(master, statusMap, naSet, big, m.mid)).filter(Boolean);
  return mergeStatesToRollup_(states);
}

function resolveMidsRollupState_(master, statusMap, naSet, big, midsTarget) {
  big = norm_(big);
  if (!big) return '';

  if (naSet && naSet[big]) return 'ー';

  const targetSet = {};
  (midsTarget || []).forEach(m => targetSet[norm_(m)] = true);

  const group = (master?.groups || []).find(g => norm_(g.big) === big);
  if (!group) return '';

  const mids = (group.mids || [])
    .filter(m => !m.is_date)
    .filter(m => targetSet[norm_(m.mid)]);

  if (!mids.length) return '';

  const states = mids.map(m => resolveMidRollupState_(master, statusMap, naSet, big, m.mid)).filter(Boolean);
  return mergeStatesToRollup_(states);
}

function resolveMidRollupState_(master, statusMap, naSet, big, mid) {
  big = norm_(big);
  mid = norm_(mid);

  if (!big || !mid) return '';
  if (naSet && naSet[big]) return 'ー';

  const stKey = 'ST_' + big + '__' + mid;
  const raw = statusMap[stKey];

  if (raw == null || raw === '') return '未着手';

  const s = String(raw).trim();
  if (!s) return '未着手';
  if (s === 'ー' || s === '-' || s === '—') return 'ー';

  const mm = findMidMaster_(master, big, mid);
  if (!mm) return normalizeSummaryStateStrict_(s);

  // 日付項目は集約対象外想定だが保険
  if (mm.is_date) {
    return s ? '対応中' : '未着手';
  }

  // code / small / display のどれでも引けるようにする
  const opt = (mm.options || []).find(o =>
    norm_(o.code) === s ||
    norm_(o.small) === s ||
    norm_(o.display) === s
  );

  // summary判定は業務ルール優先
  if (opt) {
    return normalizeSummaryStateStrict_(opt.small || opt.display || s);
  }

  return normalizeSummaryStateStrict_(s);
}

function mergeStatesToRollup_(states) {
  const arr = (states || []).filter(v => v !== '');
  if (!arr.length) return '未着手';

  if (arr.every(v => v === 'ー')) return 'ー';
  if (arr.every(v => v === '完了' || v === 'ー') && arr.some(v => v === '完了')) return '完了';
  if (arr.every(v => v === '対応無' || v === 'ー')) return '完了';
  if (arr.every(v => v === '未着手' || v === 'ー')) return '未着手';
  if (arr.some(v => v === '対応中')) return '対応中';
  if (arr.some(v => v === '完了') && arr.some(v => v === '未着手')) return '対応中';

  return '未着手';
}

function normalizeSummaryStateStrict_(s) {
  s = String(s || '').trim();
  if (!s) return '未着手';

  if (s === 'ー' || s === '-' || s === '—') return 'ー';

  // 対応無は完了扱い
  if (s.includes('対応無')) return '完了';

  // 未着手系
  if (
    s.includes('未着手') ||
    s.includes('未入荷') ||
    s.includes('未実施') ||
    s.includes('待ち')
  ) {
    return '未着手';
  }

  // 完了系
  if (
    s.includes('完了') ||
    s.includes('済') ||
    s.includes('クローズ') ||
    s.includes('機器無')
  ) {
    return '完了';
  }

  // それ以外は対応中
  return '対応中';
}

function findMidMaster_(master, big, mid) {
  const group = (master?.groups || []).find(g => norm_(g.big) === norm_(big));
  if (!group) return null;
  return (group.mids || []).find(m => norm_(m.mid) === norm_(mid)) || null;
}

function getStatusHeaderByBigMid_(master, big, mid) {
  const mm = findMidMaster_(master, big, mid);
  return mm?.status_header || '';
}

function hasBigInMaster_(master, big) {
  return (master?.groups || []).some(g => norm_(g.big) === norm_(big));
}

function getFieldWorkBigName_(type) {
  return normalizeCaseType_(type) === 'CL' ? '現地作業' : '現地・リモート作業';
}

function getAssignRelatedMids_(master, type) {
  const big = '案件確認';
  const group = (master?.groups || []).find(g => norm_(g.big) === big);
  if (!group) return ['アサイン'];

  const exists = (name) => (group.mids || []).some(m => norm_(m.mid) === norm_(name));
  const out = [];

  if (exists('アサイン')) out.push('アサイン');
  if (exists('アサインメール')) out.push('アサインメール');

  return out.length ? out : ['アサイン'];
}

function getPreCloseMids_(master, type) {
  const group = (master?.groups || []).find(g => norm_(g.big) === '案件ステータス');
  if (!group) return [];

  const names = (group.mids || []).map(m => norm_(m.mid));
  const out = [];

  if (names.includes('SE対応完了')) out.push('SE対応完了');
  if (names.includes('案件対応')) out.push('案件対応');
  if (names.includes('作業報告書')) out.push('作業報告書');
  if (names.includes('FBアンケート')) out.push('FBアンケート');

  return out;
}

function getCloseMids_(master, type) {
  const group = (master?.groups || []).find(g => norm_(g.big) === '案件ステータス');
  if (!group) return [];

  const names = (group.mids || []).map(m => norm_(m.mid));
  const out = [];

  if (names.includes('案件クローズ')) out.push('案件クローズ');

  return out;
}

/** =========================
 * コメント列関連
 * ========================= */

function getCommentTargetBigs_(type) {
  const t = normalizeCaseType_(type);
  if (t === 'CL') return ['案件管理', '社内作業', '現地作業', 'ドキュメント', '案件ステータス'];
  return ['案件管理', '社内作業', '現地・リモート作業', 'ドキュメント', '案件ステータス'];
}

function getCommentCol_(hmS, type, big) {
  const candidates = [
    norm_(big) + '_コメント',
    norm_(big) + '_備考',
    norm_(big) + '_comment',
    norm_(big) + '_Comment'
  ];

  for (let i = 0; i < candidates.length; i++) {
    if (hmS[candidates[i]]) return hmS[candidates[i]];
  }
  return 0;
}

/** =========================
 * 統合モーダル
 * ========================= */

function api_getUnifiedCaseModal(req) {
  try {
    req = req || {};
    const type = normalizeCaseType_(req.type || req.target || 'SV');
    const key = norm_(req.key || req.mitsuNo);
    if (!key) return { ok: false, error: 'key が空です' };

    const statusMasterRes = getStatusMasterCached_(type);
    if (!statusMasterRes.ok) {
      return { ok: false, error: statusMasterRes.error || 'ステータスマスタ取得失敗' };
    }

    const caseStatusRes = buildCaseStatusResponse_(type, key, statusMasterRes);
    if (!caseStatusRes.ok) {
      return { ok: false, error: caseStatusRes.error || '案件ステータス取得失敗' };
    }

    const detailRes = api_getCaseDetail(type, key);
    if (!detailRes.ok) {
      return { ok: false, error: detailRes.error || '詳細取得失敗' };
    }

    const inputRes = api_getCaseByKey({ type, key });
    if (!inputRes.ok) {
      return { ok: false, error: inputRes.error || '案件情報取得失敗' };
    }

    return {
      ok: true,
      type,
      key,
      detail: detailRes.data || {},
      statusMaster: statusMasterRes,
      caseStatus: {
        statusMap: caseStatusRes.statusMap || {},
        naGroups: caseStatusRes.naGroups || {},
        currentCode: caseStatusRes.currentCode || ''
      },
      input: inputRes.data || {}
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getUnifiedModalDetail(req) {
  try {
    req = req || {};
    const type = normalizeCaseType_(req.type || 'SV');
    const key = norm_(req.key);
    if (!key) return { ok: false, error: 'key が空です' };

    const res = api_getCaseDetail(type, key);
    if (!res.ok) return { ok: false, error: res.error || '詳細取得失敗' };

    return {
      ok: true,
      type,
      key,
      detail: res.data || {}
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getUnifiedModalInput(req) {
  try {
    req = req || {};
    const type = normalizeCaseType_(req.type || 'SV');
    const key = norm_(req.key);
    if (!key) return { ok: false, error: 'key が空です' };

    const res = api_getCaseByKey({ type, key });
    if (!res.ok) return { ok: false, error: res.error || '案件情報取得失敗' };

    return {
      ok: true,
      type,
      key,
      input: res.data || {}
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_saveUnifiedCaseModal(req) {
  try {
    req = req || {};
    const type = normalizeCaseType_(req.type || 'SV');
    const key = norm_(req.key);
    if (!key) return { ok: false, error: 'key が空です' };

    const statusReq = req.status || {};
    const inputReq = req.input || {};

    const statusRes = api_saveCaseStatus({
      type,
      key,
      current_status_code: statusReq.current_status_code || '',
      statusMap: statusReq.statusMap || {},
      naBigs: Array.isArray(statusReq.naBigs) ? statusReq.naBigs : []
    });
    if (!statusRes.ok) return { ok: false, error: statusRes.error || 'ステータス保存失敗' };

    const inputRes = api_upsertInput({
      type,
      key,
      values: inputReq.values || {}
    });
    if (!inputRes.ok) return { ok: false, error: inputRes.error || '案件情報保存失敗' };

    return api_getUnifiedCaseModal({ type, key });
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/** =========================
 * 高速化ヘルパー
 * ========================= */

function getHeaderMapCached_(sheet, headerRow) {
  const key = sheet.getParent().getId() + '::' + sheet.getSheetId() + '::' + Number(headerRow || 1);

  if (__STATUS_EXEC_CACHE.headerMap[key]) {
    return __STATUS_EXEC_CACHE.headerMap[key];
  }

  const map = buildHeaderMap_(sheet, headerRow || 1);
  __STATUS_EXEC_CACHE.headerMap[key] = map;
  return map;
}

function findRowByKeyCached_(sheet, keyCol, key, startRow) {
  key = norm_(key);
  startRow = Number(startRow || 2);

  if (!sheet || !keyCol || !key) return 0;

  const cacheKey = [
    sheet.getParent().getId(),
    sheet.getSheetId(),
    keyCol,
    startRow
  ].join('::');

  if (!__STATUS_EXEC_CACHE.rowIndex[cacheKey]) {
    __STATUS_EXEC_CACHE.rowIndex[cacheKey] = buildRowIndexMap_(sheet, keyCol, startRow);
  }

  const map = __STATUS_EXEC_CACHE.rowIndex[cacheKey];
  return Number(map[key] || 0);
}

function invalidateRowIndexCache_(sheet, keyCol, startRow) {
  if (!sheet || !keyCol) return;
  startRow = Number(startRow || 2);

  const cacheKey = [
    sheet.getParent().getId(),
    sheet.getSheetId(),
    keyCol,
    startRow
  ].join('::');

  delete __STATUS_EXEC_CACHE.rowIndex[cacheKey];
}

function buildRowIndexMap_(sheet, keyCol, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return {};

  const values = sheet.getRange(startRow, keyCol, lastRow - startRow + 1, 1).getDisplayValues();
  const map = {};

  for (let i = 0; i < values.length; i++) {
    const k = norm_(values[i][0]);
    if (!k) continue;
    if (!map[k]) map[k] = startRow + i;
  }

  return map;
}

function readSheetRowValues_(sheet, rowNo) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol || rowNo < 1) return [];
  return sheet.getRange(rowNo, 1, 1, lastCol).getValues()[0];
}

function readRowAsObjectFast_(sheet, rowNo, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol || rowNo < 1) return {};

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const row = sheet.getRange(rowNo, 1, 1, lastCol).getValues()[0];

  const obj = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (!h) continue;
    obj[h] = row[i];
  }
  return obj;
}