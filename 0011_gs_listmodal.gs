/** ======================================================================
 * Modal API
 * - 案件情報
 * - ステータス
 * - 統合取得 / 統合保存
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
    '最寄り':'最寄り',
    '住所': '住所',
    '作業内容': '作業内容',
    '台数':'台数',
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
    const hm = buildHeaderMap_(sh, headerRow);
    const keyCol = hm[INPUT_CONFIG.KEY_HEADER];
    if (!keyCol) return { ok: false, error: 'key header not found: 見積番号' };

    const rowNo = findRowByKey_(sh, keyCol, key, headerRow + 1);
    if (!rowNo) return { ok: false, error: 'not found: ' + key + ' (' + type + ')' };

    const obj = readRowAsObject_(sh, rowNo, headerRow, true);
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

    const sh = getSS_('DB').getSheetByName(STATUS_APP_CONFIG.MASTER_SHEET);
    if (!sh) {
      return { ok: false, error: 'マスタシートが見つかりません: ' + STATUS_APP_CONFIG.MASTER_SHEET };
    }

    const headerRow = STATUS_APP_CONFIG.MASTER_HEADER_ROW;
    const headerMap = buildHeaderMap_(sh, headerRow);
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

    const colTarget   = headerMap['対象'];
    const colCode     = headerMap['コード'];
    const colBig      = headerMap['大分類'];
    const colMid      = headerMap['中分類'];
    const colSmall    = headerMap['小分類'];
    const colDisp     = headerMap['表示名'];
    const colSort     = headerMap['sort'];
    const colClosed   = headerMap['is_closed'];
    const colColor    = headerMap['is_color'];
    const colForce    = headerMap['force'];
    const colDefault  = headerMap['default'];
    const colNaEnable = headerMap['na_enable'];

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
        na_enable: colNaEnable ? String(r[colNaEnable - 1]).toUpperCase() === 'TRUE' : false
      });
    }

    return {
      ok: true,
      target,
      rows,
      groups: buildGroups_(rows),
      naOptions: STATUS_APP_CONFIG.NA_OPTIONS
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function buildGroups_(rows) {
  const bigMap = {};

  (rows || []).forEach(r => {
    if (!r.big || !r.mid) return;

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

    if (!bm.midMap[r.mid]) {
      bm.midMap[r.mid] = {
        mid: r.mid,
        sort: r.sort || 0,
        force: !!r.force,
        is_date: !!r.default_value,
        options: []
      };
    }

    const mm = bm.midMap[r.mid];
    mm.sort = Math.min(mm.sort || 0, r.sort || mm.sort || 0);
    if (r.force) mm.force = true;
    if (r.default_value) mm.is_date = true;

    mm.options.push({
      code: r.code,
      small: r.small,
      sort: r.sort || 0,
      is_closed: !!r.is_closed,
      is_color: !!r.is_color,
      is_default: !!r.is_default
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

    const shS = getDisplayStatusSheet_(type);
    if (!shS) return { ok: true, statusMap: {}, naGroups: {}, currentCode: '' };

    const hmS = buildHeaderMap_(shS, 1);
    const colKey = hmS[STATUS_APP_CONFIG.KEY_HEADER];
    if (!colKey) return { ok: false, error: '見積番号列がありません' };

    const rowNo = findRowByKey_(shS, colKey, key, 2);
    if (!rowNo) return { ok: true, statusMap: {}, naGroups: {}, currentCode: '' };

    const master = api_getStatusMaster({ target: type });
    if (!master.ok) return { ok: false, error: 'マスタ取得失敗' };

    const statusMap = {};

    (master.groups || []).forEach(g => {
      (g.mids || []).forEach(m => {
        const col = hmS[statusColName_(g.big, m.mid)];
        if (!col) return;

        const v = shS.getRange(rowNo, col).getValue();
        if (v == null || v === '') return;

        const stKey = 'ST_' + g.big + '__' + m.mid;
        statusMap[stKey] = m.is_date ? toInputDateValue_(v) : String(v).trim();
      });
    });

    getCommentTargetBigs_(type).forEach(big => {
      const col = getCommentCol_(hmS, type, big);
      if (!col) return;

      const v = shS.getRange(rowNo, col).getValue();
      if (v != null && String(v).trim() !== '') {
        statusMap['CM_' + big] = String(v);
      }
    });

    let currentCode = '';
    if (hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]) {
      currentCode = norm_(shS.getRange(rowNo, hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]).getValue());
    }

    return { ok: true, statusMap, naGroups: {}, currentCode };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
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

    const hmS = buildHeaderMap_(shS, 1);
    const hmM = buildHeaderMap_(shM, STATUS_APP_CONFIG.CASE_HEADER_ROW);

    const colKeyS = hmS[STATUS_APP_CONFIG.KEY_HEADER];
    if (!colKeyS) return { ok: false, error: '見積番号列がありません' };

    const rowNo = findRowByKey_(shS, colKeyS, key, 2);
    if (!rowNo) return { ok: false, error: '見積番号が見つかりません: ' + key };

    const rowNoM = findRowByKey_(shM, hmM[STATUS_APP_CONFIG.KEY_HEADER], key, STATUS_APP_CONFIG.CASE_HEADER_ROW + 1);
    if (!rowNoM) return { ok: false, error: '案件管理表に見積番号が見つかりません: ' + key };

    const master = api_getStatusMaster({ target: type });
    if (!master.ok) return { ok: false, error: 'マスタ取得失敗' };

    const naSet = buildNaSet_(naBigs);

    clearTargetStatusCols_(shS, rowNo, hmS, master);
    applyNaMarks_(shS, rowNo, hmS, master, naSet);
    applyStatusValues_(shS, rowNo, hmS, statusMap, naSet);
    writeCommentValues_(shS, rowNo, hmS, type, statusMap);
    writeStatusMeta_(shS, rowNo, hmS, currentCode);

    syncCaseSheetAfterStatusSave_(type, rowNoM, shM, hmM, master, statusMap, naSet, currentCode);

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
  const cols = [];
  (master.groups || []).forEach(g => {
    (g.mids || []).forEach(m => {
      const col = hmS[statusColName_(g.big, m.mid)];
      if (col) cols.push(col);
    });
  });

  writeSameRowByCols_(shS, rowNo, cols, () => '');
}

function applyNaMarks_(shS, rowNo, hmS, master, naSet) {
  const cols = [];
  (master.groups || []).forEach(g => {
    const big = norm_(g.big);
    if (!big || !naSet[big]) return;

    (g.mids || []).forEach(m => {
      const col = hmS[statusColName_(big, m.mid)];
      if (col) cols.push(col);
    });
  });

  writeSameRowByCols_(shS, rowNo, cols, () => 'ー');
}

function applyStatusValues_(shS, rowNo, hmS, statusMap, naSet) {
  const setMap = {};
  const cols = [];

  Object.keys(statusMap || {}).forEach(stKey => {
    if (!String(stKey).startsWith('ST_')) return;

    const raw = String(stKey).slice(3);
    const p = raw.split('__');
    if (p.length < 2) return;

    const big = p[0];
    const mid = p.slice(1).join('__');

    if (naSet[norm_(big)]) return;

    const col = hmS[statusColName_(big, mid)];
    if (!col) return;

    cols.push(col);
    setMap[col] = statusMap[stKey];
  });

  writeSameRowByCols_(shS, rowNo, cols, col => setMap[col]);
}

function writeCommentValues_(shS, rowNo, hmS, type, statusMap) {
  getCommentTargetBigs_(type).forEach(big => {
    const col = getCommentCol_(hmS, type, big);
    if (!col) return;

    const v = statusMap['CM_' + big] == null ? '' : String(statusMap['CM_' + big]);
    shS.getRange(rowNo, col).setValue(v);
  });
}

function writeStatusMeta_(shS, rowNo, hmS, currentCode) {
  const updater = getActiveUserNameOrEmail_();
  Logger.log('[writeStatusMeta_] row=' + rowNo + ' updater=' + updater);

  if (hmS[STATUS_APP_CONFIG.UPDATE_AT_HEADER]) {
    shS.getRange(rowNo, hmS[STATUS_APP_CONFIG.UPDATE_AT_HEADER]).setValue(nowYmd_(true));
  } else {
    Logger.log('[writeStatusMeta_] updated_at col not found');
  }

  if (hmS[STATUS_APP_CONFIG.UPDATE_BY_HEADER]) {
    shS.getRange(rowNo, hmS[STATUS_APP_CONFIG.UPDATE_BY_HEADER]).setValue(updater);
  } else {
    Logger.log('[writeStatusMeta_] updated_by col not found');
  }

  if (hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]) {
    shS.getRange(rowNo, hmS[STATUS_APP_CONFIG.CURRENT_CODE_HEADER]).setValue(currentCode || '');
  }

  if (hmS['open/close']) {
    shS.getRange(rowNo, hmS['open/close']).setValue(openCloseFromCode_(currentCode));
  }
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

    const detailRes = api_getCaseDetail(type, key);
    if (!detailRes.ok) return { ok: false, error: detailRes.error || '詳細取得失敗' };

    const statusMasterRes = api_getStatusMaster({ target: type });
    if (!statusMasterRes.ok) return { ok: false, error: statusMasterRes.error || 'ステータスマスタ取得失敗' };

    const caseStatusRes = api_getCaseStatus({ type, key });
    if (!caseStatusRes.ok) return { ok: false, error: caseStatusRes.error || '案件ステータス取得失敗' };

    const inputRes = api_getCaseByKey({ type, key });
    if (!inputRes.ok) return { ok: false, error: inputRes.error || '案件情報取得失敗' };

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