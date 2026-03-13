/** ======================================================================
 * Display API / View Builder
 * ====================================================================== */

const DETAIL_APP_CONFIG = {
  HEADER_ROW: 1,
  KEY_HEADER: '見積番号',
  TZ: 'Asia/Tokyo'
};

function api_getCaseDetail(type, mitsuNo) {
  try {
    type = normalizeCaseType_(type || 'SV');
    mitsuNo = norm_(mitsuNo);

    if (!mitsuNo) {
      return { ok: false, error: '見積番号が未指定です' };
    }

    const shBase = getDisplayCaseSheet_(type);
    const shStatus = getDisplayStatusSheet_(type);

    if (!shBase) {
      return { ok: false, error: 'Displayに案件管理表がありません' };
    }
    if (!shStatus) {
      return { ok: false, error: 'Displayに案件ステータスがありません' };
    }

    const headerRow = DETAIL_APP_CONFIG.HEADER_ROW;
    const hmBase = buildHeaderMap_(shBase, headerRow);
    const hmStatus = buildHeaderMap_(shStatus, headerRow);

    const colKeyBase = hmBase[DETAIL_APP_CONFIG.KEY_HEADER];
    const colKeyStatus = hmStatus[DETAIL_APP_CONFIG.KEY_HEADER];

    if (!colKeyBase) return { ok: false, error: '案件管理表に見積番号列がありません' };
    if (!colKeyStatus) return { ok: false, error: '案件ステータスに見積番号列がありません' };

    const baseRowNo = findRowByKey_(shBase, colKeyBase, mitsuNo, headerRow + 1);
    if (!baseRowNo) return { ok: false, error: '案件が見つかりません: ' + mitsuNo };

    const statusRowNo = findRowByKey_(shStatus, colKeyStatus, mitsuNo, headerRow + 1);

    const baseRowObj = readRowAsObject_(shBase, baseRowNo, headerRow, false) || {};
    const statusRowObj = statusRowNo ? readRowAsObject_(shStatus, statusRowNo, headerRow, false) : {};

    return {
      ok: true,
      data: buildCaseDetailPayload_(type, baseRowObj, statusRowObj)
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function buildCaseDetailPayload_(type, b, s) {
  b = b || {};
  s = s || {};

  const outsidePrefix = (normalizeCaseType_(type) === 'CL') ? '現地作業' : '現地・リモート作業';

  return {
    title: '［' + (b['見積番号'] || '') + '］' + (b['顧客名'] || '') + '｜' + (b['案件概要'] || ''),

    mitsuNo: b['見積番号'] || '',
    customer: b['顧客名'] || '',
    class: b['分類'] || '',
    status: b['ステータス'] || '',
    openClose: b['open/close'] || s['open/close'] || '',
    currentStatusCode: b['current_status_code'] || s['current_status_code'] || '',

    summary: b['案件概要'] || '',
    workContent: b['作業内容'] || '',
    unitCount: b['台数'] || '',

    reqDay: fmtDetailDate_(b['作業依頼日']),
    acceptPlan: fmtDetailDate_(b['検収予定']),
    mitsuUrl: b['見積'] || '',
    workFolderUrl: b['作業フォルダ'] || '',

    sales: b['担当営業'] || '',
    pre: b['担当プリ'] || '',
    supervisor: b['監督者'] || '',
    manager: b['管理者'] || '',
    support: b['サポート'] || '',
    bpUse: b['BP利用'] || '',
    tag: b['カテゴリタグ'] || '',
    level: b['作業難度'] || '',
    mitsuKosu: b['見積工数'] || '',
    workStyle: b['作業形態'] || '',
    location: b['現地作業場所'] || '',
    address: b['住所'] || '',
    note: b['備考'] || '',

    stAssign: s['アサイン_アサイン'] || '',
    stAssignMail: s['アサイン_アサインメール'] || '',

    stConfirm: s['案件確認_案件確認'] || '',
    stHearing: s['案件確認_ヒアリングシート'] || '',
    stConfirmStart: fmtDetailDate_(s['案件確認_作業期間：開始']),
    stConfirmEnd: fmtDetailDate_(s['案件確認_作業期間：終了']),
    stConfirmComment: s['案件確認_コメント'] || '',

    stInhouseStart: fmtDetailDate_(s['社内作業_作業期間：開始']),
    stInhouseEnd: fmtDetailDate_(s['社内作業_作業期間：終了']),
    stArrival: s['社内作業_機器入荷'] || '',
    stInhouseWork: s['社内作業_作業'] || '',
    stShip: s['社内作業_機器出荷'] || '',
    stInhouseComment: s['社内作業_コメント'] || '',

    stOnsiteStart: fmtDetailDate_(s[outsidePrefix + '_作業期間：開始']),
    stOnsiteEnd: fmtDetailDate_(s[outsidePrefix + '_作業期間：終了']),
    stOnsiteWork: s[outsidePrefix + '_作業'] || '',
    stNextWork: fmtDetailDate_(s[outsidePrefix + '_次回作業']),
    stOnsiteComment: s[outsidePrefix + '_コメント'] || '',

    stDocCreate: s['ドキュメント_作成'] || '',
    stDocReview: s['ドキュメント_レビュー'] || '',

    stSeDone: s['案件ステータス_SE対応完了'] || '',
    stCaseClose: s['案件ステータス_案件クローズ'] || '',

    updatedAt: fmtDetailDate_(b['updated_at'] || s['updated_at']),
    updatedBy: b['updated_by'] || '',
    rev: b['rev'] || '',
    deleted: b['deleted'] || ''
  };
}

function fmtDetailDate_(v) {
  if (v == null || v === '') return '';

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, DETAIL_APP_CONFIG.TZ, 'yyyy/MM/dd');
  }

  return String(v);
}

/** =========================
 * 保存後に案件管理表へ反映する表示文
 * ========================= */

function syncCaseSheetAfterStatusSave_(type, rowNoM, shM, hmM, master, statusMap, naSet, currentCode) {
  const updater = getActiveUserNameOrEmail_();

  const inBig   = '社内作業';
  const outBig  = (normalizeCaseType_(type) === 'CL') ? '現地作業' : '現地・リモート作業';
  const docBig  = 'ドキュメント';

  const inSt  = resolveBigRollupState_(master, statusMap, naSet, inBig);
  const outSt = resolveBigRollupState_(master, statusMap, naSet, outBig);
  const docSt = hasBigInMaster_(master, docBig)
    ? resolveBigRollupState_(master, statusMap, naSet, docBig)
    : '';

  const caseSt = resolveCaseStatusByRule_(type, master, statusMap, naSet);

  const summaryText = [
    `案件ST：${caseSt || ''}`,
    `社内：${inSt || ''}　現地：${outSt || ''}` + (docSt ? `　Doc.：${docSt}` : '')
  ].join('\n');

  const updates = {};

  // 既存要約欄
  if (hmM[STATUS_APP_CONFIG.CASE_STATUS_HEADER]) {
    updates[hmM[STATUS_APP_CONFIG.CASE_STATUS_HEADER]] = summaryText;
  }

  // 個別欄
  if (hmM['案件ST']) updates[hmM['案件ST']] = caseSt || '';
  if (hmM['社内ST']) updates[hmM['社内ST']] = inSt || '';
  if (hmM['現地ST']) updates[hmM['現地ST']] = outSt || '';
  if (hmM['Doc.ST']) updates[hmM['Doc.ST']] = docSt || '';
  if (hmM['DocST'])  updates[hmM['DocST']]  = docSt || '';

  // メタ
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

function buildCodeToSmallMap_(masterRows) {
  const map = {};
  (masterRows || []).forEach(r => {
    map[String(r.code)] = norm_(r.small);
  });
  return map;
}

function buildCodeToRollupMap_(masterRows) {
  const map = {};
  (masterRows || []).forEach(r => {
    const code = String(r.code || '').trim();
    if (!code) return;

    const rollup = norm_(r.rollup_state || r['集計状態'] || '');
    map[code] = rollup || '未着手';
  });
  return map;
}

function getDefaultSmallForMid_(options) {
  const d = (options || []).find(o => !!o.is_default);
  return d ? norm_(d.small) : '未着手';
}

function getBigState_(master, statusMap, bigName, naSet) {
  bigName = norm_(bigName);
  if (!bigName) return '未着手';
  if (naSet && naSet[bigName]) return 'ー';

  const g = (master.groups || []).find(x => norm_(x.big) === bigName);
  if (!g) return '未着手';

  const code2small = buildCodeToSmallMap_(master.rows || []);
  const code2rollup = buildCodeToRollupMap_(master.rows || []);
  const states = [];

  (g.mids || []).forEach(m => {
    const optSmalls = (m.options || []).map(o => norm_(o.small));
    const hasStatusLike = optSmalls.length > 0;
    if (!hasStatusLike) return;

    const stKey = 'ST_' + bigName + '__' + m.mid;
    const v = statusMap ? statusMap[stKey] : '';

    let state = '';
    if (norm_(v) === 'ー') {
      state = 'ー';
    } else if (v == null || norm_(v) === '') {
      const def = (m.options || []).find(o => !!o.is_default);
      if (def) {
        state = code2rollup[String(def.code)] || normalizeStateFallback_(def.small);
      } else {
        state = '未着手';
      }
    } else {
      const code = norm_(v);
      state = code2rollup[code] || normalizeStateFallback_(code2small[code]);
    }

    states.push(state || '未着手');
  });

  if (states.length === 0) return '未着手';
  if (states.every(s => s === 'ー')) return 'ー';
  if (states.every(s => s === '完了')) return '完了';
  if (states.every(s => s === '未着手')) return '未着手';
  return '対応中';
}

function normalizeStateFallback_(s) {
  s = norm_(s);
  if (!s) return '未着手';
  if (s === 'ー' || s === '-' || s === '—') return 'ー';
  if (s.indexOf('完了') !== -1) return '完了';
  if (s.indexOf('対応中') !== -1 || s.indexOf('作業中') !== -1 || s.indexOf('レビュー') !== -1 || s.indexOf('修正中') !== -1) return '対応中';
  if (s.indexOf('未着手') !== -1 || s.indexOf('未入荷') !== -1 || s.indexOf('対応有') !== -1 || s.indexOf('固定') !== -1) return '未着手';
  return '未着手';
}

function computeCaseSt_(master, statusMap, naSet) {
  const closeSt = getMidState_(master, statusMap, '案件ステータス', '案件クローズ');
  if (closeSt === '完了') return 'CLOSE';

  const assignSt = getMidState_(master, statusMap, '案件確認', 'アサイン');

  // アサイン未完了なら最優先で「アサイン中」
  if (assignSt !== '完了') return 'アサイン中';

  // アサイン完了後の事前確認系
  const preStates = [
    getMidState_(master, statusMap, '案件確認', '詳細確認'),
    getMidState_(master, statusMap, '案件管理', 'ヒアリングシート'),
  ];

  const preSt = summarizeStates_(preStates);
  if (preSt !== '完了') return '事前確認中';

  return '作業対応中';
}

function getMidState_(master, statusMap, bigName, midName) {
  bigName = norm_(bigName);
  midName = norm_(midName);

  const g = (master.groups || []).find(x => norm_(x.big) === bigName);
  if (!g) return '未着手';

  const m = (g.mids || []).find(x => norm_(x.mid) === midName);
  if (!m) return '未着手';

  const code2small = buildCodeToSmallMap_(master.rows || []);
  const code2rollup = buildCodeToRollupMap_(master.rows || []);
  const stKey = 'ST_' + bigName + '__' + midName;
  const v = statusMap ? statusMap[stKey] : '';

  let state = '';
  if (norm_(v) === 'ー') {
    state = 'ー';
  } else if (v == null || norm_(v) === '') {
    const def = (m.options || []).find(o => !!o.is_default);
    if (def) {
      state = code2rollup[String(def.code)] || normalizeStateFallback_(def.small);
    } else {
      state = '未着手';
    }
  } else {
    const code = norm_(v);
    state = code2rollup[code] || normalizeStateFallback_(code2small[code]);
  }

  return state || '未着手';
}

function summarizeStates_(states) {
  states = (states || []).filter(Boolean);
  if (states.length === 0) return '未着手';
  if (states.every(s => s === 'ー')) return 'ー';
  if (states.every(s => s === '完了')) return '完了';
  if (states.every(s => s === '未着手')) return '未着手';
  return '対応中';
}

function buildCodeToRollupMap_(masterRows) {
  const map = {};
  (masterRows || []).forEach(r => {
    const code = String(r.code || '').trim();
    if (!code) return;
    const rollup = norm_(r.rollup_state || r['集計状態'] || '');
    map[code] = rollup || '';
  });
  return map;
}

function normalizeStateFallback_(s) {
  s = norm_(s);
  if (!s) return '未着手';
  if (s === 'ー' || s === '-' || s === '—') return 'ー';
  if (s.indexOf('完了') !== -1) return '完了';
  if (s.indexOf('対応中') !== -1 || s.indexOf('作業中') !== -1 || s.indexOf('レビュー') !== -1 || s.indexOf('修正中') !== -1) return '対応中';
  if (s.indexOf('未着手') !== -1 || s.indexOf('未入荷') !== -1 || s.indexOf('対応有') !== -1 || s.indexOf('固定') !== -1) return '未着手';
  return '未着手';
}

function buildStatusTextForView_(type, master, statusMap, naSet) {
  const caseSt = computeCaseSt_(master, statusMap, naSet);
  const inSt = getBigState_(master, statusMap, '社内作業', naSet);
  const onBig = getOutsideBigName_(type);
  const onSt = getBigState_(master, statusMap, onBig, naSet);
  const docSt = getBigState_(master, statusMap, 'ドキュメント', naSet);

  return [
    '案件ST：' + caseSt,
    '社内：' + inSt,
    '現地：' + onSt,
    'Doc.：' + docSt
  ].join('\n');
}

function buildCommentsNoteText_(type, statusMap) {
  statusMap = statusMap || {};

  const bfC = norm_(statusMap['CM_案件確認'] || '');
  const inC = norm_(statusMap['CM_社内作業'] || '');
  const onC = norm_(statusMap['CM_' + getOutsideBigName_(type)] || '');

  const lines = [];
  if (bfC) lines.push('案件確認：' + bfC);
  if (inC) lines.push('社内作業：' + inC);
  if (onC) lines.push('現地作業：' + onC);

  return lines.join('\n');
}

function getCommentTargetBigs_(type) {
  return ['案件確認', '社内作業', getOutsideBigName_(type)];
}

function getOutsideBigName_(type) {
  return normalizeCaseType_(type) === 'CL' ? '現地作業' : '現地・リモート作業';
}

function commentColName_(type, big) {
  type = normalizeCaseType_(type);
  big = norm_(big);

  if (big === '案件確認') return '案件確認_コメント';
  if (big === '社内作業') return '社内作業_コメント';
  if (big === '現地・リモート作業' || big === '現地作業') {
    return type === 'CL' ? '現地作業_コメント' : '現地・リモート作業_コメント';
  }
  return big ? (big + '_コメント') : '';
}

function getCommentCol_(hm, type, big) {
  const primary = commentColName_(type, big);
  if (primary && hm[primary]) return hm[primary];

  const alt = (primary === '現地作業_コメント')
    ? '現地・リモート作業_コメント'
    : '現地作業_コメント';

  return hm[alt] || 0;
}

function buildWorkSummary2Line_(row, headerMap) {
  const work = String(getByHeader_(row, headerMap, '作業内容') || '').trim();
  const style = String(getByHeader_(row, headerMap, '作業形態') || '').trim();
  const placeRaw = String(getByHeader_(row, headerMap, '現地作業場所') || '').trim();

  const styleText = cleanWorkStyleText_(style);
  const placeText = cleanPlaceText_(placeRaw);
  const sub = [styleText, placeText].filter(Boolean).join(' / ');

  if (work && sub) return work + '\n' + sub;
  if (work) return work;
  if (sub) return sub;
  return '';
}

function cleanWorkStyleText_(s) {
  return String(s || '').trim().replace(/[／/]/g, '・');
}

function cleanPlaceText_(s) {
  const v = String(s || '').trim();
  if (!v) return '';

  const m = v.match(/^\s*\[([^\]]*)\]\s*(.*)\s*$/);
  if (!m) return v;

  const prefs = String(m[1] || '')
    .split('/')
    .map(x => x.trim())
    .filter(Boolean)
    .join(' ');

  const free = String(m[2] || '').trim();
  return [prefs, free].filter(Boolean).join(' ');
}

function getByHeader_(row, headerMap, headerName) {
  const col = headerMap[headerName];
  if (!col) return '';
  return row[col - 1];
}

function hasBigInMaster_(master, big) {
  return (master?.groups || []).some(g => norm_(g.big) === norm_(big));
}

function resolveCurrentCaseStatusLabel_(master, currentCode, fallback) {
  const code = norm_(currentCode);
  if (!code) return fallback || '';

  const row = (master?.rows || []).find(r => norm_(r.code) === code);
  if (!row) return fallback || '';

  return row.display || row.small || fallback || '';
}

function resolveBigSummaryState_(master, statusMap, naSet, big) {
  big = norm_(big);
  if (!big) return '';

  if (naSet && naSet[big]) return 'ー';

  const group = (master?.groups || []).find(g => norm_(g.big) === big);
  if (!group) return '';

  const mids = group.mids || [];
  if (!mids.length) return '';

  const states = mids.map(m => {
    const key = 'ST_' + big + '__' + m.mid;
    const raw = statusMap[key];
    return normalizeSummaryValueForCaseSheet_(raw, !!m.is_date);
  }).filter(v => v !== '');

  if (!states.length) return '未着手';

  // すべて対応無
  if (states.every(v => v === 'ー')) return 'ー';

  // 1つでも対応中があれば対応中
  if (states.some(v => v === '対応中')) return '対応中';

  // 完了と対応無のみで構成、かつ完了が1つでもあるなら完了
  if (states.some(v => v === '完了') && states.every(v => v === '完了' || v === 'ー')) {
    return '完了';
  }

  // すべて未着手 or 対応無
  if (states.every(v => v === '未着手' || v === 'ー')) return '未着手';

  // 完了・未着手が混在 → 対応中
  if (states.some(v => v === '完了') && states.some(v => v === '未着手')) return '対応中';

  return '未着手';
}

function normalizeSummaryValueForCaseSheet_(raw, isDate) {
  if (raw == null || raw === '') return '';

  // 日付項目は値が入っていれば完了扱い
  if (isDate) return '完了';

  const s = String(raw).trim();
  if (!s) return '';
  if (s === 'ー' || s === '-' || s === '—') return 'ー';
  if (s.includes('対応無')) return 'ー';
  if (s.includes('完了')) return '完了';
  if (s.includes('未着手')) return '未着手';
  if (s.includes('未入荷')) return '未着手';
  if (s.includes('対応中')) return '対応中';

  // コードや表示名が直接入るケースに備えて保険
  return '対応中';
}