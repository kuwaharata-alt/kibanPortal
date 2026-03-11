/** ======================================================================
 * 　案件一覧表示用 API（高速化版）
 * ====================================================================== */

const WEBVIEW = {
  SHEET_SV: 'SV案件管理表',
  SHEET_SVST: 'SV案件ステータス',
  SHEET_CL: 'CL案件管理表',
  SHEET_CLST: 'CL案件ステータス',

  KEY_HEADER: '見積番号',
  STATUS_HEADER: 'ステータス',

  CLOSED_WORDS: ['クローズ'],
  MAX_ROWS: 5000,
};

const VIEW_CONFIG = {
  SV: {
    sheetName: WEBVIEW.SHEET_SV,
    keyHeader: WEBVIEW.KEY_HEADER,
    list: ['ステータス', '見積番号', '案件名', '作業概要', '担当SE', '作業期間', '備考'],
  },
  CL: {
    sheetName: WEBVIEW.SHEET_CL,
    keyHeader: WEBVIEW.KEY_HEADER,
    list: ['ステータス', '見積番号', '案件名', '作業概要', '担当SE', '作業期間', '備考'],
  }
};

const VIRTUAL_HEADERS = [
  '案件名',
  '担当SE',
  '作業期間',
  '作業概要',
  'カテゴリ'
];

/** ----------------------------------------------------
 *   メインAPI（一覧）
 *  ---------------------------------------------------- */
function api_list(params) {
  const p = params || {};
  const type = String(p.type || 'SV').toUpperCase();
  const cfg = VIEW_CONFIG[type] || VIEW_CONFIG.SV;

  const q = String(p.q || '').trim().toLowerCase();
  const hideClosed = p.hideClosed !== false;

  const pageSize = Math.min(Math.max(Number(p.pageSize || 50), 10), 500);
  const page = Math.max(Number(p.page || 1), 1);

  const sh = getSS_('Display').getSheetByName(cfg.sheetName);
  if (!sh) {
    return { ok: false, error: 'シートが見つかりません: ' + cfg.sheetName };
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return {
      ok: true,
      headers: cfg.list,
      rows: [],
      total: 0,
      page,
      pageSize,
    };
  }

  const readRows = Math.min(lastRow, WEBVIEW.MAX_ROWS);

  // 高速化ポイント1:
  // displayValues ではなく values で取得し、必要時のみ文字列化
  const values = sh.getRange(1, 1, readRows, lastCol).getValues();

  const headers = values[0].map(normCell_);
  const headerMap = buildHeaderMap0_(headers);

  const idxMap = buildIndexMap_(headerMap, [
    WEBVIEW.KEY_HEADER,
    WEBVIEW.STATUS_HEADER,
    '備考',
    '顧客名',
    '案件概要',
    '案件名',
    '作業内容',
    '作業形態',
    '現地作業場所',
    '監督者',
    '管理者',
    'サポート',
    'カテゴリタグ',
    '作業カテゴリ',
  ]);

  const statusData = buildStatusRowMapFast_(type);

  // 高速化ポイント2:
  // 検索対象は必要最小限の列だけにして、1行ごとに1回だけ連結
  const searchIndexes = buildSearchIndexes_(headerMap, cfg.list, VIRTUAL_HEADERS);

  const filtered = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    if (hideClosed) {
      const st = getByIndex_(row, idxMap[WEBVIEW.STATUS_HEADER]);
      if (WEBVIEW.CLOSED_WORDS.some(w => st.indexOf(w) !== -1)) continue;
    }

    if (q) {
      const searchText = buildSearchText_(row, searchIndexes);
      if (searchText.indexOf(q) === -1) continue;
    }

    filtered.push(r);
  }

  const total = filtered.length;
  const start = (page - 1) * pageSize;
  const pageRows = filtered.slice(start, start + pageSize);

  const rows = pageRows.map(r => {
    const row = values[r];
    const key = getByIndex_(row, idxMap[WEBVIEW.KEY_HEADER]);
    const statusRow = statusData.rowMap[key] || null;

    const virtual = buildVirtualColumnsFast_(row, idxMap, statusRow, statusData.headerMap, type);

    return cfg.list.map(h => {
      if (virtual[h] !== undefined) return virtual[h];

      const idx = headerMap[h];
      return idx !== undefined ? stringifyCell_(row[idx]) : '';
    });
  });

  return {
    ok: true,
    headers: cfg.list,
    rows,
    total,
    page,
    pageSize,
  };
}

/** ----------------------------------------------------
 *   仮想生成列
 *  ---------------------------------------------------- */
function buildVirtualColumnsFast_(caseRow, idxMap, statusRow, statusHeaderMap, type) {
  return {
    '案件名': buildProjectNameFast_(caseRow, idxMap),
    '担当SE': buildTantoSeFast_(caseRow, idxMap),
    '作業期間': buildWorkPeriodHtmlFast_(caseRow, idxMap, statusRow, statusHeaderMap, type),
    '作業概要': buildWorkSummaryFast_(caseRow, idxMap),
    'カテゴリ': buildCategoryTagsFast_(caseRow, idxMap),
  };
}

/** ----------------------------------------------------
 *   仮想列ロジック（高速化版）
 *  ---------------------------------------------------- */

/** 案件名 */
function buildProjectNameFast_(row, idxMap) {
  const key = getByIndex_(row, idxMap['顧客名']);
  const summary = getByIndex_(row, idxMap['案件概要']) || getByIndex_(row, idxMap['案件名']);

  if (key && summary) return `【${key}】\n${summary}`;
  if (key) return `【${key}】`;
  return summary || '';
}

/** 担当SE */
function buildTantoSeFast_(row, idxMap) {
  const lines = [];

  const sup = getByIndex_(row, idxMap['監督者']);
  const mgr = getByIndex_(row, idxMap['管理者']);
  const sub = getByIndex_(row, idxMap['サポート']);

  if (isValidMemberText_(sup)) lines.push(`監督者：${sup}`);
  if (isValidMemberText_(mgr)) lines.push(`管理者：${mgr}`);
  if (isValidMemberText_(sub)) lines.push(`サポート：${sub}`);

  return lines.join('\n');
}

function isValidMemberText_(v) {
  const s = String(v || '').trim();
  return s && s !== 'ー' && s !== '-';
}

/** 作業期間 */
function buildWorkPeriodHtmlFast_(caseRow, idxMap, statusRow, statusHeaderMap, type) {
  const caseStatus = getByIndex_(caseRow, idxMap['ステータス']);

  const caseBig = '案件確認';
  const inBig = '社内作業';
  const outBig = (String(type || '').toUpperCase() === 'CL') ? '現地作業' : '現地・リモート作業';

  const pCaseStart = getStatusPeriodValue_(statusRow, statusHeaderMap, caseBig, '作業期間：開始');
  const pCaseEnd   = getStatusPeriodValue_(statusRow, statusHeaderMap, caseBig, '作業期間：終了');
  const pInStart   = getStatusPeriodValue_(statusRow, statusHeaderMap, inBig, '作業期間：開始');
  const pInEnd     = getStatusPeriodValue_(statusRow, statusHeaderMap, inBig, '作業期間：終了');
  const pOutStart  = getStatusPeriodValue_(statusRow, statusHeaderMap, outBig, '作業期間：開始');
  const pOutEnd    = getStatusPeriodValue_(statusRow, statusHeaderMap, outBig, '作業期間：終了');

  const info = buildWorkPeriodInfo_(pCaseEnd, caseStatus);

  if (!pCaseStart || !pCaseEnd || pCaseStart > pCaseEnd) {
    return [
      '<div class="wp-wrap">',
        '<div class="wp-row">',
          '<div class="wp-label">案件期間</div>',
          '<div class="wp-lane"><div class="wp-none">未設定</div></div>',
        '</div>',
        '<div class="wp-row">',
          '<div class="wp-label">社内期間</div>',
          '<div class="wp-lane"><div class="wp-none">-</div></div>',
        '</div>',
        '<div class="wp-row">',
          '<div class="wp-label">社外期間</div>',
          '<div class="wp-lane"><div class="wp-none">-</div></div>',
        '</div>',
        '<div class="wp-info ' + (info.cls || '') + '">' + escapeHtml_(info.text || '完了予定未設定') + '</div>',
      '</div>'
    ].join('');
  }

  const caseBar = buildFullBarHtml_(pCaseStart, pCaseEnd, 'case');
  const inBar   = buildGanttBarHtml_(pCaseStart, pCaseEnd, pInStart, pInEnd, 'in');
  const outBar  = buildGanttBarHtml_(pCaseStart, pCaseEnd, pOutStart, pOutEnd, 'out');

  return [
    '<div class="wp-wrap">',
      '<div class="wp-row">',
        '<div class="wp-label">案件期間</div>',
        '<div class="wp-lane wp-lane-case">' + caseBar + '</div>',
      '</div>',
      '<div class="wp-row">',
        '<div class="wp-label">社内期間</div>',
        '<div class="wp-lane">' + inBar + '</div>',
      '</div>',
      '<div class="wp-row">',
        '<div class="wp-label">社外期間</div>',
        '<div class="wp-lane">' + outBar + '</div>',
      '</div>',
      '<div class="wp-info ' + (info.cls || '') + '">' + escapeHtml_(info.text || '') + '</div>',
    '</div>'
  ].join('');
}

/** 作業概要 */
function buildWorkSummaryFast_(row, idxMap) {
  const work  = getByIndex_(row, idxMap['作業内容']);
  const style = getByIndex_(row, idxMap['作業形態']);
  const place = cleanPlaceText_(getByIndex_(row, idxMap['現地作業場所']));

  const styleText = style.replace(/[／/]/g, '・');
  const sub = [styleText, place].filter(Boolean).join(' / ');

  if (work && sub) return `${work}\n${sub}`;
  if (work) return work;
  if (sub) return sub;
  return '';
}

/** カテゴリ */
function buildCategoryTagsFast_(row, idxMap) {
  const raw =
    getByIndex_(row, idxMap['カテゴリタグ']) ||
    getByIndex_(row, idxMap['作業カテゴリ']);

  if (!raw) return [];

  return raw
    .split(/[,\u3001\s]+/)
    .map(v => v.trim())
    .filter(Boolean);
}

/** ----------------------------------------------------
 *   ステータスシート高速取得
 *  ---------------------------------------------------- */
function buildStatusRowMapFast_(type) {
  const ss = getSS_('Display');
  const sh = ss.getSheetByName(type === 'CL' ? WEBVIEW.SHEET_CLST : WEBVIEW.SHEET_SVST);
  if (!sh) return { headerMap: {}, rowMap: {} };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headerMap: {}, rowMap: {} };

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(normCell_);
  const headerMap = buildHeaderMap0_(headers);

  const keyIdx = headerMap['見積番号'];
  const rowMap = {};
  if (keyIdx === undefined) return { headerMap, rowMap };

  for (let r = 1; r < values.length; r++) {
    const key = stringifyCell_(values[r][keyIdx]);
    if (!key) continue;
    rowMap[key] = values[r];
  }

  return { headerMap, rowMap };
}

/** ----------------------------------------------------
 *   ユーティリティ
 *  ---------------------------------------------------- */
function buildHeaderMap0_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[h] = i;
  });
  return map;
}

function buildIndexMap_(headerMap, names) {
  const map = {};
  (names || []).forEach(name => {
    map[name] = headerMap[name];
  });
  return map;
}

function buildSearchIndexes_(headerMap, viewList, virtualHeaders) {
  const exclude = {};
  (virtualHeaders || []).forEach(h => exclude[h] = true);

  const indexes = [];

  Object.keys(headerMap).forEach(h => {
    if (exclude[h]) return;
    indexes.push(headerMap[h]);
  });

  return uniqNums_(indexes);
}

function buildSearchText_(row, indexes) {
  const parts = [];
  for (let i = 0; i < indexes.length; i++) {
    const idx = indexes[i];
    if (idx == null) continue;

    const s = stringifyCell_(row[idx]).toLowerCase();
    if (s) parts.push(s);
  }
  return parts.join('\n');
}

function getByIndex_(row, idx) {
  if (idx === undefined || idx === null || idx < 0) return '';
  return stringifyCell_(row[idx]);
}

function stringifyCell_(v) {
  if (v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(v).trim();
}

function normCell_(v) {
  return String(v == null ? '' : v).trim();
}

function cleanPlaceText_(s) {
  const v = String(s || '').trim();
  if (!v) return '';

  const m = v.match(/^\[([^\]]*)\](.*)$/);
  if (!m) return v;

  const pref = m[1].split('/').join(' ');
  const free = String(m[2] || '').trim();

  return `${pref} ${free}`.trim();
}

function buildGanttBarHtml_(baseStart, baseEnd, start, end, kind) {
  if (!start || !end || start > end) {
    return '<div class="wp-none">-</div>';
  }

  const s = new Date(Math.max(baseStart.getTime(), start.getTime()));
  const e = new Date(Math.min(baseEnd.getTime(), end.getTime()));
  if (s > e) return '<div class="wp-none">-</div>';

  let total = diffDaysInclusive_(baseStart, baseEnd);
  if (total < 1) total = 1;

  const offset = diffDaysInclusive_(baseStart, s) - 1;
  const span   = diffDaysInclusive_(s, e);

  let left = (offset / total) * 100;
  let width = (span / total) * 100;

  const label = buildPeriodLabel_(s, e);

  const minWidth = 10;
  const isTiny = width < 22;
  if (width < minWidth) width = minWidth;

  if (left + width > 100) {
    width = Math.max(minWidth, 100 - left);
  }

  return [
    '<div class="wp-bar wp-bar-' + kind + (isTiny ? ' is-tiny' : '') + '" style="left:' + left + '%;width:' + width + '%;">',
      '<span class="wp-bar-text">' + escapeHtml_(label) + '</span>',
    '</div>'
  ].join('');
}

function buildFullBarHtml_(start, end, kind) {
  if (!start || !end || start > end) {
    return '<div class="wp-none">-</div>';
  }

  const label = buildPeriodLabel_(start, end);

  return [
    '<div class="wp-bar wp-bar-' + kind + '" style="left:0%;width:100%;">',
      '<span class="wp-bar-text">' + escapeHtml_(label) + '</span>',
    '</div>'
  ].join('');
}

function buildWorkPeriodInfo_(caseEnd, caseStatusText) {
  const endDate = parseDateSafe_(caseEnd);
  if (!endDate) return { text: '完了予定未設定', cls: '' };

  const parsed = parseCaseStatusText_(caseStatusText);
  const isClosed = (parsed.caseSt === 'CLOSE');

  const today = new Date();
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

  const diff = Math.floor((end.getTime() - base.getTime()) / 86400000);

  if (diff > 0) return { text: '⏳ 完了まで後' + diff + '日', cls: 'is-future' };
  if (diff === 0) return { text: '📌 今日完了予定', cls: 'is-today' };

  if (!isClosed) {
    return { text: '⚠ ' + Math.abs(diff) + '日超過', cls: 'is-overdue' };
  }

  return { text: '⌛ ' + Math.abs(diff) + '日経過', cls: '' };
}

function parseDateSafe_(v) {
  if (!v) return null;

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  const s = String(v).trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return null;

  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function diffDaysInclusive_(d1, d2) {
  const a = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
  const b = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  return Math.floor((b.getTime() - a.getTime()) / 86400000) + 1;
}

function shortMdLabel_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate();
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getStatusPeriodValue_(statusRow, statusHeaderMap, big, mid) {
  if (!statusRow || !statusHeaderMap) return null;

  const colName = norm_(big) + '_' + norm_(mid);
  const idx = statusHeaderMap[colName];
  if (idx === undefined) return null;

  const v = statusRow[idx];
  if (!v) return null;

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  return parseDateSafe_(v);
}

function buildPeriodLabel_(start, end) {
  if (!start || !end) return '';
  const s = shortMdLabel_(start);
  const e = shortMdLabel_(end);
  return (s === e) ? s : (s + '～' + e);
}

function parseCaseStatusText_(text) {
  const s = String(text || '').trim();
  const out = {
    caseSt: '',
    inSt: '',
    outSt: '',
    docSt: ''
  };
  if (!s) return out;

  const lines = s.split(/\r?\n/).map(v => String(v || '').trim()).filter(Boolean);

  lines.forEach(line => {
    const m1 = line.match(/^案件ST：(.+)$/);
    if (m1) {
      out.caseSt = String(m1[1] || '').trim();
      return;
    }

    const m2 = line.match(/内：([^　\s]+)[　\s]+外：([^　\s]+)[　\s]+Doc\.：([^　\s]+)$/);
    if (m2) {
      out.inSt = String(m2[1] || '').trim();
      out.outSt = String(m2[2] || '').trim();
      out.docSt = String(m2[3] || '').trim();
    }
  });

  return out;
}

/** 共通にない場合だけ追加 */
function uniqNums_(arr) {
  const seen = {};
  const out = [];
  (arr || []).forEach(n => {
    n = Number(n);
    if (!isFinite(n)) return;
    if (seen[n]) return;
    seen[n] = true;
    out.push(n);
  });
  return out;
}