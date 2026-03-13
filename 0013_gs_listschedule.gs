/** =========================================================
 * 案件全体WBS
 * Displayブックから取得
 * 実列名対応版
 * ======================================================= */

const CASE_WBS_CONFIG = {
  SS_KEY: 'Display',
  SHEET_CASE: {
    SV: 'SV案件管理表',
    CL: 'CL案件管理表',
  },
  SHEET_STATUS: {
    SV: 'SV案件ステータス',
    CL: 'CL案件ステータス',
  },
  HEADER_ROW: 1,
  MAX_DAYS: 62,
  MAX_ROWS: 5000,
};

function api_getCaseWbsRows(type) {
  type = normalizeCaseTypeServer_(type);

  const payload = buildCaseWbsPayload_(type);

  return {
    ok: true,
    type,
    maxDays: CASE_WBS_CONFIG.MAX_DAYS,
    members: payload.members,
    rows: payload.rows,
  };
}

function buildCaseWbsPayload_(type) {
  const ss = getSS_('Display');

  const shCase = ss.getSheetByName(CASE_WBS_CONFIG.SHEET_CASE[type]);
  const shSt   = ss.getSheetByName(CASE_WBS_CONFIG.SHEET_STATUS[type]);

  if (!shCase) throw new Error(`シートが見つかりません: ${CASE_WBS_CONFIG.SHEET_CASE[type]}`);
  if (!shSt)   throw new Error(`シートが見つかりません: ${CASE_WBS_CONFIG.SHEET_STATUS[type]}`);

  const caseValues = shCase.getDataRange().getValues();
  const stValues   = shSt.getDataRange().getValues();

  if (caseValues.length < 2) return { rows: [], members: [] };
  if (stValues.length < 2)   return { rows: [], members: [] };

  const caseHead = caseValues[CASE_WBS_CONFIG.HEADER_ROW - 1];
  const stHead   = stValues[CASE_WBS_CONFIG.HEADER_ROW - 1];

  const caseRows = caseValues.slice(CASE_WBS_CONFIG.HEADER_ROW);
  const stRows   = stValues.slice(CASE_WBS_CONFIG.HEADER_ROW);

  const cmap = headerMapFromRow_(caseHead);
  const smap = headerMapFromRow_(stHead);

  const stByKey = {};
  stRows.forEach(r => {
    const key = nz_(r[cmapIndex_(smap, '見積番号')]);
    if (!key) return;
    stByKey[key] = r;
  });

  const memberSet = new Set();
  const out = [];

  for (let i = 0; i < caseRows.length && out.length < CASE_WBS_CONFIG.MAX_ROWS; i++) {
    const r = caseRows[i];

    const key = nz_(r[cmapIndex_(cmap, '見積番号')]);
    if (!key) continue;

    const rowType = nz_(r[cmapIndex_(cmap, '分類')]).toUpperCase();

    if (type === 'SV') {
      if (rowType && rowType !== 'SV') continue;
    } else {
      if (rowType && rowType !== 'CL' && rowType !== 'PC') continue;
    }

    const openClose = nz_(r[cmapIndex_(cmap, 'open/close')]);
    const st = stByKey[key] || [];

    const customer   = nz_(r[cmapIndex_(cmap, '顧客名')]);
    const se         = nz_(r[cmapIndex_(cmap, '担当プリ')]);  // 実データに合わせる
    const supervisor = nz_(r[cmapIndex_(cmap, '監督者')]);
    const manager    = nz_(r[cmapIndex_(cmap, '管理者')]);
    const support    = nz_(r[cmapIndex_(cmap, 'サポート')]);

    [se, supervisor, manager, support].forEach(v => {
      if (v) splitMemberNames_(v).forEach(name => memberSet.add(name));
    });

    const projectStart = toDateStrJst_(pickFirstByNames_(st, smap, [
      '案件管理_作業期間：開始'
    ]));
    const projectEnd = toDateStrJst_(pickFirstByNames_(st, smap, [
      '案件管理_作業期間：終了'
    ]));

    const internalStart = toDateStrJst_(pickFirstByNames_(st, smap, [
      '社内作業_作業期間：開始'
    ]));
    const internalEnd = toDateStrJst_(pickFirstByNames_(st, smap, [
      '社内作業_作業期間：終了'
    ]));

    const externalStart = toDateStrJst_(pickFirstByNames_(st, smap, [
      '現地・リモート作業_作業期間：開始',
      '現地作業_作業期間：開始'
    ]));
    const externalEnd = toDateStrJst_(pickFirstByNames_(st, smap, [
      '現地・リモート作業_作業期間：終了',
      '現地作業_作業期間：終了'
    ]));

    out.push({
      key,
      customer,
      openClose,
      se,
      supervisor,
      manager,
      support,
      members: uniqueStrings_([
        ...splitMemberNames_(se),
        ...splitMemberNames_(supervisor),
        ...splitMemberNames_(manager),
        ...splitMemberNames_(support),
      ]),
      lines: [
        {
          category: '案件',
          status: nz_(pickFirstByNames_(st, smap, ['案件ST'])),
          start: projectStart,
          end: projectEnd,
        },
        {
          category: '社内',
          status: nz_(pickFirstByNames_(st, smap, ['社内ST'])),
          start: internalStart,
          end: internalEnd,
        },
        {
          category: '社外',
          status: nz_(pickFirstByNames_(st, smap, ['現地ST'])),
          start: externalStart,
          end: externalEnd,
        }
      ]
    });
  }

  return {
    rows: out,
    members: Array.from(memberSet).sort((a, b) => a.localeCompare(b, 'ja')),
  };
}

/** =========================================================
 * Helpers
 * ======================================================= */
function normalizeCaseTypeServer_(type) {
  return String(type || 'SV').toUpperCase() === 'CL' ? 'CL' : 'SV';
}

function headerMapFromRow_(row) {
  const map = {};
  row.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

function cmapIndex_(map, name) {
  return map[name] != null ? map[name] : -1;
}

function pickFirstByNames_(row, map, names) {
  if (!row || !map) return '';
  for (let i = 0; i < names.length; i++) {
    const idx = map[names[i]];
    if (idx != null) return row[idx];
  }
  return '';
}

function nz_(v) {
  return String(v == null ? '' : v).trim();
}

function toDateStrJst_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  const s = String(v).trim();
  if (!s) return '';

  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d  = ('0' + m[3]).slice(-2);
    return `${y}-${mo}-${d}`;
  }

  return '';
}

function splitMemberNames_(v) {
  const s = nz_(v);
  if (!s || s === 'ー') return [];
  return s
    .split(/[\/,、，\n\r\t・]/)
    .map(x => String(x || '').trim())
    .filter(Boolean)
    .filter(x => x !== 'ー');
}

function uniqueStrings_(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}