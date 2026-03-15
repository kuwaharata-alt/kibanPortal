/** =========================================================
 * 案件全体WBS
 * Displayブックから取得
 * 専用関数名版（衝突回避）
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
  type = caseWbsNormalizeType_(type);
  const payload = caseWbsBuildPayload_(type);

  return {
    ok: true,
    type,
    maxDays: CASE_WBS_CONFIG.MAX_DAYS,
    members: payload.members,
    rows: payload.rows,
  };
}

function caseWbsBuildPayload_(type) {
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

  const cmap = caseWbsHeaderMap_(caseHead);
  const smap = caseWbsHeaderMap_(stHead);

  const caseIdx = {
    key:        caseWbsRequireCol_(cmap, '見積番号'),
    customer:   caseWbsRequireCol_(cmap, '顧客名'),
    category:   caseWbsRequireCol_(cmap, '分類'),
    summary:    caseWbsRequireCol_(cmap, '案件概要'),
    estimate:   caseWbsRequireCol_(cmap, '見積'),
    supervisor: caseWbsRequireCol_(cmap, '監督者'),
    manager:    caseWbsRequireCol_(cmap, '管理者'),
    support:    caseWbsRequireCol_(cmap, 'サポート'),
    openClose:  caseWbsRequireCol_(cmap, 'open/close'),
  };

  const stKeyIdx = caseWbsRequireCol_(smap, '見積番号');

  Logger.log('CASE SHEET=' + CASE_WBS_CONFIG.SHEET_CASE[type]);
  Logger.log('CASE HEAD=' + JSON.stringify(caseHead));
  Logger.log('CASE ROW1=' + JSON.stringify(caseRows[0]));
  Logger.log('CASE IDX=' + JSON.stringify(caseIdx));

  const stByKey = {};
  stRows.forEach(r => {
    const key = caseWbsNz_(r[stKeyIdx]);
    if (!key) return;
    stByKey[key] = r;
  });

  const memberSet = new Set();
  const out = [];

  for (let i = 0; i < caseRows.length && out.length < CASE_WBS_CONFIG.MAX_ROWS; i++) {
    const r = caseRows[i];

    const key = caseWbsNz_(r[caseIdx.key]);
    if (!key) continue;

    const rowType = caseWbsNz_(r[caseIdx.category]).toUpperCase();

    if (type === 'SV') {
      if (rowType && rowType !== 'SV') continue;
    } else {
      if (rowType && rowType !== 'CL' && rowType !== 'PC') continue;
    }

    const customer   = caseWbsNz_(r[caseIdx.customer]);
    const summary    = caseWbsNz_(r[caseIdx.summary]);
    const estimate   = caseWbsNz_(r[caseIdx.estimate]);
    const openClose  = caseWbsNz_(r[caseIdx.openClose]);
    const supervisor = caseWbsNz_(r[caseIdx.supervisor]);
    const manager    = caseWbsNz_(r[caseIdx.manager]);
    const support    = caseWbsNz_(r[caseIdx.support]);

    const st = stByKey[key] || [];

    [supervisor, manager, support].forEach(v => {
      if (v) caseWbsSplitMembers_(v).forEach(name => memberSet.add(name));
    });

    out.push({
      key,
      customer,
      summary,
      estimate,
      openClose,
      supervisor,
      manager,
      support,
      members: caseWbsUnique_([
        ...caseWbsSplitMembers_(supervisor),
        ...caseWbsSplitMembers_(manager),
        ...caseWbsSplitMembers_(support),
      ]),
      lines: [
        {
          category: '案件',
          status: caseWbsNz_(caseWbsPickFirst_(st, smap, ['案件ST'])),
          start: caseWbsToDateStr_(caseWbsPickFirst_(st, smap, ['案件管理_作業期間：開始'])),
          end:   caseWbsToDateStr_(caseWbsPickFirst_(st, smap, ['案件管理_作業期間：終了'])),
        },
        {
          category: '社内',
          status: caseWbsNz_(caseWbsPickFirst_(st, smap, ['社内ST'])),
          start: caseWbsToDateStr_(caseWbsPickFirst_(st, smap, ['社内作業_作業期間：開始'])),
          end:   caseWbsToDateStr_(caseWbsPickFirst_(st, smap, ['社内作業_作業期間：終了'])),
        },
        {
          category: '社外',
          status: caseWbsNz_(caseWbsPickFirst_(st, smap, ['現地ST'])),
          start: caseWbsToDateStr_(caseWbsPickFirst_(st, smap, [
            '現地・リモート作業_作業期間：開始',
            '現地作業_作業期間：開始'
          ])),
          end: caseWbsToDateStr_(caseWbsPickFirst_(st, smap, [
            '現地・リモート作業_作業期間：終了',
            '現地作業_作業期間：終了'
          ])),
        }
      ]
    });
  }

  Logger.log('CASE WBS SAMPLE=' + JSON.stringify(out.slice(0, 3), null, 2));

  return {
    rows: out,
    members: Array.from(memberSet).sort((a, b) => a.localeCompare(b, 'ja')),
  };
}

/** =========================================================
 * WBS Helpers
 * ======================================================= */
function caseWbsNormalizeType_(type) {
  return String(type || 'SV').toUpperCase() === 'CL' ? 'CL' : 'SV';
}

function caseWbsHeaderMap_(row) {
  const map = {};
  row.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

function caseWbsRequireCol_(map, name) {
  if (map[name] == null) {
    throw new Error(`列が見つかりません: ${name}`);
  }
  return map[name];
}

function caseWbsPickFirst_(row, map, names) {
  if (!row || !map) return '';
  for (let i = 0; i < names.length; i++) {
    const idx = map[names[i]];
    if (idx != null) return row[idx];
  }
  return '';
}

function caseWbsNz_(v) {
  return String(v == null ? '' : v).trim();
}

function caseWbsToDateStr_(v) {
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
    const d = ('0' + m[3]).slice(-2);
    return `${y}-${mo}-${d}`;
  }
  return '';
}

function caseWbsSplitMembers_(v) {
  const s = caseWbsNz_(v);
  if (!s || s === 'ー') return [];
  return s
    .split(/[\/,、，\n\r\t・]/)
    .map(x => String(x || '').trim())
    .filter(Boolean)
    .filter(x => x !== 'ー');
}

function caseWbsUnique_(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}