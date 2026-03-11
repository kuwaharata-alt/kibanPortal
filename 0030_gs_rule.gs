/**
 * Rules Web App (GAS + Spreadsheet)
 * Optional sheet: admins (A列に編集者メール)
 *
 * Sheets:
 * - rules   : ID,Category,Section,Content,Note,Sort,Enabled,UpdateBy,UpdateAt
 * - admins  : (optional) 編集許可メール一覧（A列）
 * - history : Sort,更新内容,更新者,更新日
 */

const RULE_CONFIG = {
  SHEET_RULES: 'rules',
  SHEET_ADMINS: 'admins',
  SHEET_HISTORY: 'history',
  HEADER_ROW: 1,
};

/** ===== Data Access (rules) ===== */
function getRulesSheet_() {
  const sh = getSS_('Rule').getSheetByName(RULE_CONFIG.SHEET_RULES);
  if (!sh) throw new Error(`Sheet not found: ${RULE_CONFIG.SHEET_RULES}`);
  return sh;
}

function ensureHeader_() {
  const sh = getRulesSheet_();
  const header = sh.getRange(RULE_CONFIG.HEADER_ROW, 1, 1, 9).getValues()[0];
  const expected = ['ID', 'Category', 'Section', 'Content', 'Note', 'Sort', 'Enabled', 'UpdateBy', 'UpdateAt'];
  const ok = expected.every((v, i) => String(header[i] || '').trim() === v);
  if (!ok) throw new Error(`Header mismatch. Please set header row to: ${expected.join(', ')}`);
}

function listRules_(filter) {
  ensureHeader_();
  const sh = getRulesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= RULE_CONFIG.HEADER_ROW) return [];

  const values = sh.getRange(RULE_CONFIG.HEADER_ROW + 1, 1, lastRow - RULE_CONFIG.HEADER_ROW, 9).getValues();

  const enabledFilter = (filter.enabled || '').toLowerCase();
  const wantEnabled =
    enabledFilter === 'true' ? true :
      enabledFilter === 'false' ? false :
        null;

  const q = (filter.q || '').toLowerCase();

  return values
    .map(rowToObj_)
    .filter(o => {
      if (!o.ID) return false;
      if (filter.category && o.Category !== filter.category) return false;
      if (filter.section && o.Section !== filter.section) return false;
      if (wantEnabled !== null && o.Enabled !== wantEnabled) return false;

      if (q) {
        const hay = `${o.Category}\n${o.Section}\n${o.Content}\n${o.Note}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    })
    .sort((a, b) => {
      const c1 = (a.Category || '').localeCompare(b.Category || '');
      if (c1) return c1;
      const c2 = (a.Section || '').localeCompare(b.Section || '');
      if (c2) return c2;
      const s1 = Number(a.Sort || 0);
      const s2 = Number(b.Sort || 0);
      if (s1 !== s2) return s1 - s2;
      return String(b.UpdateAt || '').localeCompare(String(a.UpdateAt || ''));
    });
}

function upsertRules_(payload) {
  ensureHeader_();
  const sh = getRulesSheet_();

  const items = Array.isArray(payload) ? payload : [payload];
  if (!items || items.length === 0) return { records: [], warnings: [] };

  const warnings = [];
  const results = [];

  // ID→行番号
  const lastRow = sh.getLastRow();
  const idToRow = new Map();
  if (lastRow > RULE_CONFIG.HEADER_ROW) {
    const ids = sh.getRange(RULE_CONFIG.HEADER_ROW + 1, 1, lastRow - RULE_CONFIG.HEADER_ROW, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0] || '').trim();
      if (id) idToRow.set(id, RULE_CONFIG.HEADER_ROW + 1 + i);
    }
  }

  const user = getActiveEmail_();
  const nowIso = new Date().toISOString();

  items.forEach((raw) => {
    const o = normalizeInput_(raw);

    if (!o.Category) throw new Error('Category is required');
    if (!o.Section) throw new Error('Section is required');
    if (!o.Content) throw new Error('Content is required');

    const exists = !!(o.ID && idToRow.has(o.ID));
    const before = exists ? getRuleRowById_(o.ID) : null;

    let rowNumber;
    let id = o.ID;

    if (exists) {
      rowNumber = idToRow.get(o.ID);
    } else {
      id = makeId_();
      rowNumber = sh.getLastRow() + 1;
      idToRow.set(id, rowNumber);
    }

    // Sort
    let sortToUse = o.Sort;
    const isSortBlank = (raw?.Sort === '' || raw?.Sort === null || raw?.Sort === undefined);

    if (!exists) {
      if (isSortBlank || isNaN(Number(sortToUse))) {
        sortToUse = getNextSort_(o.Category, o.Section);
      }
    } else {
      if (isSortBlank) {
        const currentSort = sh.getRange(rowNumber, 6).getValue();
        sortToUse = Number(currentSort || 0);
      }
    }

    const record = {
      ID: id,
      Category: o.Category,
      Section: o.Section,
      Content: o.Content,
      Note: o.Note,
      Sort: sortToUse,
      Enabled: o.Enabled,
      UpdateBy: user,
      UpdateAt: nowIso,
    };

    sh.getRange(rowNumber, 1, 1, 9).setValues([objToRow_(record)]);
    results.push(record);

    // 履歴
    if (!exists) {
      const h = appendHistoryRow_(record.Sort, '新規', '', '', user, nowIso);
      if (!h.ok) warnings.push(`履歴書き込み失敗: ${h.error}`);
    } else {
      // appendDiffRows_ の中で appendHistoryRow_ が失敗しても落とさないならOK
      appendDiffRows_(before, record, user, nowIso);
    }
  });

  return { records: results, warnings };
}

function setEnabled_(id, enabled) {
  ensureHeader_();
  const sh = getRulesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= RULE_CONFIG.HEADER_ROW) throw new Error('No data');

  const ids = sh.getRange(RULE_CONFIG.HEADER_ROW + 1, 1, lastRow - RULE_CONFIG.HEADER_ROW, 1).getValues();
  let rowNumber = null;
  for (let i = 0; i < ids.length; i++) {
    const cur = String(ids[i][0] || '').trim();
    if (cur === id) {
      rowNumber = RULE_CONFIG.HEADER_ROW + 1 + i;
      break;
    }
  }
  if (!rowNumber) throw new Error(`ID not found: ${id}`);

  const user = getActiveEmail_();
  const nowIso = new Date().toISOString();

  sh.getRange(rowNumber, 7).setValue(enabled);
  sh.getRange(rowNumber, 8).setValue(user);
  sh.getRange(rowNumber, 9).setValue(nowIso);

  // rulesのSort取得（SortNoとして記録）
  const rule = getRuleRowById_(id);
  const sortNo = rule ? Number(rule.Sort || 0) : 0;

  appendHistoryRow_(sortNo, 'Enabled', String(!enabled), String(enabled), user, nowIso);

}

function getMeta_() {
  const rows = listRules_({ enabled: '' });
  const categories = [...new Set(rows.map(r => r.Category).filter(Boolean))].sort();
  const sections = [...new Set(rows.map(r => r.Section).filter(Boolean))].sort();
  return { categories, sections };
}

/** ===== Auth / Permission ===== */

function getActiveEmail_() {
  return Session.getActiveUser().getEmail() || '';
}

function assertCanEdit_() {
  const email = getActiveEmail_();
  if (!email) throw new Error('Cannot detect user email. Deploy as domain/internal or require sign-in.');

  const ss = getSs_();
  const sh = ss.getSheetByName(RULE_CONFIG.SHEET_ADMINS);
  if (!sh) return true;

  const lastRow = sh.getLastRow();
  if (lastRow < 1) return true;

  const values = sh.getRange(1, 1, lastRow, 1).getValues()
    .map(r => String(r[0] || '').trim())
    .filter(Boolean);

  if (values.length === 0) return true;
  if (!values.includes(email)) throw new Error(`Permission denied: ${email}`);
  return true;
}

/** ===== Utilities ===== */

function normalizeInput_(raw) {
  const o = raw || {};
  const sort = (o.Sort === '' || o.Sort === null || o.Sort === undefined) ? 0 : Number(o.Sort);
  return {
    ID: String(o.ID || '').trim(),
    Category: String(o.Category || '').trim(),
    Section: String(o.Section || '').trim(),
    Content: String(o.Content || '').trim(),
    Note: String(o.Note || '').trim(),
    Sort: isNaN(sort) ? 0 : sort,
    Enabled: (o.Enabled === '' || o.Enabled === null || o.Enabled === undefined) ? true : !!o.Enabled,
  };
}

function rowToObj_(r) {
  return {
    ID: String(r[0] || '').trim(),
    Category: String(r[1] || '').trim(),
    Section: String(r[2] || '').trim(),
    Content: String(r[3] || '').trim(),
    Note: String(r[4] || '').trim(),
    Sort: Number(r[5] || 0),
    Enabled: (String(r[6]).toLowerCase() === 'true') || r[6] === true,
    UpdateBy: String(r[7] || '').trim(),
    UpdateAt: String(r[8] || '').trim(),
  };
}

function getRuleRowById_(id) {
  ensureHeader_();
  const sh = getRulesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= RULE_CONFIG.HEADER_ROW) return null;

  const ids = sh.getRange(RULE_CONFIG.HEADER_ROW + 1, 1, lastRow - RULE_CONFIG.HEADER_ROW, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const cur = String(ids[i][0] || '').trim();
    if (cur === String(id || '').trim()) {
      const rowNum = RULE_CONFIG.HEADER_ROW + 1 + i;
      const row = sh.getRange(rowNum, 1, 1, 9).getValues()[0];
      return rowToObj_(row);
    }
  }
  return null;
}

function objToRow_(o) {
  return [o.ID, o.Category, o.Section, o.Content, o.Note, o.Sort, o.Enabled, o.UpdateBy, o.UpdateAt];
}

function makeId_() {
  const u = Utilities.getUuid().replace(/-/g, '');
  return `R_${u.slice(0, 12)}`;
}

function parseJsonBody_(e) {
  const raw = e?.postData?.contents || '';
  if (!raw) return {};
  return JSON.parse(raw);
}

function json_(obj, statusCode) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ===== UI (HtmlService) ===== */
function ui_meta() {
  return { ok: true, data: getMeta_() };
}

function ui_list(params) {
  params = params || {};
  const filter = {
    category: (params.category || '').trim(),
    section: (params.section || '').trim(),
    q: (params.q || '').trim(),
    enabled: (params.enabled || '').trim(),
  };
  return { ok: true, data: listRules_(filter) };
}

function ui_upsert(data) {
  assertCanEdit_();
  const res = upsertRules_(data); // {records,warnings}
  return { ok: true, data: res.records, warnings: res.warnings || [] };
}

function ui_toggle(id, enabled) {
  assertCanEdit_();
  return { ok: true, data: setEnabled_(String(id || '').trim(), !!enabled) };
}

function ui_history_list(params) {
  params = params || {};
  const limit = Number(params.limit || 200);
  return { ok: true, data: listHistory_(limit) };
}

/** ===== Sort auto-numbering ===== */
function getNextSort_(category, section) {
  ensureHeader_();
  const sh = getRulesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= RULE_CONFIG.HEADER_ROW) return 10;

  const values = sh.getRange(RULE_CONFIG.HEADER_ROW + 1, 1, lastRow - RULE_CONFIG.HEADER_ROW, 6).getValues();

  let maxSort = 0;
  for (const r of values) {
    const id = String(r[0] || '').trim();
    if (!id) continue;
    const cat = String(r[1] || '').trim();
    const sec = String(r[2] || '').trim();
    if (cat !== category) continue;
    if (sec !== section) continue;

    const s = Number(r[5] || 0);
    if (!isNaN(s) && s > maxSort) maxSort = s;
  }
  return maxSort + 10;
}

/** ===== History (No, SortNo, 変更箇所, 旧内容, 新内容, 更新者, 更新日) ===== */

function getHistorySheet_() {
  const sh = getSs_().getSheetByName(RULE_CONFIG.SHEET_HISTORY);
  if (!sh) throw new Error(`Sheet not found: ${RULE_CONFIG.SHEET_HISTORY}`);
  return sh;
}

function ensureHistoryHeader_() {
  const sh = getHistorySheet_();
  const header = sh.getRange(1, 1, 1, 7).getValues()[0];
  const expected = ['No', 'SortNo', '変更箇所', '旧内容', '新内容', '更新者', '更新日'];
  const ok = expected.every((v, i) => String(header[i] || '').trim() === v);
  if (!ok) throw new Error(`History header mismatch. Please set header row to: ${expected.join(', ')}`);
}

function getNextHistoryNo_() {
  const sh = getHistorySheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  const lastNo = sh.getRange(lastRow, 1).getValue(); // A:No
  const n = Number(lastNo);
  return (isNaN(n) ? (lastRow - 1) : n) + 1;
}

/**
 * 履歴1行追記
 * @param {number} sortNo rules側のSort
 * @param {string} field 変更箇所（例：内容/備考/カテゴリ…）
 * @param {string} before 旧内容
 * @param {string} after 新内容
 * @param {string} email 更新者
 * @param {string} iso 更新日（ISO）
 */
function appendHistoryRow_(sortNo, field, before, after, email, iso) {
  try {
    ensureHistoryHeader_();
    const sh = getHistorySheet_();
    const no = getNextHistoryNo_();
    sh.appendRow([
      no,
      Number(sortNo || 0),
      String(field || ''),
      String(before ?? ''),
      String(after ?? ''),
      String(email || ''),
      String(iso || ''),
    ]);
    return { ok: true };
  } catch (e) {
    // 履歴失敗はログだけ残して、保存自体は落とさない
    console.error('History append failed:', e);
    return { ok: false, error: String(e?.message || e) };
  }
}

function listHistory_(limit = 200) {
  ensureHistoryHeader_();
  const sh = getHistorySheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const rows = sh.getRange(2, 1, lastRow - 1, 7).getValues();
  const out = rows.map(r => ({
    No: Number(r[0] || 0),
    SortNo: Number(r[1] || 0),
    変更箇所: String(r[2] || '').trim(),
    旧内容: String(r[3] || ''),
    新内容: String(r[4] || ''),
    更新者: String(r[5] || '').trim(),
    更新日: String(r[6] || '').trim(),
  }));

  // 最新順（No降順）
  out.sort((a, b) => (b.No || 0) - (a.No || 0));
  return out.slice(0, Math.max(1, Number(limit) || 200));
}

function appendDiffRows_(before, after, user, nowIso) {
  const sortNo = Number(after.Sort || 0);

  const fields = [
    ['Category', 'カテゴリ'],
    ['Section', 'Section'],
    ['Content', '内容'],
    ['Note', '備考'],
    ['Sort', 'Sort'],
    ['Enabled', 'Enabled'],
  ];

  for (const [key, label] of fields) {
    const b = String(before?.[key] ?? '');
    const a = String(after?.[key] ?? '');
    if (b === a) continue;

    appendHistoryRow_(sortNo, label, b, a, user, nowIso);
  }
}
