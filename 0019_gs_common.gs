/** ======================================================================
 * Common Utilities
 * ====================================================================== */

const COMMON = {
  TZ: 'Asia/Tokyo',

  DISPLAY_SHEETS_BY_TYPE: {
    SV: 'SV案件管理表',
    CL: 'CL案件管理表',
    DEFAULT: 'SV案件管理表',
  },

  STATUS_SHEETS_BY_TYPE: {
    SV: 'SV案件ステータス',
    CL: 'CL案件ステータス',
    DEFAULT: 'SV案件ステータス',
  },

  KEY_HEADER: '見積番号',
  HEADER_ROW: 1,
};

function norm_(v) {
  return String(v == null ? '' : v).trim();
}

function isBlank_(v) {
  if (v == null) return true;
  if (Array.isArray(v)) return v.length === 0;
  return norm_(v) === '';
}

function nowYmd_(withTime) {
  const fmt = withTime ? 'yyyy/MM/dd HH:mm:ss' : 'yyyy/MM/dd';
  return Utilities.formatDate(new Date(), COMMON.TZ, fmt);
}

function normalizeCaseType_(type) {
  const t = norm_(type || 'SV').toUpperCase();
  if (t === 'PC') return 'CL';
  return t || 'SV';
}

function getSheetByType_(ss, type, map) {
  type = normalizeCaseType_(type);
  const name = map[type] || map.DEFAULT;
  return ss.getSheetByName(name);
}

function getDisplayCaseSheet_(type) {
  return getSheetByType_(getSS_('Display'), type, COMMON.DISPLAY_SHEETS_BY_TYPE);
}

function getDisplayStatusSheet_(type) {
  return getSheetByType_(getSS_('Display'), type, COMMON.STATUS_SHEETS_BY_TYPE);
}

function getMasterSheet_(sheetName) {
  return getSS_('Master').getSheetByName(sheetName);
}

function buildHeaderMap_(sh, headerRow) {
  if (!sh) throw new Error('buildHeaderMap_: sheet is null');
  headerRow = Number(headerRow || 1);

  const lastCol = sh.getLastColumn();
  if (lastCol <= 0) return {};

  const heads = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = {};
  heads.forEach((h, i) => {
    const key = norm_(h);
    if (key) map[key] = i + 1;
  });
  return map;
}

function findRowByKey_(sh, keyCol, key, startRow) {
  if (!sh) throw new Error('findRowByKey_: sheet is null');

  keyCol = Number(keyCol);
  startRow = Number(startRow || 2);

  if (!keyCol) throw new Error('findRowByKey_: keyCol is invalid');

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return 0;

  const target = norm_(key);
  const vals = sh.getRange(startRow, keyCol, lastRow - startRow + 1, 1).getDisplayValues();

  for (let i = 0; i < vals.length; i++) {
    if (norm_(vals[i][0]) === target) return startRow + i;
  }
  return 0;
}

function readRowAsObject_(sh, rowNo, headerRow, useDisplayValues) {
  if (!sh) throw new Error('readRowAsObject_: sheet is null');
  if (!rowNo) return null;

  headerRow = Number(headerRow || 1);

  const lastCol = sh.getLastColumn();
  if (lastCol <= 0) return null;

  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(norm_);
  const values = useDisplayValues
    ? sh.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0]
    : sh.getRange(rowNo, 1, 1, lastCol).getValues()[0];

  const obj = {};
  headers.forEach((h, i) => {
    if (h) obj[h] = values[i];
  });
  return obj;
}

function normalizeForClient_(obj) {
  const out = {};
  Object.entries(obj || {}).forEach(([k, v]) => {
    if (v instanceof Date) out[k] = Utilities.formatDate(v, COMMON.TZ, 'yyyy/MM/dd');
    else if (v === undefined) out[k] = '';
    else out[k] = v;
  });
  return out;
}

function upsertByMapping_(sh, opt) {
  if (!sh) throw new Error('upsertByMapping_: sheet is null');

  const headerRow = Number(opt.headerRow || 1);
  const keyHeader = norm_(opt.keyHeader || COMMON.KEY_HEADER);
  const key = norm_(opt.key);
  const mapping = opt.mapping || {};
  const values = opt.values || {};
  const always = opt.always || {};
  const transform = opt.transform;
  const insertIfNotFound = (opt.insertIfNotFound !== false);

  if (!key) throw new Error('upsertByMapping_: key is empty');

  const hm = buildHeaderMap_(sh, headerRow);
  const keyCol = hm[keyHeader];
  if (!keyCol) throw new Error('upsertByMapping_: key header not found: ' + keyHeader);

  const foundRow = findRowByKey_(sh, keyCol, key, headerRow + 1);
  const targetRow = foundRow || (insertIfNotFound ? sh.getLastRow() + 1 : 0);
  if (!targetRow) return { ok: false, error: 'row not found and insert disabled' };

  const missingHeaders = [];
  const updates = {};

  Object.keys(mapping).forEach(inputKey => {
    const headerName = mapping[inputKey];
    const col = hm[headerName];
    if (!col) {
      missingHeaders.push(headerName);
      return;
    }

    let v = values[inputKey];
    if (typeof transform === 'function') v = transform(inputKey, v);
    if (isBlank_(v)) return;

    updates[col] = v;
  });

  Object.keys(always).forEach(headerName => {
    const col = hm[headerName];
    if (!col) {
      missingHeaders.push(headerName);
      return;
    }
    updates[col] = always[headerName];
  });

  const cols = Object.keys(updates).map(Number).sort((a, b) => a - b);
  if (cols.length) {
    const minCol = cols[0];
    const maxCol = cols[cols.length - 1];
    const width = maxCol - minCol + 1;

    const rowValues = foundRow
      ? sh.getRange(targetRow, minCol, 1, width).getValues()[0]
      : new Array(width).fill('');

    cols.forEach(c => {
      rowValues[c - minCol] = updates[c];
    });

    sh.getRange(targetRow, minCol, 1, width).setValues([rowValues]);
  }

  if (!foundRow) sh.getRange(targetRow, keyCol).setValue(key);

  return { ok: true, row: targetRow, inserted: !foundRow, missingHeaders };
}

function listUniqueFromColumn_(sh, col, opt) {
  opt = opt || {};
  const startRow = Number(opt.startRow || 1);
  const skipWords = opt.skipWords || [];

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];

  const vals = sh.getRange(startRow, col, lastRow - startRow + 1, 1).getValues().flat();
  const seen = new Set();
  const out = [];

  vals.forEach(v => {
    const s = norm_(v);
    if (!s) return;
    if (skipWords.some(w => s === w)) return;
    if (seen.has(s)) return;
    seen.add(s);
    out.push(s);
  });

  return out;
}

function groupByTwoCols_(rows) {
  const map = new Map();

  (rows || []).forEach(r => {
    const area = norm_(r[0]);
    const pref = norm_(r[1]);
    if (!area || !pref) return;

    if (!map.has(area)) map.set(area, []);
    map.get(area).push(pref);
  });

  const groups = [];
  for (const [area, prefs] of map.entries()) {
    const seen = new Set();
    const uniq = [];
    prefs.forEach(p => {
      if (seen.has(p)) return;
      seen.add(p);
      uniq.push(p);
    });
    groups.push({ area, prefs: uniq });
  }

  return groups;
}

function joinMultiValue_(value, sep) {
  if (Array.isArray(value)) {
    const arr = value.map(x => norm_(x)).filter(Boolean);
    return arr.join(sep || '、');
  }
  return norm_(value);
}

function splitMultiValueSafe_(value) {
  if (Array.isArray(value)) return value;

  return String(value ?? '')
    .split(/[\n,、\/／]/)
    .map(s => s.trim())
    .filter(Boolean);
}

function normalizeDateForSheet_(value) {
  const v = norm_(value);
  if (!v) return '';

  const m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${m[1]}/${m[2]}/${m[3]}`;

  return v;
}

function ymdSlashToDashSafe_(value) {
  const v = String(value ?? '').trim();
  const m = v.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return v;

  const yyyy = m[1];
  const mm = ('0' + Number(m[2])).slice(-2);
  const dd = ('0' + Number(m[3])).slice(-2);
  return `${yyyy}-${mm}-${dd}`;
}

function toYmdSlash_(v, tz) {
  if (v == null || v === '') return '';

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz || COMMON.TZ, 'yyyy/MM/dd');
  }

  const s = String(v).trim();
  if (!s) return '';

  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return s;

  const mm = ('0' + Number(m[2])).slice(-2);
  const dd = ('0' + Number(m[3])).slice(-2);
  return `${m[1]}/${mm}/${dd}`;
}

function toInputDateValue_(v) {
  if (v == null || v === '') return '';

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, COMMON.TZ, 'yyyy-MM-dd');
  }

  const s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s.replace(/\//g, '-');

  const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!m) return s;

  return [m[1], ('0' + m[2]).slice(-2), ('0' + m[3]).slice(-2)].join('-');
}

function openCloseFromCode_(code) {
  return norm_(code) === '9999' ? 'close' : 'open';
}

function uniqNums_(arr) {
  const seen = {};
  const out = [];
  (arr || []).forEach(n => {
    n = Number(n);
    if (!n || !isFinite(n)) return;
    if (seen[n]) return;
    seen[n] = true;
    out.push(n);
  });
  return out;
}

function writeSameRowByCols_(sh, rowNo, cols, valueResolver) {
  cols = uniqNums_(cols).sort((a, b) => a - b);
  if (!cols.length) return;

  const minCol = cols[0];
  const maxCol = cols[cols.length - 1];
  const width = maxCol - minCol + 1;
  const cur = sh.getRange(rowNo, minCol, 1, width).getValues()[0];

  cols.forEach(col => {
    cur[col - minCol] = valueResolver(col);
  });

  sh.getRange(rowNo, minCol, 1, width).setValues([cur]);
}