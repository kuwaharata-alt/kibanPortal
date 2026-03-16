/** ======================================================================
 * 
    ====================================================================== */

function getPlanSheet_(name) {
  const ss = getSS_(PLAN_APP.SS_KEY);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シートが見つかりません: ${name}`);
  return sh;
}

function getHeaderMap_(sheet, headerRow = 1) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h !== '') map[String(h).trim()] = i + 1;
  });
  return map;
}

function normalizeText_(v) {
  return String(v ?? '').trim();
}

function formatDateKey_(d) {
  return Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'yyyy/MM/dd');
}

function getWeekdayJa_(d) {
  const arr = ['日', '月', '火', '水', '木', '金', '土'];
  return arr[new Date(d).getDay()];
}

function firstDayOfMonth_(ym) {
  const [y, m] = String(ym).split('/').map(Number);
  return new Date(y, m - 1, 1);
}

function addMonths_(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

function enumerateDates_(baseYm, months) {
  const from = firstDayOfMonth_(baseYm);
  const to = lastDayOfRange_(baseYm, months);
  const out = [];
  for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
    out.push(new Date(d));
  }
  return out;
}

function firstDayOfYm_(ym) {
  const [y, m] = String(ym || '').split('/').map(Number);
  if (!y || !m) throw new Error('baseYm が不正です: ' + ym);
  return new Date(y, m - 1, 1);
}

function lastDayOfRange_(ym, months) {
  const start = firstDayOfYm_(ym);
  return new Date(start.getFullYear(), start.getMonth() + Number(months), 0);
}

function parseDateFlexible_(v) {
  if (!v) return null;

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  const s = String(v).trim();

  // yyyy/MM/dd
  let m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  // yyyy-MM-dd
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}