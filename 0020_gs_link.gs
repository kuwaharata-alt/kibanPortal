/** =========================
 * Link Portal (GAS + HTML)
 * Sheet: links
 * ========================= */

const LINK_CONFIG = {
  SHEET_NAME: 'links',
  HEADER_ROW: 1,
  CACHE_SEC: 60, // アイコンURL等の軽いキャッシュ
};

/** クライアントから呼ぶ：リンク一覧取得 */
function api_listLinks() {

  const cache = CacheService.getScriptCache();
  const cached = cache.get('links_payload');
  if (cached) return JSON.parse(cached);

  const ss = getSS_('Link');
  const sh = ss.getSheetByName(LINK_CONFIG.SHEET_NAME);

  if (!sh) return { ok:false, error:'linksシートが見つかりません' };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const values = sh.getRange(LINK_CONFIG.HEADER_ROW,1,lastRow,lastCol).getValues();

  const headers = values[0].map(h=>String(h||'').trim());
  const rows = values.slice(1);

  const idx = name => headers.indexOf(name);

  const iCategory = idx('category');
  const iName = idx('name');
  const iUrl = idx('url');
  const iIconId = idx('iconId');
  const iDesc = idx('desc');
  const iSort = idx('sort');
  const iEnabled = idx('enabled');

  const items = [];
  const catSet = new Set();

  rows.forEach(r=>{

    const enabled = (iEnabled>=0)?String(r[iEnabled]).toLowerCase():'true';
    if (enabled==='false'||enabled==='0'||enabled==='') return;

    const category = (iCategory>=0)?String(r[iCategory]||'').trim():'';
    const name = (iName>=0)?String(r[iName]||'').trim():'';
    const url = (iUrl>=0)?String(r[iUrl]||'').trim():'';

    if (!name || !url) return;

    const iconId = (iIconId>=0)?String(r[iIconId]||'').trim():'';
    const desc = (iDesc>=0)?String(r[iDesc]||'').trim():'';
    const sort = (iSort>=0)?Number(r[iSort]||9999):9999;

    if (category) catSet.add(category);

    items.push({
      category,
      name,
      url,
      iconUrl: iconId ? driveImageUrl_(iconId) : '',
      desc,
      sort,
    });

  });

  items.sort((a,b)=> (a.sort-b.sort) || a.name.localeCompare(b.name,'ja'));


  /* ======================
     カテゴリソート取得
  ====================== */

  const catSheet = ss.getSheetByName('カテゴリ');

  const catMap = {};

  if (catSheet){

    const vals = catSheet.getRange(2,1,catSheet.getLastRow()-1,2).getValues();

    vals.forEach(r=>{
      const cat = String(r[0]||'').trim();
      const no  = Number(r[1]||9999);
      if (cat) catMap[cat] = no;
    });

  }

  const categories = Array.from(catSet).sort((a,b)=>{

    const sa = catMap[a] ?? 9999;
    const sb = catMap[b] ?? 9999;

    return sa - sb || a.localeCompare(b,'ja');

  });


  const payload = {
    ok:true,
    items,
    categories,
    categorySort: catMap
  };

  cache.put('links_payload', JSON.stringify(payload), LINK_CONFIG.CACHE_SEC);

  return payload;
}

function driveImageUrl_(v) {
  const id = normalizeIconId_(v);
  if (!id) return '';

  // これが一番安定（サイズも指定できる）
  return `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w256`;
}

function normalizeIconId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';

  // すでにIDだけならそのまま
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;

  // file/d/<ID>/view 形式
  let m = s.match(/\/d\/([^/]+)/);
  if (m) return m[1];

  // id=<ID> 形式
  m = s.match(/[?&]id=([^&]+)/);
  if (m) return m[1];

  return s; // 最後の保険（ダメなら空にしてもOK）
}