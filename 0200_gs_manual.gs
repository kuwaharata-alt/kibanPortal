const MANUAL_APP = {
  SS_KEY: 'Manual',
  SHEET_NAME: 'マニュアル一覧',
  HEADER_ROW: 1,
  HEADERS: {
    title: 'タイトル',
    url: 'URL',
    category: 'カテゴリ',
    order: '表示順',
  },
};

function api_getManualList() {
  try {
    const ss = getSS_('Manual');
    const sh = ss.getSheetByName(MANUAL_APP.SHEET_NAME);
    if (!sh) throw new Error(`シートが見つかりません: ${MANUAL_APP.SHEET_NAME}`);

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok: true, items: [] };

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(v => String(v || '').trim());
    const rows = values.slice(1);

    const map = {};
    headers.forEach((h, i) => map[h] = i);

    const items = rows.map((r, idx) => {
      const title = getVal_(r, map, 'タイトル');
      const url = getVal_(r, map, 'URL');
      const category = getVal_(r, map, 'カテゴリ');
      const order = Number(getVal_(r, map, '表示順')) || 9999;

      if (!title || !url) return null;

      const fileId = extractGoogleSlideId_(url);

      return {
        id: String(idx + 1),
        title,
        url,
        category,
        order,
        previewUrl: fileId
          ? `https://docs.google.com/presentation/d/${fileId}/preview`
          : url,
        embedUrl: fileId
          ? `https://docs.google.com/presentation/d/${fileId}/embed?start=false&loop=false&delayms=3000`
          : url,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.order - b.order || a.title.localeCompare(b.title, 'ja'));

    return { ok: true, items };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err),
      items: [],
    };
  }
}

function getVal_(row, map, name) {
  const idx = map[name];
  return idx == null ? '' : String(row[idx] || '').trim();
}

function extractGoogleSlideId_(url) {
  const m = String(url || '').match(/\/presentation\/d\/([a-zA-Z0-9\-_]+)/);
  return m ? m[1] : '';
}