function api_updateKado(payload) {
  try {
    const sheetName = '稼働状況';
    const histName = '更新履歴';
    const tz = 'Asia/Tokyo';

    const row = Number(payload?.row);
    const level = String(payload?.level ?? '');
    const status = String(payload?.status ?? '');

    if (!row || row < 2) return { ok: false, error: 'row が不正です' };

    const ss = getSS_('Operation');
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, error: `${sheetName} シートが見つかりません` };

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(s => String(s || '').trim());

    // 必須ヘッダー
    const idxCat = headers.indexOf('カテゴリ') + 1;
    const idxLevel = headers.indexOf('レベル') + 1;
    const idxStat = headers.indexOf('稼働状況') + 1;
    const idxUpd = headers.indexOf('更新日') + 1;

    if (!idxCat || !idxLevel || !idxStat || !idxUpd) {
      return { ok: false, error: 'ヘッダーが不足しています（カテゴリ/レベル/稼働状況/更新日）' };
    }

    const now = new Date();
    const nowStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm:ss');

    // カテゴリ（履歴用）
    const category = sh.getRange(row, idxCat).getDisplayValue();

    // 上書き（レベル・稼働状況・更新日）
    sh.getRange(row, idxLevel).setValue(level);
    sh.getRange(row, idxStat).setValue(status);
    sh.getRange(row, idxUpd).setValue(nowStr);

    // 更新履歴に追記
    const hist = ss.getSheetByName(histName);
    if (!hist) return { ok: false, error: `${histName} シートが見つかりません` };

    hist.appendRow([category, level, status, nowStr]);

    return { ok: true, row, category, level, status, updatedAt: nowStr };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function api_getOperationSheet(params) {
  try {
    params = params || {};
    const sheetName = String(params.sheetName || '').trim();
    if (!sheetName) return { ok: false, error: 'sheetName が空です' };

    const ss = getSS_('Operation');
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, error: `シートが見つかりません: ${sheetName}` };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      return { ok: true, headers: [], rows: [] };
    }

    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

    if (lastRow === 1) {
      return { ok: true, headers, rows: [] };
    }

    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    const rows = values
      .map((v, i) => ({ r: i + 2, v }))
      .filter(obj => obj.v.some(c => String(c || '').trim() !== ''));

    return { ok: true, headers, rows };

  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}