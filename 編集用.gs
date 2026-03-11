function api_getEdit(req) {
  try {
    const r = req || {};
    const type = String(r.type || 'SV').toUpperCase();
    const key = String(r.key || '').trim();

    const cfg = VIEW_CONFIG[type] || VIEW_CONFIG.SV;
    if (!key) return { ok:false, error:'key（見積番号）が空です' };

    const ss = SpreadsheetApp.openById(WEBVIEW.VIEW_SSID);
    const sh = ss.getSheetByName(cfg.sheetName);
    if (!sh) return { ok:false, error:`シートが見つかりません: ${cfg.sheetName}` };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok:false, error:'データがありません' };

    const values = sh.getRange(1, 1, Math.min(lastRow, WEBVIEW.MAX_ROWS || 5000), lastCol).getDisplayValues();
    const headersAll = values[0].map(v => String(v ?? '').trim());
    const headerMap = {};
    headersAll.forEach((h, i) => { if (h) headerMap[h] = i; }); // 0-based

    const idxKey = headerMap[cfg.keyHeader];
    if (idxKey === undefined) return { ok:false, error:`ヘッダー「${cfg.keyHeader}」が見つかりません` };

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idxKey] ?? '').trim() === key) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { ok:false, error:`キーが見つかりません: ${key}` };

    const row = values[rowIndex];

    // editableDef の値を返す
    const defs = cfg.editableDef || [];
    const fields = defs.map(d => {
      const idx = headerMap[d.key];
      return {
        key: d.key,
        label: d.label || d.key,
        type: d.type || 'text',
        options: d.options || [],
        value: (idx === undefined) ? '' : String(row[idx] ?? ''),
        missing: (idx === undefined),
      };
    });

    return {
      ok: true,
      type,
      sheet: cfg.sheetName,
      keyHeader: cfg.keyHeader,
      key,
      fields,
      rowNo: rowIndex + 1, // 1-based（デバッグ用）
    };

  } catch (err) {
    return { ok:false, error:String(err?.message || err), stack:String(err?.stack || '') };
  }
}

function api_updateFields(req) {
  try {
    const r = req || {};
    const type = String(r.type || 'SV').toUpperCase();
    const key = String(r.key || '').trim();
    const updates = r.updates || {}; // { 'ステータス':'...', '管理者':'...' }

    const cfg = VIEW_CONFIG[type] || VIEW_CONFIG.SV;
    if (!key) return { ok:false, error:'key（見積番号）が空です' };

    const ss = SpreadsheetApp.openById(WEBVIEW.VIEW_SSID);
    const sh = ss.getSheetByName(cfg.sheetName);
    if (!sh) return { ok:false, error:`シートが見つかりません: ${cfg.sheetName}` };

    const headerColMap = buildHeaderMap_(sh); // 1-based
    const keyCol = headerColMap[cfg.keyHeader];
    if (!keyCol) return { ok:false, error:`ヘッダー「${cfg.keyHeader}」が見つかりません` };

    // editable 以外は更新させない（守り）
    const editableKeys = new Set((cfg.editableDef || []).map(d => d.key));
    const targets = Object.keys(updates).filter(k => editableKeys.has(k) && headerColMap[k]);

    if (!targets.length) return { ok:false, error:'更新対象がありません（editableDef/ヘッダー存在を確認）' };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:false, error:'データがありません' };

    // キーで行検索
    const keys = sh.getRange(2, keyCol, lastRow - 1, 1).getValues();
    let rowNo = 0;
    for (let i = 0; i < keys.length; i++) {
      const k = String(keys[i][0] ?? '').trim();
      if (k === key) { rowNo = i + 2; break; }
    }
    if (!rowNo) return { ok:false, error:`キーが見つかりません: ${key}` };

    const applied = [];
    targets.forEach(h => {
      sh.getRange(rowNo, headerColMap[h]).setValue(String(updates[h] ?? ''));
      applied.push(h);
    });

    return { ok:true, type, sheet: cfg.sheetName, key, rowNo, applied };

  } catch (err) {
    return { ok:false, error:String(err?.message || err), stack:String(err?.stack || '') };
  }
}

