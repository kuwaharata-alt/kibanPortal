/**
 * ステータスマスタから、SV/CL別の候補一覧を作り、
 * 表示シート（SV案件管理表 / CL案件管理表）の「ステータス」列へ
 * データ検証（プルダウン）を自動設定する
 *
 * 前提：
 * - ステータスマスタ シートが存在し、ヘッダーに以下がある
 *   対象 / 表示名 / sort
 * - 表示シートに「ステータス」ヘッダーがある
 */
function Apply_StatusValidation_SVCL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== 設定 =====
  const MASTER_SHEET = 'ステータスマスタ';
  const DV_SHEET = '_DV_Status';

  const VIEW_SHEET_SV = 'SV案件管理表';
  const VIEW_SHEET_CL = 'CL案件管理表';

  const HEADER_STATUS = 'ステータス';

  // どこまで検証を貼るか
  // true : シートの最大行まで貼る（行追加しても検証が効く）
  // false: 最終データ行まで貼る（軽い）
  const APPLY_TO_MAX_ROWS = true;

  // ===== シート取得 =====
  const shMaster = ss.getSheetByName(MASTER_SHEET);
  if (!shMaster) throw new Error(`シート「${MASTER_SHEET}」が見つかりません`);

  let shDV = ss.getSheetByName(DV_SHEET);
  if (!shDV) shDV = ss.insertSheet(DV_SHEET);

  const shSV = ss.getSheetByName(VIEW_SHEET_SV);
  const shCL = ss.getSheetByName(VIEW_SHEET_CL);
  if (!shSV) throw new Error(`シート「${VIEW_SHEET_SV}」が見つかりません`);
  if (!shCL) throw new Error(`シート「${VIEW_SHEET_CL}」が見つかりません`);

  // ===== 1) ステータスマスタ読み込み =====
  const mLastRow = shMaster.getLastRow();
  const mLastCol = shMaster.getLastColumn();
  if (mLastRow < 2) throw new Error('ステータスマスタにデータがありません');

  const mValues = shMaster.getRange(1, 1, mLastRow, mLastCol).getValues();
  const mHeaders = mValues[0].map(v => String(v ?? '').trim());
  const mMap = headerMapFromRow_(mHeaders);

  const colType = mustCol_(mMap, '対象');
  const colDisp = mustCol_(mMap, '表示名');
  const colSort = mustCol_(mMap, 'sort');

  const svList = [];
  const clList = [];

  for (let r = 1; r < mValues.length; r++) {
    const row = mValues[r];
    const type = String(row[colType - 1] ?? '').trim().toUpperCase();
    const disp = String(row[colDisp - 1] ?? '').trim();
    const sort = Number(row[colSort - 1]) || 9999;
    if (!disp) continue;

    if (type === 'SV') svList.push({ disp, sort });
    if (type === 'CL') clList.push({ disp, sort });
  }

  // sort順にして表示名だけに
  svList.sort((a,b) => a.sort - b.sort || a.disp.localeCompare(b.disp));
  clList.sort((a,b) => a.sort - b.sort || a.disp.localeCompare(b.disp));

  const svVals = svList.map(x => [x.disp]);
  const clVals = clList.map(x => [x.disp]);

  // ===== 2) _DV_Status を更新（候補一覧）=====
  // A列:SV / B列:CL
  shDV.clearContents();
  shDV.getRange(1, 1, 1, 2).setValues([['SV候補', 'CL候補']]);

  if (svVals.length) shDV.getRange(2, 1, svVals.length, 1).setValues(svVals);
  if (clVals.length) shDV.getRange(2, 2, clVals.length, 1).setValues(clVals);

  // 参照範囲（A2:A / B2:B）
  const svRangeA1 = `'${DV_SHEET}'!A2:A`;
  const clRangeA1 = `'${DV_SHEET}'!B2:B`;

  // ===== 3) 表示シートへデータ検証を貼る =====
  applyValidationToSheet_(shSV, HEADER_STATUS, svRangeA1, APPLY_TO_MAX_ROWS);
  applyValidationToSheet_(shCL, HEADER_STATUS, clRangeA1, APPLY_TO_MAX_ROWS);

  Logger.log(`データ検証設定完了: SV候補=${svVals.length} / CL候補=${clVals.length}`);
}

/**
 * 指定シートの「指定ヘッダー列」にデータ検証を貼る
 */
function applyValidationToSheet_(sh, headerName, rangeA1, applyToMaxRows) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error(`シートが空です: ${sh.getName()}`);

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const map = headerMapFromRow_(headers);
  const col = map[headerName];
  if (!col) throw new Error(`シート「${sh.getName()}」にヘッダー「${headerName}」が見つかりません`);

  const startRow = 2;
  const endRow = applyToMaxRows ? sh.getMaxRows() : sh.getLastRow();
  const rows = Math.max(0, endRow - startRow + 1);
  if (rows <= 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getParent().getRange(rangeA1), true)
    .setAllowInvalid(false) // 手入力で逸脱を禁止（必要ならtrueに）
    .build();

  sh.getRange(startRow, col, rows, 1).setDataValidation(rule);
}

/** ===== helper ===== */

function headerMapFromRow_(row) {
  const m = {};
  for (let i = 0; i < row.length; i++) {
    const h = String(row[i] ?? '').trim();
    if (h) m[h] = i + 1;
  }
  return m;
}

function mustCol_(headerMap, name) {
  const col = headerMap[name];
  if (!col) throw new Error(`ヘッダー「${name}」が見つかりません（1行目を確認）: ${name}`);
  return col;
}
