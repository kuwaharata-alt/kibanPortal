/* =========================================
 * Task API
 * ========================================= */
const TASK_CONFIG = {
  SHEET_NAME: 'タスク',
  HEADER_ROW: 1,
  MEMBER_START_COL: 9, // I列からメンバー列
  TZ: 'Asia/Tokyo',
};

/** シート取得 */
function getTaskSheet_() {
  const ss = getSS_('Task');
  const sh = ss.getSheetByName(TASK_CONFIG.SHEET_NAME);
  if (!sh) throw new Error(`シート「${TASK_CONFIG.SHEET_NAME}」が見つかりません`);
  return sh;
}

/** ヘッダ配列取得 */
function getTaskHeaders_(sheet) {
  return sheet.getRange(
    TASK_CONFIG.HEADER_ROW,
    1,
    1,
    sheet.getLastColumn()
  ).getValues()[0].map(v => String(v || '').trim());
}

/** ヘッダマップ取得（1-based） */
function getTaskHeaderMap_(sheet) {
  const headers = getTaskHeaders_(sheet);
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i + 1;
  });
  return map;
}

/** データ全取得 */
function getTaskSheetValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/** メンバー列名一覧 */
function getTaskMemberHeaders_(headers) {
  return headers
    .slice(TASK_CONFIG.MEMBER_START_COL - 1)
    .map(v => String(v || '').trim())
    .filter(Boolean);
}

/** 日付文字列化 */
function formatTaskDate_(value) {
  if (!value) return '';

  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, TASK_CONFIG.TZ, 'yyyy/MM/dd');
  } catch (e) {
    return '';
  }
}

/** 日時文字列化 */
function formatTaskDateTimeNow_() {
  return Utilities.formatDate(new Date(), TASK_CONFIG.TZ, 'yyyy/MM/dd HH:mm:ss');
}

/** Noから行番号取得 */
function findTaskRowByNo_(values, noColIndex1Based, no) {
  for (let r = 2; r <= values.length; r++) {
    if (String(values[r - 1][noColIndex1Based - 1]) === String(no)) {
      return r;
    }
  }
  return 0;
}

/** 行データから進捗集計 */
function calcTaskProgressFromRow_(rowVals, map, memberHeaders) {
  let doneCount = 0;
  let targetCount = 0;

  memberHeaders.forEach(name => {
    const col = map[name];
    if (!col) return;

    targetCount++;
    if (rowVals[col - 1]) doneCount++;
  });

  const progress = targetCount > 0 ? doneCount / targetCount : 0;

  let status = '未着手';
  if (progress >= 1) status = '完了';
  else if (progress > 0) status = '対応中';

  return {
    doneCount,
    targetCount,
    progress,
    status,
    progressText: `${doneCount}/${targetCount}`,
  };
}

/** タスク一覧取得
 * - 完了以外のみ表示
 * - 進捗率はメンバー列の入力状況から計算
 */
function api_getTask() {
  Logger.log('api_getTask start');

  try {
    const sh = getTaskSheet_();
    const values = getTaskSheetValues_(sh);

    if (!values.length) {
      return {
        ok: true,
        data: [],
        members: [],
        loginUser: String(getActiveUserNameOrEmail_() || ''),
      };
    }

    const headers = values[0].map(v => String(v || '').trim());
    const map = {};
    headers.forEach((h, i) => { map[h] = i; }); // 0-based

    const memberHeaders = getTaskMemberHeaders_(headers);
    const loginUser = String(getActiveUserNameOrEmail_() || '');

    const data = values
      .slice(1)
      .filter(r => r[map['タスク']])
      .filter(r => String(r[map['ステータス']] || '') !== '完了')
      .map(r => {
        const members = {};
        let doneCount = 0;
        let targetCount = 0;

        memberHeaders.forEach(name => {
          const idx = map[name];
          const raw = idx != null ? r[idx] : '';
          let v = '';

          if (raw instanceof Date) {
            v = Utilities.formatDate(raw, TASK_CONFIG.TZ, 'yyyy/MM/dd HH:mm:ss');
          } else {
            v = String(raw || '');
          }

          members[String(name)] = v;
          targetCount++;
          if (v) doneCount++;
        });

        const progress = targetCount > 0 ? doneCount / targetCount : 0;

        return {
          no: String(r[map['No']] || ''),
          task: String(r[map['タスク']] || ''),
          link: String(r[map['リンク']] || ''),
          date: formatTaskDate_(r[map['期日']]),
          status: String(r[map['ステータス']] || ''),
          progress: Number(progress || 0),
          doneCount: Number(doneCount || 0),
          targetCount: Number(targetCount || 0),
          members: members,
        };
      });

    const result = {
      ok: true,
      data: data,
      members: memberHeaders.map(v => String(v || '')),
      loginUser: loginUser,
    };

    Logger.log('api_getTask result rows=' + result.data.length);
    return result;

  } catch (err) {
    Logger.log('api_getTask error: ' + err.message);
    return {
      ok: false,
      error: String(err.message || err),
    };
  }
}

/** メンバーチェック更新
 * - ON: 日時入力
 * - OFF: 空欄
 * - 進捗率 / 進捗 / ステータス 再計算
 */
function api_updateTaskMemberCheck(no, memberName, checked) {
  try {
    const sh = getTaskSheet_();
    const values = getTaskSheetValues_(sh);
    if (!values.length) return { ok: false, error: 'データがありません' };

    const headers = values[0].map(v => String(v || '').trim());
    const map = {};
    headers.forEach((h, i) => { map[h] = i + 1; }); // 1-based

    const memberHeaders = getTaskMemberHeaders_(headers);

    const noCol = map['No'];
    const statusCol = map['ステータス'];
    const progressCol = map['進捗率'];
    const progressTextCol = map['進捗'];
    const memberCol = map[String(memberName || '').trim()];

    if (!memberCol) {
      return { ok: false, error: `メンバー列が見つかりません: ${memberName}` };
    }

    const targetRow = findTaskRowByNo_(values, noCol, no);
    if (!targetRow) {
      return { ok: false, error: '対象行が見つかりません' };
    }

    const newValue = checked ? formatTaskDateTimeNow_() : '';
    sh.getRange(targetRow, memberCol).setValue(newValue);

    const rowVals = sh.getRange(targetRow, 1, 1, sh.getLastColumn()).getValues()[0];
    const summary = calcTaskProgressFromRow_(rowVals, map, memberHeaders);

    if (progressTextCol) {
      sh.getRange(targetRow, progressTextCol).setValue(summary.progressText);
    }
    if (statusCol) {
      sh.getRange(targetRow, statusCol).setValue(summary.status);
    }
    if (progressCol) {
      sh.getRange(targetRow, progressCol).setValue(summary.progress);
      sh.getRange(targetRow, progressCol).setNumberFormat('0%');
    }

    return {
      ok: true,
      value: newValue,
      status: summary.status,
      progress: summary.progress,
      doneCount: summary.doneCount,
      targetCount: summary.targetCount,
    };
  } catch (err) {
    return {
      ok: false,
      error: err.message,
    };
  }
}
