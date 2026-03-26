const PLAN_APP = {
  SS_KEY: 'Plan',
  SHEET_SUMMARY: '予定管理表',
  SHEET_MASTER: 'マスタ',
  SHEET_SETTINGS: 'Settings',
  SHEET_HOLIDAY: '祝日',
  HEADER_ROW: 1,

  MASTER_HEADERS: {
    status: 'ステータス',
    location: '勤務場所',
    vacation: '休暇',
    member: 'メンバー名',
  },

  SETTINGS_HEADERS: {
    member: '氏名',
    group: 'グループ',
  },

  MEMBER_SHEET_HEADERS: {
    date: '日付',
    week: '曜日',
    amStatus: 'AM_ステータス',
    amLocation: 'AM_勤務場所',
    amVacation: 'AM_休暇',
    amCustomer: 'AM_対応顧客名',
    pmStatus: 'PM_ステータス',
    pmLocation: 'PM_勤務場所',
    pmVacation: 'PM_休暇',
    pmCustomer: 'PM_対応顧客名',
  }
};


/** ======================================================================
 *  マスタ取得
 *  ====================================================================== */

function api_getPlanMaster() {
  const shMaster = getPlanSheet_(PLAN_APP.SHEET_MASTER);
  const hmMaster = getHeaderMap_(shMaster, PLAN_APP.HEADER_ROW);

  const shSettings = getPlanSheet_(PLAN_APP.SHEET_SETTINGS);
  const hmSettings = getHeaderMap_(shSettings, PLAN_APP.HEADER_ROW);

  const masterLastRow = shMaster.getLastRow();
  const settingsLastRow = shSettings.getLastRow();

  const statusSet = new Set();
  const locationSet = new Set();
  const vacationSet = new Set();

  // Settings側：メンバーとグループ
  const settingsMembers = [];
  if (settingsLastRow > PLAN_APP.HEADER_ROW) {
    const vals = shSettings.getRange(
      PLAN_APP.HEADER_ROW + 1,
      1,
      settingsLastRow - PLAN_APP.HEADER_ROW,
      shSettings.getLastColumn()
    ).getValues();

    vals.forEach((r, idx) => {
      const member = normalizeText_(r[hmSettings[PLAN_APP.SETTINGS_HEADERS.member] - 1]);
      const group  = normalizeText_(r[hmSettings[PLAN_APP.SETTINGS_HEADERS.group] - 1]);

      if (member) {
        settingsMembers.push({
          name: member,
          group: group || '',
          order: idx + 1, // Settingsの並び順を表示順として使う
        });
      }
    });
  }

  // マスタ側：ステータス・勤務場所・休暇・メンバー名
  const masterMemberSet = new Set();
  if (masterLastRow > PLAN_APP.HEADER_ROW) {
    const vals = shMaster.getRange(
      PLAN_APP.HEADER_ROW + 1,
      1,
      masterLastRow - PLAN_APP.HEADER_ROW,
      shMaster.getLastColumn()
    ).getValues();

    vals.forEach(r => {
      const status   = normalizeText_(r[hmMaster[PLAN_APP.MASTER_HEADERS.status] - 1]);
      const location = normalizeText_(r[hmMaster[PLAN_APP.MASTER_HEADERS.location] - 1]);
      const vacation = normalizeText_(r[hmMaster[PLAN_APP.MASTER_HEADERS.vacation] - 1]);
      const member   = normalizeText_(r[hmMaster[PLAN_APP.MASTER_HEADERS.member] - 1]);

      if (status) statusSet.add(status);
      if (location) locationSet.add(location);
      if (vacation) vacationSet.add(vacation);
      if (member) masterMemberSet.add(member);
    });
  }

  // Settingsを優先してメンバー一覧を作る
  const memberMap = new Map();
  settingsMembers.forEach(x => {
    memberMap.set(x.name, x);
  });

  // マスタにしかいないメンバーも補完
  Array.from(masterMemberSet).forEach(name => {
    if (!memberMap.has(name)) {
      memberMap.set(name, {
        name,
        group: '',
        order: 9999,
      });
    }
  });

  const uniqMembers = Array.from(memberMap.values())
    .sort((a, b) => a.order - b.order);

  return {
    ok: true,
    members: uniqMembers.map(x => x.name),
    memberGroups: uniqMembers,
    statusList: Array.from(statusSet),
    locationList: Array.from(locationSet),
    vacationList: Array.from(vacationSet),
  };
}

/** ======================================================================
 *   個人シートの月次取得
 *  ====================================================================== */

function api_getMemberMonthPlan(memberName, baseYm, months) {
  const sh = getPlanSheet_(memberName);
  const hm = getHeaderMap_(sh, 1);
  const dates = enumerateDates_(baseYm, months);

  const needHeaders = PLAN_APP.MEMBER_SHEET_HEADERS;
  const lastRow = sh.getLastRow();
  const data = [];

  const mapByDate = {};
  if (lastRow > 1) {
    const vals = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    vals.forEach(r => {
      const key = normalizePlanDateKey_(r[hm[needHeaders.date] - 1]);
      if (!key) return;

      mapByDate[key] = {
        date: key,
        week: normalizeText_(r[hm[needHeaders.week] - 1]),
        amStatus: normalizeText_(r[hm[needHeaders.amStatus] - 1]),
        amLocation: normalizeText_(r[hm[needHeaders.amLocation] - 1]),
        amVacation: normalizeText_(r[hm[needHeaders.amVacation] - 1]),
        amCustomer: normalizeText_(r[hm[needHeaders.amCustomer] - 1]),
        pmStatus: normalizeText_(r[hm[needHeaders.pmStatus] - 1]),
        pmLocation: normalizeText_(r[hm[needHeaders.pmLocation] - 1]),
        pmVacation: normalizeText_(r[hm[needHeaders.pmVacation] - 1]),
        pmCustomer: normalizeText_(r[hm[needHeaders.pmCustomer] - 1]),
      };
    });
  }

  dates.forEach(d => {
    const key = normalizePlanDateKey_(d);
    data.push(
      mapByDate[key] || {
        date: key,
        week: getWeekdayJa_(d),
        amStatus: undefined,
        amLocation: undefined,
        amVacation: undefined,
        amCustomer: undefined,
        pmStatus: undefined,
        pmLocation: undefined,
        pmVacation: undefined,
        pmCustomer: undefined,
      }
    );
  });

  return { ok: true, data };
}
/** ======================================================================
 *   個人シートの保存
 *  ====================================================================== */

function hasOwn_(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

function normalizeWriteValue_(v) {
  if (v === null) return '';
  if (v === undefined) return undefined;
  if (v === '') return undefined; // 空文字は未変更扱い
  return v;
}
function normalizePlanDateKey_(v) {
  if (v === null || v === undefined || v === '') return '';

  // Date型
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  }

  // 文字列・その他
  const s = String(v).trim();
  if (!s) return '';

  const d = parseDateFlexible_(s);
  if (d) {
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  }

  return s;
}

function isMeaningfulPlanRow_(obj) {
  return !!(
    String(obj.amStatus   || '').trim() ||
    String(obj.amLocation || '').trim() ||
    String(obj.amVacation || '').trim() ||
    String(obj.amCustomer || '').trim() ||
    String(obj.pmStatus   || '').trim() ||
    String(obj.pmLocation || '').trim() ||
    String(obj.pmVacation || '').trim() ||
    String(obj.pmCustomer || '').trim()
  );
}

function setCellIfPresent_(sh, row, col, obj, key) {
  if (!col) return;
  if (!hasOwn_(obj, key)) return; // 項目自体が来ていないなら既存値保持

  const v = normalizeWriteValue_(obj[key]);
  if (v === undefined) return;

  sh.getRange(row, col).setValue(v);
}

function getExistingEventIds_(sh, row, hm) {
  return {
    amEventId: String((hm['AM_eventId'] ? sh.getRange(row, hm['AM_eventId']).getValue() : '') || '').trim(),
    pmEventId: String((hm['PM_eventId'] ? sh.getRange(row, hm['PM_eventId']).getValue() : '') || '').trim(),
  };
}

function readPlanRowCurrentValues_(sh, row, hm) {
  return {
    date:       hm['日付']         ? sh.getRange(row, hm['日付']).getDisplayValue() : '',
    week:       hm['曜日']         ? sh.getRange(row, hm['曜日']).getDisplayValue() : '',

    amStatus:   hm['AM_ステータス']   ? sh.getRange(row, hm['AM_ステータス']).getDisplayValue() : '',
    amLocation: hm['AM_勤務場所']     ? sh.getRange(row, hm['AM_勤務場所']).getDisplayValue() : '',
    amVacation: hm['AM_休暇']         ? sh.getRange(row, hm['AM_休暇']).getDisplayValue() : '',
    amCustomer: hm['AM_対応顧客名']   ? sh.getRange(row, hm['AM_対応顧客名']).getDisplayValue() : '',

    pmStatus:   hm['PM_ステータス']   ? sh.getRange(row, hm['PM_ステータス']).getDisplayValue() : '',
    pmLocation: hm['PM_勤務場所']     ? sh.getRange(row, hm['PM_勤務場所']).getDisplayValue() : '',
    pmVacation: hm['PM_休暇']         ? sh.getRange(row, hm['PM_休暇']).getDisplayValue() : '',
    pmCustomer: hm['PM_対応顧客名']   ? sh.getRange(row, hm['PM_対応顧客名']).getDisplayValue() : '',
  };
}

function pickMergedPlanValue_(incoming, key, currentValue) {
  if (!hasOwn_(incoming, key)) return currentValue;

  const v = incoming[key];

  // undefined / '' は未変更扱い
  if (v === undefined || v === '') return currentValue;

  // null のときだけ明示クリア
  if (v === null) return '';

  return String(v).trim();
}

function mergePlanRowForSave_(current, incoming) {
  return {
    date: hasOwn_(incoming, 'date')
      ? normalizePlanDateKey_(incoming.date)
      : normalizePlanDateKey_(current.date),

    week: hasOwn_(incoming, 'week')
      ? (incoming.week === undefined || incoming.week === '' ? (current.week || '') : String(incoming.week).trim())
      : (current.week || ''),

    amStatus:   pickMergedPlanValue_(incoming, 'amStatus',   current.amStatus),
    amLocation: pickMergedPlanValue_(incoming, 'amLocation', current.amLocation),
    amVacation: pickMergedPlanValue_(incoming, 'amVacation', current.amVacation),
    amCustomer: pickMergedPlanValue_(incoming, 'amCustomer', current.amCustomer),

    pmStatus:   pickMergedPlanValue_(incoming, 'pmStatus',   current.pmStatus),
    pmLocation: pickMergedPlanValue_(incoming, 'pmLocation', current.pmLocation),
    pmVacation: pickMergedPlanValue_(incoming, 'pmVacation', current.pmVacation),
    pmCustomer: pickMergedPlanValue_(incoming, 'pmCustomer', current.pmCustomer),
  };
}


/** ======================================================================
 *   個人シートの保存
 *   - 項目未送信なら既存値保持
 *   - カレンダー同期あり
 *  ====================================================================== */

function api_saveMemberMonthPlan(payload) {
  const memberName = String(payload.memberName || '').trim();
  const rows = payload.rows || [];
  if (!memberName) throw new Error('memberName が空です');

  const sh = getPlanSheet_(memberName);
  const hm = getHeaderMap_(sh, 1);

  const required = [
    '日付','曜日','所定休日',
    'AM_ステータス','AM_勤務場所','AM_休暇','AM_対応顧客名',
    'AM_eventId','AM_syncStatus','AM_syncAt',
    'PM_ステータス','PM_勤務場所','PM_休暇','PM_対応顧客名',
    'PM_eventId','PM_syncStatus','PM_syncAt'
  ];

  required.forEach(h => {
    if (!hm[h]) throw new Error('個人シートのヘッダー不足: ' + h);
  });

  const lastRow = sh.getLastRow();
  const existingMap = {};

  if (lastRow > 1) {
    const vals = sh.getRange(2, hm['日付'], lastRow - 1, 1).getValues();
    vals.forEach((r, i) => {
      const key = normalizePlanDateKey_(r[0]);
      if (key) existingMap[key] = i + 2;
    });
  }

  const syncSourceRows = [];

  rows.forEach(obj => {
    if (!isMeaningfulPlanRow_(obj)) return;

    const key = normalizePlanDateKey_(obj.date);
    if (!key) return;

    let row = existingMap[key];
    let isNew = false;

    if (!row) {
      row = sh.getLastRow() + 1;
      existingMap[key] = row;
      isNew = true;
    }

    // 新規行なら最低限の日付・曜日だけは入れる
    if (isNew) {
      sh.getRange(row, hm['日付']).setValue(key);
      sh.getRange(row, hm['曜日']).setValue(obj.week || '');
    } else {
      // 既存行は、送られてきた項目だけ更新
      setCellIfPresent_(sh, row, hm['日付'], obj, 'date');
      setCellIfPresent_(sh, row, hm['曜日'], obj, 'week');
    }

    setCellIfPresent_(sh, row, hm['AM_ステータス'], obj, 'amStatus');
    setCellIfPresent_(sh, row, hm['AM_勤務場所'], obj, 'amLocation');
    setCellIfPresent_(sh, row, hm['AM_休暇'], obj, 'amVacation');
    setCellIfPresent_(sh, row, hm['AM_対応顧客名'], obj, 'amCustomer');

    setCellIfPresent_(sh, row, hm['PM_ステータス'], obj, 'pmStatus');
    setCellIfPresent_(sh, row, hm['PM_勤務場所'], obj, 'pmLocation');
    setCellIfPresent_(sh, row, hm['PM_休暇'], obj, 'pmVacation');
    setCellIfPresent_(sh, row, hm['PM_対応顧客名'], obj, 'pmCustomer');

    const current = readPlanRowCurrentValues_(sh, row, hm);
    const merged = mergePlanRowForSave_(current, obj);
    const eventIds = getExistingEventIds_(sh, row, hm);

    syncSourceRows.push({
      sheetRow: row,
      date: merged.date,
      week: merged.week,

      amStatus: merged.amStatus,
      amLocation: merged.amLocation,
      amVacation: merged.amVacation,
      amCustomer: merged.amCustomer,
      amEventId: eventIds.amEventId,

      pmStatus: merged.pmStatus,
      pmLocation: merged.pmLocation,
      pmVacation: merged.pmVacation,
      pmCustomer: merged.pmCustomer,
      pmEventId: eventIds.pmEventId,
    });
  });

  const syncResults = syncMemberMonthPlanToCalendar_(memberName, syncSourceRows);

  syncResults.forEach(r => {
    const row = r.sheetRow;
    if (!row) return;

    sh.getRange(row, hm['AM_eventId']).setValue(r.am?.eventId || '');
    sh.getRange(row, hm['AM_syncStatus']).setValue(r.am?.syncStatus || '');
    sh.getRange(row, hm['AM_syncAt']).setValue(r.am?.syncAt || '');

    sh.getRange(row, hm['PM_eventId']).setValue(r.pm?.eventId || '');
    sh.getRange(row, hm['PM_syncStatus']).setValue(r.pm?.syncStatus || '');
    sh.getRange(row, hm['PM_syncAt']).setValue(r.pm?.syncAt || '');
  });

  // 予定管理表は、保存後の実値で反映
  const summaryRows = syncSourceRows.map(r => ({
    date: normalizePlanDateKey_(r.date),
    week: r.week || '',
    amStatus: r.amStatus || '',
    amLocation: r.amLocation || '',
    amVacation: r.amVacation || '',
    amCustomer: r.amCustomer || '',
    pmStatus: r.pmStatus || '',
    pmLocation: r.pmLocation || '',
    pmVacation: r.pmVacation || '',
    pmCustomer: r.pmCustomer || '',
  }));

  upsertMemberToSummaryFast_(memberName, summaryRows);

  return {
    ok: true,
    syncResults,
  };
}

function upsertMemberToSummaryFast_(memberName, rows) {
  const sh = getPlanSheet_(PLAN_APP.SHEET_SUMMARY);

  const HEADER_ROW_1 = 1;
  const DATA_START_ROW = 2;

  const lastCol = Math.max(sh.getLastColumn(), 2);
  const headerValues = sh.getRange(HEADER_ROW_1, 1, 1, lastCol).getDisplayValues()[0];

  let memberCol = null;
  for (let c = 3; c <= lastCol; c++) {
    if (String(headerValues[c - 1] || '').trim() === memberName) {
      memberCol = c;
      break;
    }
  }

  if (!memberCol) {
    memberCol = sh.getLastColumn() + 1;
    sh.getRange(HEADER_ROW_1, memberCol).setValue(memberName);
  }

  const lastRow = sh.getLastRow();
  let bodyValues = [];
  const dateMap = {};

  if (lastRow >= DATA_START_ROW) {
    bodyValues = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, Math.max(memberCol, 2)).getValues();
    bodyValues.forEach((r, i) => {
      const key = normalizePlanDateKey_(r[0]);
      if (key) dateMap[key] = i;
    });
  }

  const appendRows = [];

  rows.forEach(r => {
    const dateKey = normalizePlanDateKey_(r.date);
    if (!dateKey) return;

    const mergedText = buildMergedDayText_(r);
    if (!mergedText) return;

    let idx = dateMap[dateKey];
    let rowArr;

    if (idx === undefined) {
      rowArr = new Array(Math.max(memberCol, 2)).fill('');
      rowArr[0] = dateKey;
      rowArr[1] = r.week || '';
      rowArr[memberCol - 1] = mergedText;
      appendRows.push(rowArr);

      idx = bodyValues.length + appendRows.length - 1;
      dateMap[dateKey] = idx;
    } else {
      rowArr = bodyValues[idx];
      while (rowArr.length < memberCol) rowArr.push('');

      rowArr[0] = normalizePlanDateKey_(rowArr[0] || dateKey);
      if (!rowArr[1]) rowArr[1] = r.week || '';
      rowArr[memberCol - 1] = mergedText;
    }
  });

  if (bodyValues.length) {
    const width = Math.max(memberCol, 2);
    const normalized = bodyValues.map(r => {
      const row = r.slice();
      while (row.length < width) row.push('');
      row[0] = normalizePlanDateKey_(row[0]);
      return row;
    });
    sh.getRange(DATA_START_ROW, 1, normalized.length, width).setValues(normalized);
  }

  if (appendRows.length) {
    const width = Math.max(memberCol, 2);
    const normalizedAppend = appendRows.map(r => {
      const row = r.slice();
      while (row.length < width) row.push('');
      row[0] = normalizePlanDateKey_(row[0]);
      return row;
    });
    sh.getRange(DATA_START_ROW + bodyValues.length, 1, normalizedAppend.length, width).setValues(normalizedAppend);
  }

  const finalLastRow = sh.getLastRow();
  if (finalLastRow >= DATA_START_ROW) {
    sh.getRange(DATA_START_ROW, memberCol, finalLastRow - DATA_START_ROW + 1, 1).setWrap(true);
  }
}

/** ======================================================================
 *   個人シートの高速保存
 *  ====================================================================== */

function api_saveMemberMonthPlanFast(payload) {
  try {
    const memberName = String(payload.memberName || '').trim();
    const rows = payload.rows || [];
    if (!memberName) throw new Error('memberName が空です');

    const sh = getPlanSheet_(memberName);
    const hm = getHeaderMap_(sh, 1);

    const required = [
      '日付','曜日','所定休日',
      'AM_ステータス','AM_勤務場所','AM_休暇','AM_対応顧客名',
      'PM_ステータス','PM_勤務場所','PM_休暇','PM_対応顧客名'
    ];
    required.forEach(h => {
      if (!hm[h]) throw new Error('個人シートのヘッダー不足: ' + h);
    });

    const targetRows = rows
      .map(r => ({
        ...r,
        date: normalizePlanDateKey_(r.date),
      }))
      .filter(r => r.date && isMeaningfulPlanRow_(r));

    if (!targetRows.length) {
      return { ok: true, skipped: true };
    }

    const lastCol = sh.getLastColumn();
    const lastRow = sh.getLastRow();

    let sheetValues = [];
    const existingMap = {};

    if (lastRow > 1) {
      sheetValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      sheetValues.forEach((r, i) => {
        const key = normalizePlanDateKey_(r[hm['日付'] - 1]);
        if (key) existingMap[key] = i; // 0-based
      });
    }

    const appendRows = [];
    const summaryRows = [];

    targetRows.forEach(obj => {
      const key = obj.date;
      const week = String(obj.week || '').trim();

      let idx = existingMap[key];
      let rowArr;

      if (idx === undefined) {
        rowArr = new Array(lastCol).fill('');
        rowArr[hm['日付'] - 1] = key;
        rowArr[hm['曜日'] - 1] = week;
        appendRows.push(rowArr);

        idx = sheetValues.length + appendRows.length - 1;
        existingMap[key] = idx;
      } else {
        rowArr = sheetValues[idx];
      }

      const currentRow = rowArr;

      currentRow[hm['日付'] - 1] = normalizePlanDateKey_(currentRow[hm['日付'] - 1] || key);
      if (!currentRow[hm['曜日'] - 1]) currentRow[hm['曜日'] - 1] = week;

      if (obj.amStatus   !== undefined && obj.amStatus   !== '') currentRow[hm['AM_ステータス'] - 1] = obj.amStatus;
      if (obj.amLocation !== undefined && obj.amLocation !== '') currentRow[hm['AM_勤務場所'] - 1]   = obj.amLocation;
      if (obj.amVacation !== undefined && obj.amVacation !== '') currentRow[hm['AM_休暇'] - 1]       = obj.amVacation;
      if (obj.amCustomer !== undefined && obj.amCustomer !== '') currentRow[hm['AM_対応顧客名'] - 1] = obj.amCustomer;

      if (obj.pmStatus   !== undefined && obj.pmStatus   !== '') currentRow[hm['PM_ステータス'] - 1] = obj.pmStatus;
      if (obj.pmLocation !== undefined && obj.pmLocation !== '') currentRow[hm['PM_勤務場所'] - 1]   = obj.pmLocation;
      if (obj.pmVacation !== undefined && obj.pmVacation !== '') currentRow[hm['PM_休暇'] - 1]       = obj.pmVacation;
      if (obj.pmCustomer !== undefined && obj.pmCustomer !== '') currentRow[hm['PM_対応顧客名'] - 1] = obj.pmCustomer;

      summaryRows.push({
        date: normalizePlanDateKey_(currentRow[hm['日付'] - 1]),
        week: String(currentRow[hm['曜日'] - 1] || '').trim(),
        amStatus:   String(currentRow[hm['AM_ステータス'] - 1] || '').trim(),
        amLocation: String(currentRow[hm['AM_勤務場所'] - 1] || '').trim(),
        amVacation: String(currentRow[hm['AM_休暇'] - 1] || '').trim(),
        amCustomer: String(currentRow[hm['AM_対応顧客名'] - 1] || '').trim(),
        pmStatus:   String(currentRow[hm['PM_ステータス'] - 1] || '').trim(),
        pmLocation: String(currentRow[hm['PM_勤務場所'] - 1] || '').trim(),
        pmVacation: String(currentRow[hm['PM_休暇'] - 1] || '').trim(),
        pmCustomer: String(currentRow[hm['PM_対応顧客名'] - 1] || '').trim(),
      });
    });

    if (sheetValues.length) {
      sh.getRange(2, 1, sheetValues.length, lastCol).setValues(sheetValues);
    }

    if (appendRows.length) {
      sh.getRange(sheetValues.length + 2, 1, appendRows.length, lastCol).setValues(appendRows);
    }

    upsertMemberToSummaryFast_(memberName, summaryRows);

    return {
      ok: true,
      updated: targetRows.length,
      appended: appendRows.length
    };

  } catch (e) {
    console.error(e);
    return {
      ok: false,
      error: e.message || String(e),
    };
  }
}

/** ======================================================================
 *   集計文字列作成
 *  ====================================================================== */

function buildPlanHalfText_(status, location, vacation, customer) {
  const st  = String(status   || '').trim();   // 勤務形態
  const loc = String(location || '').trim();   // 勤務場所
  const vac = String(vacation || '').trim();   // 休暇区分
  const cus = String(customer || '').trim();   // 対応顧客名

  if (!st && !loc && !vac && !cus) return '';

  // 休暇
  if (st === '休暇') {
    let text = '[[TAG:休暇]]';
    if (vac) text += vac;
    if (cus) text += `＠${cus}`;
    return text;
  }

  // 仕事系
  // タグは勤務場所で色分け
  const tag = loc || st;

  let text = '';
  if (tag) text += `[[TAG:${tag}]]`;
  if (cus) text += `＠${cus}`;

  return text;
}

function buildMergedDayText_(row) {
  const amText = buildPlanHalfText_(row.amStatus, row.amLocation, row.amVacation, row.amCustomer);
  const pmText = buildPlanHalfText_(row.pmStatus, row.pmLocation, row.pmVacation, row.pmCustomer);

  if (!pmText) return amText;
  if (amText === pmText) return amText;
  if (!amText) return pmText;

  return `AM：${amText}\nPM：${pmText}`;
}


/** ======================================================================
 *   個人登録画面の API 呼び出しイメージ
 *  ====================================================================== */

function api_getSummaryView(baseYm, months) {
  try {
    const sh = getPlanSheet_(PLAN_APP.SHEET_SUMMARY);
    const values = sh.getDataRange().getValues();

    if (!values.length) {
      return { ok: true, values: [] };
    }

    const header = values[0];
    const rows = values.slice(1).map(r => {
      const row = r.slice();
      row[0] = normalizePlanDateKey_(row[0]);
      return row;
    });

    const start = firstDayOfYm_(baseYm);
    const end   = lastDayOfRange_(baseYm, months);

    const filtered = rows.filter(r => {
      const d = parseDateFlexible_(r[0]);
      if (!d) return false;
      return d >= start && d <= end;
    });

    filtered.sort((a, b) => {
      const da = parseDateFlexible_(a[0]);
      const db = parseDateFlexible_(b[0]);
      return da - db;
    });

    return {
      ok: true,
      values: [header, ...filtered]
    };

  } catch (err) {
    return {
      ok: false,
      error: err.message
    };
  }
}

/** ======================================================================
 *   PLAN画面初期表示用
 *  ====================================================================== */

function api_getPlanInitContext() {
  try {
    const master = api_getPlanMaster();
    if (!master || !master.ok) {
      throw new Error(master?.error || 'マスタ取得失敗');
    }

    const loginName = getActiveUserNameOrEmail_();
    const members = master.members || [];

    let defaultMember = '';
    if (loginName) {
      defaultMember = members.find(x => String(x).trim() === String(loginName).trim()) || '';
    }

    if (!defaultMember && members.length) {
      defaultMember = members[0];
    }

    return {
      ok: true,
      loginName: loginName || '',
      defaultMember: defaultMember || '',
      members: master.members || [],
      memberGroups: master.memberGroups || [],
      statusList: master.statusList || [],
      locationList: master.locationList || [],
      vacationList: master.vacationList || [],
    };

  } catch (err) {
    return {
      ok: false,
      error: err.message,
    };
  }
}