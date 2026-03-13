const PLAN_APP = {
  SS_KEY: 'Plan',
  SHEET_SUMMARY: '予定管理表',
  SHEET_MASTER: 'マスタ',
  SHEET_HOLIDAY: '祝日',
  HEADER_ROW: 1,

  MASTER_HEADERS: {
    member: 'メンバー名',
    order: '表示順',
    status: 'ステータス',
    location: '勤務場所',
    vacation: '休暇',
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
  const sh = getPlanSheet_(PLAN_APP.SHEET_MASTER);
  const hm = getHeaderMap_(sh, PLAN_APP.HEADER_ROW);
  const lastRow = sh.getLastRow();
  if (lastRow <= PLAN_APP.HEADER_ROW) {
    return {
      ok: true,
      members: [],
      statusList: [],
      locationList: [],
      vacationList: [],
    };
  }

  const values = sh.getRange(PLAN_APP.HEADER_ROW + 1, 1, lastRow - PLAN_APP.HEADER_ROW, sh.getLastColumn()).getValues();

  const statusSet = new Set();
  const locationSet = new Set();
  const vacationSet = new Set();
  const members = [];

  values.forEach(r => {
    const member = normalizeText_(r[hm[PLAN_APP.MASTER_HEADERS.member] - 1]);
    const order = Number(r[hm[PLAN_APP.MASTER_HEADERS.order] - 1] || 9999);
    const status = normalizeText_(r[hm[PLAN_APP.MASTER_HEADERS.status] - 1]);
    const location = normalizeText_(r[hm[PLAN_APP.MASTER_HEADERS.location] - 1]);
    const vacation = normalizeText_(r[hm[PLAN_APP.MASTER_HEADERS.vacation] - 1]);

    if (status) statusSet.add(status);
    if (location) locationSet.add(location);
    if (vacation) vacationSet.add(vacation);
    if (member) members.push({ name: member, order });
  });

  const uniqMembers = Array.from(
    new Map(members.map(x => [x.name, x])).values()
  ).sort((a, b) => a.order - b.order);

  return {
    ok: true,
    members: uniqMembers.map(x => x.name),
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
      const key = formatDateKey_(r[hm[needHeaders.date] - 1]);
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
    const key = formatDateKey_(d);
    data.push(
      mapByDate[key] || {
        date: key,
        week: getWeekdayJa_(d),
        amStatus: '',
        amLocation: '',
        amVacation: '',
        amCustomer: '',
        pmStatus: '',
        pmLocation: '',
        pmVacation: '',
        pmCustomer: '',
      }
    );
  });

  return { ok: true, data };
}

/** ======================================================================
 *   個人シートの保存
 *  ====================================================================== */

function api_saveMemberMonthPlan(payload) {
  const memberName = String(payload.memberName || '').trim();
  const rows = payload.rows || [];
  if (!memberName) throw new Error('memberName が空です');

  const sh = getPlanSheet_(memberName);
  const hm = getHeaderMap_(sh, 1);

  const required = [
    '日付','曜日',
    'AM_ステータス','AM_勤務場所','AM_休暇','AM_対応顧客名',
    'PM_ステータス','PM_勤務場所','PM_休暇','PM_対応顧客名'
  ];
  required.forEach(h => {
    if (!hm[h]) throw new Error('個人シートのヘッダー不足: ' + h);
  });

  const lastRow = sh.getLastRow();
  const existingMap = {};

  if (lastRow > 1) {
    const vals = sh.getRange(2, hm['日付'], lastRow - 1, 1).getValues();
    vals.forEach((r, i) => {
      const key = formatDateKey_(r[0]);
      if (key) existingMap[key] = i + 2;
    });
  }

  rows.forEach(obj => {
    const key = String(obj.date || '').trim();
    if (!key) return;

    let row = existingMap[key];
    if (!row) {
      row = sh.getLastRow() + 1;
      existingMap[key] = row;
    }

    sh.getRange(row, hm['日付']).setValue(key);
    sh.getRange(row, hm['曜日']).setValue(obj.week || '');
    sh.getRange(row, hm['AM_ステータス']).setValue(obj.amStatus || '');
    sh.getRange(row, hm['AM_勤務場所']).setValue(obj.amLocation || '');
    sh.getRange(row, hm['AM_休暇']).setValue(obj.amVacation || '');
    sh.getRange(row, hm['AM_対応顧客名']).setValue(obj.amCustomer || '');
    sh.getRange(row, hm['PM_ステータス']).setValue(obj.pmStatus || '');
    sh.getRange(row, hm['PM_勤務場所']).setValue(obj.pmLocation || '');
    sh.getRange(row, hm['PM_休暇']).setValue(obj.pmVacation || '');
    sh.getRange(row, hm['PM_対応顧客名']).setValue(obj.pmCustomer || '');
  });

  // ← 保存後に予定管理表へ反映
  upsertMemberToSummary_(memberName, rows);

  return { ok: true };
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
 *   全体集計作成
 *  ====================================================================== */

function upsertMemberToSummary_(memberName, rows) {
  const sh = getPlanSheet_(PLAN_APP.SHEET_SUMMARY);

  const HEADER_ROW_1 = 1;
  const DATA_START_ROW = 2;

  const lastCol = Math.max(sh.getLastColumn(), 2);
  const h1 = sh.getRange(HEADER_ROW_1, 1, 1, lastCol).getDisplayValues()[0];

  let memberCol = null;
  for (let c = 3; c <= lastCol; c++) {
    const name1 = String(h1[c - 1] || '').trim();
    if (name1 === memberName) {
      memberCol = c;
      break;
    }
  }

  if (!memberCol) {
    memberCol = sh.getLastColumn() + 1;
    sh.getRange(HEADER_ROW_1, memberCol).setValue(memberName);
  }

  const lastRow = sh.getLastRow();
  const dateMap = {};

  if (lastRow >= DATA_START_ROW) {
    const dates = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
    dates.forEach((r, i) => {
      const key = formatDateKey_(r[0]);
      if (key) dateMap[key] = DATA_START_ROW + i;
    });
  }

  rows.forEach(r => {
    const dateKey = String(r.date || '').trim();
    if (!dateKey) return;

    let rowNum = dateMap[dateKey];
    if (!rowNum) {
      rowNum = sh.getLastRow() + 1;
      dateMap[dateKey] = rowNum;
      sh.getRange(rowNum, 1).setValue(dateKey);
      sh.getRange(rowNum, 2).setValue(r.week || '');
    }

    const mergedText = buildMergedDayText_(r);
    sh.getRange(rowNum, memberCol).setValue(mergedText);
  });

  sh.getRange(
    DATA_START_ROW,
    memberCol,
    Math.max(sh.getLastRow() - DATA_START_ROW + 1, 1),
    1
  ).setWrap(true);
}

/** ======================================================================
 *   個人登録画面の API 呼び出しイメージ
 *  ====================================================================== */

function api_getSummaryView(baseYm, months) {
  try {
    const sh = getPlanSheet_(PLAN_APP.SHEET_SUMMARY);
    const values = sh.getDataRange().getDisplayValues();

    if (!values.length) {
      return { ok: true, values: [] };
    }

    const header = values[0];
    const rows = values.slice(1);

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

    const loginName = getActiveUserNameOrEmail_(); // ← これを使う
    const members = master.members || [];

    let defaultMember = '';
    if (loginName) {
      defaultMember = members.find(x => String(x).trim() === String(loginName).trim()) || '';
    }

    // 一致しなければ先頭
    if (!defaultMember && members.length) {
      defaultMember = members[0];
    }

    return {
      ok: true,
      loginName: loginName || '',
      defaultMember: defaultMember || '',
      members: master.members || [],
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
