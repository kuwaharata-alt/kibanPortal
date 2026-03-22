function api_syncMemberMonthPlanCalendar(payload) {
  try {
    const memberName = String(payload.memberName || '').trim();
    if (!memberName) throw new Error('memberName が空です');

    const sh = getPlanSheet_(memberName);
    const hm = getHeaderMap_(sh, 1);

    const required = [
      '日付','曜日',
      'AM_ステータス','AM_勤務場所','AM_休暇','AM_対応顧客名',
      'AM_eventId','AM_syncStatus','AM_syncAt',
      'PM_ステータス','PM_勤務場所','PM_休暇','PM_対応顧客名',
      'PM_eventId','PM_syncStatus','PM_syncAt'
    ];

    required.forEach(h => {
      if (!hm[h]) throw new Error('個人シートのヘッダー不足: ' + h);
    });

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      return { ok: true, count: 0 };
    }

    const lastCol = sh.getLastColumn();
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const rows = values.map((r, i) => ({
      sheetRow: i + 2,
      date: formatDateKey_(r[hm['日付'] - 1]),
      week: r[hm['曜日'] - 1] || '',

      amStatus: r[hm['AM_ステータス'] - 1] || '',
      amLocation: r[hm['AM_勤務場所'] - 1] || '',
      amVacation: r[hm['AM_休暇'] - 1] || '',
      amCustomer: r[hm['AM_対応顧客名'] - 1] || '',
      amEventId: r[hm['AM_eventId'] - 1] || '',

      pmStatus: r[hm['PM_ステータス'] - 1] || '',
      pmLocation: r[hm['PM_勤務場所'] - 1] || '',
      pmVacation: r[hm['PM_休暇'] - 1] || '',
      pmCustomer: r[hm['PM_対応顧客名'] - 1] || '',
      pmEventId: r[hm['PM_eventId'] - 1] || '',
    })).filter(r => r.date);

    const syncResults = syncMemberMonthPlanToCalendar_(memberName, rows);

    const out = values.map((r, idx) => {
      const rowNo = idx + 2;
      const sync = syncResults.find(x => x.sheetRow === rowNo);

      if (!sync) {
        return [
          r[hm['AM_eventId'] - 1] || '',
          r[hm['AM_syncStatus'] - 1] || '',
          r[hm['AM_syncAt'] - 1] || '',
          r[hm['PM_eventId'] - 1] || '',
          r[hm['PM_syncStatus'] - 1] || '',
          r[hm['PM_syncAt'] - 1] || '',
        ];
      }

      return [
        sync.am?.eventId || '',
        sync.am?.syncStatus || '',
        sync.am?.syncAt || '',
        sync.pm?.eventId || '',
        sync.pm?.syncStatus || '',
        sync.pm?.syncAt || '',
      ];
    });

    sh.getRange(2, hm['AM_eventId'], out.length, 6).setValues(out);

    return {
      ok: true,
      count: syncResults.length,
    };

  } catch (e) {
    console.error(e);
    return {
      ok: false,
      error: e.message || String(e),
    };
  }
}

const PLAN_CAL_SYNC = {
  CALENDAR_ID: 'primary',
  TZ: 'Asia/Tokyo',
  DESC: '※予定管理表より自動登録',

  STATUS_FIXED: '確定',
  STATUS_TENTATIVE: '仮',

  ALL_START: { h: 9,  m: 0 },
  ALL_END:   { h: 18, m: 0 },

  AM_START:  { h: 9,  m: 0 },
  AM_END:    { h: 14, m: 0 },

  PM_START:  { h: 14, m: 0 },
  PM_END:    { h: 18, m: 0 },
};

function syncMemberMonthPlanToCalendar_(memberName, rows) {
  const cal = CalendarApp.getCalendarById(PLAN_CAL_SYNC.CALENDAR_ID);
  if (!cal) throw new Error('Googleカレンダーが取得できません');

  return (rows || []).map(row => syncOnePlanRowToCalendar_(cal, memberName, row));
}

function syncOnePlanRowToCalendar_(cal, memberName, row) {
  const now = formatSyncAt_(new Date());

  const amEnabled = isSyncTargetStatus_(row.amStatus);
  const pmEnabled = isSyncTargetStatus_(row.pmStatus);

  const out = {
    sheetRow: row.sheetRow,
    date: row.date,
    am: { eventId: row.amEventId || '', syncStatus: '', syncAt: '' },
    pm: { eventId: row.pmEventId || '', syncStatus: '', syncAt: '' },
  };

  if (!amEnabled && !pmEnabled) {
    if (row.amEventId) {
      const deleted = deleteCalendarEventSafe_(cal, row.amEventId);
      out.am.eventId = '';
      out.am.syncStatus = deleted ? '削除' : '削除済/未検出';
      out.am.syncAt = now;
    }

    if (row.pmEventId) {
      const deleted = deleteCalendarEventSafe_(cal, row.pmEventId);
      out.pm.eventId = '';
      out.pm.syncStatus = deleted ? '削除' : '削除済/未検出';
      out.pm.syncAt = now;
    }

    return out;
  }

  if (amEnabled && !pmEnabled) {
    const payload = buildCalendarPayload_(row, 'am_only');
    const res = upsertCalendarEvent_(cal, row.amEventId, payload);

    out.am.eventId = res.eventId || '';
    out.am.syncStatus = res.action;
    out.am.syncAt = now;

    if (row.pmEventId) {
      const deleted = deleteCalendarEventSafe_(cal, row.pmEventId);
      out.pm.eventId = '';
      out.pm.syncStatus = deleted ? '削除' : '削除済/未検出';
      out.pm.syncAt = now;
    }

    return out;
  }

  if (!amEnabled && pmEnabled) {
    if (row.amEventId) {
      const deleted = deleteCalendarEventSafe_(cal, row.amEventId);
      out.am.eventId = '';
      out.am.syncStatus = deleted ? '削除' : '削除済/未検出';
      out.am.syncAt = now;
    }

    const payload = buildCalendarPayload_(row, 'pm');
    const res = upsertCalendarEvent_(cal, row.pmEventId, payload);

    out.pm.eventId = res.eventId || '';
    out.pm.syncStatus = res.action;
    out.pm.syncAt = now;

    return out;
  }

  const amPayload = buildCalendarPayload_(row, 'am');
  const pmPayload = buildCalendarPayload_(row, 'pm');

  const amRes = upsertCalendarEvent_(cal, row.amEventId, amPayload);
  const pmRes = upsertCalendarEvent_(cal, row.pmEventId, pmPayload);

  out.am.eventId = amRes.eventId || '';
  out.am.syncStatus = amRes.action;
  out.am.syncAt = now;

  out.pm.eventId = pmRes.eventId || '';
  out.pm.syncStatus = pmRes.action;
  out.pm.syncAt = now;

  return out;
}

function isSyncTargetStatus_(status) {
  const s = String(status || '').trim();
  return s === PLAN_CAL_SYNC.STATUS_FIXED || s === PLAN_CAL_SYNC.STATUS_TENTATIVE;
}

function buildCalendarPayload_(row, mode) {
  const date = parsePlanDate_(row.date);
  if (!date) throw new Error(`日付が不正です: ${row.date}`);

  let status = '';
  let location = '';
  let vacation = '';
  let customer = '';
  let startH = 9;
  let startM = 0;
  let endH = 18;
  let endM = 0;

  if (mode === 'am_only') {
    status = row.amStatus;
    location = row.amLocation;
    vacation = row.amVacation;
    customer = row.amCustomer;
    startH = PLAN_CAL_SYNC.ALL_START.h;
    startM = PLAN_CAL_SYNC.ALL_START.m;
    endH = PLAN_CAL_SYNC.ALL_END.h;
    endM = PLAN_CAL_SYNC.ALL_END.m;
  } else if (mode === 'am') {
    status = row.amStatus;
    location = row.amLocation;
    vacation = row.amVacation;
    customer = row.amCustomer;
    startH = PLAN_CAL_SYNC.AM_START.h;
    startM = PLAN_CAL_SYNC.AM_START.m;
    endH = PLAN_CAL_SYNC.AM_END.h;
    endM = PLAN_CAL_SYNC.AM_END.m;
  } else if (mode === 'pm') {
    status = row.pmStatus;
    location = row.pmLocation;
    vacation = row.pmVacation;
    customer = row.pmCustomer;
    startH = PLAN_CAL_SYNC.PM_START.h;
    startM = PLAN_CAL_SYNC.PM_START.m;
    endH = PLAN_CAL_SYNC.PM_END.h;
    endM = PLAN_CAL_SYNC.PM_END.m;
  } else {
    throw new Error(`modeが不正です: ${mode}`);
  }

  return {
    title: buildPlanCalendarTitle_(status, location, vacation, customer),
    start: makeDateTime_(date, startH, startM),
    end: makeDateTime_(date, endH, endM),
    description: PLAN_CAL_SYNC.DESC,
  };
}

function buildPlanCalendarTitle_(status, location, vacation, customer) {
  const st = String(status || '').trim();
  const loc = String(location || '').trim();
  const vac = String(vacation || '').trim();
  const cus = String(customer || '').trim();

  let body = '';

  if (vac || /休暇/.test(loc)) {
    body = '休暇';
  } else if (/現地/.test(loc)) {
    body = cus ? `外)${cus}` : '外)顧客対応';
  } else if (loc) {
    body = cus ? `内)${cus}` : '内)顧客対応';
  } else {
    body = cus ? `内)${cus}` : '内)予定';
  }

  if (st === PLAN_CAL_SYNC.STATUS_TENTATIVE) {
    return `仮/${body}`;
  }

  return body;
}

function upsertCalendarEvent_(cal, eventId, payload) {
  const currentId = String(eventId || '').trim();

  if (currentId) {
    const ev = getCalendarEventByIdSafe_(cal, currentId);
    if (ev) {
      ev.setTitle(payload.title);
      ev.setTime(payload.start, payload.end);
      ev.setDescription(payload.description || '');
      return { action: '更新', eventId: ev.getId() };
    }
  }

  const ev = cal.createEvent(
    payload.title,
    payload.start,
    payload.end,
    { description: payload.description || '' }
  );

  return { action: '作成', eventId: ev.getId() };
}

function deleteCalendarEventSafe_(cal, eventId) {
  const ev = getCalendarEventByIdSafe_(cal, eventId);
  if (!ev) return false;
  ev.deleteEvent();
  return true;
}

function getCalendarEventByIdSafe_(cal, eventId) {
  try {
    return cal.getEventById(String(eventId || '').trim()) || null;
  } catch (e) {
    return null;
  }
}

function parsePlanDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const s = String(value || '').trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return null;

  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function makeDateTime_(baseDate, h, m) {
  return new Date(
    baseDate.getFullYear(),
    baseDate.getMonth(),
    baseDate.getDate(),
    h,
    m,
    0,
    0
  );
}

function formatSyncAt_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}
