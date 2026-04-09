/** ----------------------------------------------------
 *   シート選択
 *  ---------------------------------------------------- */

function getSS_(key) {
  const k = String(key ?? '').trim();   // ← 空白・null対策

  const map = {
    Display: '1f3Q0d3G4wEatiT_xK6hXKodoa52r5iDFG5q18iKHPFY', // Portal
    Project: '1czr0sPRCiAi9y0Yi-Dya4K5g-VLDlZFAEuZH63KWPKU', //案件管理表v3
    DB: '1y8jni2p9dg-UTP-L8rVOHJG9iZEU3A72tAW89rQdqaU', // DB
    Link: '1BDprrzWMRp96HhSi-_PqJwci26Hbpn8qjMPcJJf6nrg', // リンク集
    Rule: '1DYyx0fWAjxR9_4Gw7tPyZg-LJjSQ7qmnW0LRSVpWg2c', // 基盤ルール
    Master: '1E2geYb4aHa9Wvma_Nyvk3Dmg24jj0ENwj9wEJu3BdCE', // マスタ
    Task: '1e9MU0a0S8B-C_VIl7pUUMayU429gN1wwkWApGcSQ5vw', // タスク
    Plan: '1Rh9lQrLyre4IcQkPx_4cD_0j_ibwR4LWFL7jHExKL_k', // 予定
    Operation: '1zYeUZvWs2vPoLvksmjUHx06AZMZaSTe05ZJw5jPfLno', // 稼働状況
    Manual: '1uGB_vvDl0yBj8wo0r5koxvKQxwowgvJdbumCoY5JBGM', //マニュアル
  };

  const id = map[k];
  if (!id) {
    throw new Error(`getSS_: unknown key "${key}"`);
  }
  return SpreadsheetApp.openById(id);
}

/** ----------------------------------------------------
 *   メンバー情報取得
 *  ---------------------------------------------------- */

function api_getMemberList() {
  try {
    const ss = getSS_('Master');
    const sh = ss.getSheetByName('BSOL情報');
    if (!sh) return { ok: false, error: 'BSOL情報シートが見つかりません' };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, list: [] };

    const list = sh.getRange(2, 7, lastRow - 1, 1).getDisplayValues()
      .flat()
      .map(v => String(v || '').trim())
      .filter(Boolean);

    return { ok: true, list: [...new Set(list)] };

  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/** ----------------------------------------------------
 *  ログインユーザー取得
 *  WebAppがドメイン公開の場合のみメール取得可能
 *  ---------------------------------------------------- */

function api_getLoginUser() {
  try {
    const email = Session.getActiveUser().getEmail() || '';
    const name  = email ? email.split('@')[0] : '';

    return {
      ok: true,
      email: email,
      name: name
    };

  } catch (err) {
    return {
      ok: false,
      error: err.message
    };
  }
}

/** ----------------------------------------------------
 * 更新者取得
 * - Session.getActiveUser().getEmail()
 * - BSOL情報 F:メールアドレス / G:氏名
 *  ---------------------------------------------------- */

function getActiveUserEmail_() {
  try {
    const email = norm_(Session.getActiveUser().getEmail()).toLowerCase();
    Logger.log('[getActiveUserEmail_] email=' + email);
    return email;
  } catch (e) {
    Logger.log('[getActiveUserEmail_] error=' + e);
    return '';
  }
}

function getUserNameFromBSOL_(email) {
  email = norm_(email).toLowerCase();
  if (!email) {
    Logger.log('[getUserNameFromBSOL_] email empty');
    return '';
  }

  try {
    const sh = getSS_('DB').getSheetByName(INPUT_CONFIG.SHEET_MEMBER);
    if (!sh) {
      Logger.log('[getUserNameFromBSOL_] sheet not found: ' + INPUT_CONFIG.SHEET_MEMBER);
      return '';
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      Logger.log('[getUserNameFromBSOL_] no data');
      return '';
    }

    // F:G = メールアドレス, 氏名
    const vals = sh.getRange(2, 6, lastRow - 1, 2).getValues();

    for (let i = 0; i < vals.length; i++) {
      const rowMail = norm_(vals[i][0]).toLowerCase();
      const rowName = norm_(vals[i][1]);
      if (rowMail === email) {
        Logger.log('[getUserNameFromBSOL_] hit name=' + rowName);
        return rowName || '';
      }
    }

    Logger.log('[getUserNameFromBSOL_] no match email=' + email);
    return '';
  } catch (e) {
    Logger.log('[getUserNameFromBSOL_] error=' + e);
    return '';
  }
}

function getActiveUserNameOrEmail_() {
  const email = getActiveUserEmail_();
  const name = getUserNameFromBSOL_(email);
  const out = name || email || '';
  Logger.log('[getActiveUserNameOrEmail_] out=' + out);
  return out;
}

/** ======================================================================
 *   アクセス指定
 *  ====================================================================== */

function api_getMenuAuth() {
  try {
    const email = String(Session.getActiveUser().getEmail() || '')
      .trim()
      .toLowerCase();

    const ss = getSS_('Display');

    const shAuth = ss.getSheetByName('タブ');        // ← 権限設定
    const shAccess = ss.getSheetByName('アクセス権'); // ← 1枚に統合

    if (!shAuth) throw new Error('タブシートがない');
    if (!shAccess) throw new Error('アクセス権シートがない');

    // =============================
    // アクセス権判定
    // =============================
    const accessValues = shAccess.getDataRange().getDisplayValues();

    const headers = accessValues[0];
    const colMap = {};
    headers.forEach((h, i) => colMap[h] = i);

    const idxMail = colMap['メールアドレス'];
    const idxAdmin = colMap['管理者'];
    const idxUser = colMap['ユーザー'];

    let isAdmin = false;
    let isUser = false;

    for (let i = 1; i < accessValues.length; i++) {
      const row = accessValues[i];
      const mail = String(row[idxMail] || '').trim().toLowerCase();

      if (mail !== email) continue;

      if (row[idxAdmin] === '○') isAdmin = true;
      if (row[idxUser] === '○') isUser = true;
      break;
    }

    // =============================
    // 権限設定
    // =============================
    const values = shAuth.getDataRange().getDisplayValues();

    const h = values[0];
    const map = {};
    h.forEach((v, i) => map[v] = i);

    const allowedKeys = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const key = row[map['項目']];
      if (!key) continue;

      const adminFlag = row[map['管理者']] === '○';
      const userFlag = row[map['ユーザー']] === '○';

      const customRaw = row[map['カスタム列']];
      const customList = String(customRaw || '')
        .split(/[,\n、]/)
        .map(v => v.trim().toLowerCase())
        .filter(Boolean);

      const allow =
        (adminFlag && isAdmin) ||
        (userFlag && isUser) ||
        customList.includes(email);

      if (allow) allowedKeys.push(key);
    }

    return {
      ok: true,
      email,
      isAdmin,
      isUser,
      allowedKeys
    };

  } catch (e) {
    return {
      ok: false,
      error: e.message,
      allowedKeys: []
    };
  }
}

/** ======================================================================
 *   祝日取得
 *  ====================================================================== */

function api_getPlanHolidayList() {
  try {
    const ss = getSS_('DB');
    const sh = ss.getSheetByName('祝日');
    if (!sh) return { ok: true, dates: [] };

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      return { ok: true, dates: [] };
    }

    const vals = sh.getRange(2, 2, lastRow - 1, 1).getValues();

    const dates = vals
      .map(r => toHolidayKey_(r[0]))
      .filter(Boolean);

    return { ok: true, dates };

  } catch (err) {
    return {
      ok: false,
      error: err.message
    };
  }
}

function toHolidayKey_(v) {
  if (!v) return '';

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  const s = String(v).trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return '';

  const y = m[1];
  const mo = ('0' + m[2]).slice(-2);
  const d = ('0' + m[3]).slice(-2);
  return `${y}-${mo}-${d}`;
}