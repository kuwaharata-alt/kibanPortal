/** ----------------------------------------------------
 *   シート選択
 *  ---------------------------------------------------- */

function getSS_(key) {
  const k = String(key ?? '').trim();   // ← 空白・null対策

  const map = {
    Display: '1f3Q0d3G4wEatiT_xK6hXKodoa52r5iDFG5q18iKHPFY', // Portal
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