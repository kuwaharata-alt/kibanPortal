function doGet(e) {
  const p = e?.parameter || {};
  const action = (p.action || '').toLowerCase();

  if (action) {
    const out = handleApiGet_(e) || { ok: false, error: `unknown action: ${action}` };
    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SOL推進基盤Grポータル')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // 必要なら
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function doPost(e) {
  const res = handleApiPost_(e);
  return ContentService
    .createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiGet_(e) {
  const p = e.parameter || {};
  const action = (p.action || '').toLowerCase();

  if (action === 'case') {
    return apiGetCaseByKey_(p.key || '');
  }

  return { ok: false, error: `unknown action: ${action}` };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiPost_(e) {
  const body = JSON.parse(e.postData.contents || "{}");

  if (body.action === "updateSv") {
    updateSv_(body);
    return { ok: true };
  }

  return { ok: false, error: "unknown action" };
}


/** 疎通確認 */
function api_ping() {
  return {
    ok: true,
    ts: new Date().toISOString(),
    viewSsid: WEBVIEW.VIEW_SSID,
    sheetSV: WEBVIEW.SHEET_SV,
    sheetCL: WEBVIEW.SHEET_CL,
    config: {
      SV: VIEW_CONFIG.SV,
      CL: VIEW_CONFIG.CL,
    }
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
