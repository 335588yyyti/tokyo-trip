/**
 * 東京旅費記帳：Google 試算表雲端同步
 * 用法：在 Google 試算表的「擴充功能 → Apps Script」貼上這整段，
 *       再「部署 → 新增部署作業 → 網頁應用程式」，誰可以存取選「所有人」。
 * 匯率：每 6 小時自動抓一次最新日圓→台幣匯率（市場中間價），隨記帳一起傳給手機。
 *       改完這段程式要「部署 → 管理部署作業 → 編輯 → 版本選『新版本』→ 部署」，
 *       第一次會要求授權「連線到外部服務」，按允許即可（網址不會變）。
 */
const SHEET_NAME = '記帳';
const HEADERS = ['id', 'title', 'amount', 'currency', 'type', 'payer', 'deleted', 'updatedAt'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const values = getSheet_().getDataRange().getValues();
  const items = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    items.push({
      id: Number(r[0]),
      title: String(r[1]),
      amount: Number(r[2]),
      currency: String(r[3]),
      type: String(r[4]),
      payer: String(r[5]),
      deleted: r[6] === true || String(r[6]).toUpperCase() === 'TRUE',
      updatedAt: Number(r[7]) || 0
    });
  }
  return items;
}

// ---- 匯率（1 日圓 = ? 台幣）----
const RATE_MAX_AGE = 6 * 60 * 60 * 1000;

function fetchRate_() {
  const res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/JPY', { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('rate http ' + res.getResponseCode());
  const twd = Number(JSON.parse(res.getContentText()).rates.TWD);
  // 合理範圍檢查，避免抓到怪數字把帳算錯
  if (!(twd > 0.1 && twd < 0.5)) throw new Error('rate out of range: ' + twd);
  return twd;
}

// 可以在 Apps Script 編輯器手動執行這個函式，立刻更新匯率
function refreshRate() {
  const rate = { value: Math.round(fetchRate_() * 10000) / 10000, at: Date.now() };
  PropertiesService.getScriptProperties().setProperty('rate', JSON.stringify(rate));
  return rate;
}

function getRate_() {
  let rate = null;
  try { rate = JSON.parse(PropertiesService.getScriptProperties().getProperty('rate')); } catch (e) {}
  if (!rate || Date.now() - rate.at > RATE_MAX_AGE) {
    try { rate = refreshRate(); } catch (e) { /* 抓不到就沿用上一次的 */ }
  }
  return rate;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 防止項目名稱被試算表當成公式
function safeText_(t) {
  t = String(t || '').slice(0, 200);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function doGet() {
  return json_({ ok: true, items: readAll_(), rate: getRate_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const ops = Array.isArray(body.ops) ? body.ops : [];
    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    const rowById = {};
    for (let i = 1; i < values.length; i++) {
      if (values[i][0]) rowById[Number(values[i][0])] = { row: i + 1, updatedAt: Number(values[i][7]) || 0 };
    }
    ops.forEach(function (op) {
      const it = op && op.item;
      if (!it || !it.id || typeof it.amount !== 'number') return;
      const row = [
        Number(it.id), safeText_(it.title), Number(it.amount),
        it.currency === 'TWD' ? 'TWD' : 'JPY',
        ['split', 'single', 'other'].indexOf(it.type) >= 0 ? it.type : 'single',
        it.payer === '洪' || it.payer === '盧' ? it.payer : '',
        !!it.deleted, Number(it.updatedAt) || Date.now()
      ];
      const found = rowById[row[0]];
      if (found) {
        if (row[7] >= found.updatedAt) sh.getRange(found.row, 1, 1, HEADERS.length).setValues([row]);
      } else {
        sh.appendRow(row);
        rowById[row[0]] = { row: sh.getLastRow(), updatedAt: row[7] };
      }
    });
    return json_({ ok: true, items: readAll_(), rate: getRate_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
