/**
 * 2026 世界盃辦公室押注牆 — Google Apps Script 後端
 *
 * 功能：
 *   GET  → 回傳所有押注紀錄（JSON 陣列）
 *   POST → 寫入一筆押注 {name, team, ts} 到 Google Sheet
 *
 * 部署步驟：
 *   1. 開一個新的 Google Sheet（任意名稱）
 *   2. 擴充功能 → Apps Script，貼上這份程式碼並儲存
 *   3. 右上角「部署」→「新增部署作業」→ 類型選「網路應用程式」
 *      - 執行身分：我
 *      - 誰可以存取：任何人
 *   4. 複製產生的網址（https://script.google.com/macros/s/...../exec）
 *   5. 貼到 worldcup2026-predictor.html 裡的 const API_URL = "這裡"
 */

const SHEET_NAME = "votes";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["時間", "暱稱", "押注隊伍"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 讀取所有押注紀錄 */
function doGet() {
  const rows = getSheet_().getDataRange().getValues();
  const out = rows.slice(1).map(r => ({
    ts:   new Date(r[0]).getTime(),
    name: String(r[1]),
    team: String(r[2]),
  }));
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 寫入一筆押注 */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || "").trim().slice(0, 20);
    const team = String(d.team || "").trim().slice(0, 20);
    if (!name || !team) throw new Error("暱稱與隊伍皆為必填");

    // 簡單防灌水：同一暱稱 10 秒內只收一筆
    const cache = CacheService.getScriptCache();
    if (cache.get("v_" + name)) throw new Error("操作太頻繁，請稍候再試");
    cache.put("v_" + name, "1", 10);

    getSheet_().appendRow([new Date(), name, team]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
