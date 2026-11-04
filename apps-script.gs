// 心之使者打卡：網站回報寫入設定試算表的「回報」分頁
// 安裝：設定試算表 → 擴充功能 → Apps Script → 貼上這整段 → 部署 → 新增部署作業
//       類型選「網頁應用程式」，執行身分「我」，誰可以存取「所有人」→ 部署，把網址給網站
const SHEET = "回報";
const HEADER = ["時間戳記", "姓名", "日期", "狀態"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents || "{}");
    const name = String(p.name || "").trim().slice(0, 20);
    const status = p.status === "大堂課pass" ? "大堂課pass" : "已完成";
    const from = String(p.date || ""), to = String(p.to || p.date || "");
    const ok = /^\d{4}-\d{2}-\d{2}$/;
    if (!name || !ok.test(from) || !ok.test(to) || to < from) return out({ ok: false, error: "名字或日期不完整" });

    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
    const rows = [];
    // 大堂課 pass 可以一次填一段日期，每天寫一列（最多 14 天）
    for (let d = new Date(from + "T00:00:00Z"); rows.length < 14; d.setUTCDate(d.getUTCDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      if (k > to) break;
      rows.push([now, name, k, status]);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADER.length).setNumberFormat("@").setValues(rows);
    return out({ ok: true, ts: now, rows: rows.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
