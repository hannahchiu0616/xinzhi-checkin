// 心之使者打卡：網站「回報專區」寫入設定試算表
//   每日回報   → 「回報」分頁（沒有會自動建立）
//   大堂課回報 → 「大堂課免死」分頁
//   退出／休息 → 「成員」分頁
// 安裝：設定試算表 → 擴充功能 → Apps Script → 貼上這整段 → 部署 → 新增部署作業
//       類型選「網頁應用程式」，執行身分「我」，誰可以存取「所有人」→ 部署，把網址給網站
// 之後改了程式：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署（網址不變）

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents || "{}");
    const name = String(p.name || "").trim().slice(0, 20);
    if (!name) return out({ ok: false, error: "請填名字" });
    const ss = SpreadsheetApp.getActive();
    const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");

    if (p.type === "class") {
      const from = day(p.from), to = day(p.to);
      if (!from || !to || to < from) return out({ ok: false, error: "請確認上課日期" });
      append(ss, "大堂課免死", { "名字": name, "第一天": slash(from), "最後一天": slash(to) });
      return out({ ok: true, ts: now });
    }

    if (p.type === "pause") {
      const from = day(p.from), back = p.back ? day(p.back) : "";
      if (!from || (p.back && (!back || back <= from))) return out({ ok: false, error: "請確認日期" });
      const kind = p.kind === "退出" ? "退出" : "休息";
      const why = String(p.why || "").trim().slice(0, 50);
      append(ss, "成員", { "成員": name, "不計算開始": slash(from), "回歸": back ? slash(back) : "", "原因": why ? kind + "：" + why : kind });
      return out({ ok: true, ts: now });
    }

    // 每日回報
    const date = day(p.date);
    if (!date) return out({ ok: false, error: "請確認日期" });
    let sh = ss.getSheetByName("回報");
    if (!sh) {
      sh = ss.insertSheet("回報");
      sh.getRange(1, 1, 1, 4).setValues([["時間戳記", "姓名", "日期", "狀態"]]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat("@").setValues([[now, name, date, "已完成"]]);
    return out({ ok: true, ts: now });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 依第一列標題找欄位，加在最後一列下面
function append(ss, sheetName, values) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("設定試算表找不到「" + sheetName + "」分頁");
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const row = header.map(() => "");
  for (const k in values) {
    const i = header.findIndex(h => h.indexOf(k) >= 0);
    if (i < 0) throw new Error("「" + sheetName + "」分頁找不到「" + k + "」欄");
    row[i] = values[k];
  }
  // 只寫到最後一個有內容的欄，避免蓋到旁邊的說明
  let last = row.length;
  while (last > 0 && row[last - 1] === "") last--;
  const r = lastDataRow(sh, header) + 1;
  sh.getRange(r, 1, 1, last).setValues([row.slice(0, last)]);
}

// 第一欄最後一個有內容的列（說明文字通常在右邊，不算）
function lastDataRow(sh) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (String(col[i][0]).trim()) return i + 1;
  return 1;
}

function day(s) { s = String(s || ""); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ""; }
// 寫成 2026/11/5，試算表會當成日期（跟原本手打的一樣）
function slash(k) { const [y, m, d] = k.split("-"); return y + "/" + Number(m) + "/" + Number(d); }

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
