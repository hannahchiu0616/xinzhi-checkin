// 打卡網站後端（多群組）
// 每個群組一份自己的試算表（自動建立在擁有者的 Google 雲端硬碟），網站用 ?g=群組代碼 分開
// 安裝／更新：Apps Script 貼上這整段 → 存檔 → 部署 → 管理部署作業 → 鉛筆 → 版本選「新版本」→ 部署（網址不變）
// 第一次：上方函式選 setupXinzhi → 執行（把心之使者的資料搬進新的群組）

const SITE = "https://hannahchiu0616.github.io/xinzhi-checkin/";
const TABS = {
  "設定": ["項目", "值"],
  "成員": ["名字", "加入日期", "結束日期"],
  "回報": ["時間戳記", "姓名", "日期", "狀態"],
  "免除": ["名字", "第一天", "最後一天"],
  "退出休息": ["名字", "不計算開始日", "回歸日期", "原因"],
  "手算調整": ["名字", "月份", "金額"],
  "管理員": ["Email"],
  "修改紀錄": ["時間", "Email", "動作", "內容"],
};
const PUBLIC_TABS = ["成員", "回報", "免除", "退出休息", "手算調整"];   // 大家都看得到（管理員、修改紀錄不公開）
const EDIT_TABS = ["成員", "免除", "退出休息", "手算調整", "管理員"];   // 管理頁可以整份修改的分頁
const SETTING_KEYS = ["名稱", "副標題", "任務名稱", "罰款名稱", "開始計算日", "每次罰款", "每月免死次數", "免死從",
  "切換時間", "免除名稱", "免除提前天數", "保留月數", "規則說明"];

/* ---------- 共用 ---------- */
const props = () => PropertiesService.getScriptProperties();
const cache = () => CacheService.getScriptCache();
const owner = () => Session.getEffectiveUser().getEmail().toLowerCase();
const now = () => Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
const today = () => Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
const day = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || "")) ? String(s) : "";
const cleanName = s => String(s || "").trim().replace(/[,，\t\n]/g, " ").slice(0, 20);
const normEmail = s => String(s || "").trim().toLowerCase();

function groups() { return JSON.parse(props().getProperty("groups") || "{}"); }
function saveGroups(g) { props().setProperty("groups", JSON.stringify(g)); }
function book(g) {
  const info = groups()[g];
  if (!info) throw new Error("找不到這個群組，請確認網址");
  return SpreadsheetApp.openById(info.sheetId);
}

function tab(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), Math.min(sh.getMaxColumns(), 10)).setNumberFormat("@");
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows(ss, name) {
  const sh = tab(ss, name), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, TABS[name].length).getDisplayValues().filter(r => r.some(x => String(x).trim()));
}
function setRows(ss, name, data) {
  const sh = tab(ss, name), w = TABS[name].length, n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, w).clearContent();
  data = (data || []).map(r => TABS[name].map((_, i) => String(r[i] == null ? "" : r[i]).trim())).filter(r => r.some(Boolean));
  if (data.length) sh.getRange(2, 1, data.length, w).setNumberFormat("@").setValues(data);
  return data.length;
}
function append(ss, name, list) {
  if (!list.length) return;
  const sh = tab(ss, name);
  sh.getRange(sh.getLastRow() + 1, 1, list.length, list[0].length).setNumberFormat("@").setValues(list.map(r => r.map(String)));
}
function settingsOf(ss) { const o = {}; rows(ss, "設定").forEach(([k, v]) => { o[k] = v; }); return o; }
function log(ss, email, action, detail) { append(ss, "修改紀錄", [[now(), email, action, String(detail || "").slice(0, 500)]]); }
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// 資料保留：只留最近 N 個月（含這個月）的回報，空白或 0＝全部保留；回傳刪掉幾筆
function prune(ss) {
  const n = parseInt(settingsOf(ss)["保留月數"], 10);
  if (!n || n < 1) return 0;
  const d = new Date();
  d.setMonth(d.getMonth() - (n - 1), 1);
  const keepFrom = Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM");
  const ym = s => { const m = String(s || "").match(/(\d{4})\D+(\d{1,2})/); return m ? m[1] + "-" + ("0" + m[2]).slice(-2) : ""; };
  const all = rows(ss, "回報");
  const keep = all.filter(r => (ym(r[2]) || ym(r[0])) >= keepFrom);
  if (keep.length === all.length) return 0;
  setRows(ss, "回報", keep);
  return all.length - keep.length;
}

function ensureMember(ss, name, date) {
  if (!rows(ss, "成員").some(r => r[0] === name)) append(ss, "成員", [[name, date, ""]]);
}

/* ---------- 讀資料（大家都可以） ---------- */
function doGet(e) {
  try {
    const g = String(e.parameter.g || "");
    const ss = book(g);
    const data = { ok: true, settings: settingsOf(ss) };
    for (const t of PUBLIC_TABS) data[t] = rows(ss, t);
    return out(data);
  } catch (err) {
    return out({ ok: false, error: err.message || String(err) });
  }
}

/* ---------- 寫資料 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents || "{}");
    p.g = p.g || "xinzhi";   // 舊版網頁沒帶群組代碼
    const fn = ACTIONS[p.action || (p.type === "class" ? "class" : "report")];
    if (!fn) throw new Error("不支援的動作");
    return out(Object.assign({ ok: true }, fn(p) || {}));
  } catch (err) {
    return out({ ok: false, error: err.message || String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 管理權限：登入後拿到的通行碼（30 天），每次修改都要帶
function auth(p, ownerOnly) {
  const raw = p.token && props().getProperty("t_" + p.token);
  const t = raw && JSON.parse(raw);
  if (!t || t.exp < Date.now()) throw new Error("登入已過期，請重新登入");
  if (ownerOnly) {
    if (t.g !== "*" || t.email !== owner()) throw new Error("只有總管理可以操作");
    return { email: t.email };
  }
  if (t.g !== "*" && t.g !== p.g) throw new Error("沒有這個群組的管理權限");
  if (!canAdmin(p.g, t.email)) throw new Error("你已經不在管理員名單裡");
  return { ss: book(p.g), email: t.email };
}
function canAdmin(g, email) {
  if (email === owner()) return true;
  if (g === "*") return false;
  return rows(book(g), "管理員").some(r => normEmail(r[0]) === email);
}
function cleanTokens() {
  const all = props().getProperties();
  for (const k in all) if (k.indexOf("t_") === 0) {
    try { if (JSON.parse(all[k]).exp < Date.now()) props().deleteProperty(k); } catch (e) { props().deleteProperty(k); }
  }
}

const ACTIONS = {
  // 每日回報（大家都可以）
  report(p) {
    const ss = book(p.g), name = cleanName(p.name), date = day(p.date);
    if (!name || !date) throw new Error("名字或日期不完整");
    const ts = now();
    append(ss, "回報", [[ts, name, date, "已完成"]]);
    ensureMember(ss, name, date);
    // 每天第一次有人回報時，順便清掉超過保留月數的舊回報
    if (!cache().get("pruned:" + p.g)) { cache().put("pruned:" + p.g, "1", 86400); prune(ss); }
    return { ts };
  },
  // 大堂課回報（大家都可以）
  class(p) {
    const ss = book(p.g), name = cleanName(p.name), from = day(p.from), to = day(p.to);
    if (!name || !from || !to || to < from) throw new Error("請確認日期");
    append(ss, "免除", [[name, from, to]]);
    return {};
  },

  // 登入：Email 在管理員名單上就直接登入（不寄驗證碼）
  login(p) {
    const email = normEmail(p.email), g = p.admin ? "*" : p.g;
    if (!/^[^@\s]+@[^@\s]+$/.test(email)) throw new Error("Email 格式不對");
    if (!canAdmin(g, email)) throw new Error("這個 Email 沒有管理權限");
    cleanTokens();
    const token = Utilities.getUuid();
    props().setProperty("t_" + token, JSON.stringify({ g, email, exp: Date.now() + 30 * 864e5 }));
    if (g !== "*") log(book(g), email, "登入", "");
    return { token, email };
  },

  // （備用）寄驗證碼登入：要用的話網站改呼叫 requestCode / verifyCode
  requestCode(p) {
    const email = normEmail(p.email), g = p.admin ? "*" : p.g;
    if (!/^[^@\s]+@[^@\s]+$/.test(email)) throw new Error("Email 格式不對");
    if (!canAdmin(g, email)) throw new Error("這個 Email 沒有管理權限");
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache().put("code:" + g + ":" + email, code, 600);
    const title = g === "*" ? "打卡網站總管理" : (settingsOf(book(g))["名稱"] || "打卡網站");
    MailApp.sendEmail(email, `${title} 管理驗證碼：${code}`,
      `你的驗證碼是 ${code}，10 分鐘內有效。\n\n如果不是你本人操作，請忽略這封信。`);
    return {};
  },
  verifyCode(p) {
    const email = normEmail(p.email), g = p.admin ? "*" : p.g, k = "code:" + g + ":" + email;
    const c = cache().get(k);
    if (!c) throw new Error("驗證碼已過期，請重新寄送");
    if (c !== String(p.code || "").trim()) {
      // 錯 5 次這組驗證碼就作廢，避免被亂猜
      const tries = Number(cache().get(k + ":tries") || 0) + 1;
      if (tries >= 5) { cache().remove(k); cache().remove(k + ":tries"); throw new Error("錯太多次，請重新寄送驗證碼"); }
      cache().put(k + ":tries", String(tries), 600);
      throw new Error("驗證碼不對");
    }
    cache().remove(k);
    cache().remove(k + ":tries");
    cleanTokens();
    const token = Utilities.getUuid();
    props().setProperty("t_" + token, JSON.stringify({ g, email, exp: Date.now() + 30 * 864e5 }));
    if (g !== "*") log(book(g), email, "登入", "");
    return { token, email };
  },
  logout(p) { if (p.token) props().deleteProperty("t_" + p.token); return {}; },

  // 管理頁需要的資料（管理員名單、修改紀錄）
  adminData(p) {
    const { ss, email } = auth(p);
    return { email, admins: rows(ss, "管理員").map(r => r[0]), log: rows(ss, "修改紀錄").slice(-200).reverse() };
  },
  saveSettings(p) {
    const { ss, email } = auth(p);
    const old = settingsOf(ss), s = p.settings || {};
    const keys = SETTING_KEYS.filter(k => k in s || k in old);
    setRows(ss, "設定", keys.map(k => [k, k in s ? s[k] : old[k]]));
    const changed = keys.filter(k => k in s && String(s[k]) !== String(old[k] || ""));
    log(ss, email, "修改設定", changed.map(k => `${k}：${old[k] || "（空）"} → ${s[k] || "（空）"}`).join("；") || "沒有變更");
    const n = prune(ss);
    if (n) log(ss, email, "清除舊資料", `刪掉 ${n} 筆超過保留月數的回報`);
    return {};
  },
  saveTab(p) {
    const { ss, email } = auth(p);
    if (EDIT_TABS.indexOf(p.tab) < 0) throw new Error("不能修改這個分頁");
    if (p.tab === "管理員" && !(p.rows || []).some(r => /@/.test(String(r[0])))) throw new Error("至少要有一位管理員");
    const before = rows(ss, p.tab).length, after = setRows(ss, p.tab, p.rows);
    log(ss, email, "修改" + p.tab, `${before} 筆 → ${after} 筆${p.note ? "：" + p.note : ""}`);
    return {};
  },
  deleteReport(p) {
    const { ss, email } = auth(p);
    const sh = tab(ss, "回報"), vals = sh.getDataRange().getDisplayValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (vals[i][0] === p.ts && vals[i][1] === p.name && vals[i][2] === (p.date || "")) {
        sh.deleteRow(i + 1);
        log(ss, email, "刪除回報", `${p.name} ${p.date || p.ts}`);
        return {};
      }
    }
    throw new Error("找不到這筆回報（可能已經被刪了）");
  },
  addReport(p) {
    const { ss, email } = auth(p);
    const name = cleanName(p.name), date = day(p.date);
    if (!name || !date) throw new Error("名字或日期不完整");
    append(ss, "回報", [[now(), name, date, "補登"]]);
    ensureMember(ss, name, date);
    log(ss, email, "補登回報", `${name} ${date}`);
    return {};
  },

  // 總管理：列出／建立群組
  listGroups(p) {
    auth(p, true);
    const gs = groups();
    return { groups: Object.keys(gs).map(id => ({ id, name: gs[id].name, created: gs[id].created, url: SITE + "?g=" + id,
      sheet: "https://docs.google.com/spreadsheets/d/" + gs[id].sheetId + "/edit" })) };
  },
  createGroup(p) {
    const { email } = auth(p, true);
    const name = String(p.name || "").trim().slice(0, 40);
    const admins = String(p.admins || "").split(/[\s,，;；]+/).map(normEmail).filter(x => /@/.test(x));
    if (!name) throw new Error("請填群組名稱");
    if (!admins.length) throw new Error("請至少填一位負責人 Email");
    const gs = groups();
    let id;
    do { id = Math.random().toString(36).slice(2, 8); } while (gs[id]);
    const ss = newBook(name, {
      "名稱": name, "副標題": "每日回報狀況與罰款結算", "任務名稱": "任務", "罰款名稱": "罰款",
      "開始計算日": today(), "每次罰款": "50", "每月免死次數": "0", "免死從": "", "切換時間": "12:00",
      "免除名稱": "請假", "免除提前天數": "0", "保留月數": "12",
      "規則說明": "每天完成任務後在網站回報\n計算時間為當日中午 12 點至隔日中午 12 點\n隔天中午 12 點前沒回報視同沒做，每次罰 50 元",
    }, admins);
    gs[id] = { name, sheetId: ss.getId(), created: today() };
    saveGroups(gs);
    log(ss, email, "建立群組", name);
    return { id, url: SITE + "?g=" + id };
  },
};

function newBook(name, settings, admins) {
  const ss = SpreadsheetApp.create(name + "（打卡網站資料）");
  const first = ss.getSheets()[0];
  Object.keys(TABS).forEach(t => tab(ss, t));
  ss.deleteSheet(first);
  setRows(ss, "設定", SETTING_KEYS.map(k => [k, settings[k] == null ? "" : settings[k]]));
  setRows(ss, "管理員", admins.map(a => [a]));
  return ss;
}

/* ---------- 一次性：把心之使者搬進新的群組（在 Apps Script 編輯器裡選這個函式按「執行」） ---------- */
function setupXinzhi() {
  const gs = groups();
  if (gs.xinzhi) throw new Error("心之使者已經建立過了：" + SITE + "?g=xinzhi");
  const legacy = SpreadsheetApp.getActive();
  const read = name => {
    const sh = legacy.getSheetByName(name);
    if (!sh) return { h: [], rows: [] };
    const v = sh.getDataRange().getDisplayValues();
    return { h: v[0].map(x => String(x).trim()), rows: v.slice(1).filter(r => String(r[0]).trim()) };
  };
  const pick = (t, names) => t.rows.map(r => names.map(n => { const i = t.h.findIndex(x => x.indexOf(n) >= 0); return i < 0 ? "" : r[i]; }));

  const rule = read("規則"), col = n => { const i = rule.h.findIndex(x => x.indexOf(n) >= 0); return i < 0 ? "" : (rule.rows[0] || [])[i] || ""; };
  const ss = newBook("心之使者打卡", {
    "名稱": "心之使者打卡", "副標題": "觀心書每日回報狀況與功德金結算", "任務名稱": "觀心書", "罰款名稱": "功德金",
    "開始計算日": col("開始計算日") || "2026/6/1", "每次罰款": col("罰款") || "50", "每月免死次數": col("免死次數") || "4",
    "免死從": col("免死從") || "2026-10", "切換時間": "12:00", "免除名稱": "大堂課", "免除提前天數": "1",
    "規則說明": [
      "每日寫觀心書，沒寫一篇罰 50 元功德金",
      "計算時間為當日中午 12 點至隔日中午 12 點",
      "隔天中午 12 點前 po 出並在網站回報；沒回報視同沒寫",
      "每人每月有 4 次免死金牌，第 5 次缺交起才罰款；一個月收一次功德金",
      "有上大堂課的新生、複訓學長有 5 天免死金牌（上課前一天到上課最後一天）",
    ].join("\n"),
  }, [owner()]);

  setRows(ss, "成員", [
    ["汎汎"], ["文鴻"], ["勇達"], ["Amber"], ["俊宇"], ["季晴"], ["美琪"], ["克強"], ["千千"], ["宇熙"], ["育凱"], ["Mina"],
    ["有紳"], ["采玲"], ["享諺"], ["賜銘"], ["士軒"], ["家瑋"], ["靖貽"], ["昱愷"], ["奕昌", "2026/7/20"], ["欣潔"], ["婕綵"],
    ["承霖", "", "2026/10/27"], ["韋甄", "2026/9/20"], ["姿妤"],
  ]);
  setRows(ss, "免除", pick(read("大堂課免死"), ["名字", "第一天", "最後一天"]));
  setRows(ss, "退出休息", pick(read("成員"), ["成員", "不計算開始", "回歸", "原因"]));
  setRows(ss, "手算調整", pick(read("手算調整"), ["名字", "月份", "金額"]));

  // 舊 Google 表單的紀錄（到 11/3）＋網站回報
  const csv = Utilities.parseCsv(UrlFetchApp.fetch(SITE + "history-xinzhi.csv").getContentText("UTF-8"));
  const hist = csv.slice(1).filter(r => r[0] && r[1]).map(r => [r[0], r[1], "", ""]);
  const web = pick(read("回報"), ["時間戳記", "姓名", "日期", "狀態"]);
  setRows(ss, "回報", hist.concat(web));

  gs.xinzhi = { name: "心之使者打卡", sheetId: ss.getId(), created: today() };
  saveGroups(gs);
  log(ss, owner(), "建立群組", `從舊資料搬過來：回報 ${hist.length + web.length} 筆`);
  Logger.log("完成！網址：" + SITE + "?g=xinzhi");
  Logger.log("資料試算表：" + ss.getUrl());
}

/* ---------- 一次性：補上心之使者以前手算的金額（舊網站寫在網頁程式裡，沒在試算表） ---------- */
function fixXinzhiOverrides() {
  const ss = book("xinzhi");
  setRows(ss, "手算調整", [
    ["勇達", "2026-06", "200"],
    ["韋甄", "2026-07", "400"], ["克強", "2026-07", "450"], ["宇熙", "2026-07", "100"], ["婕綵", "2026-07", "100"],
    ["士軒", "2026-08", "250"], ["韋甄", "2026-08", "50"],
  ]);
  log(ss, owner(), "修改手算調整", "補上以前手算的 7 筆金額");
  Logger.log("完成！手算調整 7 筆");
}

/* ---------- 開啟寄信權限：在編輯器選這個函式按「執行」，Google 會問寄信權限，按允許 ---------- */
function testMail() {
  MailApp.sendEmail(owner(), "打卡網站：寄信測試", "收到這封信，代表寄驗證碼的權限已經開好了。");
  Logger.log("已寄出測試信到 " + owner());
}
