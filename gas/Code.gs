/**
 * All Blue Cebu - メール予約リクエスト受付 (Google Apps Script)
 *
 * 受け取ったリクエストを
 *   1) スプレッドシートに1行追加
 *   2) スタッフに通知メール
 *   3) お客さんに受付メール(自動返信)
 * する。設定手順は gas/セットアップ手順.md を参照。
 */

// ★ 通知先(スタッフ)のメールアドレス。ここを書き換えて再デプロイすれば変更できる。
var NOTIFY_EMAIL = "ここにメールアドレスを入れる";

var SHOP_NAME = "All Blue Cebu";
var SHEET_NAME = "Booking Requests";
var LEGACY_SHEET_NAME = "予約リクエスト";

function doPost(e) {
  try {
    var d = (e && e.parameter) || {};

    // ボット対策: 人間には見えない入力欄に値が入っていたら無視
    if (d.website) return out_({ ok: true });

    var name = clean_(d.name, 60);
    var email = clean_(d.email, 100);
    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return out_({ ok: false, error: "invalid" });

    var tel = clean_(d.tel, 30);
    var tour = clean_(d.tour, 100);
    var date = clean_(d.date, 30);
    var planId = clean_(d.planId, 60);
    var summary = clean_(d.summary, 300);
    var note = clean_(d.note, 1000);

    // 1) スプレッドシートに記録
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    migrate_(ss);
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Received", "Name", "Email", "Phone", "Tour", "Desired date", "Details", "Request / notes", "Status"]);
    }
    sh.appendRow([new Date(), safe_(name), safe_(email), safe_(tel), safe_(planLabel_(planId) || tour), safe_(date), safe_(summary), safe_(note), "New"]);
    ensureRequestSheet_(sh);

    var detail =
      "ツアー：" + tour + "\n" +
      "希望日：" + date + "\n" +
      "内容：" + summary + "\n\n" +
      "お名前：" + name + "\n" +
      "メール：" + email + "\n" +
      "電話：" + (tel || "(未入力)") + "\n" +
      "ご要望：" + (note || "(なし)");

    // 2) スタッフへ通知
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "【予約リクエスト】" + tour + " / " + date + " / " + name + "様",
      body: "Webサイトから予約リクエストが届きました。\n\n" + detail + "\n\n※返信すると、お客様のメールアドレス宛てになります。",
      replyTo: email
    });

    // 3) お客さんへ受付メール
    MailApp.sendEmail({
      to: email,
      subject: "【" + SHOP_NAME + "】予約リクエストを受け付けました",
      body:
        name + " 様\n\n" +
        "この度はお申込みありがとうございます。以下の内容でリクエストを受け付けました。\n\n" +
        detail + "\n\n" +
        "※この時点では予約は確定していません。スタッフが内容を確認し、改めてメールでご連絡します。\n\n" +
        SHOP_NAME,
      name: SHOP_NAME,
      replyTo: NOTIFY_EMAIL
    });

    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: "server" });
  }
}

/* ================= Prices & booking settings (Google Sheet -> website) =================
 * Sheet "Price Rules": Plan / Start date / End date / Adult price / Child price / Status / Notes
 *   - Plan: "All plans" or pick one from the dropdown. Empty price = normal price (plans.js).
 *   - Status: empty = Open. "Full" or "Closed" makes the date unselectable.
 *   - If several rows overlap on a date, the lower row wins.
 * Sheet "Booking Settings": Plan / Days before booking closes / Notes
 *   - "Global (all plans)" is the default; a row for a specific plan overrides it.
 *   - "Max tour types per day": max number of different tours confirmed on one day (= instructors).
 * Sheet "Booking Requests": set Status to "Confirmed" when a booking is confirmed. For LINE bookings,
 *   add a row by hand: Desired date (e.g. 2026/11/05), Tour (dropdown), Status = Confirmed.
 */
var PLAN_LIST = [
  "jinbe-tumalog-sumilon｜ジンベエザメ + ツマログの滝 + スミロン島ツアー",
  "jinbe-tumalog｜ジンベエザメ・ツマログの滝ツアー",
  "jinbe-kawasan｜ジンベエザメ・カワサン滝ツアー",
  "jinbe-sumilon｜ジンベエザメ + スミロン島ツアー",
  "island-3｜アイランドホッピング3島めぐり",
  "island-2｜アイランドホッピング2島めぐり",
  "fun-island｜離島(海洋保護区)2dive",
  "fun-mactan｜マクタン島近郊2dive",
  "ow-padi｜PADI オープンウォーターライセンス講習",
  "ow-ssi｜SSI オープンウォーターライセンス講習",
  "aow-padi｜PADI アドバンスドオープンウォーターライセンス講習",
  "aow-ssi｜SSI アドバンスドオープンウォーターライセンス講習",
  "para｜パラセーリング",
  "jet｜ジェットスキー",
  "banana｜バナナボート",
  "trial-2dive｜体験ダイビング2dive",
  "trial-island2｜体験ダイビング1dive + アイランドホッピング2島めぐり",
  "trial-island3｜体験ダイビング1dive + アイランドホッピング3島めぐり",
  "cruise-only｜クルージングのみ",
  "cruise-snorkel｜クルージング + スノーケル",
  "cruise-skin｜クルージング + スキンダイビング",
  "cruise-trial｜クルージング + 体験ダイビング",
  "cruise-fun｜クルージング + ファンダイビング"
];
var RULE_SHEET = "Price Rules";
var LEGACY_RULE_SHEET = "料金ルール";
var ALL_PLANS = "All plans";            // Price Rules: applies to every plan
var GLOBAL_ROW = "Global (all plans)";  // Booking Settings: default for every plan
var MAX_TOURS_KEY = "Max tour types per day";
var LEGACY_MAX_TOURS_KEY = "1日の最大ツアー種類数";   // シート「受付設定」のこの行の数字 = 1日に確定できるツアーの種類の上限(=稼働できるインストラクター数)
var DEFAULT_MAX_TOURS = 5;
var STATUS_LIST = ["New", "In progress", "Confirmed", "Cancelled"];
var DAY_STATUS_LIST = ["Open", "Full", "Closed"];
// English <-> Japanese values (the website uses the Japanese ones)
var DAY_STATUS_MAP = { "Open": "受付中", "Full": "満席", "Closed": "停止" };
var LEGACY_REQ_STATUS = { "未対応": "New", "対応中": "In progress", "確定": "Confirmed", "キャンセル": "Cancelled" };
var LEGACY_DAY_STATUS = { "受付中": "Open", "満席": "Full", "停止": "Closed" };
var CUTOFF_SHEET = "Booking Settings";
var LEGACY_CUTOFF_SHEET = "受付設定";

// 旧(日本語)のシート名・見出し・値を英語に移す。何度実行しても安全
function migrate_(ss) {
  var pairs = [[LEGACY_SHEET_NAME, SHEET_NAME], [LEGACY_RULE_SHEET, RULE_SHEET], [LEGACY_CUTOFF_SHEET, CUTOFF_SHEET]];
  pairs.forEach(function (p) {
    var old = ss.getSheetByName(p[0]);
    if (old && !ss.getSheetByName(p[1])) old.setName(p[1]);
  });

  // 旧ドロップダウン(日本語の選択肢)が残っていると、英語の値を書き込めないので、先に外す(あとで作り直す)
  var rs0 = ss.getSheetByName(SHEET_NAME), r0 = ss.getSheetByName(RULE_SHEET), c0 = ss.getSheetByName(CUTOFF_SHEET);
  if (rs0) rs0.getRange("A2:J1000").clearDataValidations();
  if (r0) r0.getRange("A2:G500").clearDataValidations();
  if (c0) c0.getRange("A2:C200").clearDataValidations();

  var rs = ss.getSheetByName(SHEET_NAME);
  if (rs && rs.getLastRow() >= 1) {
    rs.getRange(1, 1, 1, 9).setValues([["Received", "Name", "Email", "Phone", "Tour", "Desired date", "Details", "Request / notes", "Status"]]);
    if (rs.getLastRow() > 1) {
      var sr = rs.getRange(2, 9, rs.getLastRow() - 1, 1), sv = sr.getValues();
      sr.setValues(sv.map(function (r) { var v = String(r[0]).trim(); return [LEGACY_REQ_STATUS[v] || r[0]]; }));
    }
    rs.getRange("A1:I1").setFontWeight("bold");
  }

  var r = ss.getSheetByName(RULE_SHEET);
  if (r && r.getLastRow() >= 1) {
    r.getRange(1, 1, 1, 7).setValues([["Plan", "Start date", "End date", "Adult price (JPY)", "Child price (JPY)", "Status", "Notes"]]);
    if (r.getLastRow() > 1) {
      var n = r.getLastRow() - 1;
      var pr = r.getRange(2, 1, n, 1), pv = pr.getValues();
      pr.setValues(pv.map(function (x) { return [String(x[0]).trim() === "全プラン" ? ALL_PLANS : x[0]]; }));
      var st = r.getRange(2, 6, n, 1), stv = st.getValues();
      st.setValues(stv.map(function (x) { var v = String(x[0]).trim(); return [LEGACY_DAY_STATUS[v] || x[0]]; }));
    }
    r.getRange("A1:G1").setFontWeight("bold");
  }

  var c = ss.getSheetByName(CUTOFF_SHEET);
  if (c && c.getLastRow() >= 1) {
    c.getRange(1, 1, 1, 3).setValues([["Plan", "Days before the tour booking closes (1 = until the day before)", "Notes"]]);
    if (c.getLastRow() > 1) {
      var cr = c.getRange(2, 1, c.getLastRow() - 1, 3), cv = cr.getValues();
      cr.setValues(cv.map(function (x) {
        var k = String(x[0]).trim();
        if (k === "全体") return [GLOBAL_ROW, x[1], "Applies to all plans. A row for a specific plan overrides this."];
        if (k === LEGACY_MAX_TOURS_KEY) return [MAX_TOURS_KEY, x[1], "Max number of different tours confirmed per day (= number of available instructors)."];
        return x;
      }));
    }
    c.getRange("A1:C1").setFontWeight("bold");
  }
}

function ensureRequestSheet_(sh) {
  sh.getRange("I2:I1000").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_LIST, true).setAllowInvalid(true).build());
  // Column E "Tour" is used to count capacity. When you add a LINE booking by hand, pick the tour from this dropdown.
  sh.getRange("E2:E1000").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(PLAN_LIST, true).setAllowInvalid(true).build());
  // The old extra column J ("Plan"/"プラン") was merged into E. Clear leftovers.
  var j = sh.getRange("J1:J1000");
  j.clearContent(); j.clearDataValidations(); j.clearFormat();
}

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrate_(ss);

  var r = ss.getSheetByName(RULE_SHEET);
  if (!r) {
    r = ss.insertSheet(RULE_SHEET);
    r.appendRow(["Plan", "Start date", "End date", "Adult price (JPY)", "Child price (JPY)", "Status", "Notes"]);
    r.setFrozenRows(1);
    r.getRange("B2:C500").setNumberFormat("yyyy/mm/dd");
    r.setColumnWidth(1, 360);
    r.getRange("A1:G1").setFontWeight("bold");
  }
  r.getRange("F2:F500").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(DAY_STATUS_LIST, true).setAllowInvalid(false).build());

  var c = ss.getSheetByName(CUTOFF_SHEET);
  if (!c) {
    c = ss.insertSheet(CUTOFF_SHEET);
    c.appendRow(["Plan", "Days before the tour booking closes (1 = until the day before)", "Notes"]);
    c.appendRow([GLOBAL_ROW, 1, "Applies to all plans. A row for a specific plan overrides this."]);
    c.setFrozenRows(1);
    c.setColumnWidth(1, 360);
    c.setColumnWidth(2, 200);
    c.getRange("A1:C1").setFontWeight("bold");
  }

  // Add the "Max tour types per day" row if missing
  var has = false, cvals = c.getDataRange().getValues();
  for (var k = 1; k < cvals.length; k++) if (String(cvals[k][0]).trim() === MAX_TOURS_KEY) has = true;
  if (!has) c.appendRow([MAX_TOURS_KEY, DEFAULT_MAX_TOURS, "Max number of different tours confirmed per day (= number of available instructors)."]);

  var rs = ss.getSheetByName(SHEET_NAME);
  if (rs) ensureRequestSheet_(rs);

  // Plan dropdowns are rebuilt every time (after adding a plan, update PLAN_LIST in Code.gs)
  r.getRange("A2:A500").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([ALL_PLANS].concat(PLAN_LIST), true).setAllowInvalid(false).build());
  c.getRange("A2:A200").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([GLOBAL_ROW, MAX_TOURS_KEY].concat(PLAN_LIST), true).setAllowInvalid(false).build());
  return ss;
}

// Plan ID -> "ID｜name" (same text as the dropdown)
function planLabel_(id) {
  for (var i = 0; i < PLAN_LIST.length; i++) if (PLAN_LIST[i].split("｜")[0] === id) return PLAN_LIST[i];
  return id || "";
}

// Column E "Tour" -> plan ID. "ID｜name" gives the ID; old records (name only) are matched by name.
function planFromTour_(v) {
  var s = String(v || "").trim();
  if (!s) return "";
  if (s.indexOf("｜") !== -1) return s.split("｜")[0];
  for (var i = 0; i < PLAN_LIST.length; i++) {
    var p = PLAN_LIST[i].split("｜");
    if (s === p[0] || s === p[1]) return p[0];
  }
  return "";
}

// "ID｜name" -> ID. Both "All plans"/"Global (all plans)" (and the old Japanese labels) are mapped to the tokens the website uses.
function planId_(v) {
  var s = String(v || "").trim();
  if (s === ALL_PLANS || s === "全プラン") return "全プラン";
  if (s === GLOBAL_ROW || s === "全体") return "全体";
  if (s === MAX_TOURS_KEY || s === LEGACY_MAX_TOURS_KEY) return MAX_TOURS_KEY;
  return s.split("｜")[0];
}

function dayStatus_(v) {
  var s = String(v || "").trim();
  return DAY_STATUS_MAP[s] || s;   // Open/Full/Closed -> 受付中/満席/停止 (the website uses the Japanese words)
}

function dateKey_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  var m = String(v || "").trim().match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (!m) return "";
  return m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2);
}

function num_(v) {
  if (v === "" || v == null) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function getConfig_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get("config");
  if (hit) return JSON.parse(hit);

  var ss = ensureSheets_();
  var tz = ss.getSpreadsheetTimeZone();

  var rules = [];
  var rv = ss.getSheetByName(RULE_SHEET).getDataRange().getValues();
  for (var i = 1; i < rv.length; i++) {
    var row = rv[i];
    var plan = planId_(row[0]);
    var start = dateKey_(row[1], tz), end = dateKey_(row[2], tz) || start;
    if (!plan || !start) continue;
    rules.push({ plan: plan, start: start, end: end, adult: num_(row[3]), child: num_(row[4]), status: dayStatus_(row[5]) });
  }

  var cutoff = [], maxTours = DEFAULT_MAX_TOURS;
  var cv = ss.getSheetByName(CUTOFF_SHEET).getDataRange().getValues();
  for (var j = 1; j < cv.length; j++) {
    var p = planId_(cv[j][0]), d = num_(cv[j][1]);
    if (p === MAX_TOURS_KEY) { if (d !== null && d > 0) maxTours = d; continue; }
    if (p && d !== null) cutoff.push({ plan: p, days: d });
  }

  // 確定済みの予約から、日ごとの「確定しているツアーの種類」を数える。
  // 上限に達した日だけ { 日付: [確定しているプランID...] } で返す(他のプランはその日が選べなくなる)
  var taken = {}, byDate = {};
  var rs = ss.getSheetByName(SHEET_NAME);
  if (rs && rs.getLastRow() > 1) {
    var qv = rs.getDataRange().getValues();
    for (var q = 1; q < qv.length; q++) {
      var stt = String(qv[q][8]).trim();
      if (stt !== "Confirmed" && stt !== "確定") continue;
      var dk = dateKey_(qv[q][5], tz), pid = planFromTour_(qv[q][4]);
      if (!dk || !pid) continue;
      (byDate[dk] = byDate[dk] || {})[pid] = true;
    }
    Object.keys(byDate).forEach(function (k) {
      var ids = Object.keys(byDate[k]);
      if (ids.length >= maxTours) taken[k] = ids;
    });
  }

  var cfg = { ok: true, rules: rules, cutoff: cutoff, taken: taken, maxTours: maxTours };
  cache.put("config", JSON.stringify(cfg), 60); // 60秒キャッシュ(編集の反映は最大1分遅れる)
  return cfg;
}

// Run this from the Apps Script editor (pick "runSetup" in the function list) to set up / migrate the sheets.
// If something is wrong, the error shows up in the execution log.
function runSetup() {
  CacheService.getScriptCache().remove("config");
  ensureSheets_();
  Logger.log("Setup done.");
}

// 動作確認用(ブラウザでURLを開いたときの表示) / サイトからの設定取得
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "config") {
    try { return out_(getConfig_()); } catch (err) { return out_({ ok: false, error: String(err && err.message || err), line: err && err.lineNumber }); }
  }
  return ContentService.createTextOutput("All Blue reservation endpoint is running.");
}

function clean_(v, max) {
  return String(v == null ? "" : v).replace(/\r/g, "").trim().slice(0, max);
}

// スプレッドシートで数式として実行されないようにする
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
