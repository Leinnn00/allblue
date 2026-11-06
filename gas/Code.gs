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
  // LINE の Webhook (JSON) はこちらへ。Webサイトのフォーム送信(URLエンコード)は下の処理へ。
  if (e && e.postData && /json/i.test(String(e.postData.type || ""))) return handleLine_(e);

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
    var agree = d.agree === "yes" ? "Yes" : "";
    var summary = clean_(d.summary, 300);
    var note = clean_(d.note, 1000);

    // 1) スプレッドシートに記録
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    migrate_(ss);
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(REQ_HEADERS);
    sh.appendRow(reqRow_({ source: "Web form", name: safe_(name), email: safe_(email), phone: safe_(tel), tour: safe_(planLabel_(planId) || tour), date: safe_(date), details: safe_(summary), note: safe_(note),
      policy: agree ? "Yes (" + Utilities.formatDate(new Date(), "Asia/Manila", "yyyy/MM/dd HH:mm") + " PHT)" : "" }));
    ensureRequestSheet_(sh);

    var detail =
      "ツアー：" + tour + "\n" +
      "希望日：" + date + "\n" +
      "内容：" + summary + "\n\n" +
      "お名前：" + name + "\n" +
      "メール：" + email + "\n" +
      "電話：" + (tel || "(未入力)") + "\n" +
      "ご要望：" + (note || "(なし)") + "\n" +
      "お支払い・キャンセル(キャンセル料20%)への同意：" + (agree ? "同意済み" : "(未確認)");

    // 2) スタッフへ通知(Webフォーム分は即時。LINE分は毎日のまとめメール sendDigest_ のみ)
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
var CAP_SHEET = "Daily Capacity";   // 期間ごとの「1日に受けられるツアー種類数」の上書き
var STATUS_LIST = ["New", "In progress", "Confirmed", "Cancelled"];
var DAY_STATUS_LIST = ["Open", "Full", "Closed"];
// English <-> Japanese values (the website uses the Japanese ones)
var DAY_STATUS_MAP = { "Open": "受付中", "Full": "満席", "Closed": "停止" };
var LEGACY_REQ_STATUS = { "未対応": "New", "対応中": "In progress", "確定": "Confirmed", "キャンセル": "Cancelled" };
var LEGACY_DAY_STATUS = { "受付中": "Open", "満席": "Full", "停止": "Closed" };
var CUTOFF_SHEET = "Booking Settings";
var LEGACY_CUTOFF_SHEET = "受付設定";

// Booking Requests の列の並び(左から)。スタッフが毎日触る列を左に集めた
var REQ_HEADERS = ["Status", "Source", "Received", "Name", "Tour", "Desired date", "Details", "Request / notes", "Email", "Phone", "LINE userId", "Confirmation email", "Policy agreed"];
// 配列(0始まり)の添字
var C_STATUS = 0, C_SOURCE = 1, C_RECEIVED = 2, C_NAME = 3, C_TOUR = 4, C_DATE = 5, C_DETAILS = 6, C_NOTE = 7, C_EMAIL = 8, C_PHONE = 9, C_LINEID = 10, C_CONFIRM = 11, C_POLICY = 12;
// 旧レイアウト(Received,Name,Email,Phone,Tour,Date,Details,Note,Status,Policy,Confirm,Source,LINE)から新レイアウトへの対応
var OLD_TO_NEW = [8, 11, 0, 1, 4, 5, 6, 7, 2, 3, 12, 10, 9];

function reqRow_(o) {
  var r = new Array(REQ_HEADERS.length);
  for (var i = 0; i < r.length; i++) r[i] = "";
  r[C_STATUS] = o.status || "New"; r[C_SOURCE] = o.source || ""; r[C_RECEIVED] = o.received || new Date();
  r[C_NAME] = o.name || ""; r[C_TOUR] = o.tour || ""; r[C_DATE] = o.date || ""; r[C_DETAILS] = o.details || ""; r[C_NOTE] = o.note || "";
  r[C_EMAIL] = o.email || ""; r[C_PHONE] = o.phone || ""; r[C_LINEID] = o.lineId || ""; r[C_POLICY] = o.policy || "";
  return r;
}

// 旧(日本語)のシート名・見出し・値を英語に移す。何度実行しても安全
function migrate_(ss) {
  var pairs = [[LEGACY_SHEET_NAME, SHEET_NAME], [LEGACY_RULE_SHEET, RULE_SHEET], [LEGACY_CUTOFF_SHEET, CUTOFF_SHEET]];
  pairs.forEach(function (p) {
    var old = ss.getSheetByName(p[0]);
    if (old && !ss.getSheetByName(p[1])) old.setName(p[1]);
  });

  // 旧ドロップダウン(日本語の選択肢)が残っていると、英語の値を書き込めないので、先に外す(あとで作り直す)
  var rs0 = ss.getSheetByName(SHEET_NAME), r0 = ss.getSheetByName(RULE_SHEET), c0 = ss.getSheetByName(CUTOFF_SHEET);
  if (rs0) rs0.getRange("A2:M1000").clearDataValidations();
  if (r0) r0.getRange("A2:G500").clearDataValidations();
  if (c0) c0.getRange("A2:C200").clearDataValidations();

  var rs = ss.getSheetByName(SHEET_NAME);
  if (rs && rs.getLastRow() >= 1) {
    var n = REQ_HEADERS.length;
    if (rs.getMaxColumns() < n) rs.insertColumnsAfter(rs.getMaxColumns(), n - rs.getMaxColumns());
    // 旧レイアウト(1列目が Status ではない)なら、列を並び替える
    if (String(rs.getRange(1, 1).getValue()).trim() !== "Status") {
      var last = rs.getLastRow();
      if (last > 1) {
        var old = rs.getRange(2, 1, last - 1, n).getValues();
        var moved = old.map(function (row) { return OLD_TO_NEW.map(function (i) { return row[i]; }); });
        rs.getRange(2, 1, last - 1, n).clearContent().setValues(moved);
      }
    }
    rs.getRange(1, 1, 1, n).setValues([REQ_HEADERS]);
    if (rs.getLastRow() > 1) {
      var sr = rs.getRange(2, C_STATUS + 1, rs.getLastRow() - 1, 1), sv = sr.getValues();
      sr.setValues(sv.map(function (r) { var v = String(r[0]).trim(); return [LEGACY_REQ_STATUS[v] || r[0]]; }));
    }
    rs.getRange(1, 1, 1, n).setFontWeight("bold");
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
  sh.getRange("A2:A1000").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_LIST, true).setAllowInvalid(true).build());
  // Column E "Tour" is used to count capacity. When you add a LINE booking by hand, pick the tour from this dropdown.
  sh.getRange("E2:E1000").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(PLAN_LIST, true).setAllowInvalid(true).build());
  sh.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);   // Status, Source, Received, Name を固定
  sh.getRange("C2:C1000").setNumberFormat("yyyy/mm/dd hh:mm");
  var widths = [100, 80, 130, 150, 300, 110, 260, 220, 200, 120, 120, 170, 200];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  // Status の色分け
  var rg = sh.getRange("A2:A1000");
  function cf(t, bg) { return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setRanges([rg]).build(); }
  sh.setConditionalFormatRules([cf("New", "#FFF2CC"), cf("In progress", "#DDEBF7"), cf("Confirmed", "#D9EAD3"), cf("Cancelled", "#E7E6E6")]);
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

  var cap = ss.getSheetByName(CAP_SHEET);
  if (!cap) {
    cap = ss.insertSheet(CAP_SHEET);
    cap.appendRow(["Start date", "End date (blank = same day)", "Max tour types per day", "Notes"]);
    cap.setFrozenRows(1);
    cap.getRange("A2:B500").setNumberFormat("yyyy/mm/dd");
    cap.setColumnWidth(1, 120); cap.setColumnWidth(2, 190); cap.setColumnWidth(3, 170); cap.setColumnWidth(4, 360);
    cap.getRange("A1:D1").setFontWeight("bold");
    cap.getRange("D2").setValue("Example: instructor on leave. Fill Start/End/Max in a row. 0 = no tours that day. Later rows win when ranges overlap.");
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
  // Daily Capacity: 期間ごとの上限の上書き(重なったら下の行が優先)
  var capRules = [];
  var capSh = ss.getSheetByName(CAP_SHEET);
  if (capSh && capSh.getLastRow() > 1) {
    var cpv = capSh.getRange(2, 1, capSh.getLastRow() - 1, 3).getValues();
    cpv.forEach(function (x) {
      var a = dateKey_(x[0], tz), b = dateKey_(x[1], tz) || a, m = num_(x[2]);
      if (a && m !== null && m >= 0) capRules.push({ start: a, end: b, max: m });
    });
  }
  function limitFor_(dk) {
    var lim = maxTours;
    capRules.forEach(function (c) { if (dk >= c.start && dk <= c.end) lim = c.max; });
    return lim;
  }
  var taken = {}, byDate = {};
  // 上限0の日は、確定が無くても全プラン満席にする(最大400日分)
  capRules.forEach(function (c) {
    if (c.max !== 0) return;
    var d = new Date(c.start + "T00:00:00Z"), e = new Date(c.end + "T00:00:00Z"), n = 0;
    while (d <= e && n++ < 400) {
      var k0 = d.toISOString().slice(0, 10);
      if (limitFor_(k0) === 0) taken[k0] = [];
      d.setUTCDate(d.getUTCDate() + 1);
    }
  });
  var rs = ss.getSheetByName(SHEET_NAME);
  if (rs && rs.getLastRow() > 1) {
    var qv = rs.getDataRange().getValues();
    for (var q = 1; q < qv.length; q++) {
      var stt = String(qv[q][C_STATUS]).trim();
      if (stt !== "Confirmed" && stt !== "確定") continue;
      var dk = dateKey_(qv[q][C_DATE], tz), pid = planFromTour_(qv[q][C_TOUR]);
      if (!dk || !pid) continue;
      (byDate[dk] = byDate[dk] || {})[pid] = true;
    }
    Object.keys(byDate).forEach(function (k) {
      var ids = Object.keys(byDate[k]);
      if (ids.length >= limitFor_(k)) taken[k] = ids;
    });
  }

  var cfg = { ok: true, rules: rules, cutoff: cutoff, taken: taken, maxTours: maxTours };
  cache.put("config", JSON.stringify(cfg), 60); // 60秒キャッシュ(編集の反映は最大1分遅れる)
  return cfg;
}


/* ================= 予約確定メール(Status を Confirmed にしたら自動送信) =================
 * シート「Booking Requests」の Status(I列)を Confirmed にすると、お客様(Email列)へ予約確定メールを送る。
 * 送ったら K列「Confirmation email」に日時が入り、同じ行には二度送らない。
 * メールアドレスが空の行(LINE予約を手で足した行)には送らず、K列に "No email" と入る。
 * ※ 手で編集したときだけ動く「インストール型トリガー」が必要。runSetup を一度実行すると登録される。
 */
function onStatusEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    if (e.range.getColumn() > C_STATUS + 1 || e.range.getLastColumn() < C_STATUS + 1) return;
    var first = e.range.getRow(), last = e.range.getLastRow();
    for (var r = Math.max(first, 2); r <= last; r++) {
      var v = sh.getRange(r, 1, 1, 13).getValues()[0];
      var st = String(v[C_STATUS]).trim();
      if (st !== "Confirmed" && st !== "確定") continue;
      if (v[C_CONFIRM]) continue;   // すでに送信済み
      var email = String(v[C_EMAIL]).trim();
      var stamp = Utilities.formatDate(new Date(), "Asia/Manila", "yyyy/MM/dd HH:mm") + " PHT";
      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        sendConfirmMail_(v, email);
        sh.getRange(r, C_CONFIRM + 1).setValue("Sent (email) " + stamp);
      } else if (String(v[C_LINEID]).trim() && LINE_PUSH_ON_CONFIRM) {
        var ok = linePush_(String(v[C_LINEID]).trim(), confirmText_(v, "line"));
        sh.getRange(r, C_CONFIRM + 1).setValue(ok ? "Sent (LINE) " + stamp : "LINE push failed");
      } else {
        sh.getRange(r, C_CONFIRM + 1).setValue("No email");
      }
    }
  } catch (err) {
    console.error(err);
  }
}

// 確定メッセージの本文(メール用 / LINE用)
function confirmText_(v, kind) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var name = String(v[C_NAME]);
  var tour = String(v[C_TOUR]).split("｜").slice(1).join("｜") || String(v[C_TOUR]);
  var date = v[C_DATE] instanceof Date ? Utilities.formatDate(v[C_DATE], tz, "yyyy/MM/dd") : String(v[C_DATE]);
  var reply = kind === "line" ? "このトークにご返信" : "このメールへの返信(またはLINE)";
  return (
    name + " 様\n\n" +
    "この度はお申込みありがとうございます。以下の内容で、ご予約が確定しました。\n\n" +
    "ツアー：" + tour + "\n" +
    "ご参加日：" + date + "\n" +
    "内容：" + String(v[C_DETAILS]) + "\n\n" +
    "【お支払い】\n" +
    "ツアー当日に、現地でお支払いください(予約金はいただきません)。\n\n" +
    "【キャンセルについて】\n" +
    "確定後にお客様のご都合でキャンセルされる場合は、キャンセル料としてツアー代金の20%をいただきます。\n" +
    "悪天候・海況など当社の都合でツアーを中止する場合は、キャンセル料はいただきません。\n\n" +
    "【お迎え・集合について】\n" +
    "ご宿泊先のホテル名と、お迎えの詳細を、" + reply + "でお知らせください。\n" +
    "前日までに、集合時間などの詳細をご連絡します。\n\n" +
    "当日お会いできるのを楽しみにしています。\n" +
    SHOP_NAME
  );
}

function sendConfirmMail_(v, email) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var date = v[C_DATE] instanceof Date ? Utilities.formatDate(v[C_DATE], tz, "yyyy/MM/dd") : String(v[C_DATE]);
  MailApp.sendEmail({
    to: email,
    subject: "【" + SHOP_NAME + "】ご予約が確定しました(" + date + ")",
    body: confirmText_(v, "mail") + "\n\nご不明な点があれば、このメールにご返信ください。",
    name: SHOP_NAME,
    replyTo: NOTIFY_EMAIL
  });
}

// インストール型の編集トリガーを登録する(なければ)
function installTriggers_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "onStatusEdit"; });
  if (!has) ScriptApp.newTrigger("onStatusEdit").forSpreadsheet(ss).onEdit().create();
  var hasD = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "sendDigest_"; });
  if (!hasD) ScriptApp.newTrigger("sendDigest_").timeBased().everyDays(1).atHour(DIGEST_HOUR).inTimezone("Asia/Manila").create();
}

// 毎日 DIGEST_HOUR 時(フィリピン時間)に、Status が New の行をまとめてスタッフへメール(なければ送らない)
var DIGEST_HOUR = 18;
function sendDigest_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return;
  var tz = ss.getSpreadsheetTimeZone();
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
  var lines = [];
  rows.forEach(function (v) {
    if (String(v[C_STATUS]).trim() !== "New") return;
    var rec = v[C_RECEIVED] instanceof Date ? Utilities.formatDate(v[C_RECEIVED], tz, "MM/dd HH:mm") : String(v[C_RECEIVED]);
    lines.push("・[" + (v[C_SOURCE] || "?") + "] " + rec + " " + v[C_NAME] + "様 / " + v[C_TOUR] + " / 希望日 " + v[C_DATE] + " / " + v[C_DETAILS]);
  });
  if (!lines.length) return;
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: "【未対応の予約リクエスト】" + lines.length + "件",
    body: "Status が New のままの予約リクエストです。\n\n" + lines.join("\n") +
          "\n\nシート「" + SHEET_NAME + "」で確認し、対応したら Status を In progress / Confirmed / Cancelled に変えてください。\n" + ss.getUrl()
  });
}

/* ================= LINE公式アカウント連携(Webhook) =================
 * サイトの「この内容でLINEに送る」から届いた予約メッセージを、
 *   1) シート「Booking Requests」に1行追加(Status=New, Source=LINE)
 *   2) お客様に受付の自動返信
 * する。Status を Confirmed にすると、メールが無い行にはLINEで確定メッセージを送る(LINE_PUSH_ON_CONFIRM)。
 * 必要な設定(スクリプトのプロパティ): LINE_CHANNEL_ACCESS_TOKEN / LINE_WEBHOOK_KEY (手順は gas/セットアップ手順.md)
 */
var LINE_PUSH_ON_CONFIRM = true;   // Confirmed にしたとき、メールが無いLINE予約にはLINEで確定メッセージを送る
var LINE_MARK = "Webからのご予約リクエスト";   // サイトが作る予約メッセージの1行目に入っている言葉(apply.html と揃える)

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ""; }

function handleLine_(e) {
  try {
    // 合言葉(WebhookのURLの ?key=...)が違うものは無視
    var key = prop_("LINE_WEBHOOK_KEY");
    if (!key || !e.parameter || e.parameter.key !== key) return out_({ ok: false });

    var body = JSON.parse(e.postData.contents || "{}");
    var events = body.events || [];
    var cache = CacheService.getScriptCache();
    events.forEach(function (ev) {
      if (ev.type !== "message" || !ev.message || ev.message.type !== "text") return;
      var text = String(ev.message.text || "");
      if (text.indexOf(LINE_MARK) === -1) return;          // 予約メッセージ以外は何もしない(スタッフが手で返信)
      var eid = ev.webhookEventId || ev.message.id;
      if (cache.get("line_" + eid)) return;                // LINEの再送で二重登録しない
      cache.put("line_" + eid, "1", 600);

      var userId = ev.source && ev.source.userId ? ev.source.userId : "";
      var name = lineProfileName_(userId) || "(LINE)";
      var req = parseLineRequest_(text);
      saveLineRequest_(name, userId, req, text);
      if (ev.replyToken) lineReply_(ev.replyToken,
        "お申込みありがとうございます。内容を受け付けました。\n" +
        "スタッフが確認して、予約確定のご連絡をします。\n" +
        "※8時から17時はツアー催行中のため、返信が遅れる場合があります。");
    });
    return out_({ ok: true });
  } catch (err) {
    console.error(err);
    return out_({ ok: false });
  }
}

// サイトが作るメッセージ(ツアー：… / 希望日：… / 内容：… / プランコード：…)を読み取る
function parseLineRequest_(text) {
  function pick(label) {
    var m = text.match(new RegExp(label + "[:：]\\s*(.+)"));
    return m ? m[1].trim() : "";
  }
  var planId = pick("プランコード");
  var tour = pick("ツアー");
  if (!planId) planId = planFromTour_(tour);
  return { planId: planId, tour: tour, date: pick("希望日"), summary: pick("内容") };
}

function saveLineRequest_(name, userId, req, rawText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrate_(ss);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(REQ_HEADERS);
  var agreed = /同意して申し込みます/.test(rawText);
  var stamp = Utilities.formatDate(new Date(), "Asia/Manila", "yyyy/MM/dd HH:mm") + " PHT";
  sh.appendRow(reqRow_({ source: "LINE", name: safe_(name), tour: safe_(planLabel_(req.planId) || req.tour), date: safe_(req.date), details: safe_(req.summary),
    lineId: safe_(userId), policy: agreed ? "Yes (LINE message " + stamp + ")" : "" }));
  ensureRequestSheet_(sh);
  // スタッフへの通知は、毎日の未対応まとめメール(sendDigest_)で行う
}

function lineApi_(method, url, payload) {
  var token = prop_("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token) return null;
  var opt = { method: method, headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true };
  if (payload) { opt.contentType = "application/json"; opt.payload = JSON.stringify(payload); }
  return UrlFetchApp.fetch(url, opt);
}

function lineProfileName_(userId) {
  if (!userId) return "";
  try {
    var r = lineApi_("get", "https://api.line.me/v2/bot/profile/" + encodeURIComponent(userId));
    if (r && r.getResponseCode() === 200) return JSON.parse(r.getContentText()).displayName || "";
  } catch (err) { console.error(err); }
  return "";
}

function lineReply_(replyToken, text) {
  try {
    lineApi_("post", "https://api.line.me/v2/bot/message/reply", { replyToken: replyToken, messages: [{ type: "text", text: text }] });
  } catch (err) { console.error(err); }
}

function linePush_(userId, text) {
  try {
    var r = lineApi_("post", "https://api.line.me/v2/bot/message/push", { to: userId, messages: [{ type: "text", text: text }] });
    return !!r && r.getResponseCode() === 200;
  } catch (err) { console.error(err); return false; }
}

// Run this from the Apps Script editor (pick "runSetup" in the function list) to set up / migrate the sheets.
// If something is wrong, the error shows up in the execution log.
function runSetup() {
  CacheService.getScriptCache().remove("config");
  ensureSheets_();
  installTriggers_();
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
