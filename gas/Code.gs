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
var SHEET_NAME = "予約リクエスト";

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
    var summary = clean_(d.summary, 300);
    var note = clean_(d.note, 1000);

    // 1) スプレッドシートに記録
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["受付日時", "お名前", "メール", "電話", "ツアー", "希望日", "内容", "ご要望", "対応状況"]);
    }
    sh.appendRow([new Date(), safe_(name), safe_(email), safe_(tel), safe_(tour), safe_(date), safe_(summary), safe_(note), "未対応"]);

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

/* ================= 料金・受付設定(スプレッドシート → サイト) =================
 * シート「料金ルール」: プラン / 開始日 / 終了日 / 大人料金 / 子供料金 / 状態 / メモ
 *   - プランは「全プラン」か、ドロップダウンから選ぶ。料金が空欄なら通常料金(plans.js)を使う。
 *   - 状態が空欄なら「受付中」。「満席」「停止」にするとその日は選べなくなる。
 *   - 複数の行が重なる日は、下の行が優先される。
 * シート「受付設定」: プラン / 何日前まで受付 / メモ
 *   - 「全体」の行が全プラン共通。プランを選んだ行は、そのプランだけ上書きする。
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
var RULE_SHEET = "料金ルール";
var CUTOFF_SHEET = "受付設定";

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var r = ss.getSheetByName(RULE_SHEET);
  if (!r) {
    r = ss.insertSheet(RULE_SHEET);
    r.appendRow(["プラン", "開始日", "終了日", "大人料金(円)", "子供料金(円)", "状態", "メモ"]);
    r.setFrozenRows(1);
    r.getRange("F2:F500").setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["受付中", "満席", "停止"], true).setAllowInvalid(false).build());
    r.getRange("B2:C500").setNumberFormat("yyyy/mm/dd");
    r.setColumnWidth(1, 360);
    r.getRange("A1:G1").setFontWeight("bold");
  }

  var c = ss.getSheetByName(CUTOFF_SHEET);
  if (!c) {
    c = ss.insertSheet(CUTOFF_SHEET);
    c.appendRow(["プラン", "何日前まで受付(1=前日まで)", "メモ"]);
    c.appendRow(["全体", 1, "全プラン共通。プラン別の行があればそちらが優先"]);
    c.setFrozenRows(1);
    c.setColumnWidth(1, 360);
    c.setColumnWidth(2, 200);
    c.getRange("A1:C1").setFontWeight("bold");
  }

  // プランのドロップダウンは毎回作り直す(プランを追加したとき、Code.gs の PLAN_LIST を更新すれば反映される)
  r.getRange("A2:A500").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["全プラン"].concat(PLAN_LIST), true).setAllowInvalid(false).build());
  c.getRange("A2:A200").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["全体"].concat(PLAN_LIST), true).setAllowInvalid(false).build());
  return ss;
}

function planId_(v) {
  var s = String(v || "").trim();
  if (s === "全プラン" || s === "全体") return s;
  return s.split("｜")[0];
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
    rules.push({ plan: plan, start: start, end: end, adult: num_(row[3]), child: num_(row[4]), status: String(row[5] || "").trim() });
  }

  var cutoff = [];
  var cv = ss.getSheetByName(CUTOFF_SHEET).getDataRange().getValues();
  for (var j = 1; j < cv.length; j++) {
    var p = planId_(cv[j][0]), d = num_(cv[j][1]);
    if (p && d !== null) cutoff.push({ plan: p, days: d });
  }

  var cfg = { ok: true, rules: rules, cutoff: cutoff };
  cache.put("config", JSON.stringify(cfg), 60); // 60秒キャッシュ(編集の反映は最大1分遅れる)
  return cfg;
}

// 動作確認用(ブラウザでURLを開いたときの表示) / サイトからの設定取得
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "config") {
    try { return out_(getConfig_()); } catch (err) { return out_({ ok: false }); }
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
