/* 画面の動きと使いやすさ(トップ・パック一覧ページ共通)
   1) スクロールに合わせて、要素がふわっと表示される
   2) スマホで、画面の下に「LINEで予約 / メールで予約」の固定バーを出す
   動きを減らす設定の端末では、1) は行わない。JSが動かない場合も、内容はそのまま表示される。 */
(function () {
  var d = document, root = d.documentElement;
  var me = d.currentScript;
  var base = (me && me.getAttribute("data-base")) || "";

  /* 1) スクロール表示 */
  var reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if ("IntersectionObserver" in window && !reduce) {
    var sel = ".sec-head,.tour-card,.rank-card,.purpose,.pillar,.staff-card,.review-card,.access-info,.access-map,.ig-grid,.about,.ctaband,.pack,.page-head,.intro";
    var els = [].slice.call(d.querySelectorAll(sel));
    root.classList.add("js-reveal");
    els.forEach(function (el) { el.classList.add("rv"); });
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting) { e.target.classList.add("rv-in"); io.unobserve(e.target); }
      });
    }, { rootMargin: "0px 0px -6% 0px", threshold: 0.06 });
    els.forEach(function (el) { io.observe(el); });
    // 念のため: 一定時間後に、まだ出ていない要素を全部表示する
    setTimeout(function () { els.forEach(function (el) { el.classList.add("rv-in"); }); }, 6000);
  }

  /* 2) スマホの固定バー */
  var bar = d.createElement("div");
  bar.className = "sticky-cta";
  bar.innerHTML =
    '<a class="sc-line" href="https://line.me/R/ti/p/@513pefsy" target="_blank" rel="noopener">LINEで予約</a>' +
    '<a class="sc-mail" href="' + base + 'apply.html">メールで予約</a>';
  d.body.appendChild(bar);
})();
