/* 画像スライダー(フェード/自動切替/ドット/矢印/スワイプ)
   使い方: コンテナ内に <img class="slide"> を複数置き、initSlider(コンテナ) を呼ぶ。
   画像が1枚なら矢印・ドットは出ず、そのまま表示される。 */
(function () {
  var css = document.createElement("style");
  css.textContent =
    ".slide{position:absolute;top:0;left:0;width:100%;height:100%;object-fit:cover;opacity:0;transition:opacity .9s ease;}" +
    ".slide.active{opacity:1;}" +
    ".sl-arrow{position:absolute;top:50%;transform:translateY(-50%);z-index:5;width:42px;height:42px;border:none;border-radius:50%;background:rgba(0,0,0,.35);color:#fff;font-size:20px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px);transition:background .2s;}" +
    ".sl-arrow:hover{background:rgba(0,0,0,.6);}" +
    ".sl-prev{left:14px;}.sl-next{right:14px;}" +
    ".sl-dots{position:absolute;left:0;right:0;bottom:14px;z-index:5;display:flex;justify-content:center;gap:8px;}" +
    ".sl-dot{width:9px;height:9px;padding:0;border:none;border-radius:50%;background:rgba(255,255,255,.5);cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,.4);}" +
    ".sl-dot.active{background:#fff;transform:scale(1.25);}" +
    "@media (max-width:860px){.sl-arrow{width:34px;height:34px;font-size:16px;}.sl-prev{left:8px;}.sl-next{right:8px;}}";
  document.head.appendChild(css);

  window.initSlider = function (box, interval) {
    var slides = box.querySelectorAll(".slide");
    if (!slides.length) return;
    var cur = 0, timer = null, ms = interval || 5000;
    slides[0].classList.add("active");
    if (slides.length < 2) return;

    var prev = document.createElement("button");
    prev.className = "sl-arrow sl-prev"; prev.type = "button"; prev.setAttribute("aria-label", "前の画像"); prev.textContent = "‹";
    var next = document.createElement("button");
    next.className = "sl-arrow sl-next"; next.type = "button"; next.setAttribute("aria-label", "次の画像"); next.textContent = "›";
    var dotsBox = document.createElement("div");
    dotsBox.className = "sl-dots";
    var dots = [];
    for (var i = 0; i < slides.length; i++) {
      (function (n) {
        var d = document.createElement("button");
        d.className = "sl-dot"; d.type = "button"; d.setAttribute("aria-label", (n + 1) + "枚目");
        d.addEventListener("click", function () { go(n, true); });
        dotsBox.appendChild(d); dots.push(d);
      })(i);
    }
    box.appendChild(prev); box.appendChild(next); box.appendChild(dotsBox);

    function show(n) {
      slides[cur].classList.remove("active"); dots[cur].classList.remove("active");
      cur = (n + slides.length) % slides.length;
      slides[cur].classList.add("active"); dots[cur].classList.add("active");
    }
    function start() { stop(); timer = setInterval(function () { show(cur + 1); }, ms); }
    function stop() { if (timer) { clearInterval(timer); timer = null; } }
    function go(n, manual) { show(n); if (manual) start(); }
    dots[0].classList.add("active");

    prev.addEventListener("click", function () { go(cur - 1, true); });
    next.addEventListener("click", function () { go(cur + 1, true); });
    box.addEventListener("mouseenter", stop);
    box.addEventListener("mouseleave", start);

    var sx = null;
    box.addEventListener("touchstart", function (e) { sx = e.touches[0].clientX; }, { passive: true });
    box.addEventListener("touchend", function (e) {
      if (sx === null) return;
      var dx = e.changedTouches[0].clientX - sx; sx = null;
      if (Math.abs(dx) > 40) go(dx < 0 ? cur + 1 : cur - 1, true);
    }, { passive: true });

    start();
  };
})();
