/* プラン一覧ページ用: data-plan が付いたプランを詳細ページ(plan.html)へのリンクにする */
(function () {
  var st = document.createElement("style");
  st.textContent = ".pack[data-plan]{cursor:pointer;}.pack[data-plan]:hover{box-shadow:0 20px 40px rgba(15,40,70,.12);transform:translateY(-3px);}";
  document.head.appendChild(st);
  document.querySelectorAll(".pack[data-plan]").forEach(function (p) {
    p.classList.remove("soon");
    p.onclick = function () { location.href = "../plan.html?id=" + p.getAttribute("data-plan"); };
    var b = p.querySelector(".pack-btn");
    if (b) b.textContent = "詳細を見る";
  });
})();
