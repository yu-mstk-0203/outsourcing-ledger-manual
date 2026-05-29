(function(){

  /* ── URLコピーボタン ── */
  document.querySelectorAll(".copy-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      const text = button.getAttribute("data-copy") || "";
      const status = button.parentElement.querySelector(".copy-status");
      try{
        await navigator.clipboard.writeText(text);
        if(status){
          status.textContent = "コピーしました";
          window.setTimeout(() => { status.textContent = ""; }, 1800);
        }
      }catch(error){
        if(status){
          status.textContent = "コピーできない場合はURLを選択してコピーしてください";
        }
      }
    });
  });

  /* ── PDF保存ボタン → ガイドモーダルを表示 ── */
  const printBtn = document.getElementById("printBtn");
  const pdfGuideDialog = document.getElementById("dlg-pdf-guide");
  const pdfPrintTrigger = document.getElementById("pdfPrintTrigger");
  if(printBtn && pdfGuideDialog){
    printBtn.addEventListener("click", () => pdfGuideDialog.showModal());
  }
  if(pdfPrintTrigger){
    pdfPrintTrigger.addEventListener("click", () => {
      if(pdfGuideDialog) pdfGuideDialog.close();
      window.print();
    });
  }

  /* ── 先頭へボタン ── */
  const topBtn = document.getElementById("topBtn");
  if(topBtn){
    topBtn.addEventListener("click", () => window.scrollTo({top:0, behavior:"smooth"}));
  }

  /* ── Word(.doc)ダウンロード ── */
  const wordBtn = document.getElementById("wordBtn");
  if(wordBtn){
    wordBtn.addEventListener("click", downloadAsWord);
  }

  function downloadAsWord(){
    const contentEl = document.querySelector(".content");
    if(!contentEl) return;

    const clone = contentEl.cloneNode(true);

    /* 印刷・画面専用要素を削除 */
    clone.querySelectorAll(
      ".copy-btn, .copy-status, .fab-wrap, .screen-only, button, .jump-row, dialog, .search-overlay"
    ).forEach(el => el.remove());

    /* ルートバッジのdata-modal属性を除去（Word内でリンク扱いにならないよう） */
    clone.querySelectorAll("[data-modal]").forEach(el => el.removeAttribute("data-modal"));

    const wordStyle = `
      body {
        font-family: 'Yu Gothic', 'Meiryo', 'Hiragino Kaku Gothic ProN', sans-serif;
        font-size: 10.5pt;
        line-height: 1.8;
        color: #17202c;
        margin: 0;
      }
      h1 {
        font-size: 18pt;
        font-weight: bold;
        text-align: center;
        margin-bottom: 18pt;
        color: #fff;
        background: #18324f;
        padding: 14pt 16pt;
        mso-element: para-border-div;
      }
      h2 {
        font-size: 13pt;
        font-weight: bold;
        background: #18324f;
        color: #fff;
        padding: 8pt 12pt;
        margin: 18pt 0 8pt;
        page-break-before: always;
        mso-element: para-border-div;
      }
      h2:first-of-type { page-break-before: avoid; }
      h3 {
        font-size: 11.5pt;
        font-weight: bold;
        border-bottom: 1.5pt solid #18324f;
        padding-bottom: 3pt;
        color: #18324f;
        margin: 14pt 0 6pt;
      }
      h4, h5 {
        font-size: 10.5pt;
        font-weight: bold;
        color: #23344a;
        margin: 10pt 0 4pt;
      }
      p { margin: 4pt 0; }
      table {
        border-collapse: collapse;
        width: 100%;
        margin: 8pt 0;
        font-size: 9.5pt;
      }
      td, th {
        border: 1pt solid #aeb8c7;
        padding: 4pt 6pt;
        vertical-align: top;
      }
      th {
        background: #edf1f5;
        font-weight: bold;
        color: #18324f;
        mso-element: para-border-div;
      }
      ul, ol { padding-left: 18pt; margin: 5pt 0; }
      li { margin-bottom: 2pt; line-height: 1.7; }
      .note {
        border-left: 4pt solid #18324f;
        background: #fafafa;
        padding: 6pt 10pt;
        margin: 7pt 0;
      }
      /* 色を使わず左罫線スタイルで種別を差別化（ネイビー基調） */
      .note.warn { border-left-style: double; border-left-width: 6pt; }
      .note.check { border-left-style: dotted; border-left-width: 5pt; }
      .note.example { border-left-style: groove; border-left-width: 6pt; }
      .note strong { display: block; color: #18324f; margin-bottom: 2pt; font-weight: bold; }
      .route {
        display: inline;
        font-weight: bold;
        padding: 1pt 5pt;
        border-radius: 3pt;
      }
      .route.a { color: #fff; background: #18324f; border: 1pt solid #18324f; }
      .route.b { color: #18324f; border: 1pt solid #18324f; }
      .route.c { color: #18324f; border: 1pt dotted #18324f; }
      .steps {
        list-style: decimal;
        padding-left: 20pt;
        margin: 6pt 0;
      }
      .steps li { margin: 4pt 0; }
      .st-title { font-weight: bold; color: #18324f; }
      .chapter-summary {
        border: 1pt solid #d8dee7;
        padding: 6pt 10pt;
        margin: 6pt 0;
        background: #f7f8fa;
      }
      .pill {
        display: inline;
        border: 1pt solid #aaa;
        padding: 1pt 5pt;
        font-size: 9pt;
        border-radius: 3pt;
      }
      /* 非表示要素 */
      .toc, .fab-wrap, .search-overlay, .screen-only,
      .callout, .mock-screen, .shot-points { display: none !important; }
      img { max-width: 100%; }
      figure { margin: 8pt 0; page-break-inside: avoid; }
      figcaption { font-size: 9pt; color: #334155; margin-top: 3pt; }
    `;

    const htmlContent = `<!DOCTYPE html>
<html xmlns:o='urn:schemas-microsoft-com:office:office'
      xmlns:w='urn:schemas-microsoft-com:office:word'
      xmlns='http://www.w3.org/TR/REC-html40'
      lang='ja'>
<head>
<meta charset='UTF-8'>
<!--[if gte mso 9]><xml>
<w:WordDocument>
  <w:View>Normal</w:View>
  <w:Zoom>90</w:Zoom>
  <w:DoNotOptimizeForBrowser/>
</w:WordDocument>
</xml><![endif]-->
<style>${wordStyle}</style>
</head>
<body>
${clone.innerHTML}
</body>
</html>`;

    const blob = new Blob(["﻿", htmlContent], {type: "application/msword"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "業務委託台帳_入力マニュアル.doc";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  /* ── 検索機能 ── */
  const overlay = document.getElementById("searchOverlay");
  const searchInput = document.getElementById("searchInput");
  const searchCount = document.getElementById("searchCount");
  const searchScope = document.querySelector("main.content");
  let searchHits = [];
  let hitIndex = -1;

  const openSearch = () => {
    if(!overlay) return;
    overlay.classList.add("open");
    if(searchInput){ searchInput.focus(); searchInput.select(); }
  };
  const closeSearch = () => {
    if(!overlay) return;
    overlay.classList.remove("open");
    clearHighlights();
    if(searchInput) searchInput.value = "";
    if(searchCount) searchCount.textContent = "";
  };

  const clearHighlights = () => {
    if(!searchScope) return;
    searchScope.querySelectorAll("mark.search-hit").forEach((mark) => {
      const parent = mark.parentNode;
      if(!parent) return;
      parent.replaceChild(document.createTextNode(mark.textContent || ""), mark);
      parent.normalize();
    });
    searchHits = [];
    hitIndex = -1;
  };

  const updateCount = () => {
    if(!searchCount) return;
    if(searchHits.length){
      searchCount.textContent = (hitIndex + 1) + " / " + searchHits.length;
    } else if(searchInput && searchInput.value.trim()){
      searchCount.textContent = "0 件";
    } else {
      searchCount.textContent = "";
    }
  };

  const jumpToHit = (index) => {
    if(!searchHits.length) return;
    hitIndex = (index + searchHits.length) % searchHits.length;
    searchHits.forEach((item) => item.classList.remove("current-hit"));
    searchHits[hitIndex].classList.add("current-hit");
    searchHits[hitIndex].scrollIntoView({behavior:"smooth", block:"center"});
    updateCount();
  };

  const runSearch = () => {
    if(!searchScope || !searchInput) return;
    const query = searchInput.value.trim();
    clearHighlights();
    if(!query){ updateCount(); return; }
    const regex = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    const walker = document.createTreeWalker(searchScope, NodeFilter.SHOW_TEXT, {
      acceptNode(node){
        if(!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        const p = node.parentElement;
        if(!p) return NodeFilter.FILTER_REJECT;
        if(["SCRIPT","STYLE","MARK"].includes(p.tagName)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    const targets = [];
    while(walker.nextNode()) targets.push(walker.currentNode);
    targets.forEach((node) => {
      const text = node.nodeValue || "";
      let matched = false;
      const frag = document.createDocumentFragment();
      let lastIdx = 0;
      text.replace(regex, (m, offset) => {
        matched = true;
        frag.appendChild(document.createTextNode(text.slice(lastIdx, offset)));
        const mark = document.createElement("mark");
        mark.className = "search-hit";
        mark.textContent = m;
        frag.appendChild(mark);
        searchHits.push(mark);
        lastIdx = offset + m.length;
        return m;
      });
      if(matched){
        frag.appendChild(document.createTextNode(text.slice(lastIdx)));
        if(node.parentNode) node.parentNode.replaceChild(frag, node);
      }
    });
    hitIndex = 0;
    if(searchHits.length) jumpToHit(0); else updateCount();
  };

  if(searchInput){
    searchInput.addEventListener("input", runSearch);
    searchInput.addEventListener("keydown", (e) => {
      if(e.key === "Enter"){ e.preventDefault(); jumpToHit(e.shiftKey ? hitIndex - 1 : hitIndex + 1); }
      if(e.key === "Escape"){ e.preventDefault(); closeSearch(); }
    });
  }
  const prevBtn = document.getElementById("searchPrevBtn");
  const nextBtn = document.getElementById("searchNextBtn");
  const closeBtn = document.getElementById("searchCloseBtn");
  const searchFab = document.getElementById("searchFab");
  if(prevBtn) prevBtn.addEventListener("click", () => jumpToHit(hitIndex - 1));
  if(nextBtn) nextBtn.addEventListener("click", () => jumpToHit(hitIndex + 1));
  if(closeBtn) closeBtn.addEventListener("click", closeSearch);
  if(searchFab) searchFab.addEventListener("click", openSearch);

  document.addEventListener("keydown", (e) => {
    if((e.ctrlKey || e.metaKey) && e.key === "f"){
      e.preventDefault();
      openSearch();
    }
    if(e.key === "Escape" && overlay && overlay.classList.contains("open")){
      closeSearch();
    }
  });

  /* ── 目次アクティブリンク ── */
  const links = Array.from(document.querySelectorAll(".toc a"));
  const sections = links
    .map((link) => document.querySelector(link.getAttribute("href")))
    .filter(Boolean);
  if("IntersectionObserver" in window){
    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if(entry.isIntersecting){
          links.forEach((link) => link.classList.remove("active"));
          const active = links.find((link) => link.hash === "#" + entry.target.id);
          if(active){
            active.classList.add("active");
          }
        }
      });
    }, {rootMargin:"-25% 0px -65% 0px", threshold:0.01});
    sections.forEach((section) => observer.observe(section));
  }

  /* ── ルートモーダル ── */
  document.querySelectorAll("[data-modal]").forEach((el) => {
    el.addEventListener("click", (e) => {
      e.preventDefault();
      const dlg = document.getElementById(el.dataset.modal);
      if(dlg) dlg.showModal();
    });
  });

  /* dialogのbackdropクリックで閉じる */
  document.querySelectorAll("dialog").forEach((dlg) => {
    dlg.addEventListener("click", (e) => {
      if(e.target === dlg) dlg.close();
    });
  });

  /* .modal-close ボタン */
  document.querySelectorAll(".modal-close").forEach((btn) => {
    btn.addEventListener("click", () => {
      const dlg = btn.closest("dialog");
      if(dlg) dlg.close();
    });
  });

  /* ダイアログ内の「詳細を見る」リンクはダイアログを閉じてからジャンプ */
  document.querySelectorAll("dialog .modal-link[href]").forEach((link) => {
    link.addEventListener("click", () => {
      const dlg = link.closest("dialog");
      if(dlg) dlg.close();
    });
  });

  /* ── 読書進捗バー ── */
  const progressBar = document.querySelector(".progress-bar");
  if(progressBar){
    const updateProgress = () => {
      const scrollable = document.documentElement.scrollHeight - window.innerHeight;
      const pct = scrollable > 0 ? (window.scrollY / scrollable) * 100 : 0;
      progressBar.style.width = Math.min(pct, 100) + "%";
    };
    window.addEventListener("scroll", updateProgress, {passive:true});
    updateProgress();
  }

})();
