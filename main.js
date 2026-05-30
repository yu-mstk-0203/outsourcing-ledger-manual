(function(){
  "use strict";

  function each(nodes, callback){
    Array.prototype.forEach.call(nodes || [], callback);
  }

  function hasClass(el, className){
    return el && (" " + el.className + " ").indexOf(" " + className + " ") > -1;
  }

  function addClass(el, className){
    if(!el) return;
    if(el.classList){
      el.classList.add(className);
    } else if(!hasClass(el, className)){
      el.className += (el.className ? " " : "") + className;
    }
  }

  function removeClass(el, className){
    if(!el) return;
    if(el.classList){
      el.classList.remove(className);
    } else {
      el.className = (" " + el.className + " ").replace(" " + className + " ", " ").replace(/^\s+|\s+$/g, "");
    }
  }

  function closest(el, selector){
    if(!el) return null;
    while(el && el.nodeType === 1){
      var matches = el.matches || el.msMatchesSelector || el.webkitMatchesSelector;
      if(matches){
        if(matches.call(el, selector)) return el;
      } else if(selector.charAt(0) === "." && hasClass(el, selector.slice(1))){
        return el;
      } else if(el.tagName && el.tagName.toLowerCase() === selector.toLowerCase()){
        return el;
      }
      el = el.parentElement;
    }
    return null;
  }

  function safeScrollIntoView(el){
    if(!el || !el.scrollIntoView) return;
    try{
      el.scrollIntoView({behavior:"smooth", block:"center"});
    }catch(error){
      el.scrollIntoView();
    }
  }

  function setStatus(status, message, clearAfter){
    if(!status) return;
    status.textContent = message;
    if(clearAfter){
      window.setTimeout(function(){ status.textContent = ""; }, clearAfter);
    }
  }

  function fallbackCopy(text){
    var textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "readonly");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    var copied = false;
    try{
      copied = document.execCommand("copy");
    }catch(error){
      copied = false;
    }
    document.body.removeChild(textarea);
    return copied;
  }

  function selectVisibleUrl(button){
    var card = closest(button, ".url-card");
    var urlText = card ? card.querySelector(".url-text") : null;
    if(!urlText) return false;

    var selection = window.getSelection ? window.getSelection() : null;
    if(document.createRange && selection){
      var range = document.createRange();
      range.selectNodeContents(urlText);
      selection.removeAllRanges();
      selection.addRange(range);
      return true;
    }
    if(document.body.createTextRange){
      var textRange = document.body.createTextRange();
      textRange.moveToElementText(urlText);
      textRange.select();
      return true;
    }
    return false;
  }

  function copyText(text, done){
    if(navigator.clipboard && navigator.clipboard.writeText && window.isSecureContext){
      navigator.clipboard.writeText(text).then(function(){
        done(true);
      }).catch(function(){
        done(fallbackCopy(text));
      });
      return;
    }
    done(fallbackCopy(text));
  }

  /* URLコピーボタン */
  each(document.querySelectorAll(".copy-btn"), function(button){
    button.addEventListener("click", function(){
      var text = button.getAttribute("data-copy") || "";
      var status = button.parentElement ? button.parentElement.querySelector(".copy-status") : null;
      copyText(text, function(copied){
        if(copied){
          setStatus(status, "コピーしました", 1800);
          return;
        }
        if(selectVisibleUrl(button)){
          setStatus(status, "URLを選択しました。Ctrl+Cでコピーしてください", 0);
        } else {
          setStatus(status, "コピーできない場合はURLを選択してコピーしてください", 0);
        }
      });
    });
  });

  function supportsNativeDialog(dialog){
    return !!(dialog && typeof dialog.showModal === "function");
  }

  function openDialog(dialog, fallbackAction){
    if(!dialog){
      if(fallbackAction) fallbackAction();
      return;
    }
    if(supportsNativeDialog(dialog)){
      dialog.showModal();
      return;
    }
    if(fallbackAction){
      fallbackAction();
      return;
    }
    dialog.setAttribute("open", "open");
    addClass(dialog, "legacy-dialog-open");
    addClass(document.body, "dialog-fallback-active");
  }

  function closeDialog(dialog){
    if(!dialog) return;
    if(typeof dialog.close === "function"){
      dialog.close();
    } else {
      dialog.removeAttribute("open");
    }
    removeClass(dialog, "legacy-dialog-open");
    removeClass(document.body, "dialog-fallback-active");
  }

  /* PDF保存ボタン */
  var printBtn = document.getElementById("printBtn");
  var pdfGuideDialog = document.getElementById("dlg-pdf-guide");
  var pdfPrintTrigger = document.getElementById("pdfPrintTrigger");
  if(printBtn){
    printBtn.addEventListener("click", function(){
      openDialog(pdfGuideDialog, function(){
        window.print();
      });
    });
  }
  if(pdfPrintTrigger){
    pdfPrintTrigger.addEventListener("click", function(){
      closeDialog(pdfGuideDialog);
      window.print();
    });
  }

  /* 先頭へボタン */
  var topBtn = document.getElementById("topBtn");
  if(topBtn){
    topBtn.addEventListener("click", function(){
      try{
        window.scrollTo({top:0, behavior:"smooth"});
      }catch(error){
        window.scrollTo(0, 0);
      }
    });
  }

  /* Word(.doc)ダウンロード */
  var wordBtn = document.getElementById("wordBtn");
  if(wordBtn){
    wordBtn.addEventListener("click", downloadAsWord);
  }

  function removeNodes(parent, selector){
    each(parent.querySelectorAll(selector), function(el){
      if(el.parentNode) el.parentNode.removeChild(el);
    });
  }

  function downloadBlob(blob, filename, fallbackHtml){
    if(blob && window.navigator && window.navigator.msSaveOrOpenBlob){
      window.navigator.msSaveOrOpenBlob(blob, filename);
      return true;
    }
    if(blob && window.navigator && window.navigator.msSaveBlob){
      window.navigator.msSaveBlob(blob, filename);
      return true;
    }
    if(blob && window.URL && URL.createObjectURL){
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      if(typeof a.click === "function"){
        a.click();
      } else {
        var event = document.createEvent("MouseEvents");
        event.initMouseEvent("click", true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
        a.dispatchEvent(event);
      }
      document.body.removeChild(a);
      window.setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
      return true;
    }
    var fallbackWindow = window.open("about:blank", "_blank");
    if(fallbackWindow && fallbackWindow.document){
      fallbackWindow.document.open();
      fallbackWindow.document.write(fallbackHtml);
      fallbackWindow.document.close();
      return true;
    }
    return false;
  }

  function downloadAsWord(){
    var contentEl = document.querySelector(".content");
    if(!contentEl) return;

    var clone = contentEl.cloneNode(true);

    /* 印刷・画面専用要素を削除 */
    removeNodes(clone, ".copy-btn, .copy-status, .fab-wrap, .screen-only, button, .jump-row, dialog, .search-overlay");

    /* ルートバッジのdata-modal属性を除去（Word内でリンク扱いにならないよう） */
    each(clone.querySelectorAll("[data-modal]"), function(el){
      el.removeAttribute("data-modal");
    });

    var wordStyle =
      "body {" +
      "font-family: 'Yu Gothic', 'Meiryo', 'Hiragino Kaku Gothic ProN', sans-serif;" +
      "font-size: 10.5pt;" +
      "line-height: 1.8;" +
      "color: #17202c;" +
      "margin: 0;" +
      "}" +
      "h1 {" +
      "font-size: 18pt;" +
      "font-weight: bold;" +
      "text-align: center;" +
      "margin-bottom: 18pt;" +
      "color: #fff;" +
      "background: #18324f;" +
      "padding: 14pt 16pt;" +
      "mso-element: para-border-div;" +
      "}" +
      "h2 {" +
      "font-size: 13pt;" +
      "font-weight: bold;" +
      "background: #18324f;" +
      "color: #fff;" +
      "padding: 8pt 12pt;" +
      "margin: 18pt 0 8pt;" +
      "page-break-before: always;" +
      "mso-element: para-border-div;" +
      "}" +
      "h2:first-of-type { page-break-before: avoid; }" +
      "h3 {" +
      "font-size: 11.5pt;" +
      "font-weight: bold;" +
      "border-bottom: 1.5pt solid #18324f;" +
      "padding-bottom: 3pt;" +
      "color: #18324f;" +
      "margin: 14pt 0 6pt;" +
      "}" +
      "h4, h5 {" +
      "font-size: 10.5pt;" +
      "font-weight: bold;" +
      "color: #23344a;" +
      "margin: 10pt 0 4pt;" +
      "}" +
      "p { margin: 4pt 0; }" +
      "table {" +
      "border-collapse: collapse;" +
      "width: 100%;" +
      "margin: 8pt 0;" +
      "font-size: 9.5pt;" +
      "}" +
      "td, th {" +
      "border: 1pt solid #aeb8c7;" +
      "padding: 4pt 6pt;" +
      "vertical-align: top;" +
      "}" +
      "th {" +
      "background: #edf1f5;" +
      "font-weight: bold;" +
      "color: #18324f;" +
      "mso-element: para-border-div;" +
      "}" +
      "ul, ol { padding-left: 18pt; margin: 5pt 0; }" +
      "li { margin-bottom: 2pt; line-height: 1.7; }" +
      ".note {" +
      "border: 1pt solid #c7d4e3;" +
      "border-left: 5pt solid #18324f;" +
      "background: #eef2f6;" +
      "padding: 6pt 10pt;" +
      "margin: 7pt 0;" +
      "}" +
      ".note strong { display: block; color: #18324f; margin-bottom: 2pt; font-weight: bold; }" +
      ".route {" +
      "display: inline;" +
      "font-weight: bold;" +
      "padding: 1pt 5pt;" +
      "border-radius: 3pt;" +
      "}" +
      ".route.a { color: #fff; background: #18324f; border: 1pt solid #18324f; }" +
      ".route.b { color: #18324f; border: 1pt solid #18324f; }" +
      ".steps {" +
      "list-style: decimal;" +
      "padding-left: 20pt;" +
      "margin: 6pt 0;" +
      "}" +
      ".steps li { margin: 4pt 0; }" +
      ".st-title { font-weight: bold; color: #18324f; }" +
      ".chapter-summary {" +
      "border: 1pt solid #d8dee7;" +
      "padding: 6pt 10pt;" +
      "margin: 6pt 0;" +
      "background: #f7f8fa;" +
      "}" +
      ".pill {" +
      "display: inline;" +
      "border: 1pt solid #aaa;" +
      "padding: 1pt 5pt;" +
      "font-size: 9pt;" +
      "border-radius: 3pt;" +
      "}" +
      ".toc, .fab-wrap, .search-overlay, .screen-only, .callout, .mock-screen, .shot-points { display: none !important; }" +
      "img { max-width: 100%; }" +
      "figure { margin: 8pt 0; page-break-inside: avoid; }" +
      "figcaption { font-size: 9pt; color: #334155; margin-top: 3pt; }";

    var htmlContent = "<!DOCTYPE html>" +
      "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
      "xmlns:w='urn:schemas-microsoft-com:office:word' " +
      "xmlns='http://www.w3.org/TR/REC-html40' lang='ja'>" +
      "<head>" +
      "<meta charset='UTF-8'>" +
      "<!--[if gte mso 9]><xml>" +
      "<w:WordDocument>" +
      "<w:View>Normal</w:View>" +
      "<w:Zoom>90</w:Zoom>" +
      "<w:DoNotOptimizeForBrowser/>" +
      "</w:WordDocument>" +
      "</xml><![endif]-->" +
      "<style>" + wordStyle + "</style>" +
      "</head>" +
      "<body>" +
      clone.innerHTML +
      "</body>" +
      "</html>";

    var blob = null;
    try{
      if(window.Blob){
        blob = new Blob(["\ufeff", htmlContent], {type: "application/msword;charset=utf-8"});
      }
    }catch(error){
      blob = null;
    }
    var downloaded = downloadBlob(blob, "業務委託台帳_入力マニュアル.doc", htmlContent);
    if(!downloaded){
      alert("Word保存を開始できませんでした。ブラウザの保存制限を確認してください。");
    }
  }

  /* 検索機能 */
  var overlay = document.getElementById("searchOverlay");
  var searchInput = document.getElementById("searchInput");
  var searchCount = document.getElementById("searchCount");
  var searchScope = document.querySelector("main.content");
  var searchHits = [];
  var hitIndex = -1;

  var openSearch = function(){
    if(!overlay) return;
    addClass(overlay, "open");
    if(searchInput){
      searchInput.focus();
      searchInput.select();
    }
  };

  var closeSearch = function(){
    if(!overlay) return;
    removeClass(overlay, "open");
    clearHighlights();
    if(searchInput) searchInput.value = "";
    if(searchCount) searchCount.textContent = "";
  };

  function clearHighlights(){
    if(!searchScope) return;
    each(searchScope.querySelectorAll("mark.search-hit"), function(mark){
      var parent = mark.parentNode;
      if(!parent) return;
      parent.replaceChild(document.createTextNode(mark.textContent || ""), mark);
      parent.normalize();
    });
    searchHits = [];
    hitIndex = -1;
  }

  function updateCount(){
    if(!searchCount) return;
    if(searchHits.length){
      searchCount.textContent = (hitIndex + 1) + " / " + searchHits.length;
    } else if(searchInput && searchInput.value.replace(/^\s+|\s+$/g, "")){
      searchCount.textContent = "0 件";
    } else {
      searchCount.textContent = "";
    }
  }

  function jumpToHit(index){
    if(!searchHits.length) return;
    hitIndex = (index + searchHits.length) % searchHits.length;
    each(searchHits, function(item){
      removeClass(item, "current-hit");
    });
    addClass(searchHits[hitIndex], "current-hit");
    safeScrollIntoView(searchHits[hitIndex]);
    updateCount();
  }

  function escapeRegExp(text){
    return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function runSearch(){
    if(!searchScope || !searchInput) return;
    var query = searchInput.value.replace(/^\s+|\s+$/g, "");
    var regex;
    var walker;
    var targets = [];
    var i;

    clearHighlights();
    if(!query){
      updateCount();
      return;
    }
    regex = new RegExp(escapeRegExp(query), "gi");
    walker = document.createTreeWalker(searchScope, NodeFilter.SHOW_TEXT, {
      acceptNode:function(node){
        if(!node.nodeValue || !node.nodeValue.replace(/^\s+|\s+$/g, "")) return NodeFilter.FILTER_REJECT;
        var parent = node.parentElement;
        if(!parent) return NodeFilter.FILTER_REJECT;
        if(parent.tagName === "SCRIPT" || parent.tagName === "STYLE" || parent.tagName === "MARK") return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    while(walker.nextNode()) targets.push(walker.currentNode);
    for(i = 0; i < targets.length; i += 1){
      (function(node){
        var text = node.nodeValue || "";
        var matched = false;
        var frag = document.createDocumentFragment();
        var lastIdx = 0;
        text.replace(regex, function(match, offset){
          matched = true;
          frag.appendChild(document.createTextNode(text.slice(lastIdx, offset)));
          var mark = document.createElement("mark");
          mark.className = "search-hit";
          mark.textContent = match;
          frag.appendChild(mark);
          searchHits.push(mark);
          lastIdx = offset + match.length;
          return match;
        });
        if(matched){
          frag.appendChild(document.createTextNode(text.slice(lastIdx)));
          if(node.parentNode) node.parentNode.replaceChild(frag, node);
        }
      }(targets[i]));
    }
    hitIndex = 0;
    if(searchHits.length){
      jumpToHit(0);
    } else {
      updateCount();
    }
  }

  if(searchInput){
    searchInput.addEventListener("input", runSearch);
    searchInput.addEventListener("keydown", function(e){
      e = e || window.event;
      if(e.key === "Enter" || e.keyCode === 13){
        if(e.preventDefault) e.preventDefault();
        jumpToHit(e.shiftKey ? hitIndex - 1 : hitIndex + 1);
      }
      if(e.key === "Escape" || e.keyCode === 27){
        if(e.preventDefault) e.preventDefault();
        closeSearch();
      }
    });
  }

  var prevBtn = document.getElementById("searchPrevBtn");
  var nextBtn = document.getElementById("searchNextBtn");
  var closeBtn = document.getElementById("searchCloseBtn");
  var searchFab = document.getElementById("searchFab");
  if(prevBtn) prevBtn.addEventListener("click", function(){ jumpToHit(hitIndex - 1); });
  if(nextBtn) nextBtn.addEventListener("click", function(){ jumpToHit(hitIndex + 1); });
  if(closeBtn) closeBtn.addEventListener("click", closeSearch);
  if(searchFab) searchFab.addEventListener("click", openSearch);

  document.addEventListener("keydown", function(e){
    e = e || window.event;
    var key = e.key || String.fromCharCode(e.keyCode || 0).toLowerCase();
    if((e.ctrlKey || e.metaKey) && String(key).toLowerCase() === "f"){
      if(e.preventDefault) e.preventDefault();
      openSearch();
    }
    if((e.key === "Escape" || e.keyCode === 27) && overlay && hasClass(overlay, "open")){
      closeSearch();
    }
  });

  /* 目次アクティブリンク */
  var links = Array.prototype.slice.call(document.querySelectorAll(".toc a"));
  var sections = [];
  each(links, function(link){
    var section = document.querySelector(link.getAttribute("href"));
    if(section) sections.push(section);
  });
  if("IntersectionObserver" in window){
    var observer = new IntersectionObserver(function(entries){
      each(entries, function(entry){
        if(entry.isIntersecting){
          each(links, function(link){
            removeClass(link, "active");
          });
          var active = null;
          var i;
          for(i = 0; i < links.length; i += 1){
            if(links[i].hash === "#" + entry.target.id){
              active = links[i];
              break;
            }
          }
          if(active) addClass(active, "active");
        }
      });
    }, {rootMargin:"-25% 0px -65% 0px", threshold:0.01});
    each(sections, function(section){ observer.observe(section); });
  }

  /* ルートモーダル */
  each(document.querySelectorAll("[data-modal]"), function(el){
    el.addEventListener("click", function(e){
      e = e || window.event;
      if(e.preventDefault) e.preventDefault();
      var dlg = document.getElementById(el.getAttribute("data-modal"));
      openDialog(dlg, null);
    });
  });

  /* dialogのbackdropクリックで閉じる */
  each(document.querySelectorAll("dialog"), function(dlg){
    dlg.addEventListener("click", function(e){
      e = e || window.event;
      if(e.target === dlg) closeDialog(dlg);
    });
  });

  /* .modal-close ボタン */
  each(document.querySelectorAll(".modal-close"), function(btn){
    btn.addEventListener("click", function(){
      closeDialog(closest(btn, "dialog"));
    });
  });

  /* ダイアログ内の「詳細を見る」リンクはダイアログを閉じてからジャンプ */
  each(document.querySelectorAll("dialog .modal-link[href]"), function(link){
    link.addEventListener("click", function(){
      closeDialog(closest(link, "dialog"));
    });
  });

  /* 読書進捗バー */
  var progressBar = document.querySelector(".progress-bar");
  if(progressBar){
    var updateProgress = function(){
      var scrollable = document.documentElement.scrollHeight - window.innerHeight;
      var pct = scrollable > 0 ? (window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0) / scrollable * 100 : 0;
      progressBar.style.width = Math.min(pct, 100) + "%";
    };
    window.addEventListener("scroll", updateProgress, {passive:true});
    updateProgress();
  }

}());
