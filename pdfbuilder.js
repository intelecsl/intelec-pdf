/* Intelec · Correo a PDF — generador de PDF (compartido por taskpane y pruebas)
 * Requiere: lib/jspdf.umd.min.js, lib/html2canvas.min.js, lib/purify.min.js
 */
(function (global) {
  "use strict";

  var APP_VERSION = "1.3";
  var NAVY = "#293A5F", YELLOW = "#F8DB2C";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function fmtPerson(p) {
    if (!p) return "";
    var n = p.displayName || p.name || "", e = p.emailAddress || p.address || "";
    if (n && e && n.toLowerCase() !== e.toLowerCase()) return n + " <" + e + ">";
    return n || e;
  }

  function fmtList(arr) {
    return (arr || []).map(fmtPerson).filter(Boolean).join("; ");
  }

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  function fmtDate(d) {
    if (!d) return "";
    d = new Date(d);
    if (isNaN(d)) return "";
    var dias = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
    return dias[d.getDay()] + ", " + pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" +
      d.getFullYear() + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  function fmtSize(b) {
    if (!b && b !== 0) return "";
    if (b < 1024) return b + " B";
    if (b < 1048576) return Math.round(b / 1024) + " KB";
    return (b / 1048576).toFixed(1).replace(".", ",") + " MB";
  }

  /** Nombre: AAAA-MM-DD_Remitente_Asunto.pdf */
  function buildFileName(mail) {
    var d = new Date(mail.date || Date.now());
    if (isNaN(d)) d = new Date();
    var fecha = d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
    var rem = (mail.from && (mail.from.displayName || mail.from.emailAddress)) || "Sin remitente";
    var asunto = mail.subject || "Sin asunto";
    asunto = asunto.replace(/^\s*((re|rv|fw|fwd|reenv|resp)\s*:\s*)+/i, "");
    function clean(s, max) {
      // sin tildes ni caracteres especiales: máxima compatibilidad con descargas y SharePoint
      s = String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^\x20-\x7E]/g, "").replace(/[\\\/:*?"<>|#%&{}~]/g, " ")
        .replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
      return s.length > max ? s.slice(0, max).trim() : s;
    }
    return fecha + "_" + clean(rem, 40) + "_" + clean(asunto || "Sin asunto", 80) + ".pdf";
  }

  /** Limpia el HTML del cuerpo y sustituye imágenes cid: por data URIs */
  function prepareBody(html, inlineImages) {
    inlineImages = inlineImages || [];
    var clean = global.DOMPurify.sanitize(html || "", {
      WHOLE_DOCUMENT: false,
      ADD_TAGS: ["style"],
      FORBID_TAGS: ["script", "iframe", "object", "embed", "form", "input", "button", "link", "meta", "base"],
      FORBID_ATTR: ["onerror", "onload", "onclick"],
      ALLOW_UNKNOWN_PROTOCOLS: true
    });
    var tmp = document.createElement("div");
    tmp.innerHTML = clean;

    // Estilos del correo: se limitan al contenedor para no romper la UI
    tmp.querySelectorAll("style").forEach(function (st) {
      st.textContent = st.textContent
        .replace(/(^|[},])\s*(html|body)\b/gi, "$1 .ipdf-body")
        .replace(/@page[^{]*\{[^}]*\}/gi, "");
    });

    var used = {};
    tmp.querySelectorAll("img").forEach(function (img, idx) {
      var src = img.getAttribute("src") || "";
      var m = /^cid:([^@>\s]+)(@[^>\s]*)?/i.exec(src);
      var found = null;
      if (m) {
        var key = decodeURIComponent(m[1]).toLowerCase();
        var full = decodeURIComponent(src.slice(4)).toLowerCase();
        found = inlineImages.find(function (a) {
          return (a.contentId && a.contentId.toLowerCase().replace(/[<>]/g, "") === full) ||
                 (a.name && a.name.toLowerCase() === key);
        });
      } else if (!/^(data:|https?:)/i.test(src) || /attachment\.outlook|\/owa\/|service\.svc/i.test(src)) {
        // Outlook web a veces reescribe cid: por URLs internas: se asigna por orden
        found = inlineImages.find(function (a) { return !used[a.name]; });
      }
      if (found && found.dataUrl) {
        used[found.name] = true;
        img.setAttribute("src", found.dataUrl);
      } else if (!/^(data:|https?:)/i.test(src)) {
        img.remove();
      }
      img.removeAttribute("srcset");
    });
    return tmp.innerHTML;
  }

  function headerHtml(mail) {
    var rows = [
      ["De", fmtPerson(mail.from)],
      ["Enviado", fmtDate(mail.date)],
      ["Para", fmtList(mail.to)],
      ["CC", fmtList(mail.cc)],
      ["Asunto", mail.subject || "(sin asunto)"]
    ].filter(function (r) { return r[1]; });

    var att = (mail.attachments || []).filter(function (a) { return !a.isInline; });
    var h = '<div class="ipdf-head">' +
      '<div class="ipdf-subject">' + esc(mail.subject || "(sin asunto)") + '</div>' +
      '<table class="ipdf-meta">';
    rows.forEach(function (r) {
      h += '<tr><td class="k">' + r[0] + ':</td><td class="v">' + esc(r[1]) + '</td></tr>';
    });
    if (att.length) {
      h += '<tr><td class="k">Adjuntos:</td><td class="v">' + att.map(function (a) {
        return esc(a.name) + (a.size ? " (" + fmtSize(a.size) + ")" : "");
      }).join("<br>") + '</td></tr>';
    }
    h += '</table></div>';
    return h;
  }

  var BASE_CSS =
    ".ipdf-root{width:720px;background:#fff;color:#000;font-family:Helvetica,Arial,sans-serif;font-size:13px;line-height:1.35;padding-bottom:40px}" +
    ".ipdf-head{border-bottom:3px solid " + YELLOW + ";padding:0 0 10px;margin:0 0 14px}" +
    ".ipdf-subject{font-size:17px;font-weight:bold;color:" + NAVY + ";margin:0 0 8px}" +
    ".ipdf-meta{border-collapse:collapse;width:100%}" +
    ".ipdf-meta td{padding:2px 0;vertical-align:top;font-size:12px}" +
    ".ipdf-meta td.k{width:75px;font-weight:bold;color:" + NAVY + ";white-space:nowrap}" +
    ".ipdf-meta td.v{word-break:break-word}" +
    ".ipdf-meta .sz{color:#666;font-size:12px;vertical-align:baseline}" +
    ".ipdf-body{word-wrap:break-word;overflow-wrap:break-word}" +
    ".ipdf-body img{max-width:100%!important;height:auto}" +
    ".ipdf-body table{max-width:100%}" +
    ".ipdf-body pre{white-space:pre-wrap}";

  // jsPDF.html (autoPaging) deja "Q" sin su "q" en las páginas 2+: se equilibran
  function balanceGraphicsState(d) {
    var pages = d.internal.pages;
    for (var i = 1; i < pages.length; i++) {
      var pg = pages[i]; if (!pg || !pg.length) continue;
      var depth = 0, min = 0;
      for (var j = 0; j < pg.length; j++) {
        var l = String(pg[j]).trim();
        if (l === "q") depth++;
        else if (l === "Q") { depth--; if (depth < min) min = depth; }
      }
      for (var k = 0; k < -min; k++) pg.unshift("q");
      for (var r = depth - min; r > 0; r--) pg.push("Q");
    }
  }

  /**
   * mail = { subject, from:{displayName,emailAddress}, to:[], cc:[], date, bodyHtml,
   *          attachments:[{name,size,isInline}], inlineImages:[{name,contentId,dataUrl}] }
   * Devuelve Promise<{ bytes:ArrayBuffer, fileName, pages }>
   */
  function buildPdf(mail, opts) {
    opts = opts || {};
    var jsPDF = global.jspdf.jsPDF;

    var host = document.createElement("div");
    host.style.cssText = "position:fixed;left:-10000px;top:0;width:720px;z-index:-1;";
    var style = document.createElement("style");
    style.textContent = BASE_CSS;
    var root = document.createElement("div");
    root.className = "ipdf-root";
    root.innerHTML = headerHtml(mail) +
      '<div class="ipdf-body">' + prepareBody(mail.bodyHtml, mail.inlineImages) + '</div>';
    host.appendChild(style);
    host.appendChild(root);
    document.body.appendChild(host);

    // Espera a que carguen las imágenes (máx. 8 s)
    var imgs = Array.prototype.slice.call(root.querySelectorAll("img"));
    var waitImgs = Promise.all(imgs.map(function (im) {
      if (im.complete) return Promise.resolve();
      return new Promise(function (res) {
        im.addEventListener("load", res); im.addEventListener("error", function () { im.remove(); res(); });
        setTimeout(res, 8000);
      });
    }));

    return waitImgs.then(function () {
      var doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait", compress: true });
      doc.setProperties({
        title: mail.subject || "Correo",
        subject: "Correo de " + fmtPerson(mail.from),
        author: fmtPerson(mail.from),
        creator: "Intelec · Correo a PDF v" + APP_VERSION
      });
      return new Promise(function (resolve, reject) {
        try {
          doc.html(root, {
            x: 0, y: 0,
            margin: [14, 12, 16, 12],
            width: 186,              // 210 - 12 - 12 mm
            windowWidth: 720,
            autoPaging: "text",
            html2canvas: { useCORS: true, allowTaint: false, logging: false, backgroundColor: null },
            callback: function (d) { resolve(d); }
          });
        } catch (e) { reject(e); }
      });
    }).then(function (d) {
      balanceGraphicsState(d);
      // Pie de página
      var n = d.getNumberOfPages(), stamp = buildFileName(mail).replace(/\.pdf$/, "");
      for (var i = 1; i <= n; i++) {
        d.setPage(i);
        d.setFont("helvetica", "normal");
        d.setFontSize(7.5);
        d.setTextColor(120);
        d.text(stamp.length > 95 ? stamp.slice(0, 95) + "…" : stamp, 12, 290);
        d.text(i + " / " + n, 198, 290, { align: "right" });
      }
      global.__PDF_PRODUCER__ = "Intelec Correo a PDF v" + APP_VERSION;
      var out = { bytes: d.output("arraybuffer"), fileName: buildFileName(mail), pages: n };
      host.remove();
      return out;
    }, function (e) { host.remove(); throw e; });
  }

  global.IntelecPdf = {
    VERSION: APP_VERSION,
    buildPdf: buildPdf,
    buildFileName: buildFileName,
    fmtSize: fmtSize
  };
})(window);
