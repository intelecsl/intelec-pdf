/* Intelec · Correo a PDF — panel de Outlook */
(function () {
  "use strict";
  var VERSION = "1.3";
  var BASE = location.href.replace(/[^\/]*$/, "");
  var CHUNK = 30000;
  var last = null; // { bytes, fileName }

  var $ = function (id) { return document.getElementById(id); };

  function setStatus(kind, text) {
    var s = $("status");
    s.innerHTML = (kind === "busy" ? '<div class="spin"></div>' : "") +
      '<span class="' + (kind === "ok" ? "ok" : kind === "err" ? "err" : "") + '"></span>';
    s.lastChild.textContent = text;
  }

  function wrap(fn) {
    return new Promise(function (res, rej) {
      fn(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) res(r.value);
        else rej(r.error || new Error("Error de Office"));
      });
    });
  }

  function toBase64(buf) {
    var bytes = new Uint8Array(buf), bin = "", step = 0x8000;
    for (var i = 0; i < bytes.length; i += step)
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
    return btoa(bin);
  }

  function collectMail() {
    var item = Office.context.mailbox.item;
    var atts = (item.attachments || []).map(function (a) {
      return { id: a.id, name: a.name, size: a.size, isInline: a.isInline, contentType: a.contentType || "" };
    });
    var canContent = Office.context.requirements.isSetSupported("Mailbox", "1.8");

    var inlineP = Promise.all(atts.filter(function (a) {
      return a.isInline && canContent && /^image\//i.test(a.contentType || "image/") &&
             (!a.size || a.size < 8 * 1024 * 1024);
    }).map(function (a) {
      return wrap(function (cb) { item.getAttachmentContentAsync(a.id, cb); }).then(function (c) {
        if (c.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) return null;
        var type = a.contentType || ("image/" + ((a.name.split(".").pop() || "png").toLowerCase()));
        return { name: a.name, dataUrl: "data:" + type + ";base64," + c.content };
      }).catch(function () { return null; });
    }));

    var bodyP = wrap(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); });

    return Promise.all([bodyP, inlineP]).then(function (r) {
      return {
        subject: item.subject,
        from: item.from,
        to: item.to,
        cc: item.cc,
        date: item.dateTimeCreated,
        bodyHtml: r[0],
        attachments: atts,
        inlineImages: r[1].filter(Boolean)
      };
    });
  }

  function downloadDirect(bytes, fileName) {
    var blob = new Blob([bytes], { type: "application/pdf" });
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, fileName); return Promise.resolve();
    }
    var url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = fileName; a.style.display = "none";
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 4000);
    return Promise.resolve();
  }

  function downloadWindow(bytes, fileName) {
    return new Promise(function (resolve, reject) {
      var b64 = toBase64(bytes);
      Office.context.ui.displayDialogAsync(BASE + "dialog.html?v=" + VERSION,
        { height: 32, width: 28, displayInIframe: false, promptBeforeOpen: false },
        function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("No se pudo abrir la ventana (" + (r.error && r.error.code) + ")"));
            return;
          }
          var dlg = r.value, closed = false;
          function close() { if (!closed) { closed = true; try { dlg.close(); } catch (e) {} } }
          dlg.addEventHandler(Office.EventType.DialogEventReceived, function () { closed = true; resolve(); });
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (ev) {
            var msg = {}; try { msg = JSON.parse(ev.message); } catch (e) {}
            if (msg.t === "ready") {
              var total = Math.ceil(b64.length / CHUNK);
              dlg.messageChild(JSON.stringify({ t: "meta", name: fileName, total: total }));
              for (var i = 0; i < total; i++)
                dlg.messageChild(JSON.stringify({ t: "chunk", i: i, d: b64.substr(i * CHUNK, CHUNK) }));
              dlg.messageChild(JSON.stringify({ t: "end" }));
            } else if (msg.t === "done") {
              setTimeout(close, 1500); resolve();
            } else if (msg.t === "keep") {
              resolve(); // el usuario descargará desde la ventana
            }
          });
        });
    });
  }

  function download() {
    if (!last) return Promise.resolve();
    return downloadDirect(last.bytes, last.fileName).then(function () {
      setStatus("ok", "PDF descargado.");
    });
  }

  function run() {
    setStatus("busy", "Leyendo el correo…");
    var item = Office.context.mailbox.item;
    $("subj").textContent = item.subject || "(sin asunto)";
    $("from").textContent = item.from ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "";

    collectMail().then(function (mail) {
      setStatus("busy", "Generando PDF…");
      return IntelecPdf.buildPdf(mail);
    }).then(function (out) {
      last = out;
      $("fname").textContent = out.fileName + "  ·  " + out.pages + " pág. · " + IntelecPdf.fmtSize(out.bytes.byteLength);
      $("fname").classList.remove("hidden");
      $("btnAgain").classList.remove("hidden");
      $("btnWindow").classList.remove("hidden");
      setStatus("busy", "Descargando…");
      return download();
    }).catch(function (e) {
      console.error(e);
      setStatus("err", "Error: " + (e && e.message ? e.message : e));
    });
  }

  Office.onReady(function () {
    $("ver").textContent = "v" + VERSION;
    $("btnAgain").onclick = function () {
      download().catch(function (e) { setStatus("err", "Error: " + e.message); });
    };
    $("btnWindow").onclick = function () {
      if (!last) return;
      downloadWindow(last.bytes, last.fileName).then(function () { setStatus("ok", "PDF descargado (ventana aparte)."); })
        .catch(function (e) { setStatus("err", "Error: " + e.message); });
    };
    // Al cambiar de correo con el panel anclado, se regenera
    if (Office.context.mailbox.addHandlerAsync && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        last = null; if (Office.context.mailbox.item) run();
      });
    }
    run();
  });
})();
