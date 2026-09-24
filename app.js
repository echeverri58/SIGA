/* ============================================================
   SIGA · FORMATO PLANILLA DE ASISTENCIA
   Generación del Excel 100% en el navegador (sin servidor).
   ============================================================ */
(function () {
  "use strict";

  /* ---------- Constantes de la plantilla ---------- */
  var TEMPLATE_URL = "Siga.xlsx";            // plantilla completa (Excel que descarga el usuario)
  var TEMPLATE_PDF = "Siga_planilla.xlsx";   // plantilla con SOLO la hoja de la planilla (para el PDF)

  /* URL del backend que convierte el Excel a PDF.
     - "" (vacío) = usa el mismo servidor donde está la app (node server.js).
     - Si despliegas el backend en la nube, pega aquí su URL, por ejemplo:
         var CONVERSOR_URL = "https://siga-backend.onrender.com";
     Con eso, el PDF que se descargue DESDE GITHUB también será el Excel convertido. */
  var CONVERSOR_URL = "";

  var SHEET2 = "xl/worksheets/sheet2.xml";            // hoja "asistencia_mod"
  var DRAWING2 = "xl/drawings/drawing2.xml";          // logo + casillas
  var DRAWING2_RELS = "xl/drawings/_rels/drawing2.xml.rels";
  var FILA_INICIO = 11;          // primera fila de datos (1-based)
  var MAX_PARTICIPANTES = 20;    // filas 11 a 30 de la planilla
  var FIRMA_COL = 11;            // columna L (indice base 0)
  var FIRMA_W = 150, FIRMA_H = 55, EMU = 9525;

  var IMG_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

  // columnas del participante -> letra de columna en Excel
  var COLS = {
    tipo_doc: "B",
    num_doc: "C",
    nombres: "D",
    apellidos: "F",
    direccion: "H",
    correo: "I",
    telefono: "K"
  };

  /* ---------- Estado ---------- */
  var participantes = [];
  var formato = "excel";   // "excel" | "pdf"

  /* ---------- Helpers DOM ---------- */
  function $(id) { return document.getElementById(id); }
  function escapeXml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /* ---------- Firma a pantalla completa ---------- */
  var canvas = $("canvasFirma");
  var ctx = canvas.getContext("2d");
  var dibujando = false;
  var firmaHecha = false;
  var firmaActual = "";
  var scrollPrevio = 0;

  function medidaCanvas() {
    var r = canvas.getBoundingClientRect();
    return { w: Math.max(1, Math.round(r.width)), h: Math.max(1, Math.round(r.height)) };
  }

  function ajustarCanvas() {
    var m = medidaCanvas();
    if (m.w < 4 || m.h < 4) return;              // el modal está oculto
    var dpr = window.devicePixelRatio || 1;
    var w = Math.round(m.w * dpr), h = Math.round(m.h * dpr);
    if (canvas.width !== w || canvas.height !== h) {
      canvas.width = w;
      canvas.height = h;
    }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineWidth = 3;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.strokeStyle = "#16302a";
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, m.w, m.h);
    firmaHecha = false;
  }

  function punto(e) {
    var r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }
  function iniciar(e) {
    e.preventDefault();
    dibujando = true;
    var p = punto(e);
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
    ctx.lineTo(p.x + 0.01, p.y + 0.01);
    ctx.stroke();
    firmaHecha = true;
    $("padHint").classList.add("hidden");
    if (typeof actualizarAvisoGiro === "function") actualizarAvisoGiro();
  }
  function mover(e) {
    if (!dibujando) return;
    e.preventDefault();
    var p = punto(e);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
  }
  function terminar() { dibujando = false; }

  canvas.addEventListener("pointerdown", iniciar);
  canvas.addEventListener("pointermove", mover);
  canvas.addEventListener("pointerup", terminar);
  canvas.addEventListener("pointerleave", terminar);
  canvas.addEventListener("pointercancel", terminar);

  function borrarFirma() {
    var m = medidaCanvas();
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, m.w, m.h);
    firmaHecha = false;
    $("padHint").classList.remove("hidden");
    if (typeof actualizarAvisoGiro === "function") actualizarAvisoGiro();
  }

  /* Dibuja una imagen (data URL) dentro del lienzo sin deformarla */
  function dibujarEnCanvas(dataUrl, cb) {
    var img = new Image();
    img.onload = function () {
      var m = medidaCanvas();
      var escala = Math.min(m.w / img.width, m.h / img.height);
      var w = img.width * escala, h = img.height * escala;
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, m.w, m.h);
      ctx.drawImage(img, (m.w - w) / 2, (m.h - h) / 2, w, h);
      firmaHecha = true;
      if (cb) cb();
    };
    img.onerror = function () { toast("No se pudo leer la imagen.", "error"); };
    img.src = dataUrl;
  }

  /* Recorta el espacio en blanco sobrante de la firma */
  function firmaRecortada() {
    if (!firmaHecha) return "";
    var w = canvas.width, h = canvas.height;
    var datos;
    try {
      datos = ctx.getImageData(0, 0, w, h).data;
    } catch (e) {
      return canvas.toDataURL("image/png");
    }
    var minX = w, minY = h, maxX = -1, maxY = -1;
    for (var y = 0; y < h; y++) {
      var base = y * w * 4;
      for (var x = 0; x < w; x++) {
        var i = base + x * 4;
        if (datos[i] < 235 || datos[i + 1] < 235 || datos[i + 2] < 235) {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }
    if (maxX < 0) return "";
    var pad = Math.round(Math.min(w, h) * 0.04) + 4;
    minX = Math.max(0, minX - pad); minY = Math.max(0, minY - pad);
    maxX = Math.min(w - 1, maxX + pad); maxY = Math.min(h - 1, maxY + pad);
    var cw = maxX - minX + 1, ch = maxY - minY + 1;
    // Limita el tamaño para que el Excel/PDF no pesen demasiado
    var MAX_LADO = 520;
    var esc = Math.min(1, MAX_LADO / Math.max(cw, ch));
    var fw = Math.max(1, Math.round(cw * esc));
    var fh = Math.max(1, Math.round(ch * esc));
    var tmp = document.createElement("canvas");
    tmp.width = fw; tmp.height = fh;
    var t = tmp.getContext("2d");
    t.fillStyle = "#fff";
    t.fillRect(0, 0, fw, fh);
    t.drawImage(canvas, minX, minY, cw, ch, 0, 0, fw, fh);
    return tmp.toDataURL("image/png");
  }

  /* Evita que la página se desplace mientras se firma */
  function bloquearScroll() {
    scrollPrevio = window.scrollY || window.pageYOffset || 0;
    document.body.style.position = "fixed";
    document.body.style.top = "-" + scrollPrevio + "px";
    document.body.style.left = "0";
    document.body.style.right = "0";
    document.body.style.width = "100%";
    document.body.classList.add("modal-open");
  }
  function liberarScroll() {
    document.body.classList.remove("modal-open");
    document.body.style.position = "";
    document.body.style.top = "";
    document.body.style.left = "";
    document.body.style.right = "";
    document.body.style.width = "";
    window.scrollTo(0, scrollPrevio);
  }

  function actualizarPreviewFirma() {
    var img = $("firmaPreview");
    var vacio = $("firmaEmpty");
    if (firmaActual) {
      img.src = firmaActual;
      img.hidden = false;
      vacio.hidden = true;
      $("btnQuitarFirma").hidden = false;
    } else {
      img.hidden = true;
      img.removeAttribute("src");
      vacio.hidden = false;
      $("btnQuitarFirma").hidden = true;
    }
  }

  /* Muestra u oculta el aviso de girar el celular (solo en vertical) */
  function actualizarAvisoGiro() {
    var hint = $("rotateHint");
    if (!hint) return;
    var vertical = window.innerHeight >= window.innerWidth;
    hint.hidden = !(vertical && !firmaHecha);
  }

  /* Redimensiona el lienzo conservando lo ya dibujado (al girar / ampliar) */
  function reajustarCanvas() {
    if ($("firmaModal").hidden) return;
    var previa = firmaHecha ? canvas.toDataURL() : "";
    ajustarCanvas();
    if (previa) dibujarEnCanvas(previa, actualizarAvisoGiro);
    else actualizarAvisoGiro();
  }

  /* Intenta poner la firma en pantalla completa y en horizontal */
  function alternarPantallaCompleta() {
    var el = $("firmaModal");
    var salir = document.exitFullscreen || document.webkitExitFullscreen || document.msExitFullscreen;
    var entrar = el.requestFullscreen || el.webkitRequestFullscreen || el.msRequestFullscreen;

    if (document.fullscreenElement || document.webkitFullscreenElement) {
      if (salir) salir.call(document);
      if (screen.orientation && screen.orientation.unlock) {
        try { screen.orientation.unlock(); } catch (e) {}
      }
      return;
    }
    if (!entrar) {
      toast("Gira el celular a horizontal para tener más espacio.");
      return;
    }
    Promise.resolve(entrar.call(el)).then(function () {
      if (screen.orientation && screen.orientation.lock) {
        screen.orientation.lock("landscape").catch(function () {});
      }
      setTimeout(reajustarCanvas, 350);
    }).catch(function () {
      toast("Gira el celular a horizontal para tener más espacio.");
    });
  }

  function abrirModalFirma() {
    bloquearScroll();
    $("firmaModal").hidden = false;
    requestAnimationFrame(function () {
      ajustarCanvas();
      if (firmaActual) {
        dibujarEnCanvas(firmaActual, function () {
          $("padHint").classList.add("hidden");
          actualizarAvisoGiro();
        });
      } else {
        $("padHint").classList.remove("hidden");
        actualizarAvisoGiro();
      }
    });
  }

  function cerrarModalFirma() {
    $("firmaModal").hidden = true;
    liberarScroll();
  }

  $("btnCargarFirma").addEventListener("click", abrirModalFirma);
  $("btnCerrarFirma").addEventListener("click", cerrarModalFirma);
  $("btnBorrarFirma").addEventListener("click", borrarFirma);

  $("btnGuardarFirma").addEventListener("click", function () {
    var data = firmaRecortada();
    if (!data) { toast("Dibuja o carga la firma antes de guardarla.", "error"); return; }
    firmaActual = data;
    actualizarPreviewFirma();
    cerrarModalFirma();
    toast("Firma guardada.", "success");
  });

  $("btnQuitarFirma").addEventListener("click", function () {
    firmaActual = "";
    actualizarPreviewFirma();
  });

  $("btnSubirFirma").addEventListener("click", function () { $("inputFirmaArchivo").click(); });
  $("inputFirmaArchivo").addEventListener("change", function () {
    var f = this.files && this.files[0];
    this.value = "";
    if (!f) return;
    var lector = new FileReader();
    lector.onload = function (ev) {
      dibujarEnCanvas(ev.target.result, function () {
        $("padHint").classList.add("hidden");
        toast("Imagen cargada. Pulsa Guardar firma.", "success");
      });
    };
    lector.readAsDataURL(f);
  });

  $("btnAmpliarFirma").addEventListener("click", alternarPantallaCompleta);

  // Al girar el celular o cambiar de tamaño, el lienzo se adapta y conserva la firma
  var temporizadorGiro = null;
  function alCambiarTamano() {
    clearTimeout(temporizadorGiro);
    temporizadorGiro = setTimeout(reajustarCanvas, 120);
  }
  window.addEventListener("resize", alCambiarTamano);
  window.addEventListener("orientationchange", alCambiarTamano);

  /* ---------- UI: lista y contadores ---------- */
  function iniciales(nombres, apellidos) {
    var a = (nombres || "").trim().charAt(0);
    var b = (apellidos || "").trim().charAt(0);
    return ((a || "") + (b || "")).toUpperCase() || "?";
  }

  function actualizarContadores() {
    var n = participantes.length;
    $("chipContador").textContent = n + " de " + MAX_PARTICIPANTES;
    $("chipContador").classList.toggle("full", n >= MAX_PARTICIPANTES);
    $("totalFinal").textContent = n;
    $("btnGenerar").disabled = n === 0;
    $("btnVaciar").hidden = n === 0;
  }

  function renderLista() {
    var lista = $("lista");
    lista.innerHTML = "";
    if (!participantes.length) {
      var empty = document.createElement("div");
      empty.className = "empty";
      empty.innerHTML = '<div class="empty-ico">👥</div><p>Aún no hay participantes.</p>' +
        '<span>Agrega el primero para comenzar a construir la planilla.</span>';
      lista.appendChild(empty);
      return;
    }
    participantes.forEach(function (p, i) {
      var item = document.createElement("div");
      item.className = "item";

      var avatar = document.createElement("div");
      avatar.className = "avatar";
      avatar.textContent = iniciales(p.nombres, p.apellidos);

      var info = document.createElement("div");
      info.className = "info";
      var nombre = document.createElement("span");
      nombre.className = "nombre";
      nombre.textContent = (p.nombres + " " + p.apellidos).trim();
      var meta = document.createElement("span");
      meta.className = "meta";
      meta.textContent = (p.tipo_doc + " " + p.num_doc) +
        (p.correo ? " · " + p.correo : "") +
        (p.telefono ? " · " + p.telefono : "");
      info.appendChild(nombre);
      info.appendChild(meta);

      var firma = document.createElement("img");
      firma.className = "firma-mini";
      firma.src = p.firma;
      firma.alt = "Firma";

      var quitar = document.createElement("button");
      quitar.className = "btn-quitar";
      quitar.type = "button";
      quitar.title = "Quitar";
      quitar.textContent = "✕";
      quitar.addEventListener("click", function () {
        participantes.splice(i, 1);
        renderLista();
        actualizarContadores();
      });

      item.appendChild(avatar);
      item.appendChild(info);
      item.appendChild(firma);
      item.appendChild(quitar);
      lista.appendChild(item);
    });
  }

  /* ---------- Toast ---------- */
  var toastTimer;
  function toast(msg, tipo) {
    var t = $("toast");
    t.textContent = msg;
    t.className = "toast show " + (tipo || "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.className = "toast"; }, 3200);
  }

  /* ---------- Agregar participante ---------- */
  function vaciarFormulario() {
    $("tipoDoc").value = "CC";
    ["numDoc", "nombres", "apellidos", "direccion", "correo", "telefono"]
      .forEach(function (id) { $(id).value = ""; });
    firmaActual = "";
    actualizarPreviewFirma();
  }

  $("formParticipante").addEventListener("submit", function (e) {
    e.preventDefault();
    if (participantes.length >= MAX_PARTICIPANTES) {
      return toast("La planilla admite máximo " + MAX_PARTICIPANTES + " participantes. Genera este archivo y crea otro para los siguientes.", "error");
    }
    var tipo = $("tipoDoc").value.trim().toUpperCase();
    var num = $("numDoc").value.trim().toUpperCase();
    var nom = $("nombres").value.trim().toUpperCase();
    var ape = $("apellidos").value.trim().toUpperCase();
    var dir = $("direccion").value.trim().toUpperCase();
    var cor = $("correo").value.trim().toUpperCase();
    var tel = $("telefono").value.trim().toUpperCase();
    var firma = firmaActual;

    if (!num || !nom || !ape) {
      return toast("Completa al menos: número de documento, nombres y apellidos.", "error");
    }
    if (!firma) {
      return toast("Dibuja la firma del participante antes de agregarlo.", "error");
    }

    participantes.push({
      tipo_doc: tipo, num_doc: num, nombres: nom, apellidos: ape,
      direccion: dir, correo: cor, telefono: tel, firma: firma
    });

    vaciarFormulario();
    renderLista();
    actualizarContadores();
    toast("Participante agregado correctamente.", "success");
  });

  $("btnVaciar").addEventListener("click", function () {
    participantes = [];
    renderLista();
    actualizarContadores();
    toast("Lista vaciada.", "success");
  });

  /* ---------- Edición XML de la plantilla ---------- */
  function setCellInline(xml, ref, value) {
    // Los datos ingresados van en MAYÚSCULA y con fuente tamaño 12 (sz val="12")
    var inline = '<is><r><rPr><sz val="12"/></rPr><t xml:space="preserve">' + escapeXml(value) + '</t></r></is>';
    var re = new RegExp('<c r="' + ref + '"([^>]*?)/>');
    if (re.test(xml)) {
      return xml.replace(re, '<c r="' + ref + '"$1 t="inlineStr">' + inline + '</c>');
    }
    var re2 = new RegExp('<c r="' + ref + '"([^>]*)>[\\s\\S]*?<\\/c>');
    if (re2.test(xml)) {
      return xml.replace(re2, '<c r="' + ref + '"$1 t="inlineStr">' + inline + '</c>');
    }
    return xml;
  }

  function setRowHeight(xml, row, height) {
    var re = new RegExp('(<row r="' + row + '"[^>]*?\\sht=")([^"]*)(")');
    return xml.replace(re, '$1' + height + '$3');
  }

  function dataURLtoUint8(dataUrl) {
    var b64 = dataUrl.split(",")[1];
    var bin = atob(b64);
    var u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  function firmaAnchor(rid, index) {
    var row0 = (FILA_INICIO - 1) + index;
    var cid = 1000 + index;
    var padX = 47625;   // ~5px de margen horizontal (EMU)
    var padY = 38100;   // ~4px de margen vertical (EMU)
    // twoCellAnchor: la firma queda contenida DENTRO de la celda L:M (FIRMA)
    return '<xdr:twoCellAnchor editAs="oneCell">' +
      '<xdr:from><xdr:col>' + FIRMA_COL + '</xdr:col><xdr:colOff>' + padX + '</xdr:colOff>' +
      '<xdr:row>' + row0 + '</xdr:row><xdr:rowOff>' + padY + '</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + (FIRMA_COL + 2) + '</xdr:col><xdr:colOff>-' + padX + '</xdr:colOff>' +
      '<xdr:row>' + (row0 + 1) + '</xdr:row><xdr:rowOff>-' + padY + '</xdr:rowOff></xdr:to>' +
      '<xdr:pic><xdr:nvPicPr>' +
      '<xdr:cNvPr id="' + cid + '" name="Firma ' + (index + 1) + '" descr="Firma"/>' +
      '<xdr:cNvPicPr/></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="' + rid + '"/>' +
      '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
      '</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';
  }

  function firmaRel(rid, n) {
    return '<Relationship Id="' + rid + '" Type="' + IMG_REL_TYPE +
      '" Target="../media/siga_firma_' + n + '.png"/>';
  }

  /* ---------- Generar Excel ---------- */
  function nombreArchivo(ext) {
    var d = new Date();
    function p(n) { return (n < 10 ? "0" : "") + n; }
    return "SIGA_FORMATO_PLANILLA_ASISTENCIA_" +
      d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "_" +
      p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()) + "." + ext;
  }

  function descargarBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1000);
  }

  async function generarExcel(sinComprimir, plantilla) {
    var resp = await fetch(plantilla || TEMPLATE_URL, { cache: "no-store" });
    if (!resp.ok) throw new Error("No se pudo cargar la plantilla " + (plantilla || TEMPLATE_URL) + " (HTTP " + resp.status + ")");
    var buf = await resp.arrayBuffer();
    var zip = await JSZip.loadAsync(buf);

    // 1) Hoja de datos
    var sheet2 = await zip.file(SHEET2).async("string");
    for (var i = 0; i < participantes.length; i++) {
      var p = participantes[i];
      var row = FILA_INICIO + i;
      for (var k in COLS) {
        sheet2 = setCellInline(sheet2, COLS[k] + row, p[k] || "");
      }
      sheet2 = setRowHeight(sheet2, row, 60);
    }
    zip.file(SHEET2, sheet2);

    // 2) Dibujo: anclas de firmas + relaciones
    var drawing = await zip.file(DRAWING2).async("string");
    var rels = await zip.file(DRAWING2_RELS).async("string");
    var anchors = "";
    var newRels = "";
    var n = 0;
    for (var j = 0; j < participantes.length; j++) {
      if (participantes[j].firma) {
        var rid = "rId" + (n + 2); // rId1 ya es el logo de la hoja
        anchors += firmaAnchor(rid, j);
        newRels += firmaRel(rid, n);
        n++;
      }
    }
    drawing = drawing.replace(/<\/xdr:wsDr>\s*$/, anchors + "</xdr:wsDr>");
    rels = rels.replace(/<\/Relationships>\s*$/, newRels + "</Relationships>");
    zip.file(DRAWING2, drawing);
    zip.file(DRAWING2_RELS, rels);

    // 3) Imágenes de las firmas
    n = 0;
    for (var m = 0; m < participantes.length; m++) {
      if (participantes[m].firma) {
        zip.file("xl/media/siga_firma_" + n + ".png", dataURLtoUint8(participantes[m].firma));
        n++;
      }
    }

    // 4) Empaquetar (sin comprimir cuando solo se usara para convertir a PDF: es mas rapido)
    var blob = await zip.generateAsync({
      type: "blob",
      compression: sinComprimir ? "STORE" : "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    return blob;
  }

  /* ---------- PDF generado en el navegador (respaldo sin servidor) ---------- */
  function fechaHoy() {
    var d = new Date();
    var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    return d.getDate() + " de " + meses[d.getMonth()] + " de " + d.getFullYear();
  }

  var logoCache = null;
  function logoDataURL() {
    if (logoCache !== null) return Promise.resolve(logoCache);
    return fetch("sena_logo.png", { cache: "force-cache" })
      .then(function (r) { return r.ok ? r.blob() : null; })
      .then(function (b) {
        if (!b) { logoCache = ""; return ""; }
        return new Promise(function (res) {
          var fr = new FileReader();
          fr.onload = function () {
            var img = new Image();
            img.onload = function () {
              // Se reduce el logo para que el PDF no pese demasiado
              var MAX = 150;
              var esc = Math.min(1, MAX / Math.max(img.width, img.height));
              var cv = document.createElement("canvas");
              cv.width = Math.max(1, Math.round(img.width * esc));
              cv.height = Math.max(1, Math.round(img.height * esc));
              var c = cv.getContext("2d");
              c.fillStyle = "#fff";
              c.fillRect(0, 0, cv.width, cv.height);
              c.drawImage(img, 0, 0, cv.width, cv.height);
              logoCache = cv.toDataURL("image/png");
              res(logoCache);
            };
            img.onerror = function () { logoCache = ""; res(""); };
            img.src = fr.result;
          };
          fr.onerror = function () { logoCache = ""; res(""); };
          fr.readAsDataURL(b);
        });
      })
      .catch(function () { logoCache = ""; return ""; });
  }

  /* ---------- Réplica fiel del formato (PDF generado en el navegador) ----------
     Geometría tomada de la plantilla oficial: área de impresión A1:M33,
     anchos de columna y altos de fila reales. */

  // Anchos de columna A..M (px a 96 dpi, de la plantilla)
  var GEO_COLS = [55, 157, 155, 141, 64, 140, 109, 302, 116, 64, 203, 83, 161];

  // Alto de cada fila (px). Las filas 11-30 se calculan según la firma.
  var GEO_ALTO = { 1: 110, 2: 39, 3: 59, 4: 37, 5: 48, 6: 72, 7: 37, 8: 37, 9: 37, 10: 55, 31: 20, 32: 24, 33: 24 };

  // Bloques fijos del formato: [columnaInicial, cuantasColumnas, texto, estilo]
  var GEO_BLOQUES = {
    2: [[0, 12, "PROCESO DE DIRECCIÓN DE FORMACIÓN PROFESIONAL INTEGRAL", { b: 1, c: 1 }],
        [12, 1, "VERSIÓN: 3", { b: 1, c: 1, s: 7 }]],
    3: [[0, 12, "FORMATO PLANILLA  DE ASISTENCIA", { b: 1, c: 1, s: 12 }],
        [12, 1, "CÓDIGO: GFPI-PL-001", { b: 1, c: 1, s: 7 }]],
    4: [[0, 13, "FECHA DE DILIGENCIAMIENTO:", { b: 1 }]],
    5: [[0, 2, "REGIONAL:", { b: 1 }], [2, 3, "5 Antioquia", {}],
        [5, 2, "CENTRO DE FORMACIÓN:", { b: 1, s: 8 }], [7, 3, "9401 Centro de Servicios de Salud", {}],
        [10, 1, "CIUDAD/MUNICIPIO:", { b: 1, s: 6 }], [11, 2, "Medellin", {}]],
    6: [[0, 9, "NOMBRE DEL PROGRAMA DE FORMACIÓN:", { b: 1 }],
        [9, 2, "NÚMERO DE FICHA DE CARACTERIZACIÓN", { b: 1, s: 6 }],
        [11, 2, "", {}]],
    7: [[0, 13, "A CONTINUACIÓN SELECCIONE EL PROCESO QUE SE VA A REALIZAR", { b: 1, s: 8, c: 1 }]],
    8: [[0, 3, "CHARLAS INFORMATIVAS", { s: 8 }], [3, 1, "", { chk: 1 }],
        [4, 4, "PRESENTACIÓN PRUEBAS PRESENCIALES", { s: 8 }], [8, 1, "", { chk: 1 }],
        [9, 2, "MATRICULA", { s: 8 }], [11, 2, "", { chk: 1 }]],
    9: [[0, 13, "DATOS DE LOS  PARTICIPANTES", { b: 1, c: 1 }]],
    10: [[0, 1, "No:", { b: 1, s: 8, c: 1 }],
         [1, 1, "TIPO DE DOCUMENTO DE IDENTIDAD ASPIRANTE", { b: 1, s: 8, c: 1 }],
         [2, 1, "NÚMERO DOCUMENTO IDENTIDAD ASPIRANTE", { b: 1, s: 8, c: 1 }],
         [3, 2, "NOMBRES DEL PARTICIPANTE", { b: 1, s: 8, c: 1 }],
         [5, 2, "APELLIDOS DEL PARTICIPANTE", { b: 1, s: 8, c: 1 }],
         [7, 1, "DIRECCIÓN / DEPENDENCIA / CARGO", { b: 1, s: 7, c: 1 }],
         [8, 2, "CORREO ELECTRONICO", { b: 1, s: 8, c: 1 }],
         [10, 1, "TELÉFONO", { b: 1, s: 8, c: 1 }],
         [11, 2, "FIRMA", { b: 1, s: 8, c: 1 }]]
  };

  async function generarPDFNavegador() {
    var doc = new jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    var pageW = doc.internal.pageSize.getWidth();
    var pageH = doc.internal.pageSize.getHeight();
    var logo = await logoDataURL();

    // --- alturas de fila (px) ---
    var alto = [];
    var r;
    for (r = 1; r <= 33; r++) {
      if (r >= 11 && r <= 30) {
        var pp = participantes[r - 11];
        alto[r] = (pp && pp.firma) ? 80 : 47;
      } else {
        alto[r] = GEO_ALTO[r] || 20;
      }
    }

    var totalW = 0, totalH = 0, i;
    for (i = 0; i < GEO_COLS.length; i++) totalW += GEO_COLS[i];
    for (r = 1; r <= 33; r++) totalH += alto[r];

    // --- escala para que quepa en la hoja, centrado ---
    var MG = 8;
    var esc = Math.min((pageW - MG * 2) / totalW, (pageH - MG * 2) / totalH);
    var ox = (pageW - totalW * esc) / 2;
    var oy = (pageH - totalH * esc) / 2;

    var xs = [0];
    for (i = 0; i < GEO_COLS.length; i++) xs.push(xs[i] + GEO_COLS[i]);
    var ys = [0, 0];
    for (r = 1; r <= 33; r++) ys[r + 1] = ys[r] + alto[r];

    function X(col) { return ox + xs[col] * esc; }
    function Y(fila) { return oy + ys[fila] * esc; }
    function ANCHO(span, col) { return (xs[col + span] - xs[col]) * esc; }
    function ALTO(fila, span) { return (ys[fila + span] - ys[fila]) * esc; }
    function PT(puntos) { return Math.max(3.2, puntos * esc * 3.7795); }  // tamaño de fuente escalado

    doc.setTextColor(20, 38, 28);

    function celda(fila, col, span, texto, est) {
      est = est || {};
      var x = X(col), y = Y(fila), w = ANCHO(span, col), h = ALTO(fila, 1);
      doc.setDrawColor(95, 110, 102);
      doc.setLineWidth(0.15);
      doc.rect(x, y, w, h);
      if (est.chk) {
        var lado = Math.min(w, h) * 0.55;
        doc.setLineWidth(0.3);
        doc.rect(x + (w - lado) / 2, y + (h - lado) / 2, lado, lado);
      }
      if (texto) {
        var fs = PT(est.s || 9);
        doc.setFont("helvetica", est.b ? "bold" : "normal");
        doc.setFontSize(fs);
        var pad = Math.max(0.25, 1.2 * esc * 3.7795);
        var lineas = doc.splitTextToSize(String(texto), Math.max(0.6, w - pad * 2));
        var lh = fs * 0.3528 * 1.15;
        var ty = y + (h - lineas.length * lh) / 2 + fs * 0.3528 * 0.85;
        if (est.c) doc.text(lineas, x + w / 2, ty, { align: "center" });
        else doc.text(lineas, x + pad, ty);
      }
    }

    // --- fila 1: logo ---
    doc.setDrawColor(95, 110, 102);
    doc.setLineWidth(0.15);
    doc.rect(X(0), Y(1), ANCHO(13, 0), ALTO(1, 1));
    if (logo) {
      var lh = ALTO(1, 1) * 0.75;
      try {
        doc.addImage(logo, "PNG", X(5) + (ANCHO(3, 5) - lh) / 2, Y(1) + (ALTO(1, 1) - lh) / 2, lh, lh);
      } catch (e) { /* logo no válido */ }
    }

    // --- bloques fijos (filas 2 a 10) ---
    for (r = 2; r <= 10; r++) {
      var celdas = GEO_BLOQUES[r] || [];
      for (i = 0; i < celdas.length; i++) {
        celda(r, celdas[i][0], celdas[i][1], celdas[i][2], celdas[i][3]);
      }
    }

    // --- filas de datos (11 a 30) ---
    for (r = 11; r <= 30; r++) {
      var p = participantes[r - 11];
      celda(r, 0, 1, String(r - 10), { s: 11, c: 1 });
      celda(r, 1, 1, p ? p.tipo_doc : "", { s: 11, c: 1 });
      celda(r, 2, 1, p ? p.num_doc : "", { s: 11, c: 1 });
      celda(r, 3, 2, p ? p.nombres : "", { s: 11, c: 1 });
      celda(r, 5, 2, p ? p.apellidos : "", { s: 11, c: 1 });
      celda(r, 7, 1, p ? p.direccion : "", { s: 11, c: 1 });
      celda(r, 8, 2, p ? p.correo : "", { s: 11, c: 1 });
      celda(r, 10, 1, p ? p.telefono : "", { s: 11, c: 1 });
      celda(r, 11, 2, "", { s: 11 });
      if (p && p.firma) {
        try {
          doc.addImage(p.firma, "PNG", X(11) + 0.7, Y(r) + 0.7,
                       ANCHO(2, 11) - 1.4, ALTO(r, 1) - 1.4, undefined, "FAST");
        } catch (e) { /* firma no válida */ }
      }
    }

    // --- fila 31 (separación) ---
    celda(31, 0, 13, "", {});

    // --- consentimiento (filas 32-33, combinadas) ---
    var xc = X(0), yc = Y(32), wc = ANCHO(13, 0), hc = ALTO(32, 2);
    doc.setDrawColor(95, 110, 102);
    doc.setLineWidth(0.15);
    doc.rect(xc, yc, wc, hc);
    doc.setFont("helvetica", "normal");
    var fsc = PT(7);
    doc.setFontSize(fsc);
    var nota = "Consentimiento de prueba: es la manifestación libre, voluntaria y expresa, que da por escrito o vía web, " +
               "un aspirante a la formación en el SENA, autorizando que se le realice una prueba de selección, " +
               "cuyo resultado deberá consignarse en su registro.";
    doc.text(doc.splitTextToSize(nota, wc - 2), xc + 1, yc + 1.5 + fsc * 0.3528 * 0.85);

    // --- pie con la fecha ---
    doc.setFont("helvetica", "normal");
    doc.setFontSize(5);
    doc.setTextColor(130, 140, 134);
    doc.text("Generado con SIGA · " + fechaHoy(), pageW / 2, pageH - 1.5, { align: "center" });

    return doc.output("blob");
  }

  /* ---------- Generar PDF ----------
     Si está el servidor local con Excel, convierte el Excel real (idéntico).
     Si no (por ejemplo GitHub Pages), usa el PDF del navegador. */
  var servidorPdf = null;      // null = sin probar
  var motivoFalloPdf = "";     // por qué no se pudo usar la conversión

  async function generarPDF() {
    if (servidorPdf !== false) {
      try {
        // Se usa la plantilla de UNA sola hoja para que el PDF salga solo con la planilla
        var xlsxBlob = await generarExcel(true, TEMPLATE_PDF);
        var resp = await fetch((CONVERSOR_URL || "") + "/api/pdf", { method: "POST", body: xlsxBlob });
        if (resp.ok) {
          servidorPdf = true;
          motivoFalloPdf = "";
          return await resp.blob();
        }
        var detalle = "";
        try { var j = await resp.json(); detalle = j && j.error ? j.error : ""; } catch (e2) {}
        servidorPdf = false;
        motivoFalloPdf = detalle || ("el servidor respondió HTTP " + resp.status);
      } catch (e) {
        servidorPdf = false;
        motivoFalloPdf = e && e.message ? e.message : "no se pudo contactar el conversor";
      }
    }
    return generarPDFNavegador();
  }

  /* ---------- Selector de formato ---------- */
  var fmtButtons = document.querySelectorAll("#formatToggle .fmt");
  fmtButtons.forEach(function (b) {
    b.addEventListener("click", function () {
      formato = b.getAttribute("data-format");
      fmtButtons.forEach(function (x) { x.classList.remove("active"); });
      b.classList.add("active");
    });
  });

  /* ---------- Botón Descargar ---------- */
  var primerPdf = true;

  $("btnGenerar").addEventListener("click", function () {
    if (!participantes.length) return;
    var btn = $("btnGenerar");
    var txt = $("btnGenerarTxt");
    var esPdf = (formato === "pdf");
    btn.disabled = true;

    // Indicador de progreso con segundos transcurridos (el PDF usa Excel)
    var seg = 0;
    txt.textContent = esPdf ? "Generando PDF… 0s" : "Generando…";
    var reloj = null;
    if (esPdf) {
      reloj = setInterval(function () {
        seg++;
        txt.textContent = "Generando PDF… " + seg + "s";
      }, 1000);
      if (primerPdf) {
        primerPdf = false;
        toast("Abriendo Excel para el PDF. La primera vez puede tardar unos segundos…");
      }
    }

    var promesa = esPdf ? generarPDF() : generarExcel();
    promesa
      .then(function (blob) {
        descargarBlob(blob, nombreArchivo(esPdf ? "pdf" : "xlsx"));
        if (!esPdf) {
          return toast("Descarga lista: " + participantes.length + " participante(s) — Excel.", "success");
        }
        if (servidorPdf) {
          return toast("Descarga lista — PDF IDÉNTICO al Excel (convertido).", "success");
        }
        // Se usó la réplica: se muestra el motivo para poder corregirlo
        toast("PDF aproximado (no se pudo convertir el Excel). Motivo: " + (motivoFalloPdf || "conversor no disponible"), "error");
      })
      .catch(function (err) {
        toast("Error al generar el archivo: " + err.message, "error");
      })
      .finally(function () {
        if (reloj) clearInterval(reloj);
        btn.disabled = participantes.length === 0;
        txt.textContent = "Descargar";
      });
  });

  /* ---------- Inicialización ---------- */
  actualizarPreviewFirma();
  actualizarContadores();

  // Detecta si hay un conversor (Excel/LibreOffice) disponible: solo entonces el PDF es idéntico al Excel
  function mostrarEstadoPdf() {
    var el = $("pdfEstado");
    if (!el) return;
    el.hidden = false;
    if (servidorPdf === true) {
      el.className = "pdf-estado ok";
      el.textContent = "✅ PDF EXACTO ACTIVO — la descarga en PDF será el Excel convertido: idéntico al Excel.";
    } else {
      el.className = "pdf-estado no";
      el.textContent = "⚠️ PDF APROXIMADO — no se detectó el conversor de Excel. Abre la app con 'node server.js' " +
                       "en un PC con Excel y reinícialo; así el PDF será idéntico al Excel.";
    }
  }

  fetch((CONVERSOR_URL || "") + "/api/pdf", { method: "GET", cache: "no-store" })
    .then(function (r) { return r.ok ? r.json() : null; })
    .then(function (d) { servidorPdf = !!(d && d.ok); mostrarEstadoPdf(); })
    .catch(function () { servidorPdf = false; mostrarEstadoPdf(); });

  // Service worker: permite instalar la app en el celular y usarla sin conexión
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", function () {
      navigator.serviceWorker.register("sw.js").catch(function () { /* sin soporte */ });
    });
  }
})();
