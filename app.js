/* ============================================================
   SIGA · FORMATO PLANILLA DE ASISTENCIA
   Generación del Excel 100% en el navegador (sin servidor).
   ============================================================ */
(function () {
  "use strict";

  /* ---------- Constantes de la plantilla ---------- */
  var TEMPLATE_URL = "Siga.xlsx";
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

  /* ---------- Firma (canvas táctil + mouse) ---------- */
  var canvas = $("canvasFirma");
  var ctx = canvas.getContext("2d");
  var dibujando = false;
  var firmaHecha = false;

  function ajustarCanvas() {
    var dpr = window.devicePixelRatio || 1;
    var rect = canvas.getBoundingClientRect();
    var w = Math.max(1, Math.round(rect.width * dpr));
    var h = Math.max(1, Math.round(rect.height * dpr));
    if (canvas.width !== w || canvas.height !== h) {
      var prev = null;
      if (canvas.width) { prev = canvas.toDataURL(); }
      canvas.width = w;
      canvas.height = h;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.lineWidth = 2.6;
      ctx.lineCap = "round";
      ctx.lineJoin = "round";
      ctx.strokeStyle = "#16302a";
      if (prev) {
        var img = new Image();
        img.onload = function () { ctx.drawImage(img, 0, 0, rect.width, rect.height); };
        img.src = prev;
      }
    }
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
    $("padHint").classList.add("hidden");
  }
  function mover(e) {
    if (!dibujando) return;
    e.preventDefault();
    var p = punto(e);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
  }
  function terminar() {
    if (dibujando) { dibujando = false; firmaHecha = true; }
  }

  canvas.addEventListener("pointerdown", iniciar);
  canvas.addEventListener("pointermove", mover);
  canvas.addEventListener("pointerup", terminar);
  canvas.addEventListener("pointerleave", terminar);
  canvas.addEventListener("pointercancel", terminar);

  function borrarFirma() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    firmaHecha = false;
    $("padHint").classList.remove("hidden");
  }

  function firmaDataURL() {
    if (!firmaHecha) return "";
    var w = canvas.width, h = canvas.height;
    var tmp = document.createElement("canvas");
    tmp.width = w; tmp.height = h;
    var t = tmp.getContext("2d");
    t.fillStyle = "#ffffff";
    t.fillRect(0, 0, w, h);
    t.drawImage(canvas, 0, 0);
    return tmp.toDataURL("image/png");
  }

  $("btnBorrarFirma").addEventListener("click", borrarFirma);
  window.addEventListener("resize", ajustarCanvas);

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
    borrarFirma();
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
    var firma = firmaDataURL();

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

  async function generarExcel() {
    var resp = await fetch(TEMPLATE_URL, { cache: "no-store" });
    if (!resp.ok) throw new Error("No se pudo cargar la plantilla Siga.xlsx (HTTP " + resp.status + ")");
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

    // 4) Comprimir y descargar
    var blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    return blob;
  }

  /* ---------- Generar PDF en el navegador (aproximado, respaldo) ---------- */
  function fechaHoy() {
    var d = new Date();
    var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    return d.getDate() + " de " + meses[d.getMonth()] + " de " + d.getFullYear();
  }

  function generarPDFNavegador() {
    var doc = new jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    var pageW = doc.internal.pageSize.getWidth();

    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text("SERVICIO NACIONAL DE APRENDIZAJE SENA", pageW / 2, 12, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text("SISTEMA INTEGRADO DE GESTIÓN · PROCEDIMIENTO PLANEACIÓN Y PUBLICACIÓN DE LA OFERTA EDUCATIVA", pageW / 2, 17, { align: "center" });
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("FORMATO PLANILLA DE ASISTENCIA", pageW / 2, 23, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text("VERSIÓN: 3     CÓDIGO: GFPI-PL-001     Fecha de generación: " + fechaHoy(), pageW / 2, 28, { align: "center" });

    var body = participantes.map(function (p, i) {
      return [String(i + 1), p.tipo_doc, p.num_doc, p.nombres, p.apellidos,
              p.direccion, p.correo, p.telefono, ""];
    });

    doc.autoTable({
      startY: 32,
      margin: { left: 10, right: 10 },
      head: [["No", "Tipo doc.", "Número doc.", "Nombres", "Apellidos",
              "Dirección / Dependencia / Cargo", "Correo electrónico", "Teléfono", "Firma"]],
      body: body,
      theme: "grid",
      styles: { font: "helvetica", fontSize: 8, cellPadding: 2, valign: "middle",
                halign: "center", textColor: [20, 38, 28], lineColor: [150, 162, 156], lineWidth: 0.25 },
      headStyles: { fillColor: [15, 122, 61], textColor: [255, 255, 255],
                    fontStyle: "bold", halign: "center", valign: "middle" },
      alternateRowStyles: { fillColor: [240, 246, 242] },
      columnStyles: {
        0: { cellWidth: 10 },
        1: { cellWidth: 20 },
        2: { cellWidth: 24, halign: "left" },
        3: { cellWidth: 32, halign: "left" },
        4: { cellWidth: 32, halign: "left" },
        5: { cellWidth: 37, halign: "left" },
        6: { cellWidth: 40, halign: "left" },
        7: { cellWidth: 24, halign: "left" },
        8: { cellWidth: 58 }
      },
      bodyStyles: { minCellHeight: 18 },
      didDrawCell: function (data) {
        if (data.section === "body" && data.column.index === 8) {
          var p = participantes[data.row.index];
          if (p && p.firma) {
            try {
              var x = data.cell.x + 2;
              var y = data.cell.y + 2;
              var w = data.cell.width - 4;
              var h = data.cell.height - 4;
              doc.addImage(p.firma, "PNG", x, y, w, h, undefined, "FAST");
            } catch (e) { /* firma no válida */ }
          }
        }
      }
    });

    return doc.output("blob");
  }

  /* ---------- Generar PDF fiel: convierte el Excel real a PDF ----------
     Requiere el servidor local (server.js) y Microsoft Excel instalado.
     Si no está disponible, cae al PDF aproximado del navegador. */
  async function generarPDF() {
    try {
      var xlsxBlob = await generarExcel();
      var resp = await fetch("/api/pdf", { method: "POST", body: xlsxBlob });
      if (resp.ok) {
        return await resp.blob();
      }
    } catch (e) { /* sin servidor: usa el respaldo */ }
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
  $("btnGenerar").addEventListener("click", function () {
    if (!participantes.length) return;
    var btn = $("btnGenerar");
    var txt = $("btnGenerarTxt");
    btn.disabled = true;
    txt.textContent = "Generando…";

    var promesa = (formato === "pdf") ? generarPDF() : generarExcel();
    promesa
      .then(function (blob) {
        descargarBlob(blob, nombreArchivo(formato === "pdf" ? "pdf" : "xlsx"));
        toast("Descarga lista: " + participantes.length + " participante(s) en " +
              (formato === "pdf" ? "PDF" : "Excel") + ".", "success");
      })
      .catch(function (err) {
        toast("Error al generar el archivo: " + err.message, "error");
      })
      .finally(function () {
        btn.disabled = participantes.length === 0;
        txt.textContent = "Descargar";
      });
  });

  /* ---------- Inicialización ---------- */
  ajustarCanvas();
  borrarFirma();
  actualizarContadores();
})();
