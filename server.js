/* Mini servidor estático para probar SIGA localmente (sin Python).
   Usa solo módulos nativos de Node.js (no necesita instalar nada).
   Uso:  node server.js          (puerto 8080)
         node server.js 3000     (puerto personalizado)
*/
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { spawn } = require("child_process");

const PORT = parseInt(process.argv[2], 10) || 8080;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".md": "text/markdown; charset=utf-8"
};

function localIP() {
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === "IPv4" && !net.internal) return net.address;
    }
  }
  return "127.0.0.1";
}

/* ============================================================
   Convertidor Excel -> PDF en segundo plano.
   Mantiene una sola instancia de Excel abierta para que las
   conversiones siguientes sean mucho más rápidas.
   ============================================================ */
let conv = null;          // { proc, ready }
let convBuf = "";
let convArrancando = null;
let convPendiente = null; // { resolve, reject, timer }
let convCola = Promise.resolve();

function arrancarConvertidor() {
  if (conv && conv.ready) return Promise.resolve();
  if (convArrancando) return convArrancando;

  const promesa = new Promise((resolve, reject) => {
    const script = path.join(ROOT, "convertir_server.ps1");
    const proc = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
      { windowsHide: true }
    );
    conv = { proc, ready: false };
    convBuf = "";

    const timeout = setTimeout(() => {
      try { proc.kill(); } catch (e) {}
      conv = null;
      reject(new Error("Excel tardó demasiado en iniciar"));
    }, 90000);

    proc.stdout.setEncoding("utf8");
    proc.stdout.on("data", (chunk) => {
      convBuf += chunk;
      let i;
      while ((i = convBuf.indexOf("\n")) >= 0) {
        const line = convBuf.slice(0, i).replace(/\r$/, "").trim();
        convBuf = convBuf.slice(i + 1);
        if (!line) continue;

        if (line === "READY") {
          conv.ready = true;
          clearTimeout(timeout);
          console.log("  [Excel] listo para convertir a PDF");
          resolve();
        } else if (convPendiente) {
          const p = convPendiente;
          convPendiente = null;
          clearTimeout(p.timer);
          if (line.indexOf("OK") === 0) p.resolve();
          else p.reject(new Error(line.replace(/^ERR\s*/, "") || "Error al convertir"));
        }
      }
    });

    proc.stderr.on("data", () => {});
    proc.on("exit", () => {
      conv = null;
      if (convPendiente) {
        const p = convPendiente;
        convPendiente = null;
        clearTimeout(p.timer);
        p.reject(new Error("El convertidor de Excel se cerró"));
      }
    });
  });

  convArrancando = promesa;
  const limpiar = () => { convArrancando = null; };
  promesa.then(limpiar, limpiar);
  return promesa;
}

function convertirUna(xlsxPath, pdfPath) {
  return arrancarConvertidor().then(() => new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      convPendiente = null;
      try { if (conv) conv.proc.kill(); } catch (e) {}
      reject(new Error("La conversión tardó demasiado"));
    }, 90000);
    convPendiente = { resolve, reject, timer };
    conv.proc.stdin.write(xlsxPath + "|" + pdfPath + "\n");
  }));
}

// Serializa las conversiones (una a la vez) reutilizando el mismo Excel.
function convertir(xlsxPath, pdfPath) {
  const tarea = convCola.then(() => convertirUna(xlsxPath, pdfPath));
  convCola = tarea.catch(() => {});
  return tarea;
}

function cerrarConvertidor() {
  if (conv) {
    try { conv.proc.stdin.write("QUIT\n"); } catch (e) {}
    setTimeout(() => { try { if (conv) conv.proc.kill(); } catch (e) {} }, 800);
  }
}

// Convierte un .xlsx (recibido en el body) a PDF.
function handlePdfConversion(req, res) {
  const chunks = [];
  req.on("data", (c) => chunks.push(c));
  req.on("end", () => {
    const body = Buffer.concat(chunks);
    if (!body.length) {
      res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ ok: false, error: "Archivo vacío" }));
    }
    const ts = Date.now();
    const xlsxPath = path.join(ROOT, `_tmp_${ts}.xlsx`);
    const pdfPath = path.join(ROOT, `_tmp_${ts}.pdf`);
    fs.writeFileSync(xlsxPath, body);

    const t0 = Date.now();
    convertir(xlsxPath, pdfPath)
      .then(() => {
        if (!fs.existsSync(pdfPath)) throw new Error("No se generó el PDF");
        const pdf = fs.readFileSync(pdfPath);
        res.writeHead(200, {
          "Content-Type": "application/pdf",
          "Content-Length": pdf.length
        });
        res.end(pdf);
        console.log("  [PDF] convertido en " + ((Date.now() - t0) / 1000).toFixed(1) + " s");
      })
      .catch((err) => {
        res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({
          ok: false,
          error: "No se pudo convertir a PDF: " + err.message
        }));
        console.log("  [PDF] error: " + err.message);
      })
      .finally(() => {
        try { fs.unlinkSync(xlsxPath); } catch (e) {}
        try { fs.unlinkSync(pdfPath); } catch (e) {}
      });
  });
}

const server = http.createServer((req, res) => {
  try {
    const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);

    // Endpoint para convertir Excel -> PDF (PDF idéntico al Excel)
    if (req.method === "POST" && urlPath === "/api/pdf") {
      return handlePdfConversion(req, res);
    }

    let filePath = urlPath;
    if (filePath === "/") filePath = "/index.html";

    // Evita salir de la carpeta (seguridad básica)
    const fullPath = path.normalize(path.join(ROOT, filePath));
    if (!fullPath.startsWith(ROOT)) {
      res.writeHead(403);
      return res.end("Acceso denegado");
    }

    fs.readFile(fullPath, (err, data) => {
      if (err) {
        res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
        return res.end("No encontrado: " + filePath);
      }
      const ext = path.extname(fullPath).toLowerCase();
      res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
      res.end(data);
    });
  } catch (e) {
    res.writeHead(500);
    res.end("Error del servidor");
  }
});

server.listen(PORT, "0.0.0.0", () => {
  console.log("");
  console.log("  SIGA · Formato Planilla de Asistencia");
  console.log("  --------------------------------------");
  console.log("  En esta computadora:  http://localhost:" + PORT);
  console.log("  Desde el celular (misma red WiFi):  http://" + localIP() + ":" + PORT);
  console.log("  Detener con Ctrl+C");
  console.log("");
  // Prepara Excel en segundo plano para que el primer PDF también sea rápido
  arrancarConvertidor()
    .then(() => console.log("  Listo para generar PDF."))
    .catch((e) => console.log("  [Aviso] No se pudo preparar Excel para PDF: " + e.message));
  console.log("");
});

process.on("SIGINT", () => {
  console.log("\nCerrando…");
  cerrarConvertidor();
  setTimeout(() => process.exit(0), 1200);
});
