/* Mini servidor estático para probar SIGA localmente (sin Python).
   Usa solo módulos nativos de Node.js (no necesita instalar nada).
   Uso:  node server.js          (puerto 8080)
         node server.js 3000     (puerto personalizado)
*/
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFile } = require("child_process");

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

// Convierte un .xlsx (recibido en el body) a PDF usando Excel vía PowerShell.
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

    const script = path.join(ROOT, "convertir.ps1");
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, xlsxPath, pdfPath],
      { timeout: 90000, windowsHide: true },
      (err) => {
        if (err || !fs.existsSync(pdfPath)) {
          res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({
            ok: false,
            error: "No se pudo convertir a PDF (¿tienes Microsoft Excel instalado?)"
          }));
        } else {
          const pdf = fs.readFileSync(pdfPath);
          res.writeHead(200, {
            "Content-Type": "application/pdf",
            "Content-Length": pdf.length
          });
          res.end(pdf);
        }
        try { fs.unlinkSync(xlsxPath); } catch (e) {}
        try { fs.unlinkSync(pdfPath); } catch (e) {}
      }
    );
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
});
