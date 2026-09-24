/* ============================================================
   SIGA · Servidor para la NUBE (todo en uno)
   - Sirve la aplicación (index.html, app.js, plantillas, etc.)
   - Convierte Excel -> PDF con LibreOffice   (POST /api/pdf)
   - Sonda del conversor                      (GET  /api/pdf)
   ============================================================ */
const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFile } = require("child_process");

const PORT = process.env.PORT || 3000;
const APP_DIR = path.join(__dirname, "..");      // carpeta con index.html, app.js, Siga.xlsx...
const MAX_BYTES = 40 * 1024 * 1024;              // 40 MB por archivo

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".md": "text/markdown; charset=utf-8"
};

function cors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function json(res, code, obj) {
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(obj));
}

function convertir(req, res) {
  const partes = [];
  let total = 0;
  req.on("data", (c) => {
    total += c.length;
    if (total > MAX_BYTES) { req.destroy(); return; }
    partes.push(c);
  });
  req.on("end", () => {
    const cuerpo = Buffer.concat(partes);
    if (!cuerpo.length) return json(res, 400, { ok: false, error: "Archivo vacío" });

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "siga-"));
    const xlsx = path.join(dir, "planilla.xlsx");
    const pdf = path.join(dir, "planilla.pdf");
    fs.writeFileSync(xlsx, cuerpo);

    const t0 = Date.now();
    execFile(
      "soffice",
      ["--headless", "--norestore", "--nolockcheck", "--nodefault",
       "--convert-to", "pdf", "--outdir", dir, xlsx],
      { timeout: 120000, env: Object.assign({}, process.env, { HOME: dir }) },
      (err) => {
        if (err || !fs.existsSync(pdf)) {
          json(res, 500, { ok: false, error: "No se pudo convertir: " + (err ? err.message : "sin PDF") });
        } else {
          const datos = fs.readFileSync(pdf);
          res.writeHead(200, { "Content-Type": "application/pdf", "Content-Length": datos.length });
          res.end(datos);
          console.log("  [PDF] convertido en " + ((Date.now() - t0) / 1000).toFixed(1) + " s");
        }
        try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) {}
      }
    );
  });
}

const server = http.createServer((req, res) => {
  const url = decodeURIComponent((req.url || "/").split("?")[0]);
  cors(res);

  if (req.method === "OPTIONS") { res.writeHead(204); return res.end(); }

  if (url === "/api/pdf") {
    if (req.method === "GET") return json(res, 200, { ok: true, conversor: "libreoffice" });
    if (req.method === "POST") return convertir(req, res);
  }

  // --- Archivos de la aplicación ---
  const rel = url === "/" ? "/index.html" : url;
  const full = path.normalize(path.join(APP_DIR, rel));
  if (!full.startsWith(APP_DIR)) { res.writeHead(403); return res.end("Acceso denegado"); }

  fs.readFile(full, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      return res.end("No encontrado: " + rel);
    }
    res.writeHead(200, {
      "Content-Type": MIME[path.extname(full).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-cache"
    });
    res.end(data);
  });
});

server.listen(PORT, "0.0.0.0", () => {
  console.log("SIGA (nube) escuchando en el puerto " + PORT);
  console.log("App + conversión Excel->PDF (LibreOffice) listos.");
});
