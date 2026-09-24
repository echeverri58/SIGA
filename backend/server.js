/* ============================================================
   SIGA · Backend de conversión Excel -> PDF (LibreOffice)
   Recibe un .xlsx y devuelve el PDF convertido (idéntico al Excel).

   Endpoints:
     GET  /api/pdf   -> { ok:true, conversor:"libreoffice" }  (sonda)
     POST /api/pdf   -> cuerpo = archivo .xlsx ; respuesta = PDF
   ============================================================ */
const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFile } = require("child_process");

const PORT = process.env.PORT || 3000;
const MAX_BYTES = 40 * 1024 * 1024;

function ponerCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function json(res, code, obj) {
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(obj));
}

const server = http.createServer((req, res) => {
  const url = (req.url || "/").split("?")[0];
  ponerCors(res);

  if (req.method === "OPTIONS") { res.writeHead(204); return res.end(); }

  if (req.method === "GET" && url === "/api/pdf") {
    return json(res, 200, { ok: true, conversor: "libreoffice" });
  }

  if (req.method === "POST" && url === "/api/pdf") {
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
            res.writeHead(200, {
              "Content-Type": "application/pdf",
              "Content-Length": datos.length
            });
            res.end(datos);
            console.log("  [PDF] convertido en " + ((Date.now() - t0) / 1000).toFixed(1) + " s");
          }
          try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) {}
        }
      );
    });
    return;
  }

  res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
  res.end("Backend SIGA: usa POST /api/pdf");
});

server.listen(PORT, "0.0.0.0", () => {
  console.log("Conversor SIGA (LibreOffice) escuchando en el puerto " + PORT);
});
