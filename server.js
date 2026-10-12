// Server statis untuk folder docs/.
//   node server.js                 HTTPS di https://localhost:3000 (cara resmi untuk Office)
//   node server.js --http          HTTP biasa: tanpa sertifikat, cukup untuk Office di Windows
//   --mock                         tambah AI tiruan di /mock/... untuk mencoba tanpa API key
//   --dev                          tambah jembatan uji (lihat tools/dev-bridge-server.js)
const fs = require("fs");
const http = require("http");
const https = require("https");
const os = require("os");
const path = require("path");

const ROOT = path.join(__dirname, "docs");
// Nomor proses server dicatat di sini supaya tools/stop.js tahu proses mana yang boleh dimatikan.
const PID_FILE = path.join(os.tmpdir(), "office-ai-addin-server.pid");
const args = process.argv.slice(2);
const PORT = Number((args.find((a) => a.startsWith("--port=")) || "").slice(7)) || Number(process.env.PORT) || 3000;
const mock = args.includes("--mock") ? require("./tools/mock-llm") : null;
const dev = args.includes("--dev") ? require("./tools/dev-bridge-server")(__dirname) : null;

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
};

function handler(req, res) {
  let url;
  let pathname;
  try {
    url = new URL(req.url, "http://localhost");
    pathname = decodeURIComponent(url.pathname);
  } catch {
    res.writeHead(400).end();
    return;
  }
  if (mock && pathname.startsWith("/mock/")) {
    mock(req, res, pathname);
    return;
  }
  if (dev && pathname.startsWith("/dev/")) {
    dev.handle(req, res, url);
    return;
  }
  if (dev && pathname === "/dev-bridge.js") {
    res.writeHead(200, { "Content-Type": TYPES[".js"], "Cache-Control": "no-store" }).end(dev.script());
    return;
  }
  const file = path.join(ROOT, pathname === "/" ? "taskpane.html" : pathname);
  if (!file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(file, (error, data) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Tidak ditemukan");
      return;
    }
    if (dev) {
      console.log(`${req.method} ${pathname}  [${(req.headers["user-agent"] || "").slice(-60)}]`);
      if (path.basename(file) === "taskpane.html") data = dev.inject(data.toString());
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

async function main() {
  const plain = args.includes("--http");
  // getHttpsServerOptions membuat sertifikat localhost dan, saat pertama kali, meminta izin memasangnya.
  const server = plain
    ? http.createServer(handler)
    : https.createServer(await require("office-addin-dev-certs").getHttpsServerOptions(), handler);
  server.listen(PORT, "localhost", () => {
    fs.writeFileSync(PID_FILE, `${process.pid} ${PORT}`);
    const extras = [mock && "AI tiruan", dev && "jembatan uji"].filter(Boolean).join(", ");
    console.log(`Add-in tersaji di ${plain ? "http" : "https"}://localhost:${PORT}/taskpane.html${extras ? ` (${extras})` : ""}`);
  });
}

module.exports = { PID_FILE };

if (require.main === module) main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
