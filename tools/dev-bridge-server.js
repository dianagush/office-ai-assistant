// Jembatan uji (hanya untuk pengembangan, aktif dengan `node server.js --dev`).
// Menjalankan potongan JS di dalam task pane yang sedang terbuka di Office lalu
// mengembalikan hasilnya, supaya panggilan Office.js bisa diuji dari terminal:
//
//   curl -X POST "http://localhost:3000/dev/run" -H "x-dev-token: $(cat .dev-token)" \
//        --data-binary "return Office.context.document.url"
//
// Pengaman: setiap endpoint butuh token acak (disimpan di .dev-token dan disisipkan ke
// taskpane.html, bukan ke berkas skrip) dan hanya menerima Host localhost, sehingga halaman
// web lain di komputer ini tidak bisa menyuruh task pane menjalankan kode.
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const POLL_MS = 25000;
const ALIVE_MS = 40000;

module.exports = function createDevBridge(root) {
  // Token dipakai ulang antar-restart supaya task pane yang sudah terbuka tetap tersambung.
  const tokenFile = path.join(root, ".dev-token");
  const token = fs.existsSync(tokenFile) ? fs.readFileSync(tokenFile, "utf8").trim() : crypto.randomBytes(16).toString("hex");
  fs.writeFileSync(tokenFile, token);

  const clients = new Map(); // id → { info, seen, queue, waiter }
  const pending = new Map(); // id perintah → { resolve, timer }
  let nextId = 1;

  const json = (res, status, body) => {
    res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
    res.end(JSON.stringify(body));
  };
  const readBody = (req) =>
    new Promise((resolve) => {
      let raw = "";
      req.on("data", (chunk) => (raw += chunk));
      req.on("end", () => resolve(raw));
    });
  const alive = () => [...clients].filter(([, c]) => Date.now() - c.seen < ALIVE_MS);

  function deliver(client) {
    if (!client.waiter || !client.queue.length) return;
    const { res, timer } = client.waiter;
    client.waiter = null;
    clearTimeout(timer);
    json(res, 200, client.queue.shift());
  }

  async function handle(req, res, url) {
    const local = /^(localhost|127\.0\.0\.1)(:\d+)?$/.test(req.headers.host || "");
    if (!local || req.headers["x-dev-token"] !== token) return json(res, 403, { error: "token salah" });

    if (url.pathname === "/dev/poll") {
      const id = url.searchParams.get("client");
      const client = clients.get(id) || { queue: [], waiter: null };
      clients.set(id, client);
      client.seen = Date.now();
      try {
        client.info = JSON.parse(url.searchParams.get("info") || "{}");
      } catch {
        client.info = {};
      }
      if (client.waiter) client.waiter.res.writeHead(204).end();
      const timer = setTimeout(() => {
        client.waiter = null;
        res.writeHead(204).end();
      }, POLL_MS);
      client.waiter = { res, timer };
      req.on("close", () => {
        if (client.waiter && client.waiter.res === res) {
          clearTimeout(timer);
          client.waiter = null;
        }
      });
      return deliver(client);
    }

    if (url.pathname === "/dev/result") {
      const result = JSON.parse(await readBody(req));
      const waiting = pending.get(result.id);
      if (waiting) {
        pending.delete(result.id);
        clearTimeout(waiting.timer);
        waiting.resolve(result);
      }
      return json(res, 200, {});
    }

    if (url.pathname === "/dev/clients") {
      return json(res, 200, alive().map(([id, c]) => ({ id, ...c.info })));
    }

    if (url.pathname === "/dev/run") {
      const code = await readBody(req);
      const live = alive();
      const wanted = url.searchParams.get("client");
      const match = wanted ? live.filter(([id, c]) => id === wanted || c.info.host === wanted) : live;
      if (match.length !== 1) {
        return json(res, 409, {
          error: match.length ? "lebih dari satu task pane cocok; pilih dengan ?client=" : "tidak ada task pane yang cocok",
          clients: live.map(([id, c]) => ({ id, ...c.info })),
        });
      }
      const [, client] = match[0];
      const id = nextId++;
      const timeout = Number(url.searchParams.get("timeout")) || 60000;
      const result = await new Promise((resolve) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          resolve({ id, ok: false, error: `tidak ada jawaban dalam ${timeout} ms` });
        }, timeout);
        pending.set(id, { resolve, timer });
        client.queue.push({ id, code });
        deliver(client);
      });
      return json(res, 200, result);
    }

    json(res, 404, { error: "tidak dikenal" });
  }

  // Token ditaruh di dalam HTML (yang tidak bisa dibaca lintas-origin), bukan di berkas skrip.
  const inject = (html) =>
    html.replace(
      "</body>",
      `<script>window.__DEV_TOKEN__ = "${token}";</script>\n<script type="module" src="/dev-bridge.js"></script>\n</body>`,
    );

  const script = () => fs.readFileSync(path.join(__dirname, "dev-bridge.js"));

  return { handle, inject, script };
};
