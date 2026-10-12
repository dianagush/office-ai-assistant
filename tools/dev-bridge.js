/* global Office */
// Sisi task pane dari jembatan uji. Hanya disajikan oleh `node server.js --dev`;
// tidak ada di docs/, jadi tidak pernah ikut ter-host. Lihat dev-bridge-server.js.

const headers = { "x-dev-token": window.__DEV_TOKEN__, "Content-Type": "application/json" };
// id bertahan saat panel dimuat ulang, supaya tidak muncul sebagai dua klien
const id = sessionStorage.getItem("dev-bridge-id") || Math.random().toString(36).slice(2, 8);
sessionStorage.setItem("dev-bridge-id", id);
const AsyncFunction = (async () => {}).constructor;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function plain(value) {
  try {
    return JSON.parse(JSON.stringify(value ?? null));
  } catch {
    return String(value);
  }
}

async function describe() {
  const info = typeof Office === "undefined" ? {} : await Office.onReady();
  return {
    host: info.host ? String(info.host) : null,
    platform: info.platform ? String(info.platform) : null,
    url: (info.host && Office.context.document && Office.context.document.url) || "",
  };
}

(async () => {
  const info = encodeURIComponent(JSON.stringify(await describe()));
  for (;;) {
    try {
      const res = await fetch(`/dev/poll?client=${id}&info=${info}`, { headers });
      if (res.status !== 200) {
        if (res.status !== 204) await sleep(3000);
        continue;
      }
      const command = await res.json();
      let result;
      try {
        result = { id: command.id, ok: true, value: plain(await new AsyncFunction(command.code)()) };
      } catch (error) {
        result = {
          id: command.id,
          ok: false,
          error: String((error && error.stack) || error),
          debug: plain(error && error.debugInfo),
        };
      }
      await fetch("/dev/result", { method: "POST", headers, body: JSON.stringify(result) });
    } catch {
      await sleep(3000);
    }
  }
})();
