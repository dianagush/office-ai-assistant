// Lintas aplikasi. Panel AI Assistant di Word, Excel, dan PowerPoint dimuat dari alamat yang sama,
// jadi bisa saling berkirim pesan lewat BroadcastChannel tanpa server. Dengan itu agen di satu
// aplikasi bisa menjalankan alat milik aplikasi lain yang panelnya sedang terbuka, misalnya membaca
// angka dari Excel lalu menuliskannya ke Word.
//
// Terbukti jalan di Office untuk Windows. Microsoft tidak menjaminnya di platform lain; di sana
// panel lain tidak akan terdeteksi dan fitur ini diam saja.
//
// Batasnya: alat berisiko "danger" (yang tidak bisa dibatalkan) tidak bisa dijalankan dari aplikasi lain.

const CHANNEL = "office-ai-addin-peers";
const BEAT_MS = 4000;
const ALIVE_MS = 12000;
const CALL_TIMEOUT_MS = 120000;
const KEEP_UNDO = 50; // jumlah panggilan terakhir yang masih bisa dibatalkan dari jauh

// options: { host, title, tools, enabled() → boolean, context() → { trackChanges, revert(since) | null }, onActivity(teks) }
// `channel` hanya diganti saat pengujian, supaya tidak tercampur dengan panel yang sedang dipakai.
export function startPeers({ host, title, tools, enabled, context, onActivity, channel: name = CHANNEL }) {
  if (!host || typeof BroadcastChannel === "undefined") return null;

  const channel = new BroadcastChannel(name);
  const me = { id: Math.random().toString(36).slice(2, 10), host, title };
  const peers = new Map(); // id → { id, host, title, focused, seen }
  const calls = new Map(); // sisi peminta: id panggilan → { resolve, reject, confirm, arm }
  const asks = new Map(); // sisi pelaksana: id pertanyaan izin → resolve
  const undos = new Map(); // sisi pelaksana: id panggilan → langkah-langkah pembatalannya
  let focused = Date.now();
  let counter = 0;

  const send = (message) => channel.postMessage({ ...message, from: { ...me, focused } });
  const announce = () => enabled() && send({ type: "here" });

  // ---- sisi pelaksana: menjalankan alat sendiri atas permintaan panel lain ----

  async function serve({ id, from, name, args }) {
    const tool = tools.find((t) => t.name === name);
    const steps = [];
    let changes = 0;
    let asked = 0;
    const { trackChanges, revert } = context();
    const since = new Date();
    onActivity(`${from.host}: ${tool ? tool.label(args) : name}`);
    try {
      if (!tool || tool.risk === "danger") throw new Error("Alat ini tidak bisa dijalankan dari aplikasi lain.");
      const output = await tool.run(args, {
        trackChanges,
        // izin diminta di panel tempat pengguna sedang bercakap
        confirm: (text) =>
          new Promise((resolve) => {
            const ask = `${id}:${asked++}`;
            asks.set(ask, resolve);
            send({ type: "ask", to: from.id, id, ask, text });
          }),
        undo: (step) => (steps.push(step), changes++),
        changed: () => changes++,
      });
      // Word: perubahan yang tercatat sebagai tracked changes dibatalkan dengan menolaknya
      if (changes && !steps.length && revert) steps.push(() => revert(since));
      if (steps.length) {
        undos.set(id, steps);
        if (undos.size > KEEP_UNDO) undos.delete(undos.keys().next().value);
      }
      send({ type: "result", to: from.id, id, ok: true, output, changes, undoable: steps.length > 0 });
    } catch (error) {
      send({ type: "result", to: from.id, id, ok: false, error: error.message });
    } finally {
      onActivity("");
    }
  }

  async function serveUndo({ id, from, call }) {
    try {
      for (const step of [...(undos.get(call) || [])].reverse()) await step();
      undos.delete(call);
      send({ type: "undone", to: from.id, id, ok: true });
    } catch (error) {
      send({ type: "undone", to: from.id, id, ok: false, error: error.message });
    }
  }

  // ---- sisi peminta ----

  function request(peer, message, confirm) {
    return new Promise((resolve, reject) => {
      const id = `${me.id}-${++counter}`;
      let timer;
      const entry = {
        confirm,
        // jam batas waktu dihentikan selama menunggu pengguna menjawab permintaan izin
        arm(on) {
          clearTimeout(timer);
          if (on) {
            timer = setTimeout(() => {
              calls.delete(id);
              reject(new Error(`${peer.host} tidak menjawab. Pastikan panel AI Assistant di ${peer.host} masih terbuka.`));
            }, CALL_TIMEOUT_MS);
          }
        },
        resolve,
        reject,
      };
      calls.set(id, entry);
      entry.arm(true);
      send({ ...message, to: peer.id, id });
    });
  }

  channel.onmessage = async ({ data }) => {
    const { type, from, to } = data;
    if (!from || from.id === me.id) return;
    if (type === "bye") return void peers.delete(from.id);
    if (!enabled()) return;
    if (type === "here" || type === "who") {
      peers.set(from.id, { ...from, seen: Date.now() });
      if (type === "who") announce();
      return;
    }
    if (to !== me.id) return;
    if (type === "call") return void serve(data);
    if (type === "undo") return void serveUndo(data);
    if (type === "answer") {
      asks.get(data.ask)?.(data.allowed);
      asks.delete(data.ask);
      return;
    }
    const call = calls.get(data.id);
    if (!call) return;
    if (type === "ask") {
      call.arm(false);
      const allowed = await call.confirm(data.text).catch(() => false);
      call.arm(true);
      send({ type: "answer", to: from.id, id: data.id, ask: data.ask, allowed });
    } else if (type === "result" || type === "undone") {
      call.arm(false);
      calls.delete(data.id);
      if (data.ok) call.resolve(data);
      else call.reject(new Error(data.error));
    }
  };

  const touch = () => (focused = Date.now());
  const leave = () => send({ type: "bye" });
  window.addEventListener("focus", touch);
  window.addEventListener("pointerdown", touch);
  window.addEventListener("pagehide", leave);
  const beat = setInterval(announce, BEAT_MS);
  if (enabled()) send({ type: "who" });

  return {
    stop() {
      clearInterval(beat);
      window.removeEventListener("focus", touch);
      window.removeEventListener("pointerdown", touch);
      window.removeEventListener("pagehide", leave);
      leave();
      channel.close();
    },
    // Panel aplikasi lain yang masih hidup: satu per aplikasi, yang terakhir dipakai pengguna.
    list() {
      if (!enabled()) return [];
      const byHost = new Map();
      for (const peer of peers.values()) {
        if (Date.now() - peer.seen > ALIVE_MS || peer.host === host) continue;
        const best = byHost.get(peer.host);
        if (!best || peer.focused > best.focused) byHost.set(peer.host, peer);
      }
      return [...byHost.values()];
    },
    // Menjalankan alat di panel lain. `context` adalah konteks alat panel ini (izin, pencatatan perubahan).
    async call(peer, name, args, context) {
      const result = await request(peer, { type: "call", name, args }, context.confirm);
      if (result.undoable) context.undo(() => request(peer, { type: "undo", call: result.id }, null).then(() => {}));
      else for (let i = 0; i < result.changes; i++) context.changed();
      return result.output;
    },
  };
}
