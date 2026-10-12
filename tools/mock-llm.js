// AI tiruan untuk mencoba tampilan tanpa API key (npm run preview).
// Meniru tiga format streaming yang dipakai add-in, termasuk pemanggilan alat. Pakai sebagai Base URL:
//   http://localhost:3000/mock/openai      (penyedia "Kustom", DeepSeek, dst.)
//   http://localhost:3000/mock/gemini      (penyedia "Google Gemini")
//   http://localhost:3000/mock/anthropic   (penyedia "Anthropic Claude")
//
// Kelakuannya ditentukan kata di pesan pengguna:
//   "galat"  → dijawab 401, untuk melihat tampilan error
//   "kuota"  → sekali dijawab 429 (coba lagi dalam 2 detik), lalu normal
//   "tulis"  → memanggil alat tulis (write_range / insert_content / add_slide)
//   "hapus"  → memanggil alat hapus yang butuh izin (delete_sheet / delete_slide)
//   "lintas" → memanggil alat milik aplikasi Office lain yang panelnya terbuka
//              ("lintas tulis" untuk menulis, "lintas timpa" untuk menimpa 12 sel Excel)
//   lainnya  → kalau ada alat get_*_overview, memanggilnya dulu lalu menjawab
//
// Seperti API aslinya, tiruan ini menolak (400) kalau giliran model tidak dikirim balik utuh:
// thoughtSignature Gemini, blok thinking Claude, atau tool_call_id yang tidak cocok.

let calls = 0;
let throttled = false;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const bad = (message) => Object.assign(new Error(message), { status: 400 });

const FORMATS = {
  openai: {
    names: (body) => (body.tools || []).map((t) => t.function.name),
    turns: (body) => body.messages.length,
    last(body) {
      const messages = body.messages;
      const message = messages.at(-1);
      if (message.role !== "tool") return { prompt: String(message.content) };
      const caller = messages.findLast((m) => m.role === "assistant");
      if (!caller?.tool_calls?.some((c) => c.id === message.tool_call_id)) throw bad("tool_call_id tidak cocok dengan panggilan sebelumnya.");
      return { result: message.content };
    },
    text: (text) => [{ choices: [{ index: 0, delta: { content: text } }] }],
    call: (id, name, args) => [
      { choices: [{ index: 0, delta: { tool_calls: [{ index: 0, id, type: "function", function: { name, arguments: "" } }] } }] },
      { choices: [{ index: 0, delta: { tool_calls: [{ index: 0, function: { arguments: JSON.stringify(args) } }] } }] },
      { choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }] },
    ],
    done: "[DONE]",
    models: { data: [{ id: "tiruan-kecil" }, { id: "tiruan-besar" }] },
  },
  gemini: {
    names: (body) => (body.tools?.[0]?.functionDeclarations || []).map((f) => f.name),
    turns: (body) => body.contents.length,
    last(body) {
      const part = body.contents.at(-1).parts[0];
      if (!part.functionResponse) return { prompt: String(part.text) };
      const caller = body.contents.at(-2);
      if (!caller?.parts.some((p) => p.functionCall && p.thoughtSignature)) throw bad("Function call is missing a thought_signature.");
      return { result: JSON.stringify(part.functionResponse.response) };
    },
    text: (text) => [{ candidates: [{ content: { role: "model", parts: [{ text }] } }] }],
    call: (id, name, args) => [
      { candidates: [{ content: { role: "model", parts: [{ functionCall: { name, args }, thoughtSignature: "dGlydWFu" }] }, finishReason: "STOP" }] },
    ],
    models: { models: [{ name: "models/tiruan-flash", supportedGenerationMethods: ["generateContent"] }] },
  },
  anthropic: {
    names: (body) => (body.tools || []).map((t) => t.name),
    turns: (body) => body.messages.length,
    last(body) {
      const content = body.messages.at(-1).content;
      if (!Array.isArray(content)) return { prompt: String(content) };
      const caller = body.messages.at(-2)?.content;
      if (!Array.isArray(caller) || caller[0]?.type !== "thinking" || !caller[0].signature) {
        throw bad("Expected `thinking` block before `tool_use` in the assistant turn.");
      }
      if (!caller.some((b) => b.type === "tool_use" && b.id === content[0].tool_use_id)) throw bad("tool_use_id tidak cocok.");
      return { result: content[0].content };
    },
    start: [{ type: "content_block_start", index: 0, content_block: { type: "text", text: "" } }],
    text: (text) => [{ type: "content_block_delta", index: 0, delta: { type: "text_delta", text } }],
    call: (id, name, args) => [
      { type: "content_block_start", index: 0, content_block: { type: "thinking", thinking: "", signature: "" } },
      { type: "content_block_delta", index: 0, delta: { type: "signature_delta", signature: "tiruan" } },
      { type: "content_block_stop", index: 0 },
      { type: "content_block_start", index: 1, content_block: { type: "tool_use", id, name, input: {} } },
      { type: "content_block_delta", index: 1, delta: { type: "input_json_delta", partial_json: JSON.stringify(args) } },
      { type: "content_block_stop", index: 1 },
      { type: "message_delta", delta: { stop_reason: "tool_use" } },
    ],
    models: { data: [{ id: "tiruan-claude" }] },
  },
};

const WRITES = {
  write_range: { address: "H1", values: [["Sumber", "Nilai"], ["tiruan", "42"]] },
  insert_content: { where: "end", markdown: "## Ditulis AI tiruan\n\nParagraf contoh dari agen." },
  add_slide: { title: "Slide dari AI tiruan", body: "Poin satu\nPoin dua" },
};
const DELETES = { delete_sheet: { sheet: "SheetYangTidakAda" }, delete_slide: { slide: 99 } };
const OVERWRITE = { address: "K1", values: Array.from({ length: 4 }, (_, r) => [`x${r}`, `y${r}`, `z${r}`]) };

// Alat yang dipanggil untuk pesan ini, atau null kalau cukup dijawab dengan teks.
function plan(prompt, names) {
  const first = prompt.split("\n")[0];
  if (/lintas/i.test(first)) {
    // alat milik aplikasi lain bernama "<aplikasi>__<alat>"
    const remote = (base) => names.find((name) => name.endsWith(`__${base}`));
    const write = /tulis/i.test(first) && Object.keys(WRITES).find(remote);
    // "timpa": menulis 12 sel mulai K1, untuk memancing permintaan izin menimpa data
    if (/timpa/i.test(first) && remote("write_range")) return { name: remote("write_range"), args: OVERWRITE };
    if (write) return { name: remote(write), args: WRITES[write] };
    const overview = names.find((name) => /__get_\w+_overview$/.test(name));
    if (overview) return { name: overview, args: {} };
  }
  const pick = (table) => Object.keys(table).find((name) => names.includes(name));
  if (/tulis/i.test(first) && pick(WRITES)) return { name: pick(WRITES), args: WRITES[pick(WRITES)] };
  if (/hapus/i.test(first) && pick(DELETES)) return { name: pick(DELETES), args: DELETES[pick(DELETES)] };
  const overview = names.find((name) => /^get_\w+_overview$/.test(name));
  return overview ? { name: overview, args: {} } : null;
}

const answer = (prompt, turns) =>
  [
    `Ini jawaban **tiruan** (giliran ke-${turns}) untuk: *${prompt.split("\n")[0].slice(0, 120)}*`,
    "",
    "## Contoh format",
    "",
    "1. Daftar bernomor",
    "2. Dengan `kode sebaris`",
    "   - dan butir bersarang",
    "",
    "| Bulan | Penjualan |",
    "| --- | --- |",
    "| Januari | 1200 |",
    "| Februari | 1350 |",
    "",
    "```",
    "=SUM(B2:B3)",
    "```",
  ].join("\n");

function readBody(req) {
  return new Promise((resolve) => {
    let raw = "";
    req.on("data", (chunk) => (raw += chunk));
    req.on("end", () => {
      try {
        resolve(JSON.parse(raw || "{}"));
      } catch {
        resolve({});
      }
    });
  });
}

function reject(res, status, message, headers = {}) {
  res.writeHead(status, { "Content-Type": "application/json", ...headers });
  res.end(JSON.stringify({ error: { message } }));
}

module.exports = async function mock(req, res, pathname) {
  const format = FORMATS[pathname.split("/")[2]];
  if (!format) {
    res.writeHead(404).end();
    return;
  }
  if (req.method === "GET") {
    res.writeHead(200, { "Content-Type": "application/json" }).end(JSON.stringify(format.models));
    return;
  }
  const body = await readBody(req);
  let last;
  try {
    last = format.last(body);
  } catch (error) {
    reject(res, error.status || 500, error.message);
    return;
  }
  const first = (last.prompt || "").split("\n")[0];
  if (/galat/i.test(first)) return reject(res, 401, "API key tiruan tidak valid.");
  if (/kuota/i.test(first) && !throttled) {
    throttled = true;
    return reject(res, 429, "Kuota tiruan habis.", { "Retry-After": "2" });
  }
  throttled = false;

  const step = last.prompt !== undefined ? plan(last.prompt, format.names(body)) : null;
  let frames;
  if (step) {
    const call = format.call(`call_tiruan_${++calls}`, step.name, step.args);
    // Claude membuka dengan blok thinking, jadi tidak diberi teks pembuka
    frames = format.start ? call : [...format.text("Saya periksa dulu. "), ...call];
  } else {
    const reply =
      last.result !== undefined
        ? `Alat sudah dijalankan. Hasilnya ${last.result.length} karakter, diawali: \`${last.result.slice(0, 80).replace(/`/g, "'")}\``
        : answer(last.prompt, format.turns(body));
    frames = [...(format.start || [])];
    for (let i = 0; i < reply.length; i += 7) frames.push(...format.text(reply.slice(i, i + 7)));
  }

  let open = true;
  res.on("close", () => (open = false));
  res.writeHead(200, { "Content-Type": "text/event-stream", "Cache-Control": "no-cache" });
  for (const frame of frames) {
    if (!open) return;
    res.write(`data: ${JSON.stringify(frame)}\n\n`);
    await sleep(20);
  }
  if (format.done) res.write(`data: ${format.done}\n\n`);
  res.end();
};
