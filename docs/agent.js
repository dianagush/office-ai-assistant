// Loop agen: model menjawab, memanggil alat, menerima hasilnya, dan mengulang sampai selesai.
// Tidak tahu apa-apa soal tampilan; semua yang perlu ditunjukkan ke pengguna lewat `hooks`.

import { streamChat } from "./providers.js";

const MAX_STEPS = 40;
// Hasil alat di atas ini dipotong supaya satu pembacaan besar tidak menghabiskan konteks model.
const RESULT_LIMIT = 30000;
// Di atas ini (karakter riwayat) percakapan dipadatkan menjadi ringkasan sebelum giliran berikutnya.
const COMPACT_AT = 300000;
const RETRY_STATUS = [429, 503, 529];
const MAX_WAIT = 60;

const COMPACT_PROMPT =
  "Ringkas percakapan di atas supaya bisa dilanjutkan tanpa riwayat aslinya. Tulis: tujuan pengguna, " +
  "apa saja yang sudah dikerjakan dan diubah di dokumen (sebut lokasinya), keputusan dan preferensi pengguna, " +
  "serta apa yang masih tersisa. Jangan panggil alat.";

export class Denied extends Error {}

function sleep(seconds, signal) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, seconds * 1000);
    signal.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(new DOMException("Dibatalkan", "AbortError"));
    });
  });
}

// Pemeriksaan ringan atas argumen alat: wajib ada dan tipe dasarnya cocok.
function validate(schema, args) {
  for (const key of schema.required || []) {
    if (args[key] === undefined || args[key] === null) return `Argumen wajib "${key}" tidak ada.`;
  }
  for (const [key, value] of Object.entries(args)) {
    const expected = schema.properties?.[key]?.type;
    if (!expected || value === null) continue;
    const actual = Array.isArray(value) ? "array" : typeof value;
    const fits = expected === actual || (expected === "integer" && actual === "number") || (expected === "string" && actual !== "object" && actual !== "array");
    if (!fits) return `Argumen "${key}" harus bertipe ${expected}, bukan ${actual}.`;
  }
  return "";
}

function clip(output) {
  const text = JSON.stringify(output);
  if (text.length <= RESULT_LIMIT) return output;
  return {
    terpotong: true,
    catatan: `Hasil terlalu besar (${text.length} karakter) dan dipotong. Minta bagian yang lebih kecil.`,
    awal: text.slice(0, RESULT_LIMIT),
  };
}

// Satu permintaan ke model, dicoba ulang kalau kena batas kuota sebelum ada teks yang keluar.
async function requestTurn({ config, system, history, tools, hooks, signal }) {
  const wire = tools.map(({ name, description, parameters }) => ({ name, description, parameters }));
  for (let attempt = 0; ; attempt++) {
    let streamed = false;
    try {
      let turn;
      for await (const event of streamChat(config.kind, { ...config, system, messages: history, tools: wire, signal })) {
        if (event.text) {
          streamed = true;
          hooks.onText(event.text);
        } else if (event.turn) {
          turn = event.turn;
        }
      }
      return turn || { content: "", calls: [], stop: "end" };
    } catch (error) {
      const wait = error.retryAfter || 15 * (attempt + 1);
      if (streamed || attempt >= 2 || !RETRY_STATUS.includes(error.status) || wait > MAX_WAIT) throw error;
      hooks.onStatus(`Kuota per menit habis atau server sibuk. Mencoba lagi dalam ${Math.ceil(wait)} detik…`);
      await sleep(wait, signal);
      hooks.onStatus("");
    }
  }
}

// Menjalankan agen sampai model berhenti memanggil alat. `history` diubah di tempat dan selalu
// ditinggalkan dalam keadaan sah (giliran berselang, setiap panggilan alat punya hasilnya).
//
// hooks: onTurnStart(), onText(potongan), onAssistant(entri), onCall(call, tool) → pegangan tampilan,
//        onResult(pegangan, hasil), onStatus(teks), confirm(teks) → Promise<boolean>
// context diteruskan ke tool.run(args, context).
export async function runAgent({ config, system, history, tools, hooks, context, signal }) {
  const byName = new Map(tools.map((tool) => [tool.name, tool]));
  const fail = (call, message) => ({ id: call.id, name: call.name, output: { error: message }, error: true });
  let partial = "";
  const textHooks = { ...hooks, onText: (text) => ((partial += text), hooks.onText(text)) };

  async function execute(call) {
    const tool = byName.get(call.name);
    if (!tool) return fail(call, `Alat "${call.name}" tidak tersedia.`);
    const problem = call.invalid ? "Argumen bukan JSON yang sah." : validate(tool.parameters, call.args);
    call.label = problem ? call.name : tool.label(call.args);
    const handle = hooks.onCall(call, tool);
    let result;
    try {
      if (problem) throw new Error(problem);
      if (tool.risk === "danger" || (tool.risk === "write" && context.confirmWrites)) {
        if (!(await hooks.confirm(call.label))) throw new Denied("Pengguna menolak tindakan ini.");
      }
      result = { id: call.id, name: call.name, output: clip(await tool.run(call.args, context)) };
    } catch (error) {
      if (error.name === "AbortError") throw error;
      result = fail(call, error.message);
    }
    hooks.onResult(handle, result);
    return result;
  }

  try {
    for (let step = 0; ; step++) {
      if (step >= MAX_STEPS) {
        history.push({ role: "assistant", content: `(Berhenti setelah ${MAX_STEPS} langkah. Minta saya melanjutkan kalau belum selesai.)` });
        hooks.onAssistant(history[history.length - 1]);
        return;
      }
      partial = "";
      hooks.onTurnStart();
      const turn = await requestTurn({ config, system, history, tools, hooks: textHooks, signal });
      partial = "";
      const calls = turn.stop === "refusal" ? [] : turn.calls;
      const entry = { role: "assistant", content: turn.content, ...(calls.length && { calls }), ...(turn.raw && { raw: turn.raw }) };
      if (turn.stop === "refusal") entry.content += `${entry.content ? "\n\n" : ""}(Model menolak permintaan ini.)`;
      if (!entry.content && !calls.length) entry.content = "(Model tidak memberi jawaban.)";
      history.push(entry);
      hooks.onAssistant(entry);
      if (!calls.length) return;

      const results = [];
      history.push({ role: "tool", results });
      for (const call of calls) {
        if (turn.stop === "length") {
          // argumen bisa terpotong di tengah: jangan dijalankan
          results.push(fail(call, "Jawaban terpotong batas token sebelum argumen selesai. Ulangi dengan potongan yang lebih kecil."));
          continue;
        }
        try {
          results.push(await execute(call));
        } catch (error) {
          for (const rest of calls.slice(results.length)) results.push(fail(rest, "Dibatalkan pengguna."));
          throw error;
        }
      }
    }
  } catch (error) {
    // Tinggalkan riwayat dalam keadaan sah: teks yang sudah keluar disimpan, dan giliran alat
    // yang tidak dijawab model ditutup dengan catatan.
    const aborted = error.name === "AbortError";
    const last = history[history.length - 1];
    if (partial) history.push({ role: "assistant", content: partial });
    else if (last.role === "tool") history.push({ role: "assistant", content: aborted ? "(Dihentikan pengguna.)" : "(Terhenti karena galat.)" });
    throw error;
  }
}

const sizeOf = (history) => history.reduce((n, m) => n + JSON.stringify(m).length, 0);

export const needsCompaction = (history) => sizeOf(history) > COMPACT_AT;

// Pemadatan sederhana: seluruh riwayat diganti satu ringkasan. Tidak ada giliran lama yang diputar
// ulang, jadi tidak ada thought signature atau blok thinking yang tertinggal tanpa konteksnya.
export async function compact({ config, system, history, signal }) {
  let summary = "";
  const messages = [...history, { role: "user", content: COMPACT_PROMPT }];
  for await (const event of streamChat(config.kind, { ...config, system, messages, signal })) {
    if (event.text) summary += event.text;
  }
  if (!summary.trim()) throw new Error("Pemadatan gagal: model tidak memberi ringkasan.");
  return [
    { role: "user", content: `<ringkasan_percakapan_sebelumnya>\n${summary.trim()}\n</ringkasan_percakapan_sebelumnya>`, text: "(Percakapan sebelumnya dipadatkan menjadi ringkasan.)" },
    { role: "assistant", content: "Baik, saya lanjutkan dari ringkasan itu." },
  ];
}
