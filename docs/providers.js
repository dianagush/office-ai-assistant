// Penyedia AI. Semua dipanggil langsung dari task pane (tanpa backend), jadi API key
// tidak pernah lewat server lain selain milik penyedia itu sendiri.
//
// Menambah penyedia baru: kalau API-nya OpenAI-compatible, cukup tambah satu baris di PRESETS.
//
// Riwayat percakapan memakai satu bentuk netral, lalu tiap adapter menerjemahkannya:
//   { role: "user", content }
//   { role: "assistant", content, calls: [{ id, name, args }], raw }
//   { role: "tool", results: [{ id, name, output, error }] }
// `raw` menyimpan giliran model persis seperti diterima (thought signature Gemini, blok thinking
// Claude, reasoning DeepSeek). Penyedia mewajibkan bagian itu dikirim balik utuh selama alat dipakai.

export const PRESETS = {
  gemini: {
    label: "Google Gemini",
    kind: "gemini",
    baseUrl: "https://generativelanguage.googleapis.com/v1beta",
    // Flash-Lite sebagai bawaan: mode agen memakai beberapa permintaan per tugas, sedangkan kuota paket gratis
    // untuk Flash sangat kecil (20 permintaan per hari saat diuji Oktober 2026). Yang punya paket berbayar
    // bisa pindah ke Flash atau Pro.
    model: "gemini-flash-lite-latest",
    keyUrl: "https://aistudio.google.com/apikey",
    note: "Punya paket gratis, tetapi kuotanya dihitung per model per hari dan mode agen memakai beberapa permintaan per tugas. Di paket gratis, Google memakai isi permintaan untuk meningkatkan produknya, jadi hindari dokumen rahasia.",
  },
  deepseek: {
    label: "DeepSeek",
    kind: "openai",
    baseUrl: "https://api.deepseek.com",
    model: "deepseek-flash",
    keyUrl: "https://platform.deepseek.com/api_keys",
    note: "Prabayar: isi saldo dulu, lalu dipotong sesuai pemakaian.",
  },
  anthropic: {
    label: "Anthropic Claude",
    kind: "anthropic",
    baseUrl: "https://api.anthropic.com/v1",
    model: "claude-opus-5-5",
    keyUrl: "https://platform.claude.com/settings/keys",
  },
  openai: {
    label: "OpenAI",
    kind: "openai",
    baseUrl: "https://api.openai.com/v1",
    model: "",
    keyUrl: "https://platform.openai.com/api-keys",
  },
  openrouter: {
    label: "OpenRouter",
    kind: "openai",
    baseUrl: "https://openrouter.ai/api/v1",
    model: "",
    keyUrl: "https://openrouter.ai/keys",
    note: "Satu API key untuk banyak model, termasuk beberapa model gratis.",
  },
  groq: {
    label: "Groq",
    kind: "openai",
    baseUrl: "https://api.groq.com/openai/v1",
    model: "",
    keyUrl: "https://console.groq.com/keys",
  },
  ollama: {
    label: "Ollama (lokal)",
    kind: "openai",
    baseUrl: "http://localhost:11434/v1",
    model: "",
    keyOptional: true,
    note: "Model berjalan di komputer sendiri, tanpa biaya dan tanpa internet. Jalankan Ollama dengan OLLAMA_ORIGINS=* supaya add-in boleh mengaksesnya.",
  },
  ninerouter: {
    label: "9router (lokal)",
    kind: "openai",
    baseUrl: "http://localhost:20128/v1",
    model: "",
    note: "Jalankan 9router dulu (npx 9router) dan hubungkan minimal satu provider di dashboard-nya. Secara bawaan 9router mewajibkan API key: buat atau salin dari dashboard (http://localhost:20128/dashboard) lalu tempel di sini. Kalau fitur key dimatikan di 9router, isi teks apa saja. Nama model memakai awalan provider (mis. cc/claude-opus-4-7), atau nama combo buatan sendiri.",
  },
  custom: {
    label: "Kustom (OpenAI-compatible)",
    kind: "openai",
    baseUrl: "",
    model: "",
    keyOptional: true,
    note: "Untuk layanan apa pun yang punya endpoint /chat/completions dan mengizinkan akses dari browser (CORS).",
  },
};

// Batas menunggu. Model yang kelebihan beban kadang menerima sambungan lalu diam; tanpa batas ini
// panel akan menunggu selamanya.
const FIRST_BYTE_MS = 90000; // sampai server mulai menjawab
const IDLE_MS = 180000; // jeda terlama di tengah jawaban

// fetch untuk permintaan ke model: menyerah kalau server tidak mulai menjawab dalam FIRST_BYTE_MS.
async function post(url, init) {
  const controller = new AbortController();
  const forward = () => controller.abort();
  if (init.signal?.aborted) forward();
  else init.signal?.addEventListener("abort", forward);
  let late = false;
  const timer = setTimeout(() => {
    late = true;
    controller.abort();
  }, FIRST_BYTE_MS);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } catch (error) {
    if (!late) throw error;
    throw new Error(
      `Model tidak mulai menjawab dalam ${FIRST_BYTE_MS / 1000} detik, kemungkinan sedang kelebihan beban. Coba lagi, atau pilih model lain di Pengaturan.`,
    );
  } finally {
    clearTimeout(timer);
  }
}

// Membaca stream Server-Sent Events dan menghasilkan isi tiap baris "data:".
async function* sse(res) {
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const read = () =>
    new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reader.cancel().catch(() => {});
        reject(new Error("Jawaban terhenti di tengah jalan: server diam terlalu lama. Coba lagi."));
      }, IDLE_MS);
      reader.read().then(resolve, reject).finally(() => clearTimeout(timer));
    });
  for (;;) {
    const { value, done } = await read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let end;
    while ((end = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, end).replace(/\r$/, "");
      buffer = buffer.slice(end + 1);
      if (line.startsWith("data:")) yield line.slice(5).trim();
    }
  }
  if (buffer.startsWith("data:")) yield buffer.slice(5).trim();
}

// Durasi seperti "23s", "57168.5s", atau "15h52m48s" → detik.
function seconds(duration) {
  const part = (unit) => parseFloat((new RegExp(`([0-9.]+)${unit}`).exec(duration) || [])[1]) || 0;
  return part("h") * 3600 + part("m") * 60 + part("s");
}

// Error dari respons gagal. `status` dan `retryAfter` (detik) dipakai untuk mencoba ulang saat kena batas kuota.
async function failure(res) {
  let detail = await res.text().catch(() => "");
  let retryAfter = Number(res.headers.get("retry-after")) || 0;
  try {
    const json = JSON.parse(detail);
    const error = (Array.isArray(json) ? json[0] : json).error;
    const retryInfo = error?.details?.find?.((d) => String(d["@type"]).endsWith("RetryInfo"));
    if (retryInfo) retryAfter = seconds(String(retryInfo.retryDelay)) || retryAfter;
    detail = error?.message || (typeof error === "string" ? error : "") || json.message || detail;
  } catch {
    // isi bukan JSON: tampilkan apa adanya
  }
  const message = `${res.status} ${res.statusText}${detail ? ` — ${String(detail).slice(0, 400)}` : ""}`.trim();
  return Object.assign(new Error(message), { status: res.status, retryAfter });
}

async function getJson(url, headers, signal) {
  const res = await fetch(url, { headers, signal });
  if (!res.ok) throw await failure(res);
  return res.json();
}

let localIds = 0;
// ID panggilan alat buatan sendiri, untuk penyedia yang tidak memberi ID.
const localId = () => `local_${Date.now().toString(36)}_${localIds++}`;

function parseArgs(json) {
  if (!json) return { args: {} };
  try {
    const args = JSON.parse(json);
    return args && typeof args === "object" && !Array.isArray(args) ? { args } : { args: {}, invalid: true };
  } catch {
    return { args: {}, invalid: true };
  }
}

const asObject = (output) => (output && typeof output === "object" && !Array.isArray(output) ? output : { result: output });

// ---------- OpenAI-compatible (DeepSeek, OpenAI, OpenRouter, Groq, Ollama, …) ----------

const bearer = (apiKey) => (apiKey ? { Authorization: `Bearer ${apiKey}` } : {});

function toOpenAI(messages, origin) {
  const out = [];
  for (const m of messages) {
    if (m.role === "user") {
      out.push({ role: "user", content: m.content });
    } else if (m.role === "assistant") {
      const message = { role: "assistant", content: m.content || (m.calls?.length ? null : "") };
      if (m.calls?.length) {
        message.tool_calls = m.calls.map((c) => ({
          id: c.id,
          type: "function",
          function: { name: c.name, arguments: JSON.stringify(c.args) },
        }));
      }
      // DeepSeek mewajibkan reasoning_content dikirim balik di setiap giliran selama alat dipakai.
      // Hanya dikirim ke layanan yang menghasilkannya; layanan lain bisa menolak field ini.
      if (m.raw?.kind === "openai" && m.raw.origin === origin) message.reasoning_content = m.raw.reasoning;
      out.push(message);
    } else {
      for (const r of m.results) out.push({ role: "tool", tool_call_id: r.id, content: JSON.stringify(r.output) });
    }
  }
  return out;
}

const openai = {
  async *stream({ baseUrl, apiKey, model, system, messages, tools, signal }) {
    const res = await post(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...bearer(apiKey) },
      body: JSON.stringify({
        model,
        stream: true,
        messages: [...(system ? [{ role: "system", content: system }] : []), ...toOpenAI(messages, baseUrl)],
        ...(tools.length && {
          tools: tools.map((t) => ({
            type: "function",
            function: { name: t.name, description: t.description, parameters: t.parameters },
          })),
        }),
      }),
      signal,
    });
    if (!res.ok) throw await failure(res);
    let content = "";
    let reasoning = "";
    let finish = "";
    const calls = [];
    for await (const data of sse(res)) {
      if (!data) continue;
      if (data === "[DONE]") break;
      const json = JSON.parse(data);
      if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));
      const choice = json.choices?.[0];
      if (!choice) continue;
      const delta = choice.delta || {};
      if (delta.content) {
        content += delta.content;
        yield { text: delta.content };
      }
      if (delta.reasoning_content) reasoning += delta.reasoning_content;
      for (const part of delta.tool_calls || []) {
        // tanpa `index`, anggap potongan ber-ID sebagai panggilan baru dan sisanya lanjutan yang terakhir
        const at = part.index ?? (part.id ? calls.length : Math.max(0, calls.length - 1));
        const call = (calls[at] ||= { id: "", name: "", json: "" });
        if (part.id) call.id = part.id;
        if (part.function?.name) call.name = part.function.name;
        if (part.function?.arguments) call.json += part.function.arguments;
      }
      if (choice.finish_reason) finish = choice.finish_reason;
    }
    yield {
      turn: {
        content,
        calls: calls.filter(Boolean).map((c) => ({ id: c.id || localId(), name: c.name, ...parseArgs(c.json) })),
        stop: finish === "length" ? "length" : "end",
        raw: reasoning ? { kind: "openai", reasoning, origin: baseUrl } : undefined,
      },
    };
  },
  async listModels({ baseUrl, apiKey, signal }) {
    const json = await getJson(`${baseUrl}/models`, bearer(apiKey), signal);
    return (json.data || []).map((m) => m.id).sort();
  },
};

// ---------- Google Gemini (generateContent) ----------

// Nilai resmi untuk melewati pemeriksaan thought signature pada panggilan alat yang tidak
// berasal dari Gemini (riwayat dari penyedia lain).
const NO_SIGNATURE = "skip_thought_signature_validator";

// origin: alamat API yang sedang dipakai. Giliran model yang dihasilkan di alamat itu dikirim balik
// persis (thoughtSignature menempel di part-nya). Giliran dari tempat lain disusun ulang dari isinya,
// karena signature-nya tidak berlaku di sini dan akan ditolak sebagai rusak. origin null = susun ulang semua.
function toGemini(messages, origin) {
  const contents = [];
  for (const m of messages) {
    if (m.role === "user") {
      contents.push({ role: "user", parts: [{ text: m.content }] });
    } else if (m.role === "assistant") {
      const parts =
        m.raw?.kind === "gemini" && m.raw.origin === origin
          ? m.raw.parts
          : [
              ...(m.content ? [{ text: m.content }] : []),
              ...(m.calls || []).map((c, i) => ({
                functionCall: { name: c.name, args: c.args },
                ...(i === 0 && { thoughtSignature: NO_SIGNATURE }),
              })),
            ];
      if (parts.length) contents.push({ role: "model", parts });
    } else {
      contents.push({
        role: "user",
        parts: m.results.map((r) => ({
          functionResponse: {
            ...(!r.id.startsWith("local_") && { id: r.id }),
            name: r.name,
            response: asObject(r.output),
          },
        })),
      });
    }
  }
  return contents;
}

const isPlainText = (part) => typeof part.text === "string" && Object.keys(part).length === 1;

const gemini = {
  async *stream({ baseUrl, apiKey, model, system, messages, tools, signal }) {
    const id = encodeURIComponent(model.replace(/^models\//, ""));
    const request = (origin) =>
      post(`${baseUrl}/models/${id}:streamGenerateContent?alt=sse`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": apiKey },
        body: JSON.stringify({
          ...(system && { systemInstruction: { parts: [{ text: system }] } }),
          contents: toGemini(messages, origin),
          ...(tools.length && {
            tools: [
              {
                functionDeclarations: tools.map((t) => ({
                  name: t.name,
                  description: t.description,
                  // alat tanpa parameter: Gemini menolak skema objek kosong, jadi dihilangkan
                  ...(Object.keys(t.parameters.properties || {}).length && { parameters: t.parameters }),
                })),
              },
            ],
          }),
        }),
        signal,
      });

    let res = await request(baseUrl);
    if (res.status === 400) {
      // Signature bisa kedaluwarsa atau tidak cocok (mis. model diganti di tengah percakapan).
      // Pemulihannya: kirim ulang sekali dengan riwayat yang disusun ulang tanpa signature asli.
      const error = await failure(res);
      if (!/thought.?signature/i.test(error.message)) throw error;
      res = await request(null);
    }
    if (!res.ok) throw await failure(res);
    let content = "";
    let stop = "end";
    const parts = [];
    const calls = [];
    for await (const data of sse(res)) {
      if (!data) continue;
      const json = JSON.parse(data);
      if (json.error) throw new Error(json.error.message);
      const blocked = json.promptFeedback?.blockReason;
      if (blocked) throw new Error(`Permintaan diblokir Gemini (${blocked}).`);
      const candidate = json.candidates?.[0];
      for (const part of candidate?.content?.parts || []) {
        const last = parts[parts.length - 1];
        // hanya teks polos yang digabung; part ber-signature harus tetap utuh
        if (last && isPlainText(last) && isPlainText(part)) last.text += part.text;
        else parts.push({ ...part });
        if (part.functionCall) {
          calls.push({ id: part.functionCall.id || localId(), name: part.functionCall.name, args: part.functionCall.args || {} });
        } else if (part.text && !part.thought) {
          content += part.text;
          yield { text: part.text };
        }
      }
      const finish = candidate?.finishReason;
      if (finish === "MAX_TOKENS") stop = "length";
      else if (finish && finish !== "STOP") throw new Error(`Jawaban dihentikan Gemini (${finish}).`);
    }
    yield { turn: { content, calls, stop, raw: { kind: "gemini", origin: baseUrl, parts } } };
  },
  async listModels({ baseUrl, apiKey, signal }) {
    const json = await getJson(`${baseUrl}/models?pageSize=1000`, { "x-goog-api-key": apiKey }, signal);
    return (json.models || [])
      .filter((m) => m.supportedGenerationMethods?.includes("generateContent"))
      .map((m) => m.name.replace(/^models\//, ""))
      .sort();
  },
};

// ---------- Anthropic Claude (Messages API) ----------

// Model yang mendukung pengalihan otomatis di sisi server saat sebuah permintaan ditolak
// pengaman (parameter `fallbacks`). Model lain akan menolak parameter ini.
const CLAUDE_FALLBACK = /^claude-(fable-5-1|opus-5-5|opus-5|sonnet-5-5)$/;
const ANTHROPIC_API = /^https:\/\/api\.anthropic\.com\//;
const isThinking = (block) => block.type === "thinking" || block.type === "redacted_thinking";

const claudeHeaders = (apiKey) => ({
  "x-api-key": apiKey,
  "anthropic-version": "2023-06-01",
  "anthropic-dangerous-direct-browser-access": "true",
});

function toClaude(messages, keepThinking) {
  const out = [];
  for (const m of messages) {
    if (m.role === "user") {
      out.push({ role: "user", content: m.content });
    } else if (m.role === "assistant") {
      const blocks =
        m.raw?.kind === "anthropic"
          ? m.raw.blocks.filter((b) => keepThinking || !isThinking(b))
          : [
              ...(m.content ? [{ type: "text", text: m.content }] : []),
              ...(m.calls || []).map((c) => ({ type: "tool_use", id: c.id, name: c.name, input: c.args })),
            ];
      if (blocks.length) out.push({ role: "assistant", content: blocks });
    } else {
      out.push({
        role: "user",
        content: m.results.map((r) => ({
          type: "tool_result",
          tool_use_id: r.id,
          content: JSON.stringify(r.output),
          ...(r.error && { is_error: true }),
        })),
      });
    }
  }
  return out;
}

const anthropic = {
  async *stream(options) {
    const { baseUrl, apiKey, model, system, messages, tools, signal } = options;
    const fallback = CLAUDE_FALLBACK.test(model);
    // Lewat proxy, input alat dibiarkan di-buffer server (lebih aman untuk proxy yang belum kenal field ini).
    const eager = ANTHROPIC_API.test(`${baseUrl}/`);
    const request = (keepThinking) =>
      post(`${baseUrl}/messages`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...claudeHeaders(apiKey),
          ...(fallback && { "anthropic-beta": "server-side-fallback-2026-07-01" }),
        },
        body: JSON.stringify({
          model,
          max_tokens: 64000,
          stream: true,
          ...(system && { system }),
          ...(fallback && { fallbacks: "default" }),
          ...(tools.length && {
            tools: tools.map((t) => ({
              name: t.name,
              description: t.description,
              input_schema: t.parameters,
              ...(eager && { eager_input_streaming: true }),
            })),
          }),
          messages: toClaude(messages, keepThinking),
        }),
        signal,
      });

    let res = await request(true);
    if (res.status === 400) {
      // Blok thinking terikat pada riwayat persisnya. Kalau riwayat berubah (mis. instruksi diedit di
      // tengah percakapan) API menolaknya; pemulihannya: kirim ulang sekali tanpa blok thinking.
      const error = await failure(res);
      if (!/signature|different conversation/i.test(error.message)) throw error;
      res = await request(false);
    }
    if (!res.ok) throw await failure(res);

    let content = "";
    let stop = "end";
    const blocks = [];
    const partial = [];
    for await (const data of sse(res)) {
      if (!data) continue;
      const event = JSON.parse(data);
      const delta = event.delta || {};
      if (event.type === "content_block_start") {
        blocks[event.index] = { ...event.content_block };
        partial[event.index] = "";
      } else if (event.type === "content_block_delta") {
        const block = blocks[event.index];
        if (delta.type === "text_delta") {
          block.text += delta.text;
          content += delta.text;
          yield { text: delta.text };
        } else if (delta.type === "thinking_delta") {
          block.thinking += delta.thinking;
        } else if (delta.type === "signature_delta") {
          block.signature = delta.signature;
        } else if (delta.type === "input_json_delta") {
          partial[event.index] += delta.partial_json;
        }
      } else if (event.type === "content_block_stop") {
        const block = blocks[event.index];
        if (block?.type === "tool_use") {
          const parsed = parseArgs(partial[event.index]);
          block.input = parsed.args;
          block.invalid = parsed.invalid;
        }
      } else if (event.type === "message_delta") {
        if (delta.stop_reason === "refusal") stop = "refusal";
        else if (delta.stop_reason === "max_tokens") stop = "length";
      } else if (event.type === "error") {
        throw new Error(event.error?.message || "Stream terputus.");
      }
    }
    const done = blocks.filter(Boolean);
    const calls = done
      .filter((b) => b.type === "tool_use")
      .map((b) => ({ id: b.id, name: b.name, args: b.input || {}, ...(b.invalid && { invalid: true }) }));
    for (const block of done) delete block.invalid;
    yield { turn: { content, calls, stop, raw: { kind: "anthropic", blocks: done } } };
  },
  async listModels({ baseUrl, apiKey, signal }) {
    const json = await getJson(`${baseUrl}/models?limit=1000`, claudeHeaders(apiKey), signal);
    return (json.data || []).map((m) => m.id);
  },
};

const ADAPTERS = { openai, gemini, anthropic };
const trimUrl = (options) => ({ ...options, baseUrl: options.baseUrl.replace(/\/+$/, "") });

// Satu giliran model. Menghasilkan { text } untuk tiap potongan teks, lalu sekali di akhir
// { turn: { content, calls, stop: "end" | "length" | "refusal", raw } }.
// options: { baseUrl, apiKey, model, system, messages (bentuk netral), tools: [{ name, description, parameters }], signal }
export const streamChat = (kind, options) => ADAPTERS[kind].stream(trimUrl({ tools: [], ...options }));

// Daftar ID model yang bisa dipakai API key ini.
export const listModels = (kind, options) => ADAPTERS[kind].listModels(trimUrl(options));
