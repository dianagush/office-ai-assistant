// Markdown ringkas. Satu pengurai (parseBlocks + inlineRuns) dipakai dua kali:
// renderMarkdown mengubahnya ke HTML untuk panel, dan tools/word.js mengubahnya ke paragraf Word.
// Teks selalu di-escape sebelum dijadikan HTML, jadi keluaran model tidak bisa menyisipkan tag sendiri.

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const FENCE = /^[ \t]*```/;
const HEADING = /^(#{1,6})\s+(.*)$/;
const RULE = /^\s*([-*_])(\s*\1){2,}\s*$/;
// model kadang menulis butir dengan "•" alih-alih "-": diperlakukan sama
const LIST = /^(\s*)([-*+•◦▪‣]|\d+[.)])\s+(.*)$/;
const QUOTE = /^\s*>\s?(.*)$/;
const DIVIDER = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;
// tebal | kode | miring (hanya kalau berdiri sebagai kata, supaya 2*3*4 tidak ikut) | tautan | rujukan sel
const INLINE = /\*\*([^*]+)\*\*|`([^`]+)`|(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])|\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)|\[\[([^[\]]{1,120})\]\]/g;

// Teks sebaris → potongan berformat: { text, bold?, italic?, code?, link?, cite? }
export function inlineRuns(text) {
  const runs = [];
  let last = 0;
  for (const m of text.matchAll(INLINE)) {
    if (m.index > last) runs.push({ text: text.slice(last, m.index) });
    if (m[1] !== undefined) runs.push({ text: m[1], bold: true });
    else if (m[2] !== undefined) {
      // model kadang membungkus rujukan sel dengan backtick: tetap dijadikan rujukan yang bisa diklik
      const cite = /^\[\[([^[\]]{1,120})\]\]$/.exec(m[2]);
      runs.push(cite ? { text: cite[1], cite: true } : { text: m[2], code: true });
    }
    else if (m[3] !== undefined) runs.push({ text: m[3], italic: true });
    else if (m[4] !== undefined) runs.push({ text: m[4], link: m[5] });
    else runs.push({ text: m[6], cite: true });
    last = m.index + m[0].length;
  }
  if (last < text.length) runs.push({ text: text.slice(last) });
  return runs;
}

function inlineHtml(text) {
  return inlineRuns(text)
    .map((run) => {
      const body = esc(run.text);
      if (run.bold) return `<strong>${body}</strong>`;
      if (run.code) return `<code>${body}</code>`;
      if (run.italic) return `<em>${body}</em>`;
      if (run.link) return `<a href="${esc(run.link)}" target="_blank" rel="noopener">${body}</a>`;
      // [[Sheet1!A1:B5]] → rujukan sel yang bisa diklik (ditangani task pane)
      if (run.cite) return `<a class="cite" href="#" data-ref="${body}">${body}</a>`;
      return body;
    })
    .join("");
}

const cells = (line) => line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
const isRow = (line) => !!line && line.includes("|");

// Tabel Markdown yang mulai di baris i → { rows, next }, atau null kalau bukan tabel.
export function tableAt(lines, i) {
  const divider = lines[i + 1];
  if (!isRow(lines[i]) || !isRow(divider) || !DIVIDER.test(divider)) return null;
  const rows = [cells(lines[i])];
  let next = i + 2;
  while (next < lines.length && isRow(lines[next])) rows.push(cells(lines[next++]));
  return { rows, next };
}

// Markdown → daftar blok:
//   { type: "heading", level, text } | { type: "paragraph", text } | { type: "quote", text }
//   { type: "item", ordered, depth, text } | { type: "table", rows } | { type: "code", text } | { type: "rule" }
export function parseBlocks(source) {
  const lines = source.replace(/\r\n?/g, "\n").split("\n");
  const blocks = [];
  const startsBlock = (i) =>
    FENCE.test(lines[i]) || HEADING.test(lines[i]) || RULE.test(lines[i]) || LIST.test(lines[i]) || QUOTE.test(lines[i]) || tableAt(lines, i);

  for (let i = 0; i < lines.length; ) {
    const line = lines[i];
    let m;
    let table;
    if (!line.trim()) {
      i++;
    } else if (FENCE.test(line)) {
      const body = [];
      for (i++; i < lines.length && !FENCE.test(lines[i]); i++) body.push(lines[i]);
      i++;
      blocks.push({ type: "code", text: body.join("\n") });
    } else if ((m = HEADING.exec(line))) {
      blocks.push({ type: "heading", level: m[1].length, text: m[2] });
      i++;
    } else if (RULE.test(line)) {
      blocks.push({ type: "rule" });
      i++;
    } else if ((table = tableAt(lines, i))) {
      blocks.push({ type: "table", rows: table.rows });
      i = table.next;
    } else if (LIST.test(line)) {
      const indents = [];
      for (; i < lines.length; i++) {
        m = LIST.exec(lines[i]);
        if (!m) {
          // baris kosong di antara dua butir tidak memutus daftar
          if (!lines[i].trim() && LIST.test(lines[i + 1] || "")) continue;
          break;
        }
        const indent = m[1].length;
        while (indents.length > 1 && indent < indents[indents.length - 1]) indents.pop();
        if (!indents.length || indent > indents[indents.length - 1]) indents.push(indent);
        blocks.push({ type: "item", ordered: /\d/.test(m[2]), depth: indents.length - 1, text: m[3] });
      }
    } else if (QUOTE.test(line)) {
      const quote = [];
      for (; i < lines.length && (m = QUOTE.exec(lines[i])); i++) quote.push(m[1]);
      blocks.push({ type: "quote", text: quote.join("\n") });
    } else {
      const para = [lines[i++]];
      while (i < lines.length && lines[i].trim() && !startsBlock(i)) para.push(lines[i++]);
      blocks.push({ type: "paragraph", text: para.join("\n") });
    }
  }
  return blocks;
}

const lines = (text) => text.split("\n").map(inlineHtml).join("<br>");

export function renderMarkdown(source) {
  const out = [];
  const lists = []; // tag daftar yang sedang terbuka, dari luar ke dalam
  const closeLists = (keep) => {
    while (lists.length > keep) out.push(`</li></${lists.pop()}>`);
  };
  for (const block of parseBlocks(source)) {
    if (block.type === "item") {
      const tag = block.ordered ? "ol" : "ul";
      if (block.depth < lists.length) {
        closeLists(block.depth + 1);
        // daftar bernomor langsung sesudah daftar butir (atau sebaliknya) adalah daftar baru
        if (lists[block.depth] === tag) out.push("</li><li>");
        else closeLists(block.depth);
      }
      if (block.depth >= lists.length) {
        lists.push(tag);
        out.push(`<${tag}><li>`);
      }
      out.push(inlineHtml(block.text));
      continue;
    }
    closeLists(0);
    if (block.type === "heading") {
      out.push(`<h${block.level}>${inlineHtml(block.text)}</h${block.level}>`);
    } else if (block.type === "code") {
      out.push(`<pre><code>${esc(block.text)}</code></pre>`);
    } else if (block.type === "rule") {
      out.push("<hr>");
    } else if (block.type === "quote") {
      out.push(`<blockquote>${lines(block.text)}</blockquote>`);
    } else if (block.type === "table") {
      const [head, ...body] = block.rows;
      const row = (tag) => (r) => `<tr>${r.map((c) => `<${tag}>${inlineHtml(c)}</${tag}>`).join("")}</tr>`;
      out.push(`<table><thead>${row("th")(head)}</thead><tbody>${body.map(row("td")).join("")}</tbody></table>`);
    } else {
      out.push(`<p>${lines(block.text)}</p>`);
    }
  }
  closeLists(0);
  return out.join("");
}

// Membuang penanda Markdown untuk tempat yang hanya menerima teks polos (sel Excel, kotak teks slide).
// Miring satu-bintang (*teks*) sengaja dibiarkan: bentuknya sama dengan perkalian di rumus.
export function toPlain(source) {
  return source
    .replace(/\r\n?/g, "\n")
    .replace(/^[ \t]*```.*\n?/gm, "")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/^([ \t]*)[-*+•◦▪‣]\s+/gm, "$1")
    .replace(/^[ \t]*>\s?/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, "$1 ($2)")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
