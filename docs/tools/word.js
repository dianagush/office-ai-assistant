/* global Office, Word */
// Alat agen untuk Word. Bentuk alat sama dengan tools/excel.js.
// Paragraf dirujuk dengan nomor urut mulai 0. Nomor bergeser setelah menyisip atau menghapus,
// jadi model diminta membaca ulang sesudah perubahan struktur.
//
// Kalau context.trackChanges aktif, setiap perubahan dicatat sebagai tracked changes, sehingga
// pengguna bisa menerima atau menolaknya satu per satu di Word.

import { inlineRuns, parseBlocks, toPlain } from "../markdown.js";

const MAX_TEXT = 20000; // karakter per pembacaan
const MAX_MATCHES = 30;
const CONFIRM_DELETE = 10; // menghapus lebih dari sekian paragraf butuh izin
// "ListBullet" dan "ListNumber" bukan gaya bawaan Word API; keduanya dibuat sebagai daftar sungguhan (lihat makeList).
const LIST_STYLES = ["ListBullet", "ListNumber"];
const STYLES = ["Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Title", "Subtitle", "Quote", ...LIST_STYLES];

const text = (description) => ({ type: "string", description });
const int = (description) => ({ type: "integer", description });
const flag = (description) => ({ type: "boolean", description });

function define(name, risk, description, properties, required, label, run) {
  return { name, risk, description, parameters: { type: "object", properties, ...(required.length && { required }) }, label, run };
}

const supports = (version) => Office.context.requirements.isSetSupported("WordApi", version);

// Menjalankan perubahan dengan pelacakan aktif, lalu mengembalikan mode pelacakan semula.
function edit(context, work) {
  return Word.run(async (ctx) => {
    const track = context.trackChanges && supports("1.4");
    let previous;
    if (track) {
      ctx.document.load("changeTrackingMode");
      await ctx.sync();
      previous = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }
    try {
      const result = await work(ctx);
      await ctx.sync();
      context.changed();
      return result;
    } finally {
      if (track) {
        ctx.document.changeTrackingMode = previous;
        await ctx.sync().catch(() => {});
      }
    }
  });
}

async function paragraphsOf(ctx, fields = "items/text") {
  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load(fields);
  await ctx.sync();
  return paragraphs.items;
}

function at(items, index) {
  if (!Number.isInteger(index) || index < 0 || index >= items.length) {
    throw new Error(`Paragraf ${index} tidak ada. Dokumen punya ${items.length} paragraf (nomor 0 sampai ${items.length - 1}).`);
  }
  return items[index];
}

async function find(ctx, query, matchCase = true) {
  if (String(query).length > 255) throw new Error("Teks pencarian paling panjang 255 karakter. Pakai potongan yang lebih pendek tetapi unik.");
  const results = ctx.document.body.search(String(query), { matchCase });
  results.load("items/text");
  await ctx.sync();
  return results.items;
}

const notFound = (query) =>
  new Error(`Teks "${String(query).slice(0, 80)}" tidak ditemukan. Salin persis dari dokumen, dalam satu paragraf, tanpa tanda Markdown.`);

async function overview() {
  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load("items/text,items/outlineLevel");
    const tables = ctx.document.body.tables;
    tables.load("items/rowCount");
    const selection = ctx.document.getSelection();
    selection.load("text");
    if (supports("1.4")) ctx.document.load("changeTrackingMode");
    await ctx.sync();
    const items = paragraphs.items;
    return {
      paragraphs: items.length,
      words: items.reduce((n, p) => n + (p.text.trim() ? p.text.trim().split(/\s+/).length : 0), 0),
      headings: items
        .map((p, index) => ({ index, level: p.outlineLevel, text: p.text.slice(0, 120) }))
        .filter((h) => h.level >= 1 && h.level <= 9 && h.text.trim()),
      tables: tables.items.length,
      selection: selection.text.slice(0, 500),
      ...(supports("1.4") && { change_tracking: ctx.document.changeTrackingMode }),
    };
  });
}

async function readParagraphs({ start, count }) {
  return Word.run(async (ctx) => {
    const items = await paragraphsOf(ctx, "items/text,items/outlineLevel,items/tableNestingLevel,items/isListItem");
    const from = Math.max(0, start || 0);
    const picked = [];
    let size = 0;
    for (let i = from; i < items.length && picked.length < (count || 60); i++) {
      size += items[i].text.length;
      if (size > MAX_TEXT && picked.length) break;
      picked.push(i);
    }
    // penanda daftar ("•", "1.") ikut dibaca, supaya jelas paragraf mana yang sudah berupa butir daftar
    const markers = picked.map((i) => {
      if (!items[i].isListItem) return null;
      const item = items[i].listItem;
      item.load("listString,level");
      return item;
    });
    await ctx.sync();
    const out = picked.map((i, n) => ({
      index: i,
      text: items[i].text,
      ...(items[i].outlineLevel <= 9 && { heading: items[i].outlineLevel }),
      // butir (bukan nomor) memakai glif font simbol yang tidak terbaca sebagai teks: tampilkan sebagai "•"
      ...(markers[n] && { list: /[0-9a-z]/i.test(markers[n].listString) ? markers[n].listString : "•", ...(markers[n].level > 0 && { level: markers[n].level }) }),
      ...(items[i].tableNestingLevel > 0 && { in_table: true }),
    }));
    const next = from + out.length;
    return { total_paragraphs: items.length, paragraphs: out, ...(next < items.length && { next_start: next }) };
  });
}

async function searchText({ query, match_case }) {
  return Word.run(async (ctx) => {
    const found = await find(ctx, query, !!match_case);
    const shown = found.slice(0, MAX_MATCHES).map((range) => {
      const paragraph = range.paragraphs.getFirst();
      paragraph.load("text");
      return { range, paragraph };
    });
    await ctx.sync();
    return { count: found.length, matches: shown.map(({ range, paragraph }) => ({ text: range.text, paragraph: paragraph.text.slice(0, 300) })) };
  });
}

function replaceText({ replacements, match_case }, context) {
  return edit(context, async (ctx) => {
    const results = [];
    for (const { find: query, replace, all } of replacements) {
      const found = await find(ctx, query, match_case !== false);
      const targets = all ? found : found.slice(0, 1);
      for (const range of targets) range.insertText(String(replace), Word.InsertLocation.replace);
      // diselesaikan satu per satu, supaya pencarian berikutnya melihat teks yang sudah berubah
      await ctx.sync();
      results.push({ find: String(query).slice(0, 60), replaced: targets.length, ...(!found.length && { error: "tidak ditemukan" }) });
    }
    if (results.every((result) => !result.replaced)) throw notFound(replacements[0]?.find);
    return { results };
  });
}

function editParagraphs({ edits }, context) {
  return edit(context, async (ctx) => {
    const items = await paragraphsOf(ctx);
    for (const { index } of edits) at(items, index);
    for (const { index, text: value } of edits) items[index].insertText(toPlain(String(value)), Word.InsertLocation.replace);
    return { edited_paragraphs: edits.map((e) => e.index) };
  });
}

function styleOf(block) {
  if (block.type === "heading") return `Heading${Math.min(block.level, 9)}`;
  if (block.type === "quote") return "Quote";
  return "Normal";
}

const rectangular = (rows) => {
  const width = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => [...row, ...Array(width - row.length).fill("")].map((cell) => String(cell ?? "")));
};

// Mengisi satu paragraf dengan isi sebuah blok. Mengembalikan true kalau paragraf berakhir dengan
// teks tebal/miring, karena paragraf berikutnya mewarisi format itu dan harus menetralkannya.
function writeBlock(paragraph, block, inherited) {
  // butir daftar memakai gaya yang diberikan Word saat daftar dibuat
  if (block.type !== "item") paragraph.styleBuiltIn = styleOf(block);
  if (block.type === "rule") return false;
  if (block.type === "code") {
    paragraph.insertText(block.text, Word.InsertLocation.end);
    return false;
  }
  // judul tidak dinetralkan: tebalnya berasal dari gaya judul itu sendiri
  let dirty = inherited && block.type !== "heading";
  let styled = false;
  for (const run of inlineRuns(block.text.replace(/\n/g, " "))) {
    const range = paragraph.insertText(run.link ? `${run.text} (${run.link})` : run.text, Word.InsertLocation.end);
    styled = !!(run.bold || run.italic);
    if (run.bold || dirty) range.font.bold = !!run.bold;
    if (run.italic || dirty) range.font.italic = !!run.italic;
    dirty = dirty || styled;
  }
  return styled;
}

// Tabel yang baru disisipkan mewarisi gaya paragraf di atasnya: sesudah judul, tiap selnya bergaya
// judul; sesudah butir daftar, tiap selnya ikut bernomor. Begitu pula paragraf kosong yang Word
// buat sesudah tabel di akhir dokumen. Semuanya dikembalikan ke gaya biasa.
async function tidyTable(ctx, table, header) {
  const inside = table.getRange().paragraphs;
  inside.load("items/isListItem");
  const next = table.getParagraphAfterOrNullObject();
  next.load("isListItem,text");
  await ctx.sync();
  const strays = !next.isNullObject && !next.text.trim() ? [...inside.items, next] : inside.items;
  await detach(ctx, strays);
  for (const paragraph of strays) paragraph.styleBuiltIn = "Normal";
  if (header) {
    table.headerRowCount = 1;
    table.rows.getFirst().font.bold = true;
  }
}

// Menyisipkan blok-blok Markdown sebagai paragraf Word yang memakai gaya bawaan dokumen (Heading,
// Normal, List Bullet, …), satu per satu di samping `anchor`. Sengaja tidak lewat insertHtml: Word
// meleburkan blok HTML terakhir ke paragraf tetangganya dan memberi gaya "Normal (Web)".
// Paragraf acuan yang kosong dipakai untuk blok pertama, supaya tidak tertinggal baris kosong.
//
// Daftar: paragraf baru mewarisi keanggotaan daftar dari paragraf di sebelahnya. Jadi butir pertama
// memulai daftar baru (kecuali menyambung daftar yang sudah ada), butir berikutnya tinggal mewarisi,
// dan paragraf biasa sesudah daftar harus dilepas dari daftar itu.
async function insertBlocks(ctx, anchor, location, blocks, { empty, inList }) {
  let cursor = anchor;
  let where = location;
  let inherited = false;
  let list = null; // daftar yang dimulai di sini: { object, ordered, levels }
  let listed = inList; // apakah paragraf baru akan mewarisi daftar

  const paragraphFor = (i) => (i === 0 && empty ? anchor : cursor.insertParagraph("", where));

  for (const [i, block] of blocks.entries()) {
    if (block.type === "table") {
      const rows = rectangular(block.rows.map((row) => row.map(toPlain)));
      const table = cursor.insertTable(rows.length, rows[0].length, where, rows);
      await tidyTable(ctx, table, true);
      cursor = table;
      listed = false;
      list = null;
    } else if (block.type === "item") {
      const paragraph = paragraphFor(i);
      // daftar bernomor sesudah daftar butir (atau sebaliknya) adalah daftar baru
      const switched = list && block.depth === 0 && list.ordered !== block.ordered;
      if (!listed || switched) {
        if (listed) paragraph.detachFromList();
        // butir pertama sesudah judul mewarisi gaya judul; dikembalikan ke biasa sebelum daftar dimulai
        paragraph.styleBuiltIn = "Normal";
        list = { object: paragraph.startNewList(), ordered: block.ordered, levels: new Set() };
      }
      if (list && !list.levels.has(block.depth)) {
        list.levels.add(block.depth);
        if (block.ordered) list.object.setLevelNumbering(block.depth, Word.ListNumbering.arabic, [block.depth, "."]);
        else list.object.setLevelBullet(block.depth, Word.ListBullet.solid);
      }
      paragraph.listItem.level = block.depth;
      inherited = writeBlock(paragraph, block, inherited);
      cursor = paragraph;
      listed = true;
    } else {
      // blok kode: satu paragraf per baris
      const pieces = block.type === "code" ? block.text.split("\n") : [block.text];
      pieces.forEach((piece, n) => {
        const paragraph = n === 0 ? paragraphFor(i) : cursor.insertParagraph("", Word.InsertLocation.after);
        if (listed) paragraph.detachFromList();
        inherited = writeBlock(paragraph, { ...block, text: piece }, inherited);
        cursor = paragraph;
        listed = false;
      });
      list = null;
    }
    where = Word.InsertLocation.after;
  }
  return blocks.length;
}

async function insertMarkdown(ctx, anchor, location, markdown) {
  anchor.load("text,isListItem");
  await ctx.sync();
  return insertBlocks(ctx, anchor, location, parseBlocks(String(markdown)), { empty: anchor.text.trim() === "", inList: anchor.isListItem });
}

function insertContent({ where, index, markdown }, context) {
  return edit(context, async (ctx) => {
    const { before, after } = Word.InsertLocation;
    const paragraphs = ctx.document.body.paragraphs;
    let blocks;
    if (where === "selection") {
      blocks = await insertMarkdown(ctx, ctx.document.getSelection().paragraphs.getLast(), after, markdown);
    } else if (where === "start") {
      blocks = await insertMarkdown(ctx, paragraphs.getFirst(), before, markdown);
    } else if (where === "after_paragraph" || where === "before_paragraph") {
      blocks = await insertMarkdown(ctx, at(await paragraphsOf(ctx), index), where === "after_paragraph" ? after : before, markdown);
    } else {
      blocks = await insertMarkdown(ctx, paragraphs.getLast(), after, markdown);
    }
    return {
      inserted_blocks: blocks,
      where: where || "end",
      note: "Sudah tersisip sebagai judul, daftar, dan tabel asli Word; tidak perlu dibaca ulang untuk memeriksa. Nomor paragraf sesudah titik sisip bergeser.",
    };
  });
}

// Menulis ke pilihan pengguna. Satu baris disisipkan sebaris (format sekitarnya tetap);
// beberapa blok menjadi paragraf-paragraf baru.
async function writeSelection(ctx, markdown, replace) {
  const selection = ctx.document.getSelection();
  const blocks = parseBlocks(String(markdown).trim());
  const inline = blocks.length <= 1 && !["table", "code"].includes(blocks[0]?.type);
  if (inline) {
    selection.insertText(toPlain(String(markdown).trim()), replace ? Word.InsertLocation.replace : Word.InsertLocation.end);
    return;
  }
  const anchor = replace ? selection.paragraphs.getFirst() : selection.paragraphs.getLast();
  if (replace) selection.delete();
  anchor.load("text,isListItem");
  await ctx.sync();
  await insertBlocks(ctx, anchor, Word.InsertLocation.after, blocks, { empty: anchor.text.trim() === "", inList: anchor.isListItem });
}

// Dipakai tombol "Sisipkan" / "Ganti pilihan" di panel (cara manual, tanpa agen).
export function insertAtSelection(markdown, replace) {
  return Word.run(async (ctx) => {
    await writeSelection(ctx, markdown, replace);
    await ctx.sync();
  });
}

async function deleteParagraphs({ start, count }, context) {
  const total = count || 1;
  if (total > CONFIRM_DELETE && !(await context.confirm(`Menghapus ${total} paragraf mulai paragraf ${start}`))) {
    throw new Error("Pengguna menolak tindakan ini.");
  }
  return edit(context, async (ctx) => {
    const items = await paragraphsOf(ctx);
    at(items, start);
    const last = Math.min(items.length, start + total) - 1;
    for (let i = last; i >= start; i--) items[i].delete();
    return { deleted_paragraphs: last - start + 1, note: "Nomor paragraf sesudahnya sudah bergeser." };
  });
}

// Melepas paragraf dari daftarnya, satu per satu: melepas beberapa butir dari daftar yang sama
// dalam satu kiriman gagal, karena daftar itu berubah susunan setiap kali satu butir keluar.
async function detach(ctx, paragraphs) {
  for (const paragraph of paragraphs) {
    if (!paragraph.isListItem) continue;
    paragraph.detachFromList();
    await ctx.sync();
  }
}

// Menjadikan paragraf-paragraf berurutan satu daftar: yang pertama memulai daftar, sisanya bergabung.
async function makeList(ctx, paragraphs, ordered) {
  await detach(ctx, paragraphs);
  const list = paragraphs[0].startNewList();
  if (ordered) list.setLevelNumbering(0, Word.ListNumbering.arabic, [0, "."]);
  else list.setLevelBullet(0, Word.ListBullet.solid);
  list.load("id");
  await ctx.sync();
  for (const paragraph of paragraphs.slice(1)) paragraph.attachToList(list.id, 0);
}

function setStyle({ start, count, style, custom_style }, context) {
  return edit(context, async (ctx) => {
    const items = await paragraphsOf(ctx, "items/text,items/isListItem");
    at(items, start);
    const targets = items.slice(start, start + (count || 1));
    if (!custom_style && LIST_STYLES.includes(style)) {
      await makeList(ctx, targets, style === "ListNumber");
    } else {
      // "Normal" pada butir daftar berarti keluar dari daftarnya
      if (!custom_style && style === "Normal") await detach(ctx, targets);
      for (const paragraph of targets) {
        if (custom_style) paragraph.style = custom_style;
        else paragraph.styleBuiltIn = style;
      }
    }
    return { styled_paragraphs: targets.length, style: custom_style || style };
  });
}

function formatText({ find: query, all, bold, italic, underline, highlight_color, font_color, size }, context) {
  return edit(context, async (ctx) => {
    const found = await find(ctx, query, true);
    if (!found.length) throw notFound(query);
    const targets = all ? found : found.slice(0, 1);
    for (const range of targets) {
      if (bold !== undefined) range.font.bold = bold;
      if (italic !== undefined) range.font.italic = italic;
      if (underline !== undefined) range.font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      if (highlight_color) range.font.highlightColor = highlight_color;
      if (font_color) range.font.color = font_color;
      if (size) range.font.size = size;
    }
    return { formatted: targets.length };
  });
}

function insertTable({ after_paragraph, values, header }, context) {
  return edit(context, async (ctx) => {
    const rows = rectangular(values);
    const anchor = Number.isInteger(after_paragraph) ? at(await paragraphsOf(ctx), after_paragraph) : ctx.document.body.paragraphs.getLast();
    const table = anchor.insertTable(rows.length, rows[0].length, Word.InsertLocation.after, rows);
    await tidyTable(ctx, table, header !== false);
    return { table_rows: rows.length, table_columns: rows[0].length };
  });
}

async function readTable({ index }) {
  return Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load("items/values");
    await ctx.sync();
    const table = tables.items[index || 0];
    if (!table) throw new Error(`Tabel ${index || 0} tidak ada. Dokumen punya ${tables.items.length} tabel.`);
    return { index: index || 0, values: table.values };
  });
}

function addComment({ find: query, comment }, context) {
  if (!supports("1.4")) throw new Error("Versi Word ini belum mendukung komentar dari add-in.");
  return edit({ ...context, trackChanges: false }, async (ctx) => {
    const found = await find(ctx, query, true);
    if (!found.length) throw notFound(query);
    found[0].insertComment(String(comment));
    return { commented_on: found[0].text };
  });
}

async function readComments() {
  if (!supports("1.4")) throw new Error("Versi Word ini belum mendukung komentar dari add-in.");
  return Word.run(async (ctx) => {
    const comments = ctx.document.body.getComments();
    comments.load("items/content,items/authorName,items/resolved");
    await ctx.sync();
    return { comments: comments.items.map((c) => ({ author: c.authorName, text: c.content, ...(c.resolved && { resolved: true }) })) };
  });
}

function replaceSelection({ text: value }, context) {
  return edit(context, async (ctx) => {
    await writeSelection(ctx, value, true);
    return { replaced: "selection" };
  });
}

// Menerima atau menolak tracked changes yang dibuat sejak `since` (dipakai tombol di panel).
// Sebuah penggantian tercatat sebagai pasangan sisip + hapus, dan separuhnya baru muncul di daftar
// setelah separuh lainnya diproses, jadi daftar dibaca berulang sampai habis.
export async function reviewChanges(since, accept) {
  // tanggal tracked change dibulatkan ke bawah ke menit, jadi batasnya dilonggarkan satu menit
  const from = since.getTime() - 60000;
  let total = 0;
  for (let pass = 0; pass < 6; pass++) {
    const handled = await Word.run(async (ctx) => {
      const changes = ctx.document.body.getTrackedChanges();
      changes.load("items/date");
      await ctx.sync();
      const mine = changes.items.filter((change) => new Date(change.date).getTime() >= from);
      for (const change of mine) accept ? change.accept() : change.reject();
      await ctx.sync();
      return mine.length;
    });
    if (!handled) break;
    total += handled;
  }
  return total;
}

export const canReview = () => supports("1.6");

export const tools = [
  define(
    "get_document_overview",
    "read",
    "Gambaran dokumen: jumlah paragraf dan kata, daftar judul beserta nomor paragrafnya, jumlah tabel, dan teks yang sedang dipilih. Panggil ini dulu sebelum bekerja.",
    {},
    [],
    () => "Melihat struktur dokumen",
    overview,
  ),
  define(
    "read_paragraphs",
    "read",
    'Membaca isi paragraf secara berurutan. Tiap paragraf punya nomor (index) yang dipakai alat lain. Paragraf yang berupa butir daftar bertanda list (mis. "•" atau "1.").',
    { start: int("Nomor paragraf pertama. Bawaan: 0."), count: int("Jumlah paragraf. Bawaan: 60.") },
    [],
    (a) => `Membaca paragraf mulai ${a.start || 0}`,
    readParagraphs,
  ),
  define(
    "search_text",
    "read",
    "Mencari teks di dokumen dan menampilkan paragraf tempatnya berada.",
    { query: text("Teks yang dicari, paling panjang 255 karakter."), match_case: flag("Peka huruf besar-kecil. Bawaan: tidak.") },
    ["query"],
    (a) => `Mencari "${a.query}"`,
    searchText,
  ),
  define(
    "replace_text",
    "write",
    "Mengganti potongan teks dengan teks baru, format di sekitarnya tetap. Cara utama untuk menyunting: ganti hanya bagian yang berubah (kata, frasa, kalimat), bukan seluruh paragraf. Kirim semua penggantian sekaligus dalam satu panggilan.",
    {
      replacements: {
        type: "array",
        items: {
          type: "object",
          properties: {
            find: text("Teks yang diganti, persis seperti di dokumen, dalam satu paragraf, paling panjang 255 karakter."),
            replace: text("Teks pengganti (teks polos)."),
            all: flag("Ganti semua kemunculan. Bawaan: hanya yang pertama."),
          },
          required: ["find", "replace"],
        },
        description: "Daftar penggantian, dikerjakan berurutan.",
      },
      match_case: flag("Peka huruf besar-kecil. Bawaan: ya."),
    },
    ["replacements"],
    (a) => `Mengganti ${a.replacements.length} potongan teks`,
    replaceText,
  ),
  define(
    "edit_paragraphs",
    "write",
    "Mengganti seluruh teks satu atau beberapa paragraf (gaya paragrafnya tetap). Pakai hanya kalau hampir seluruh paragraf berubah; kirim semuanya dalam satu panggilan.",
    {
      edits: {
        type: "array",
        items: { type: "object", properties: { index: int("Nomor paragraf."), text: text("Teks baru paragraf itu.") }, required: ["index", "text"] },
        description: "Paragraf yang ditulis ulang.",
      },
    },
    ["edits"],
    (a) => `Menulis ulang ${a.edits.length} paragraf`,
    editParagraphs,
  ),
  define(
    "insert_content",
    "write",
    'Menyisipkan konten baru yang ditulis dalam Markdown. Judul (#), daftar ("- " atau "1. " di awal baris, satu butir per baris), tebal, dan tabel menjadi format asli Word mengikuti gaya dokumen.',
    {
      where: { type: "string", enum: ["end", "start", "after_paragraph", "before_paragraph", "selection"], description: "Tempat menyisipkan. Bawaan: end." },
      index: int("Nomor paragraf acuan untuk after_paragraph / before_paragraph."),
      markdown: text("Konten dalam Markdown."),
    },
    ["markdown"],
    (a) => `Menyisipkan konten (${a.where || "end"})`,
    insertContent,
  ),
  define(
    "delete_paragraphs",
    "write",
    "Menghapus satu atau beberapa paragraf berurutan.",
    { start: int("Nomor paragraf pertama yang dihapus."), count: int("Jumlah paragraf. Bawaan: 1.") },
    ["start"],
    (a) => `Menghapus ${a.count || 1} paragraf mulai ${a.start}`,
    deleteParagraphs,
  ),
  define(
    "set_paragraph_style",
    "write",
    "Menerapkan gaya pada paragraf-paragraf berurutan: judul, normal, kutipan, atau menjadikannya daftar butir (ListBullet) / bernomor (ListNumber).",
    {
      start: int("Nomor paragraf pertama."),
      count: int("Jumlah paragraf. Bawaan: 1."),
      style: { type: "string", enum: STYLES, description: "Gaya bawaan Word." },
      custom_style: text("Nama gaya kustom di dokumen, sebagai pengganti style."),
    },
    ["start"],
    (a) => `Menerapkan gaya ${a.custom_style || a.style}`,
    setStyle,
  ),
  define(
    "format_text",
    "write",
    "Mengatur format huruf pada potongan teks: tebal, miring, garis bawah, stabilo, warna, ukuran.",
    {
      find: text("Teks yang diformat, persis seperti di dokumen."),
      all: flag("Semua kemunculan. Bawaan: hanya yang pertama."),
      bold: flag("Tebal."),
      italic: flag("Miring."),
      underline: flag("Garis bawah."),
      highlight_color: text('Warna stabilo, mis. "#FFFF00".'),
      font_color: text('Warna huruf, mis. "#C00000".'),
      size: { type: "number", description: "Ukuran huruf dalam poin." },
    },
    ["find"],
    (a) => `Memformat "${String(a.find).slice(0, 40)}"`,
    formatText,
  ),
  define(
    "insert_table",
    "write",
    "Menyisipkan tabel.",
    {
      values: { type: "array", items: { type: "array", items: { type: "string" } }, description: "Baris-baris sel." },
      after_paragraph: int("Nomor paragraf; tabel diletakkan sesudahnya. Kosongkan untuk akhir dokumen."),
      header: flag("Baris pertama adalah judul kolom. Bawaan: ya."),
    },
    ["values"],
    () => "Menyisipkan tabel",
    insertTable,
  ),
  define("read_table", "read", "Membaca isi sebuah tabel.", { index: int("Nomor tabel, mulai 0.") }, [], (a) => `Membaca tabel ${a.index || 0}`, readTable),
  define(
    "add_comment",
    "write",
    "Menambah komentar pada potongan teks, tanpa mengubah isinya. Cocok untuk saran atau pertanyaan.",
    { find: text("Teks yang dikomentari, persis seperti di dokumen."), comment: text("Isi komentar.") },
    ["find", "comment"],
    (a) => `Mengomentari "${String(a.find).slice(0, 40)}"`,
    addComment,
  ),
  define("read_comments", "read", "Membaca semua komentar di dokumen.", {}, [], () => "Membaca komentar", readComments),
  define(
    "replace_selection",
    "write",
    "Mengganti teks yang sedang dipilih pengguna. Kalau tidak ada yang dipilih, teks disisipkan di kursor.",
    { text: text("Teks pengganti. Satu baris disisipkan sebagai teks polos; beberapa baris boleh Markdown.") },
    ["text"],
    () => "Mengganti teks terpilih",
    replaceSelection,
  ),
];
