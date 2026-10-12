/* global Office, Word, Excel, PowerPoint */
// Jembatan ke dokumen: membaca pilihan dan menyisipkan jawaban, per aplikasi Office.

import { tableAt, toPlain } from "./markdown.js";
import { insertAtSelection } from "./tools/word.js";

// Label dan tombol sisip yang ditampilkan untuk tiap aplikasi.
export const HOSTS = {
  Word: {
    pick: "Blok teks di dokumen",
    selection: "Teks terpilih",
    document: "Seluruh dokumen",
    insert: [
      { mode: "insert", label: "Sisipkan", title: "Sisipkan di kursor / setelah teks terpilih" },
      { mode: "replace", label: "Ganti pilihan", title: "Ganti teks terpilih dengan jawaban ini" },
    ],
  },
  Excel: {
    pick: "Pilih sel di lembar kerja",
    selection: "Sel terpilih",
    document: "Sheet aktif",
    insert: [{ mode: "insert", label: "Tulis ke sel", title: "Tulis mulai dari sel aktif. Tabel dan rumus dipecah ke sel-selnya." }],
  },
  PowerPoint: {
    pick: "Blok teks di slide",
    selection: "Teks terpilih",
    document: "Semua slide",
    insert: [{ mode: "replace", label: "Sisipkan", title: "Sisipkan ke kotak teks terpilih (teks yang diblok akan diganti)" }],
  },
};

let host = null;

// Mengembalikan "Word" | "Excel" | "PowerPoint", atau null kalau halaman dibuka di luar Office.
export async function initHost() {
  if (typeof Office === "undefined") return null;
  const info = await Office.onReady();
  const name = info && info.host ? String(info.host) : "";
  host = HOSTS[name] ? name : null;
  return host;
}

const asyncResult = (resolve, reject) => (result) =>
  result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(new Error(result.error.message));

const getSelectedText = () =>
  new Promise((resolve, reject) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, asyncResult(resolve, reject)),
  );

const setSelectedText = (text) =>
  new Promise((resolve, reject) =>
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, asyncResult(resolve, reject)),
  );

function columnName(index) {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

// Rentang Excel → teks: nilai seperti yang tampil di layar (dipisah tab), plus daftar rumusnya.
function describeRange(range) {
  if (range.text.every((row) => row.every((cell) => cell === ""))) return "";
  const formulas = [];
  range.formulas.forEach((row, r) =>
    row.forEach((formula, c) => {
      if (typeof formula === "string" && formula.startsWith("=")) {
        formulas.push(`${columnName(range.columnIndex + c)}${range.rowIndex + r + 1}: ${formula}`);
      }
    }),
  );
  return [
    `Rentang: ${range.address}`,
    range.text.map((row) => row.join("\t")).join("\n"),
    ...(formulas.length ? ["Rumus:", ...formulas] : []),
  ].join("\n");
}

const RANGE_FIELDS = ["address", "text", "formulas", "rowIndex", "columnIndex"];

export async function readSelection() {
  if (host === "Word") {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text;
    });
  }
  if (host === "Excel") {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(RANGE_FIELDS);
      await context.sync();
      return describeRange(range);
    });
  }
  if (host === "PowerPoint") {
    // gagal kalau yang dipilih bukan teks (mis. slide atau gambar): anggap tidak ada pilihan
    return getSelectedText().then((text) => text || "", () => "");
  }
  return "";
}

export async function readDocument() {
  if (host === "Word") {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text;
    });
  }
  if (host === "Excel") {
    return Excel.run(async (context) => {
      const used = context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true);
      used.load(RANGE_FIELDS);
      await context.sync();
      return used.isNullObject ? "" : describeRange(used);
    });
  }
  if (host === "PowerPoint") {
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
      throw new Error("Versi PowerPoint ini belum bisa membaca semua slide. Pakai “Teks terpilih”.");
    }
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      const out = [];
      for (const [n, slide] of slides.items.entries()) {
        slide.shapes.load("items");
        await context.sync();
        const texts = [];
        for (const shape of slide.shapes.items) {
          const range = shape.textFrame.textRange;
          range.load("text");
          try {
            await context.sync();
            if (range.text.trim()) texts.push(range.text.trim());
          } catch {
            // bentuk tanpa teks (gambar, tabel, grafik): lewati
          }
        }
        out.push(`Slide ${n + 1}:\n${texts.join("\n")}`);
      }
      return out.join("\n\n");
    });
  }
  return "";
}

// Jawaban AI → kisi sel. Urutan: tabel Markdown, lalu isi blok kode (biasanya rumus),
// lalu teks biasa (satu baris per baris sel, kolom dipisah tab).
export function toGrid(text) {
  const source = text.replace(/\r\n?/g, "\n");
  const lines = source.split("\n");
  let rows;
  const table = lines.map((_, i) => tableAt(lines, i)).find(Boolean);
  const fence = /```[^\n]*\n([\s\S]*?)```/.exec(source);
  if (table) {
    rows = table.rows.map((row) => row.map(toPlain));
  } else if (fence) {
    rows = fence[1].trim().split("\n").map((line) => line.split("\t"));
  } else {
    rows = toPlain(source).split("\n").filter((line) => line.trim()).map((line) => line.split("\t"));
  }
  const width = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => [...row, ...Array(width - row.length).fill("")]);
}

function needsConfirm(message) {
  return Object.assign(new Error(message), { needsConfirm: true });
}

// mode: "insert" | "replace". Melempar error ber-flag needsConfirm kalau akan menimpa
// sel Excel yang sudah berisi; panggil lagi dengan { overwrite: true } setelah pengguna setuju.
export async function insertReply(text, mode, { overwrite = false } = {}) {
  const reply = text.trim();
  if (host === "Word") return insertAtSelection(reply, mode === "replace");
  if (host === "Excel") {
    return Excel.run(async (context) => {
      const grid = toGrid(reply);
      const selection = context.workbook.getSelectedRange();
      selection.load(["rowIndex", "columnIndex"]);
      await context.sync();
      const target = selection.worksheet.getRangeByIndexes(selection.rowIndex, selection.columnIndex, grid.length, grid[0].length);
      target.load(["address", "text"]);
      await context.sync();
      if (!overwrite && target.text.some((row) => row.some((cell) => cell !== ""))) {
        throw needsConfirm(`${target.address} sudah berisi data.`);
      }
      target.formulas = grid;
      target.select();
      await context.sync();
    });
  }
  if (host === "PowerPoint") {
    return setSelectedText(toPlain(reply)).catch((error) => {
      throw new Error(`Klik dulu di dalam kotak teks pada slide, lalu coba lagi. (${error.message})`);
    });
  }
  throw new Error("Penyisipan hanya tersedia di dalam Word, Excel, atau PowerPoint.");
}
