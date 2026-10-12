/* global Office, PowerPoint */
// Alat agen untuk PowerPoint. Bentuk alat sama dengan tools/excel.js.
// Slide dirujuk dengan nomor urut mulai 1 (seperti di PowerPoint); bentuk dirujuk dengan id-nya.
// Posisi dan ukuran dalam poin; slide 16:9 berukuran 960 x 540.

import { toPlain } from "../markdown.js";

const TITLE_TYPES = ["Title", "CenterTitle"];
const BODY_TYPES = ["Body", "Content", "Subtitle", "VerticalBody", "VerticalContent"];
const CHART_KINDS = ["bar", "line", "pie", "doughnut", "area"];
const PPTX_LIB = "vendor/pptxgen.bundle.js";

const text = (description) => ({ type: "string", description });
const int = (description) => ({ type: "integer", description });
const num = (description) => ({ type: "number", description });
const flag = (description) => ({ type: "boolean", description });
const SLIDE = int("Nomor slide, mulai 1.");
const SHAPE = text("Id bentuk, dari get_presentation_overview atau read_slide.");

function define(name, risk, description, properties, required, label, run) {
  return { name, risk, description, parameters: { type: "object", properties, ...(required.length && { required }) }, label, run };
}

const supports = (version) => Office.context.requirements.isSetSupported("PowerPointApi", version);

function need(version, what) {
  if (!supports(version)) throw new Error(`Versi PowerPoint ini belum mendukung ${what}.`);
}

async function slidesOf(ctx) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  return slides.items;
}

function slideAt(slides, number) {
  if (!Number.isInteger(number) || number < 1 || number > slides.length) {
    throw new Error(`Slide ${number} tidak ada. Presentasi punya ${slides.length} slide.`);
  }
  return slides[number - 1];
}

// Bentuk-bentuk sebuah slide beserta teksnya. Bentuk tanpa bingkai teks (gambar, grafik) bertanda text: null.
async function shapesOf(ctx, slide, details = false) {
  slide.shapes.load(`items/id,items/name,items/type${details ? ",items/left,items/top,items/width,items/height" : ""}`);
  await ctx.sync();
  const shapes = slide.shapes.items;
  const placeholders = supports("1.8") ? shapes.filter((s) => s.type === "Placeholder") : [];
  for (const shape of placeholders) shape.placeholderFormat.load("type");
  let frames = null;
  if (supports("1.10")) {
    frames = shapes.map((shape) => {
      const frame = shape.getTextFrameOrNullObject();
      frame.textRange.load("text");
      return frame;
    });
  }
  await ctx.sync();
  const out = [];
  for (const [i, shape] of shapes.entries()) {
    let content = null;
    if (frames) {
      content = frames[i].isNullObject ? null : frames[i].textRange.text;
    } else {
      // API lama: satu per satu, karena bentuk tanpa teks membuat seluruh permintaan gagal
      const range = shape.textFrame.textRange;
      range.load("text");
      try {
        await ctx.sync();
        content = range.text;
      } catch {
        content = null;
      }
    }
    out.push({ shape, text: content, placeholder: placeholders.includes(shape) ? shape.placeholderFormat.type : undefined });
  }
  return out;
}

const describe = ({ shape, text: content, placeholder }, limit) => ({
  id: shape.id,
  name: shape.name,
  type: placeholder ? `Placeholder:${placeholder}` : shape.type,
  ...(content !== null && { text: limit ? content.slice(0, limit) : content }),
});

async function overview() {
  return PowerPoint.run(async (ctx) => {
    const slides = await slidesOf(ctx);
    const masters = ctx.presentation.slideMasters;
    masters.load("items/id,items/name");
    for (const slide of slides) slide.layout.load("name");
    await ctx.sync();
    for (const master of masters.items) master.layouts.load("items/name");
    await ctx.sync();
    const out = [];
    for (const [i, slide] of slides.entries()) {
      const shapes = await shapesOf(ctx, slide);
      out.push({ slide: i + 1, layout: slide.layout.name, shapes: shapes.map((s) => describe(s, 160)) });
    }
    return {
      slide_count: slides.length,
      slide_size: "960 x 540 poin (16:9), kecuali presentasinya memakai ukuran lain",
      layouts: [...new Set(masters.items.flatMap((m) => m.layouts.items.map((l) => l.name)))],
      slides: out,
    };
  });
}

async function readSlide({ slide: number }) {
  return PowerPoint.run(async (ctx) => {
    const slide = slideAt(await slidesOf(ctx), number);
    slide.layout.load("name");
    const shapes = await shapesOf(ctx, slide, true);
    const tables = supports("1.8") ? shapes.filter((s) => s.shape.type === "Table") : [];
    const loaded = tables.map((s) => {
      const table = s.shape.getTable();
      table.load("values");
      return table;
    });
    await ctx.sync();
    return {
      slide: number,
      layout: slide.layout.name,
      shapes: shapes.map((s) => ({
        ...describe(s),
        left: Math.round(s.shape.left),
        top: Math.round(s.shape.top),
        width: Math.round(s.shape.width),
        height: Math.round(s.shape.height),
        ...(tables.includes(s) && { table: loaded[tables.indexOf(s)].values }),
      })),
    };
  });
}

async function findLayout(ctx, name) {
  const masters = ctx.presentation.slideMasters;
  masters.load("items/id");
  await ctx.sync();
  for (const master of masters.items) master.layouts.load("items/id,items/name");
  await ctx.sync();
  const all = masters.items.flatMap((m) => m.layouts.items.map((l) => ({ name: l.name, layoutId: l.id, slideMasterId: m.id })));
  const wanted = String(name).trim().toLowerCase();
  const found = all.find((l) => l.name.toLowerCase() === wanted);
  if (!found) throw new Error(`Layout "${name}" tidak ada. Pilihan: ${all.map((l) => l.name).join(", ")}.`);
  return found;
}

const removeSlide = (id) => () =>
  PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.getItem(id).delete();
    await ctx.sync();
  });

const removeShape = (slideId, shapeId) => () =>
  PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).delete();
    await ctx.sync();
  });

async function addSlide({ layout, title, body, position }, context) {
  return PowerPoint.run(async (ctx) => {
    const options = layout ? await findLayout(ctx, layout) : null;
    ctx.presentation.slides.add(options ? { layoutId: options.layoutId, slideMasterId: options.slideMasterId } : undefined);
    await ctx.sync();
    const slides = await slidesOf(ctx);
    const slide = slides[slides.length - 1];
    const shapes = await shapesOf(ctx, slide);
    const fill = (types, value) => {
      const target = shapes.find((s) => types.includes(s.placeholder) && s.text !== null);
      if (target && value) target.shape.textFrame.textRange.text = toPlain(String(value));
      return !!target || !value;
    };
    const placed = { title: fill(TITLE_TYPES, title), body: fill(BODY_TYPES, body) };
    let number = slides.length;
    if (Number.isInteger(position) && position >= 1 && position < slides.length && supports("1.8")) {
      slide.moveTo(position - 1);
      number = position;
    }
    await ctx.sync();
    context.undo(removeSlide(slide.id));
    return {
      slide: number,
      shapes: shapes.map((s) => describe(s, 0)),
      ...(!placed.title && { note_title: "Layout ini tidak punya tempat judul; judul tidak ditulis." }),
      ...(!placed.body && { note_body: "Layout ini tidak punya tempat isi; pakai add_text_box." }),
    };
  });
}

async function setShapeText({ slide: number, shape_id, text: value }, context) {
  return PowerPoint.run(async (ctx) => {
    const slide = slideAt(await slidesOf(ctx), number);
    const range = slide.shapes.getItem(String(shape_id)).textFrame.textRange;
    range.load("text");
    await ctx.sync();
    const before = range.text;
    range.text = toPlain(String(value));
    await ctx.sync();
    context.undo(() =>
      PowerPoint.run(async (undoCtx) => {
        undoCtx.presentation.slides.getItem(slide.id).shapes.getItem(String(shape_id)).textFrame.textRange.text = before;
        await undoCtx.sync();
      }),
    );
    return { slide: number, shape_id };
  });
}

async function addTextBox({ slide: number, text: value, left, top, width, height, font_size, bold, font_color }, context) {
  need("1.4", "penambahan bentuk");
  return PowerPoint.run(async (ctx) => {
    const slide = slideAt(await slidesOf(ctx), number);
    const box = slide.shapes.addTextBox(toPlain(String(value)), { left: left ?? 60, top: top ?? 120, width: width ?? 840, height: height ?? 80 });
    if (font_size) box.textFrame.textRange.font.size = font_size;
    if (bold !== undefined) box.textFrame.textRange.font.bold = bold;
    if (font_color) box.textFrame.textRange.font.color = font_color;
    box.load("id");
    await ctx.sync();
    context.undo(removeShape(slide.id, box.id));
    return { slide: number, shape_id: box.id };
  });
}

async function addTable({ slide: number, values, left, top, width, height }, context) {
  need("1.8", "tabel");
  return PowerPoint.run(async (ctx) => {
    const slide = slideAt(await slidesOf(ctx), number);
    const rows = values.map((row) => row.map((cell) => String(cell ?? "")));
    const columns = Math.max(...rows.map((row) => row.length));
    for (const row of rows) while (row.length < columns) row.push("");
    const shape = slide.shapes.addTable(rows.length, columns, {
      values: rows,
      left: left ?? 60,
      top: top ?? 140,
      ...(width && { width }),
      ...(height && { height }),
    });
    shape.load("id");
    await ctx.sync();
    context.undo(removeShape(slide.id, shape.id));
    return { slide: number, shape_id: shape.id, rows: rows.length, columns };
  });
}

async function formatShape(args, context) {
  const { slide: number, shape_id, font_size, bold, italic, font_color, fill_color, left, top, width, height } = args;
  return PowerPoint.run(async (ctx) => {
    const slide = slideAt(await slidesOf(ctx), number);
    const shape = slide.shapes.getItem(String(shape_id));
    shape.load("left,top,width,height");
    await ctx.sync();
    const before = { left: shape.left, top: shape.top, width: shape.width, height: shape.height };
    for (const [key, value] of Object.entries({ left, top, width, height })) if (value !== undefined) shape[key] = value;
    const font = shape.textFrame.textRange.font;
    if (font_size) font.size = font_size;
    if (bold !== undefined) font.bold = bold;
    if (italic !== undefined) font.italic = italic;
    if (font_color) font.color = font_color;
    if (fill_color) shape.fill.setSolidColor(fill_color);
    await ctx.sync();
    context.undo(() =>
      PowerPoint.run(async (undoCtx) => {
        // posisi dan ukuran dikembalikan; warna dan huruf tidak
        Object.assign(undoCtx.presentation.slides.getItem(slide.id).shapes.getItem(String(shape_id)), before);
        await undoCtx.sync();
      }),
    );
    return { slide: number, shape_id };
  });
}

async function moveSlide({ slide: number, to }, context) {
  need("1.8", "pemindahan slide");
  return PowerPoint.run(async (ctx) => {
    const slides = await slidesOf(ctx);
    const slide = slideAt(slides, number);
    slideAt(slides, to);
    slide.moveTo(to - 1);
    await ctx.sync();
    context.undo(() =>
      PowerPoint.run(async (undoCtx) => {
        undoCtx.presentation.slides.getItem(slide.id).moveTo(number - 1);
        await undoCtx.sync();
      }),
    );
    return { moved_to: to };
  });
}

async function deleteShape({ slide: number, shape_id }) {
  return PowerPoint.run(async (ctx) => {
    slideAt(await slidesOf(ctx), number).shapes.getItem(String(shape_id)).delete();
    await ctx.sync();
    return { deleted_shape: shape_id };
  });
}

async function deleteSlide({ slide: number }) {
  return PowerPoint.run(async (ctx) => {
    slideAt(await slidesOf(ctx), number).delete();
    await ctx.sync();
    return { deleted_slide: number };
  });
}

function replaceSelection({ text: value }, context) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(toPlain(String(value)), { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(`Tidak ada teks atau kotak teks yang dipilih di slide. (${result.error.message})`));
        return;
      }
      context.changed();
      resolve({ replaced: "selection" });
    });
  });
}

let pptxLibrary;
// PptxGenJS dimuat hanya saat pertama kali grafik diminta.
function loadPptxLibrary() {
  pptxLibrary ||= new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = PPTX_LIB;
    script.onload = () => resolve(window.PptxGenJS);
    script.onerror = () => reject(new Error("Pustaka pembuat grafik gagal dimuat."));
    document.head.append(script);
  });
  return pptxLibrary;
}

// PowerPoint tidak punya API grafik untuk add-in. Jalan memutarnya: slide bergrafik dibuat sebagai
// berkas .pptx di dalam panel, lalu disisipkan ke presentasi mengikuti tema tujuannya.
async function addChartSlide({ title, kind, categories, series, after_slide }, context) {
  const PptxGenJS = await loadPptxLibrary();
  const deck = new PptxGenJS();
  deck.layout = "LAYOUT_WIDE";
  const slide = deck.addSlide();
  if (title) slide.addText(String(title), { x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 28, bold: true });
  const data = series.map((s) => ({ name: String(s.name), labels: categories.map(String), values: s.values.map(Number) }));
  slide.addChart(deck.ChartType[kind] || deck.ChartType.bar, data, {
    x: 0.6,
    y: 1.4,
    w: 12.1,
    h: 5.6,
    showLegend: data.length > 1 || kind === "pie" || kind === "doughnut",
    legendPos: "b",
    showValue: kind === "pie" || kind === "doughnut",
  });
  const base64 = await deck.write({ outputType: "base64" });
  return PowerPoint.run(async (ctx) => {
    const before = await slidesOf(ctx);
    const anchor = Number.isInteger(after_slide) ? slideAt(before, after_slide) : before[before.length - 1];
    ctx.presentation.insertSlidesFromBase64(base64, {
      formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
      ...(anchor && { targetSlideId: anchor.id }),
    });
    await ctx.sync();
    const known = new Set(before.map((s) => s.id));
    const after = await slidesOf(ctx);
    const added = after.findIndex((s) => !known.has(s.id));
    if (added >= 0) context.undo(removeSlide(after[added].id));
    return { slide: added + 1, note: "Grafik ini asli PowerPoint dan datanya bisa diedit, tetapi tidak tertaut ke berkas Excel." };
  });
}

export const tools = [
  define(
    "get_presentation_overview",
    "read",
    "Gambaran presentasi: jumlah slide, layout yang tersedia di template, dan tiap slide beserta bentuk-bentuknya (id, jenis, cuplikan teks). Panggil ini dulu sebelum bekerja.",
    {},
    [],
    () => "Melihat struktur presentasi",
    overview,
  ),
  define("read_slide", "read", "Membaca satu slide secara lengkap: teks utuh, posisi dan ukuran tiap bentuk, dan isi tabel.", { slide: SLIDE }, ["slide"], (a) => `Membaca slide ${a.slide}`, readSlide),
  define(
    "add_slide",
    "write",
    "Menambah slide baru dari salah satu layout template, sekaligus mengisi judul dan isinya. Pakai nama layout persis seperti di get_presentation_overview.",
    {
      layout: text('Nama layout, mis. "Title and Content". Kosongkan untuk layout bawaan.'),
      title: text("Judul slide."),
      body: text("Isi slide. Tiap baris menjadi satu poin; tulis ringkas, tanpa simbol bullet."),
      position: int("Nomor urut slide baru. Kosongkan untuk menaruhnya di akhir."),
    },
    [],
    (a) => `Menambah slide${a.title ? ` "${String(a.title).slice(0, 40)}"` : ""}`,
    addSlide,
  ),
  define(
    "set_shape_text",
    "write",
    "Mengganti teks sebuah bentuk atau placeholder. Tiap baris menjadi satu paragraf.",
    { slide: SLIDE, shape_id: SHAPE, text: text("Teks baru.") },
    ["slide", "shape_id", "text"],
    (a) => `Mengubah teks di slide ${a.slide}`,
    setShapeText,
  ),
  define(
    "add_text_box",
    "write",
    "Menambah kotak teks di posisi tertentu. Pakai kalau layout tidak punya placeholder yang cocok.",
    {
      slide: SLIDE,
      text: text("Isi kotak teks."),
      left: num("Jarak dari kiri, poin."),
      top: num("Jarak dari atas, poin."),
      width: num("Lebar, poin."),
      height: num("Tinggi, poin."),
      font_size: num("Ukuran huruf, poin."),
      bold: flag("Tebal."),
      font_color: text('Warna huruf, mis. "#1F3864".'),
    },
    ["slide", "text"],
    (a) => `Menambah kotak teks di slide ${a.slide}`,
    addTextBox,
  ),
  define(
    "add_table",
    "write",
    "Menambah tabel ke sebuah slide.",
    {
      slide: SLIDE,
      values: { type: "array", items: { type: "array", items: { type: "string" } }, description: "Baris-baris sel; baris pertama biasanya judul kolom." },
      left: num("Jarak dari kiri, poin."),
      top: num("Jarak dari atas, poin."),
      width: num("Lebar, poin."),
      height: num("Tinggi, poin."),
    },
    ["slide", "values"],
    (a) => `Menambah tabel di slide ${a.slide}`,
    addTable,
  ),
  define(
    "add_chart_slide",
    "write",
    "Menambah slide baru berisi grafik asli PowerPoint (datanya bisa diedit) dari angka yang diberikan.",
    {
      title: text("Judul slide."),
      kind: { type: "string", enum: CHART_KINDS, description: "Jenis grafik. Bawaan: bar." },
      categories: { type: "array", items: { type: "string" }, description: 'Label sumbu kategori, mis. ["Jan","Feb","Mar"].' },
      series: {
        type: "array",
        items: { type: "object", properties: { name: text("Nama seri."), values: { type: "array", items: { type: "number" }, description: "Satu angka per kategori." } }, required: ["name", "values"] },
        description: "Satu atau beberapa seri data.",
      },
      after_slide: int("Nomor slide; slide grafik diletakkan sesudahnya. Kosongkan untuk akhir."),
    },
    ["categories", "series"],
    (a) => `Menambah slide grafik${a.title ? ` "${String(a.title).slice(0, 40)}"` : ""}`,
    addChartSlide,
  ),
  define(
    "format_shape",
    "write",
    "Mengatur sebuah bentuk: posisi, ukuran, huruf, dan warna isian. Isi hanya yang ingin diubah.",
    {
      slide: SLIDE,
      shape_id: SHAPE,
      left: num("Jarak dari kiri, poin."),
      top: num("Jarak dari atas, poin."),
      width: num("Lebar, poin."),
      height: num("Tinggi, poin."),
      font_size: num("Ukuran huruf, poin."),
      bold: flag("Tebal."),
      italic: flag("Miring."),
      font_color: text("Warna huruf."),
      fill_color: text("Warna isian bentuk."),
    },
    ["slide", "shape_id"],
    (a) => `Memformat bentuk di slide ${a.slide}`,
    formatShape,
  ),
  define("move_slide", "write", "Memindahkan slide ke urutan lain.", { slide: SLIDE, to: int("Nomor urut tujuan, mulai 1.") }, ["slide", "to"], (a) => `Memindahkan slide ${a.slide} ke urutan ${a.to}`, moveSlide),
  define("delete_shape", "danger", "Menghapus sebuah bentuk dari slide. Tidak bisa dibatalkan.", { slide: SLIDE, shape_id: SHAPE }, ["slide", "shape_id"], (a) => `Menghapus bentuk di slide ${a.slide}`, deleteShape),
  define("delete_slide", "danger", "Menghapus sebuah slide. Tidak bisa dibatalkan.", { slide: SLIDE }, ["slide"], (a) => `Menghapus slide ${a.slide}`, deleteSlide),
  define(
    "replace_selection",
    "write",
    "Mengganti teks yang sedang dipilih pengguna di slide.",
    { text: text("Teks pengganti.") },
    ["text"],
    () => "Mengganti teks terpilih",
    replaceSelection,
  ),
];
