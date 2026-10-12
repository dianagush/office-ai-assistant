/* global Excel */
// Alat agen untuk Excel. Setiap alat: { name, risk, description, parameters, label(args), run(args, context) }.
//   risk "read"   → langsung jalan
//   risk "write"  → jalan, dan mendaftarkan cara membatalkannya lewat context.undo(fn)
//   risk "danger" → selalu minta izin pengguna dulu (tidak bisa dibatalkan)

// Batas sel per pembacaan, supaya satu panggilan tidak menghabiskan konteks model.
const MAX_READ = 4000;
// Di atas ini keadaan sebelumnya tidak disimpan (terlalu besar), jadi tidak bisa dibatalkan.
const MAX_SNAPSHOT = 20000;
// Menimpa lebih dari sekian sel berisi dianggap berisiko dan butuh izin.
const CONFIRM_OVERWRITE = 10;

const CHART_TYPES = ["ColumnClustered", "ColumnStacked", "BarClustered", "BarStacked", "Line", "LineMarkers", "Pie", "Doughnut", "Area", "XYScatter"];
const OPERATORS = ["GreaterThan", "GreaterThanOrEqual", "LessThan", "LessThanOrEqual", "EqualTo", "NotEqualTo", "Between", "NotBetween"];
const BORDER_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];

const text = (description) => ({ type: "string", description });
const flag = (description) => ({ type: "boolean", description });
const list = (items, description) => ({ type: "array", items, description });
const ADDRESS = text('Alamat rentang, boleh dengan nama sheet. Contoh: "A1:C10" (sheet aktif) atau "Data!B2".');

function define(name, risk, description, properties, required, label, run) {
  return { name, risk, description, parameters: { type: "object", properties, ...(required.length && { required }) }, label, run };
}

function split(address) {
  const full = String(address).trim();
  const bang = full.lastIndexOf("!");
  if (bang < 0) return { sheet: "", cells: full };
  return { sheet: full.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'"), cells: full.slice(bang + 1) };
}

const sheetOf = (ctx, name) => (name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet());
const rangeOf = (ctx, address) => sheetOf(ctx, split(address).sheet).getRange(split(address).cells);

function columnName(index) {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

function rectangular(values) {
  const rows = (Array.isArray(values) ? values : [[values]]).map((row) => (Array.isArray(row) ? row : [row]));
  if (!rows.length) throw new Error("values kosong.");
  const width = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => [...row, ...Array(width - row.length).fill("")].map((cell) => cell ?? ""));
}

// Keadaan sebuah rentang sebelum diubah, untuk dikembalikan saat Batalkan.
async function snapshot(address) {
  return Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address);
    range.load(["address", "cellCount"]);
    await ctx.sync();
    if (range.cellCount > MAX_SNAPSHOT) return { address: range.address, cells: range.cellCount };
    range.load(["formulas", "numberFormat"]);
    await ctx.sync();
    return { address: range.address, cells: range.cellCount, formulas: range.formulas, numberFormat: range.numberFormat };
  });
}

const filledCells = (before) => (before.formulas ? before.formulas.flat().filter((v) => v !== "").length : before.cells);

function restore(before) {
  return () =>
    Excel.run(async (ctx) => {
      const range = rangeOf(ctx, before.address);
      range.numberFormat = before.numberFormat;
      range.formulas = before.formulas;
      await ctx.sync();
    });
}

// Izin untuk perubahan yang menimpa banyak data atau terlalu besar untuk dibatalkan.
async function guard(before, action, context) {
  const filled = filledCells(before);
  const reasons = [];
  if (filled > CONFIRM_OVERWRITE) reasons.push(`${filled} sel berisi`);
  if (!before.formulas) reasons.push("tidak bisa dibatalkan");
  if (reasons.length && !(await context.confirm(`${action} ${before.address} (${reasons.join(", ")})`))) {
    throw new Error("Pengguna menolak tindakan ini.");
  }
  if (before.formulas) context.undo(restore(before));
}

async function overview() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    const names = ctx.workbook.names;
    names.load("items/name,items/formula");
    await ctx.sync();
    const parts = sheets.items.map((sheet) => {
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("address,rowCount,columnCount");
      sheet.tables.load("items/name");
      sheet.charts.load("items/name");
      sheet.pivotTables.load("items/name");
      return { sheet, used };
    });
    await ctx.sync();
    const tables = parts.flatMap(({ sheet }) =>
      sheet.tables.items.map((table) => {
        const range = table.getRange();
        range.load("address");
        return { table, range };
      }),
    );
    let selection = "";
    try {
      const selected = ctx.workbook.getSelectedRange();
      selected.load("address");
      await ctx.sync();
      selection = selected.address;
    } catch {
      await ctx.sync(); // yang dipilih bukan sel (mis. grafik)
    }
    return {
      active_sheet: active.name,
      selection,
      sheets: parts.map(({ sheet, used }) => ({
        name: sheet.name,
        ...(sheet.visibility !== "Visible" && { hidden: true }),
        used_range: used.isNullObject ? null : used.address,
        rows: used.isNullObject ? 0 : used.rowCount,
        columns: used.isNullObject ? 0 : used.columnCount,
        charts: sheet.charts.items.map((c) => c.name),
        pivot_tables: sheet.pivotTables.items.map((p) => p.name),
      })),
      tables: tables.map(({ table, range }) => ({ name: table.name, address: range.address })),
      named_ranges: names.items.map((n) => ({ name: n.name, refers_to: n.formula })),
    };
  });
}

async function readRange({ address, as }) {
  return Excel.run(async (ctx) => {
    // dipersempit ke bagian yang berisi, supaya "A:A" tidak berarti sejuta baris
    let range = rangeOf(ctx, address).getUsedRangeOrNullObject(true);
    range.load(["address", "rowCount", "columnCount"]);
    await ctx.sync();
    if (range.isNullObject) return { address, empty: true };
    const total = { rows: range.rowCount, columns: range.columnCount };
    let note;
    if (total.rows * total.columns > MAX_READ) {
      const rows = Math.max(1, Math.floor(MAX_READ / total.columns));
      const columns = Math.min(total.columns, MAX_READ);
      range = range.getCell(0, 0).getAbsoluteResizedRange(rows, columns);
      note = `Rentang berisi ${total.rows} baris x ${total.columns} kolom; hanya ${rows} baris x ${columns} kolom pertama yang dibaca. Baca sisanya dengan alamat lanjutan.`;
    }
    range.load(["address", "rowIndex", "columnIndex", "values", "text", "formulas"]);
    await ctx.sync();
    const formulas = {};
    range.formulas.forEach((row, r) =>
      row.forEach((formula, c) => {
        if (typeof formula === "string" && formula.startsWith("=")) {
          formulas[`${columnName(range.columnIndex + c)}${range.rowIndex + r + 1}`] = formula;
        }
      }),
    );
    return { address: range.address, ...(note && { note }), values: as === "text" ? range.text : range.values, formulas };
  });
}

async function findCells({ query, sheet }) {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const targets = sheets.items.filter((s) => !sheet || s.name === sheet);
    if (!targets.length) throw new Error(`Sheet "${sheet}" tidak ada.`);
    const found = targets.map((s) => {
      const areas = s.findAllOrNullObject(String(query), { completeMatch: false, matchCase: false });
      areas.load("address,cellCount");
      return areas;
    });
    await ctx.sync();
    const matches = found.filter((a) => !a.isNullObject).map((a) => ({ cells: a.cellCount, addresses: a.address.slice(0, 1500) }));
    return { query, matches };
  });
}

async function findErrors({ sheet }) {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const targets = sheets.items.filter((s) => !sheet || s.name === sheet);
    if (!targets.length) throw new Error(`Sheet "${sheet}" tidak ada.`);
    const found = targets.map((s) => {
      const areas = s.getUsedRange(true).getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas, Excel.SpecialCellValueType.errors);
      areas.load("address,cellCount");
      return areas;
    });
    await ctx.sync().catch(() => {}); // sheet kosong: tidak ada used range
    const errors = [];
    for (const areas of found) {
      if (areas.isNullObject !== false || errors.length >= 60) continue;
      areas.areas.load("items/address,items/values,items/formulas,items/rowIndex,items/columnIndex");
      await ctx.sync();
      for (const area of areas.areas.items) {
        const sheetName = split(area.address).sheet;
        area.formulas.forEach((row, r) =>
          row.forEach((formula, c) => {
            if (errors.length < 60) {
              errors.push({ cell: `${sheetName}!${columnName(area.columnIndex + c)}${area.rowIndex + r + 1}`, error: area.values[r][c], formula });
            }
          }),
        );
      }
    }
    return { errors, ...(errors.length >= 60 && { note: "Hanya 60 galat pertama yang ditampilkan." }) };
  });
}

async function traceCell({ address, direction }) {
  return Excel.run(async (ctx) => {
    const cell = rangeOf(ctx, address);
    const linked = direction === "dependents" ? cell.getDirectDependents() : cell.getDirectPrecedents();
    linked.load("addresses");
    cell.load("address,formulas");
    try {
      await ctx.sync();
    } catch {
      return { cell: address, direction: direction || "precedents", cells: [] };
    }
    return { cell: cell.address, formula: cell.formulas[0][0], direction: direction || "precedents", cells: linked.addresses };
  });
}

async function writeRange({ address, values }, context) {
  const grid = rectangular(values);
  const target = await Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address).getCell(0, 0).getAbsoluteResizedRange(grid.length, grid[0].length);
    range.load("address");
    await ctx.sync();
    return range.address;
  });
  await guard(await snapshot(target), "Menimpa", context);
  await Excel.run(async (ctx) => {
    rangeOf(ctx, target).formulas = grid;
    await ctx.sync();
  });
  return { written: target, rows: grid.length, columns: grid[0].length };
}

async function clearRange({ address, what }, context) {
  await guard(await snapshot(address), "Mengosongkan", context);
  await Excel.run(async (ctx) => {
    const apply = { formats: Excel.ClearApplyTo.formats, all: Excel.ClearApplyTo.all }[what] || Excel.ClearApplyTo.contents;
    rangeOf(ctx, address).clear(apply);
    await ctx.sync();
  });
  return { cleared: address, what: what || "contents" };
}

async function formatRange(args, context) {
  const { address, bold, italic, font_color, fill_color, number_format, align, wrap_text, border, autofit } = args;
  return Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address);
    range.load(["address", "cellCount", "rowCount", "columnCount", "numberFormat"]);
    const wanted = {
      format: {
        font: { bold: bold !== undefined, italic: italic !== undefined, color: !!font_color },
        fill: { color: !!fill_color },
        horizontalAlignment: !!align,
        wrapText: wrap_text !== undefined,
      },
    };
    await ctx.sync();
    const small = range.cellCount <= 5000;
    const props = small ? range.getCellProperties(wanted) : null;
    await ctx.sync();
    const before = { address: range.address, numberFormat: range.numberFormat, props: props && props.value };

    if (bold !== undefined) range.format.font.bold = bold;
    if (italic !== undefined) range.format.font.italic = italic;
    if (font_color) range.format.font.color = font_color;
    if (fill_color) range.format.fill.color = fill_color;
    if (align) range.format.horizontalAlignment = align;
    if (wrap_text !== undefined) range.format.wrapText = wrap_text;
    if (number_format) range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(number_format));
    if (border) {
      for (const edge of BORDER_EDGES) {
        if (border === "outline" && edge.startsWith("Inside")) continue;
        range.format.borders.getItem(edge).style = border === "none" ? "None" : "Continuous";
      }
    }
    if (autofit) range.format.autofitColumns();
    await ctx.sync();

    if (small) {
      context.undo(() =>
        Excel.run(async (undoCtx) => {
          const target = rangeOf(undoCtx, before.address);
          // "tanpa isian" terbaca sebagai warna kosong, yang ditolak saat ditulis balik: sel itu dikosongkan isiannya
          const unfilled = [];
          before.props.forEach((row, r) =>
            row.forEach((cell, c) => {
              if (cell.format.fill?.color === "") {
                delete cell.format.fill;
                unfilled.push([r, c]);
              }
            }),
          );
          target.setCellProperties(before.props);
          for (const [r, c] of unfilled) target.getCell(r, c).format.fill.clear();
          if (number_format) target.numberFormat = before.numberFormat;
          await undoCtx.sync();
        }),
      );
    }
    return { formatted: range.address };
  });
}

async function insertCells({ address }, context) {
  return Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address);
    const rows = /^\d+:\d+$/.test(split(address).cells.replace(/\$/g, ""));
    // alamat dicatat sebelum menyisipkan: sesudahnya objek rentang ini ikut bergeser
    range.load("address");
    await ctx.sync();
    const inserted = range.address;
    range.insert(rows ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
    await ctx.sync();
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        rangeOf(undoCtx, inserted).delete(rows ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
        await undoCtx.sync();
      }),
    );
    return { inserted };
  });
}

async function deleteCells({ address }) {
  return Excel.run(async (ctx) => {
    const rows = /^\d+:\d+$/.test(split(address).cells.replace(/\$/g, ""));
    rangeOf(ctx, address).delete(rows ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
    await ctx.sync();
    return { deleted: address };
  });
}

async function addSheet({ name }, context) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.add(name || undefined);
    sheet.activate();
    sheet.load("name");
    await ctx.sync();
    const created = sheet.name;
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.worksheets.getItem(created).delete();
        await undoCtx.sync();
      }),
    );
    return { sheet: created };
  });
}

async function renameSheet({ sheet, new_name }, context) {
  return Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(sheet).name = new_name;
    await ctx.sync();
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.worksheets.getItem(new_name).name = sheet;
        await undoCtx.sync();
      }),
    );
    return { sheet: new_name };
  });
}

async function deleteSheet({ sheet }) {
  return Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(sheet).delete();
    await ctx.sync();
    return { deleted: sheet };
  });
}

async function createTable({ address, has_headers, name, style }, context) {
  return Excel.run(async (ctx) => {
    const table = sheetOf(ctx, split(address).sheet).tables.add(address, has_headers !== false);
    if (name) table.name = name;
    if (style) table.style = style;
    table.load("name");
    await ctx.sync();
    const created = table.name;
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.tables.getItem(created).convertToRange();
        await undoCtx.sync();
      }),
    );
    return { table: created };
  });
}

async function sortRange({ address, keys, has_headers }, context) {
  const before = await snapshot(address);
  if (before.formulas) context.undo(restore(before));
  else if (!(await context.confirm(`Mengurutkan ${before.address} (tidak bisa dibatalkan)`))) throw new Error("Pengguna menolak tindakan ini.");
  await Excel.run(async (ctx) => {
    const fields = keys.map((k) => ({ key: Number(k.column), ascending: k.ascending !== false }));
    rangeOf(ctx, address).sort.apply(fields, false, has_headers !== false);
    await ctx.sync();
  });
  return { sorted: before.address };
}

async function filterRange({ address, column, values, criterion, clear }, context) {
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, split(address || "").sheet);
    if (clear) {
      sheet.autoFilter.remove();
      await ctx.sync();
      return { filter: "removed" };
    }
    const criteria = values?.length
      ? { filterOn: Excel.FilterOn.values, values: values.map(String) }
      : { filterOn: Excel.FilterOn.custom, criterion1: String(criterion) };
    sheet.autoFilter.apply(rangeOf(ctx, address), Number(column), criteria);
    sheet.load("name");
    await ctx.sync();
    const name = sheet.name;
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.worksheets.getItem(name).autoFilter.remove();
        await undoCtx.sync();
      }),
    );
    return { filtered: address, column };
  });
}

async function createChart({ source_address, type, title, anchor }, context) {
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, split(anchor || source_address).sheet);
    const chart = sheet.charts.add(type || "ColumnClustered", rangeOf(ctx, source_address), Excel.ChartSeriesBy.auto);
    if (title) chart.title.text = title;
    if (anchor) chart.setPosition(split(anchor).cells);
    chart.load("name");
    sheet.load("name");
    await ctx.sync();
    const created = { sheet: sheet.name, chart: chart.name };
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.worksheets.getItem(created.sheet).charts.getItem(created.chart).delete();
        await undoCtx.sync();
      }),
    );
    return created;
  });
}

async function createPivotTable({ name, source_address, destination_address, rows, columns, values, filters }, context) {
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, split(destination_address).sheet);
    const pivot = sheet.pivotTables.add(name, rangeOf(ctx, source_address), rangeOf(ctx, destination_address));
    for (const field of rows || []) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
    for (const field of columns || []) pivot.columnHierarchies.add(pivot.hierarchies.getItem(field));
    for (const field of filters || []) pivot.filterHierarchies.add(pivot.hierarchies.getItem(field));
    for (const value of values || []) {
      const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(value.field));
      if (value.summarize_by) data.summarizeBy = value.summarize_by;
    }
    sheet.load("name");
    await ctx.sync();
    const created = { sheet: sheet.name, pivot_table: name };
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        undoCtx.workbook.worksheets.getItem(created.sheet).pivotTables.getItem(name).delete();
        await undoCtx.sync();
      }),
    );
    return created;
  });
}

async function addConditionalFormat(args, context) {
  const { address, kind, operator, value1, value2, fill_color, font_color } = args;
  return Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address);
    let format;
    if (kind === "color_scale") {
      format = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      const point = Excel.ConditionalFormatColorCriterionType;
      format.colorScale.criteria = {
        minimum: { formula: null, type: point.lowestValue, color: "#F8696B" },
        midpoint: { formula: "50", type: point.percent, color: "#FFEB84" },
        maximum: { formula: null, type: point.highestValue, color: "#63BE7B" },
      };
    } else if (kind === "data_bar") {
      format = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    } else {
      format = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      if (fill_color) format.cellValue.format.fill.color = fill_color;
      if (font_color) format.cellValue.format.font.color = font_color;
      format.cellValue.rule = {
        formula1: String(value1),
        ...(value2 !== undefined && { formula2: String(value2) }),
        operator: operator || "GreaterThan",
      };
    }
    format.load("id");
    range.load("address");
    await ctx.sync();
    const created = { address: range.address, id: format.id };
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        rangeOf(undoCtx, created.address).conditionalFormats.getItem(created.id).delete();
        await undoCtx.sync();
      }),
    );
    return { conditional_format: kind, address: created.address };
  });
}

async function addDataValidation({ address, kind, source, min, max }, context) {
  return Excel.run(async (ctx) => {
    const range = rangeOf(ctx, address);
    if (kind === "list") {
      range.dataValidation.rule = { list: { inCellDropDown: true, source: String(source) } };
    } else {
      const numeric = { formula1: Number(min), formula2: Number(max), operator: Excel.DataValidationOperator.between };
      range.dataValidation.rule = kind === "whole_number" ? { wholeNumber: numeric } : { decimal: numeric };
    }
    range.load("address");
    await ctx.sync();
    const target = range.address;
    context.undo(() =>
      Excel.run(async (undoCtx) => {
        rangeOf(undoCtx, target).dataValidation.clear();
        await undoCtx.sync();
      }),
    );
    return { validation: kind, address: target };
  });
}

// Dipakai juga oleh tampilan saat sitasi sel diklik.
export async function selectRange(address) {
  return Excel.run(async (ctx) => {
    const { sheet } = split(address);
    if (sheet) sheetOf(ctx, sheet).activate();
    rangeOf(ctx, address).select();
    await ctx.sync();
  });
}

export const tools = [
  define(
    "get_workbook_overview",
    "read",
    "Struktur workbook: daftar sheet dengan rentang terpakainya, tabel, grafik, pivot table, nama rentang, sheet aktif, dan sel yang sedang dipilih. Panggil ini dulu sebelum bekerja.",
    {},
    [],
    () => "Melihat struktur workbook",
    overview,
  ),
  define(
    "read_range",
    "read",
    `Membaca isi sebuah rentang: nilai tiap sel dan rumus sel-sel yang berumus. Paling banyak ${MAX_READ} sel per panggilan.`,
    { address: ADDRESS, as: { type: "string", enum: ["values", "text"], description: 'Bentuk nilai. "values" (bawaan) = nilai mentah; "text" = seperti tampil di layar (tanggal, mata uang).' } },
    ["address"],
    (a) => `Membaca ${a.address}`,
    readRange,
  ),
  define(
    "find_cells",
    "read",
    "Mencari teks di sel (nilai maupun rumus) dan mengembalikan alamat sel yang cocok.",
    { query: text("Teks yang dicari, tidak peka huruf besar-kecil."), sheet: text("Nama sheet. Kosongkan untuk semua sheet.") },
    ["query"],
    (a) => `Mencari "${a.query}"`,
    findCells,
  ),
  define(
    "find_errors",
    "read",
    "Mencari sel berumus yang menghasilkan galat (#REF!, #DIV/0!, #N/A, dst.) beserta rumusnya.",
    { sheet: text("Nama sheet. Kosongkan untuk semua sheet.") },
    [],
    (a) => `Mencari galat rumus${a.sheet ? ` di ${a.sheet}` : ""}`,
    findErrors,
  ),
  define(
    "trace_cell",
    "read",
    "Melacak hubungan rumus satu langkah: sel yang dipakai sebuah rumus (precedents) atau sel yang bergantung padanya (dependents). Ulangi untuk menelusuri lebih jauh.",
    { address: text('Alamat satu sel, mis. "Ringkasan!C42".'), direction: { type: "string", enum: ["precedents", "dependents"], description: "Arah pelacakan. Bawaan: precedents." } },
    ["address"],
    (a) => `Melacak ${a.direction === "dependents" ? "sel yang bergantung pada" : "sumber"} ${a.address}`,
    traceCell,
  ),
  define(
    "write_range",
    "write",
    'Menulis nilai atau rumus ke sel. values adalah larik baris; tiap sel berupa string. Rumus diawali "=" dan memakai nama fungsi bahasa Inggris dengan pemisah koma. Untuk mengubah asumsi, tulis ke sel masukannya saja dan biarkan sel berumus menghitung ulang.',
    {
      address: text("Sel kiri-atas tujuan (atau rentangnya). Ukuran mengikuti values."),
      values: list({ type: "array", items: { type: "string" } }, 'Baris-baris sel, mis. [["Bulan","Total"],["Jan","=SUM(B2:B9)"]].'),
    },
    ["address", "values"],
    (a) => `Menulis ke ${a.address}`,
    writeRange,
  ),
  define(
    "clear_range",
    "write",
    "Mengosongkan isi, format, atau keduanya pada sebuah rentang.",
    { address: ADDRESS, what: { type: "string", enum: ["contents", "formats", "all"], description: "Yang dihapus. Bawaan: contents." } },
    ["address"],
    (a) => `Mengosongkan ${a.address}`,
    clearRange,
  ),
  define(
    "format_range",
    "write",
    "Mengatur tampilan sel: tebal, miring, warna, format angka, perataan, bungkus teks, garis tepi, lebar kolom otomatis. Isi hanya yang ingin diubah.",
    {
      address: ADDRESS,
      bold: flag("Tebal."),
      italic: flag("Miring."),
      font_color: text('Warna huruf, mis. "#C00000".'),
      fill_color: text('Warna isian sel, mis. "#FFF2CC".'),
      number_format: text('Kode format angka Excel, mis. "#,##0", "0.0%", "dd/mm/yyyy".'),
      align: { type: "string", enum: ["Left", "Center", "Right"], description: "Perataan mendatar." },
      wrap_text: flag("Bungkus teks."),
      border: { type: "string", enum: ["all", "outline", "none"], description: "Garis tepi." },
      autofit: flag("Sesuaikan lebar kolom dengan isinya."),
    },
    ["address"],
    (a) => `Memformat ${a.address}`,
    formatRange,
  ),
  define(
    "insert_rows_or_columns",
    "write",
    'Menyisipkan baris atau kolom kosong. Isi yang ada bergeser ke bawah atau ke kanan.',
    { address: text('Baris atau kolom utuh yang disisipkan. Contoh: "5:7" (tiga baris mulai baris 5), "C:D" (dua kolom mulai C), boleh dengan nama sheet.') },
    ["address"],
    (a) => `Menyisipkan ${a.address}`,
    insertCells,
  ),
  define(
    "delete_rows_or_columns",
    "danger",
    "Menghapus baris atau kolom utuh beserta isinya. Tidak bisa dibatalkan.",
    { address: text('Baris atau kolom utuh, mis. "5:7" atau "C:D", boleh dengan nama sheet.') },
    ["address"],
    (a) => `Menghapus baris/kolom ${a.address}`,
    deleteCells,
  ),
  define("add_sheet", "write", "Menambah sheet baru dan mengaktifkannya.", { name: text("Nama sheet. Kosongkan untuk nama bawaan.") }, [], (a) => `Menambah sheet${a.name ? ` "${a.name}"` : ""}`, addSheet),
  define("rename_sheet", "write", "Mengganti nama sheet.", { sheet: text("Nama sekarang."), new_name: text("Nama baru.") }, ["sheet", "new_name"], (a) => `Mengganti nama sheet "${a.sheet}" menjadi "${a.new_name}"`, renameSheet),
  define("delete_sheet", "danger", "Menghapus sheet beserta seluruh isinya. Tidak bisa dibatalkan.", { sheet: text("Nama sheet.") }, ["sheet"], (a) => `Menghapus sheet "${a.sheet}"`, deleteSheet),
  define(
    "create_table",
    "write",
    "Mengubah rentang data menjadi tabel Excel (dengan tombol filter dan gaya tabel).",
    { address: ADDRESS, has_headers: flag("Baris pertama adalah judul kolom. Bawaan: true."), name: text("Nama tabel, tanpa spasi."), style: text('Gaya tabel, mis. "TableStyleMedium2".') },
    ["address"],
    (a) => `Membuat tabel di ${a.address}`,
    createTable,
  ),
  define(
    "sort_range",
    "write",
    "Mengurutkan baris sebuah rentang menurut satu atau beberapa kolom.",
    {
      address: ADDRESS,
      keys: list(
        { type: "object", properties: { column: { type: "integer", description: "Nomor kolom di dalam rentang, mulai 0." }, ascending: flag("Naik (bawaan) atau turun.") }, required: ["column"] },
        "Kunci urut, dari yang utama.",
      ),
      has_headers: flag("Baris pertama adalah judul kolom dan tidak ikut diurutkan. Bawaan: true."),
    },
    ["address", "keys"],
    (a) => `Mengurutkan ${a.address}`,
    sortRange,
  ),
  define(
    "filter_range",
    "write",
    "Memasang filter pada satu kolom rentang, atau melepas filter sheet (clear: true). Baris yang tidak cocok disembunyikan, tidak dihapus.",
    {
      address: text("Rentang data beserta baris judulnya."),
      column: { type: "integer", description: "Nomor kolom di dalam rentang, mulai 0." },
      values: list({ type: "string" }, "Nilai yang ditampilkan."),
      criterion: text('Syarat sebagai pengganti values, mis. ">100" atau "=*Jakarta*".'),
      clear: flag("Lepas filter dari sheet."),
    },
    [],
    (a) => (a.clear ? "Melepas filter" : `Memfilter ${a.address}`),
    filterRange,
  ),
  define(
    "create_chart",
    "write",
    "Membuat grafik dari sebuah rentang data (termasuk baris judulnya).",
    {
      source_address: text("Rentang data sumber."),
      type: { type: "string", enum: CHART_TYPES, description: "Jenis grafik. Bawaan: ColumnClustered." },
      title: text("Judul grafik."),
      anchor: text('Sel tempat pojok kiri-atas grafik diletakkan, mis. "E2".'),
    },
    ["source_address"],
    (a) => `Membuat grafik dari ${a.source_address}`,
    createChart,
  ),
  define(
    "create_pivot_table",
    "write",
    "Membuat pivot table dari rentang data. Nama field harus sama persis dengan judul kolom di data sumber.",
    {
      name: text("Nama pivot table."),
      source_address: text("Rentang data sumber beserta baris judulnya."),
      destination_address: text('Sel kiri-atas tujuan, sebaiknya di sheet lain, mis. "Pivot!A3".'),
      rows: list({ type: "string" }, "Field untuk baris."),
      columns: list({ type: "string" }, "Field untuk kolom."),
      filters: list({ type: "string" }, "Field untuk filter."),
      values: list(
        { type: "object", properties: { field: text("Judul kolom yang dihitung."), summarize_by: { type: "string", enum: ["Sum", "Count", "Average", "Max", "Min"], description: "Cara menghitung. Bawaan: Sum." } }, required: ["field"] },
        "Field nilai.",
      ),
    },
    ["name", "source_address", "destination_address", "values"],
    (a) => `Membuat pivot table "${a.name}"`,
    createPivotTable,
  ),
  define(
    "add_conditional_format",
    "write",
    "Menambah format bersyarat: skala warna, batang data, atau sorotan sel yang memenuhi syarat.",
    {
      address: ADDRESS,
      kind: { type: "string", enum: ["cell_value", "color_scale", "data_bar"], description: "Jenis format bersyarat." },
      operator: { type: "string", enum: OPERATORS, description: "Untuk cell_value: pembanding." },
      value1: text('Untuk cell_value: nilai atau rumus pembanding, mis. "0" atau "=$B$1".'),
      value2: text("Untuk Between/NotBetween: batas kedua."),
      fill_color: text("Untuk cell_value: warna isian sel yang cocok."),
      font_color: text("Untuk cell_value: warna huruf sel yang cocok."),
    },
    ["address", "kind"],
    (a) => `Menambah format bersyarat di ${a.address}`,
    addConditionalFormat,
  ),
  define(
    "add_data_validation",
    "write",
    "Memasang validasi data: daftar pilihan (dropdown) atau batas angka.",
    {
      address: ADDRESS,
      kind: { type: "string", enum: ["list", "whole_number", "decimal"], description: "Jenis validasi." },
      source: text('Untuk list: pilihan dipisah koma ("Ya,Tidak") atau rujukan rentang ("=Data!A2:A9").'),
      min: { type: "number", description: "Untuk angka: batas bawah." },
      max: { type: "number", description: "Untuk angka: batas atas." },
    },
    ["address", "kind"],
    (a) => `Memasang validasi data di ${a.address}`,
    addDataValidation,
  ),
];
