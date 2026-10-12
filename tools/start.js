// Mengaktifkan add-in: menyalakan server lokal (HTTP, tanpa sertifikat), mendaftarkan add-in ke
// Office, lalu membuka aplikasinya dengan panel AI Assistant.
//   npm start                                  → Word, server di jendela ini (Ctrl+C untuk berhenti)
//   node tools/start.js word excel             → beberapa aplikasi sekaligus
//   node tools/start.js word --background      → server di latar belakang; matikan dengan tools/stop.js
//   node tools/start.js word --hosted          → tanpa server: add-in dimuat dari GitHub Pages
//                                                (alamatnya dibaca dari manifest.prod.xml)
//
// Aplikasinya harus dibuka lewat perintah ini: add-in pengembang hanya termuat pada dokumen yang
// ikut menyalakan aplikasi, bukan pada dokumen yang dibuka ke aplikasi yang sudah berjalan.
const { spawn, spawnSync } = require("child_process");
const fs = require("fs");
const http = require("http");
const https = require("https");
const net = require("net");
const os = require("os");
const path = require("path");

const PORT = 3000;
const PROCESS = { word: "WINWORD.EXE", excel: "EXCEL.EXE", powerpoint: "POWERPNT.EXE" };
const DOCUMENT = { word: ["Word", ".docx"], excel: ["Excel", ".xlsx"], powerpoint: ["PowerPoint", ".pptx"] };
const root = path.join(__dirname, "..");
const KEPT = path.join(root, "Dokumen tersimpan");
const STAMPS = path.join(os.tmpdir(), "office-ai-addin-documents.json");
const args = process.argv.slice(2);
const background = args.includes("--background");
const hosted = args.includes("--hosted");
const manifest = path.join(root, hosted ? "manifest.prod.xml" : "manifest.http.xml");
const apps = args.filter((a) => !a.startsWith("--")).map((a) => a.toLowerCase());
if (!apps.length) apps.push("word");

const node = (script, ...rest) => spawnSync(process.execPath, [script, ...rest], { stdio: "inherit" });
const settings = (...rest) => node(require.resolve("office-addin-dev-settings/cli.js"), ...rest);

const listening = () =>
  new Promise((resolve) => {
    const socket = net.connect(PORT, "localhost");
    socket.once("connect", () => (socket.destroy(), resolve(true)));
    socket.once("error", () => resolve(false));
  });

// Yang menjawab di port itu benar server add-in ini, bukan program lain yang kebetulan memakainya?
const serving = () =>
  new Promise((resolve) => {
    const request = http.get(`http://localhost:${PORT}/taskpane.html`, (response) => {
      let body = "";
      response.on("data", (chunk) => (body += chunk));
      response.on("end", () => resolve(response.statusCode === 200 && body.includes("taskpane.js")));
    });
    request.on("error", () => resolve(false));
    request.setTimeout(3000, () => request.destroy());
  });

const reachable = (url) =>
  new Promise((resolve) => {
    const request = https.get(url, (response) => (response.resume(), resolve(response.statusCode === 200)));
    request.on("error", () => resolve(false));
    request.setTimeout(5000, () => request.destroy());
  });

function running(image) {
  if (process.platform !== "win32") return false;
  const result = spawnSync("tasklist", ["/FI", `IMAGENAME eq ${image}`, "/NH"], { encoding: "utf8" });
  return (result.stdout || "").toUpperCase().includes(image);
}

// Dokumen kosong yang terbuka bersama add-in adalah berkas sementara: sideload menghapus dan membuatnya
// ulang setiap kali. Supaya pekerjaan yang telanjur disimpan di situ tidak ikut terhapus, waktu ubah
// berkasnya dicatat sesudah dibuat; kalau kemudian berbeda, berkasnya disalin dulu ke "Dokumen tersimpan".
function sideloadDocument(app) {
  const id = (fs.readFileSync(manifest, "utf8").match(/<Id>([^<]+)<\/Id>/) || [])[1];
  const [name, extension] = DOCUMENT[app];
  return path.join(os.tmpdir(), `${name} add-in ${id}${extension}`);
}

function readStamps() {
  try {
    return JSON.parse(fs.readFileSync(STAMPS, "utf8"));
  } catch {
    return {};
  }
}

function rescue(app) {
  const file = sideloadDocument(app);
  const stamp = readStamps()[file];
  if (!stamp || !fs.existsSync(file)) return;
  const changed = fs.statSync(file).mtime;
  if (changed.getTime() === stamp) return;
  const two = (n) => String(n).padStart(2, "0");
  const when = `${changed.getFullYear()}-${two(changed.getMonth() + 1)}-${two(changed.getDate())} ${two(changed.getHours())}.${two(changed.getMinutes())}.${two(changed.getSeconds())}`;
  const copy = path.join(KEPT, `${DOCUMENT[app][0]} ${when}${DOCUMENT[app][1]}`);
  fs.mkdirSync(KEPT, { recursive: true });
  fs.copyFileSync(file, copy);
  console.log(`\nAnda sempat menyimpan pekerjaan di dokumen sementara add-in. Salinannya diamankan di:\n  ${copy}\n`);
}

function remember(app) {
  const file = sideloadDocument(app);
  if (!fs.existsSync(file)) return;
  fs.writeFileSync(STAMPS, JSON.stringify({ ...readStamps(), [file]: fs.statSync(file).mtime.getTime() }));
}

(async () => {
  const unknown = apps.filter((app) => !PROCESS[app]);
  if (unknown.length) {
    console.error(`Tidak dikenal: ${unknown.join(", ")}. Pilihannya: word, excel, powerpoint.`);
    process.exit(1);
  }

  let server = null;
  if (hosted) {
    if (!fs.existsSync(manifest)) {
      console.error("manifest.prod.xml tidak ada. Buat dulu: npm run manifest -- https://NAMA.github.io/REPO");
      process.exit(1);
    }
    const page = (fs.readFileSync(manifest, "utf8").match(/<SourceLocation DefaultValue="([^"]+)"/) || [])[1];
    if (page && (await reachable(page))) console.log(`Add-in dimuat dari ${page}`);
    else console.log(`Peringatan: ${page || "alamat add-in"} tidak bisa dijangkau. Periksa koneksi internet; panelnya tidak akan terbuka tanpa itu.`);
  } else if (await listening()) {
    if (!(await serving())) {
      console.error(`Port ${PORT} sedang dipakai program lain, jadi server add-in tidak bisa menyala. Tutup program itu lalu coba lagi.`);
      process.exit(1);
    }
    console.log(`Server add-in sudah menyala di port ${PORT}.`);
  } else {
    server = spawn(process.execPath, [path.join(root, "server.js"), "--http"], background ? { detached: true, stdio: "ignore", windowsHide: true } : { stdio: "inherit" });
    for (let i = 0; i < 25 && !(await listening()); i++) await new Promise((resolve) => setTimeout(resolve, 200));
    if (!(await listening())) {
      console.error(`Server add-in gagal menyala di port ${PORT}.`);
      process.exit(1);
    }
    if (background) {
      server.unref();
      console.log(`Server add-in menyala di latar belakang (http://localhost:${PORT}).`);
    }
  }

  if (!hosted) node(path.join(__dirname, "make-manifest.js"), `http://localhost:${PORT}`);
  settings("register", manifest);
  const skipped = [];
  for (const app of apps) {
    if (running(PROCESS[app])) {
      skipped.push(app);
      continue;
    }
    rescue(app);
    settings("sideload", manifest, "desktop", "--app", app);
    remember(app);
  }

  if (skipped.length) {
    // Tanda untuk "Aktifkan AI Assistant.cmd": jangan menutup jendela sebelum pesan ini terbaca.
    if (background) process.exitCode = 3;
    console.log(
      `\n${skipped.join(", ")} sedang berjalan, jadi tidak dibuka ulang. Kalau panel AI Assistant sudah terbuka di sana, ` +
        "panel itu bisa dipakai lagi. Kalau belum ada, simpan pekerjaan Anda, tutup aplikasinya, lalu jalankan ini lagi.",
    );
  }
  console.log(
    "\nDokumen kosong yang terbuka bersama add-in bersifat sementara dan dibuat ulang setiap aktivasi. " +
      "Simpan pekerjaan Anda dengan Save As ke folder lain, atau buka berkas Anda sendiri lewat File > Open.",
  );
  console.log("Buka aplikasinya lewat skrip ini setiap kali. Kalau dibuka dengan cara biasa (menu Start, klik dua kali dokumen), tombol AI Assistant tidak muncul.");
  if (background) console.log('\nAdd-in aktif. Untuk melepasnya dari Office, jalankan "Nonaktifkan AI Assistant.cmd".');
  else if (server) console.log("\nBiarkan jendela ini terbuka selama add-in dipakai. Ctrl+C untuk mematikan server.");
})();
