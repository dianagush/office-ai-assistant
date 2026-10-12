// Mengirim pembaruan add-in ke GitHub, dan dari sana ke GitHub Pages: periksa rahasia, commit, push.
//   npm run rilis                          → pesan commit otomatis (tanggal dan jam)
//   npm run rilis -- "Perbaiki tabel Excel"
//   npm run rilis -- --dry-run             → hanya menampilkan apa yang akan dikirim
//   npm run rilis -- --yes                 → tanpa tanya konfirmasi
//
// Pages butuh sekitar 1 menit untuk membangun ulang. Panel yang sedang terbuka perlu ditutup lalu
// dibuka lagi. Kalau manifest.xml berubah, manifest.prod.xml dibuat ulang otomatis; setiap komputer
// pengguna lalu perlu sideload ulang (npm run sideload).
const { spawnSync } = require("child_process");
const path = require("path");
const readline = require("readline");

const root = path.join(__dirname, "..");

// Berkas yang tidak boleh ikut terunggah, walaupun .gitignore sudah menahannya.
const FORBIDDEN = [
  /(^|\/)\.dev-token$/,
  /(^|\/)\.env(\..*)?$/,
  /(^|\/)manifest\.http\.xml$/,
  /\.(pem|key|pfx|crt)$/i,
  /^node_modules\//,
  /^Dokumen tersimpan\//,
];

// Isi yang tidak diperiksa: berkas pihak ketiga yang besar dan biner.
const SKIP_CONTENT = /^docs\/vendor\/|^package-lock\.json$|\.(png|ico|jpe?g|gif)$/i;

const SECRETS = [
  ["API key Google/Gemini", /AIza[0-9A-Za-z_-]{35}/],
  ["API key bergaya OpenAI/Claude (sk-...)", /\bsk-(?:ant-|proj-|live-)?[A-Za-z0-9_-]{20,}/],
  ["token GitHub", /\b(?:gh[pousr]_[A-Za-z0-9]{36,}|github_pat_[A-Za-z0-9_]{20,})/],
  ["token Slack", /\bxox[abprs]-[A-Za-z0-9-]{10,}/],
  ["kunci akses AWS", /\bAKIA[0-9A-Z]{16}\b/],
  ["kunci privat", /-----BEGIN [A-Z ]*PRIVATE KEY-----/],
  ["isian key/secret/token/password yang berisi nilai", /(?:api[_-]?key|secret|token|password)["']?\s*[:=]\s*["'][A-Za-z0-9_\-/+=.]{16,}["']/i],
];

function git(...args) {
  const result = spawnSync("git", ["-c", "core.quotepath=off", ...args], { cwd: root, encoding: "utf8", maxBuffer: 256 * 1024 * 1024 });
  return { ok: result.status === 0, out: (result.stdout || "").trim(), err: (result.stderr || "").trim() };
}

function fail(message, code = 1) {
  console.error(`\n${message}`);
  process.exit(code);
}

function scanLine(line) {
  return SECRETS.filter(([, pattern]) => pattern.test(line)).map(([name]) => name);
}

// Baris yang ditambahkan oleh diff, lengkap dengan berkas dan nomor barisnya.
function addedLines(diff) {
  const added = [];
  let file = null;
  let line = 0;
  for (const text of diff.split(/\r?\n/)) {
    if (text.startsWith("+++ b/") || text === "+++ /dev/null") {
      file = text === "+++ /dev/null" ? null : text.slice(6);
      continue;
    }
    const hunk = text.match(/^@@ -\S+ \+(\d+)/);
    if (hunk) {
      line = Number(hunk[1]);
      continue;
    }
    if (file && text.startsWith("+")) added.push({ file, line: line++, text: text.slice(1) });
  }
  return added;
}

// https://github.com/PEMILIK/REPO.git → https://pemilik.github.io/REPO
function pagesUrl(remote) {
  const match = remote.match(/github\.com[:/]([^/]+)\/([^/]+?)(?:\.git)?$/);
  return match ? `https://${match[1].toLowerCase()}.github.io/${match[2]}` : null;
}

function ask(question) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(question, (answer) => (rl.close(), resolve(answer.trim().toLowerCase())));
  });
}

async function main() {
  const args = process.argv.slice(2);
  const dryRun = args.includes("--dry-run");
  const yes = args.includes("--yes");
  const message = args.filter((a) => !a.startsWith("--")).join(" ").trim();

  if (!git("rev-parse", "--is-inside-work-tree").ok) fail("Folder ini belum berupa repositori git.");
  const remote = git("remote", "get-url", "origin");
  if (!remote.ok) fail("Belum ada remote 'origin'. Hubungkan dulu: git remote add origin https://github.com/NAMA/REPO.git");
  const branch = git("rev-parse", "--abbrev-ref", "HEAD").out;
  const site = pagesUrl(remote.out);

  const manifestChanged = Boolean(git("status", "--porcelain", "--", "manifest.xml").out);
  if (manifestChanged && site && !dryRun) {
    spawnSync(process.execPath, [path.join(__dirname, "make-manifest.js"), site], { cwd: root, stdio: "inherit" });
  }

  git("add", "-A");
  const unstage = () => git("reset", "-q");
  const names = git("diff", "--cached", "--name-only").out.split("\n").filter(Boolean);

  if (!names.length) {
    const upstream = git("rev-list", "--count", `origin/${branch}..HEAD`);
    const pending = upstream.ok ? Number(upstream.out) : 0;
    if (!pending) {
      console.log("Tidak ada perubahan untuk dikirim.");
      return;
    }
    console.log(`Tidak ada perubahan baru, tetapi ada ${pending} commit yang belum terkirim ke GitHub.`);
    if (dryRun) return console.log("(dry-run: tidak ada yang dikirim)");
    if (!yes && !(await confirm())) return console.log("Dibatalkan.");
    return push(branch, site, false);
  }

  const forbidden = names.filter((name) => FORBIDDEN.some((pattern) => pattern.test(name)));
  const found = addedLines(git("diff", "--cached", "-U0", "--no-color", "--src-prefix=a/", "--dst-prefix=b/").out)
    .filter(({ file }) => !SKIP_CONTENT.test(file))
    .flatMap(({ file, line, text }) => scanLine(text).map((name) => ({ file, line, name })));
  if (forbidden.length || found.length) {
    unstage();
    console.error("\nDIBATALKAN: ada yang tampak seperti rahasia. Tidak ada yang di-commit atau dikirim.\n");
    for (const name of forbidden) console.error(`  Berkas terlarang: ${name}`);
    for (const { file, line, name } of found) console.error(`  ${file}:${line}  ${name}`);
    console.error("\nHapus rahasianya dari berkas (atau masukkan berkasnya ke .gitignore), lalu jalankan lagi.");
    process.exit(1);
  }

  console.log("Yang akan dikirim:\n");
  console.log(git("diff", "--cached", "--stat=100", "--stat-name-width=70").out);
  console.log("\nPemeriksaan rahasia: bersih.");
  if (names.some((name) => name.startsWith("docs/"))) console.log("Perubahan di docs/ langsung tayang ke semua pengguna setelah Pages selesai membangun.");
  if (manifestChanged) console.log("manifest.xml berubah: komputer pengguna perlu sideload ulang (npm run sideload) dengan Word ditutup.");

  if (dryRun) {
    unstage();
    return console.log("\n(dry-run: tidak ada yang di-commit atau dikirim)");
  }
  if (!yes && !(await confirm())) {
    unstage();
    return console.log("Dibatalkan. Tidak ada yang di-commit.");
  }

  const stamp = new Date().toLocaleString("sv-SE").slice(0, 16);
  const commit = git("commit", "-q", "-m", message || `Pembaruan ${stamp}`);
  if (!commit.ok) fail(`Commit gagal:\n${commit.err || commit.out}`);
  push(branch, site, manifestChanged);
}

async function confirm() {
  if (!process.stdin.isTTY) {
    git("reset", "-q");
    fail("Tidak bisa bertanya konfirmasi di sini, jadi tidak ada yang di-commit. Tambahkan --yes kalau memang mau langsung dikirim.");
  }
  return (await ask("\nKirim ke GitHub? (y/n): ")) === "y";
}

function push(branch, site, manifestChanged) {
  const result = spawnSync("git", ["push", "-u", "origin", branch], { cwd: root, stdio: "inherit" });
  if (result.status !== 0) {
    fail("Push gagal. Perubahannya sudah ter-commit di komputer ini; jalankan perintah ini lagi untuk mencoba mengirim ulang.");
  }
  console.log("\nTerkirim ke GitHub.");
  if (site) console.log(`GitHub Pages membangun ulang dalam sekitar 1 menit: ${site}/taskpane.html\nTutup panel AI Assistant lalu buka lagi untuk memuat versi baru.`);
  if (manifestChanged) console.log("Ingat: karena manifest berubah, jalankan npm run sideload lagi di tiap komputer pengguna.");
}

module.exports = { scanLine, addedLines, pagesUrl, FORBIDDEN };
if (require.main === module) main();
