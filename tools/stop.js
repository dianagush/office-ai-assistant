// Menonaktifkan add-in: mematikan server lokal dan melepas add-in dari Office.
//   npm stop
// Aplikasi Office yang sedang terbuka tidak ditutup; tombol AI Assistant hilang setelah
// aplikasinya dibuka ulang.
const { spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const { PID_FILE } = require("../server");

const root = path.join(__dirname, "..");
const manifests = ["manifest.http.xml", "manifest.prod.xml"].map((name) => path.join(root, name)).filter((file) => fs.existsSync(file));

// Proses yang sedang mendengarkan di sebuah port.
function listeners(port) {
  const output = spawnSync("netstat", ["-ano"], { encoding: "utf8" }).stdout || "";
  const pids = output
    .split(/\r?\n/)
    .filter((line) => /^\s*TCP\s/i.test(line) && /LISTENING/i.test(line) && line.trim().split(/\s+/)[1].endsWith(`:${port}`))
    .map((line) => Number(line.trim().split(/\s+/).pop()));
  return [...new Set(pids)].filter(Boolean);
}

// Server mencatat "<pid> <port>" saat menyala. Catatan itu hanya dipercaya kalau proses tersebut
// memang masih mendengarkan di port itu; kalau tidak, catatannya sudah basi dan tidak ada yang dimatikan.
function stopServer() {
  if (!fs.existsSync(PID_FILE)) return false;
  const [pid, port] = fs.readFileSync(PID_FILE, "utf8").trim().split(/\s+/).map(Number);
  fs.rmSync(PID_FILE, { force: true });
  if (!pid || !port || !listeners(port).includes(pid)) return false;
  try {
    process.kill(pid);
    return true;
  } catch (error) {
    console.error(`Server add-in (PID ${pid}) gagal dimatikan: ${error.message}`);
    return false;
  }
}

console.log(stopServer() ? "Server add-in dimatikan." : "Server add-in memang tidak sedang menyala.");

for (const manifest of manifests) {
  spawnSync(process.execPath, [require.resolve("office-addin-dev-settings/cli.js"), "unregister", manifest], { stdio: "ignore" });
}
if (manifests.length) console.log("Add-in dilepas dari Office.");
console.log("\nAdd-in nonaktif. Aplikasi Office yang sedang terbuka tidak ditutup; tombol AI Assistant hilang setelah aplikasinya dibuka ulang.");
