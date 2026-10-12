// Membuat manifest untuk alamat lain dari manifest.xml (yang menunjuk https://localhost:3000).
//   npm run manifest -- https://NAMA.github.io/REPO   → manifest.prod.xml (add-in yang di-host)
//   npm run manifest -- http://localhost:3000         → manifest.http.xml (lokal tanpa sertifikat)
const fs = require("fs");
const path = require("path");

const DEV_URL = "https://localhost:3000";
const DEV_ID = "6445989a-167b-4ac2-800f-22e01ecb4e53";
// ID berbeda per varian supaya semuanya bisa terpasang bersamaan.
const VARIANTS = {
  hosted: { id: "de80e411-c17e-4588-9645-646436fc62f5", file: "manifest.prod.xml" },
  http: { id: "3f1c2a58-9b6e-4d0a-8c41-7e5d2b9a6f10", file: "manifest.http.xml" },
};

const base = (process.argv[2] || "").replace(/\/+$/, "");
const localHttp = /^http:\/\/localhost(:\d+)?$/.test(base);
if (!localHttp && !/^https:\/\/[^/]+/.test(base)) {
  console.error(
    "Beri alamat tempat folder docs/ tersaji: HTTPS, atau http://localhost untuk percobaan lokal.\n" +
      "Contoh: npm run manifest -- https://NAMA.github.io/REPO",
  );
  process.exit(1);
}

const variant = localHttp ? VARIANTS.http : VARIANTS.hosted;
const root = path.join(__dirname, "..");
const manifest = fs
  .readFileSync(path.join(root, "manifest.xml"), "utf8")
  .replace(/<!--[\s\S]*?-->\r?\n/, "")
  .replaceAll(DEV_URL, base)
  .replace(DEV_ID, variant.id);
fs.writeFileSync(path.join(root, variant.file), manifest);
console.log(`${variant.file} dibuat untuk ${base}`);
