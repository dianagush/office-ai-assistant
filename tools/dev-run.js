// Menjalankan potongan JS di task pane lewat jembatan uji (butuh `node server.js --dev`).
//   node tools/dev-run.js [Word|Excel|PowerPoint|id-klien] < potongan.js
//   node tools/dev-run.js --clients
const fs = require("fs");
const path = require("path");

const base = `http://localhost:${process.env.PORT || 3000}/dev`;
const headers = { "x-dev-token": fs.readFileSync(path.join(__dirname, "..", ".dev-token"), "utf8") };
const target = process.argv[2] || "";

const request =
  target === "--clients"
    ? fetch(`${base}/clients`, { headers })
    : fetch(`${base}/run?client=${encodeURIComponent(target)}&timeout=${process.env.TIMEOUT || 60000}`, {
        method: "POST",
        headers,
        body: fs.readFileSync(0, "utf8"),
      });

request
  .then((res) => res.json())
  .then((result) => {
    const failed = result && result.ok === false;
    console.log(JSON.stringify(result && result.ok ? result.value : result, null, 2));
    process.exit(failed || result.error ? 1 : 0);
  })
  .catch((error) => {
    console.error(error.message);
    process.exit(1);
  });
