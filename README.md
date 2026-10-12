# AI Assistant untuk Word, Excel, dan PowerPoint

Panel samping di Office tempat AI membaca dan mengubah dokumen langsung, memakai API key Anda sendiri: Gemini,
DeepSeek, Claude, OpenAI, OpenRouter, Groq, Ollama, 9router, atau API lain yang OpenAI-compatible.

## Cara kerja

```
Word / Excel / PowerPoint
└─ task pane (HTML + JS statis dari folder docs/)
   ├─ loop agen   → model meminta alat, panel menjalankannya, hasilnya dikirim balik, berulang
   ├─ Office.js   → alat-alat itu: baca, tulis, format, tabel, grafik, slide, tracked changes
   └─ fetch()     → langsung ke API penyedia AI (streaming)
```

Tidak ada server perantara. Add-in hanyalah berkas statis, jadi bisa di-host gratis, dan API key tidak pernah lewat
pihak ketiga.

## Yang bisa dilakukan

| Aplikasi | Kemampuan agen |
| --- | --- |
| Excel | Melihat struktur workbook lintas sheet, membaca rentang beserta rumusnya, mencari sel, mencari dan melacak galat rumus, menulis nilai/rumus, format, sisip/hapus baris dan kolom, sheet, tabel, urut, filter, grafik, pivot table, format bersyarat, validasi data. Rujukan sel di jawaban bisa diklik. |
| Word | Melihat kerangka dokumen, membaca paragraf, mencari, mengganti potongan teks, menulis ulang paragraf, menyisipkan konten dengan style dokumen (judul, daftar, tabel), gaya paragraf, format huruf, tabel, komentar. Perubahan dicatat sebagai tracked changes. |
| PowerPoint | Melihat slide dan layout template, membaca slide, menambah slide dari layout, mengisi placeholder, kotak teks, tabel, slide bergrafik, format dan posisi bentuk, memindah dan menghapus slide. |

Di semua aplikasi:

- **Izin.** Tindakan yang tidak bisa dibatalkan (menghapus sheet, baris, slide, bentuk) dan penimpaan banyak data
  selalu minta izin dulu. Bisa juga diatur supaya setiap perubahan minta izin.
- **Batalkan.** Setelah AI mengubah sesuatu, muncul ringkasan perubahan dengan tombol **Batalkan** (Excel,
  PowerPoint) atau **Terima / Tolak** tracked changes (Word).
- **Riwayat percakapan.** Semua percakapan tersimpan di perangkat ini (60 terbaru, untuk Word, Excel, dan PowerPoint
  sekaligus). Tombol jam di bilah atas membuka daftarnya, yang otomatis hanya menampilkan riwayat aplikasi yang
  sedang dipakai (tombol **Tampilkan semua** membuka riwayat aplikasi lain). Pilih untuk melanjutkan, atau hapus
  satu per satu atau semua yang sedang tampil. **Chat baru** tidak menghapus percakapan sebelumnya. Saat berkas yang sama dibuka lagi, percakapan
  terakhirnya muncul otomatis. Percakapan yang sangat panjang dipadatkan otomatis menjadi ringkasan.
- **Lintas aplikasi.** Kalau panel AI Assistant juga terbuka di aplikasi Office lain, agen bisa memakai alat aplikasi
  itu dalam percakapan yang sama, misalnya mengambil angka dari Excel lalu menuliskannya sebagai tabel di Word.
  Aplikasi yang terhubung tampil di bawah bilah atas. Tindakan yang tidak bisa dibatalkan hanya bisa dilakukan dari
  panel aplikasinya sendiri. Bisa dimatikan di Pengaturan.
- **Instruksi tetap** per aplikasi, dan **Skills**: resep instruksi buatan sendiri yang dibaca AI saat tugasnya cocok.
- **Tombol cepat** untuk teks terpilih: Ringkas, Perbaiki, Tulis ulang, Persingkat, terjemahan.
- Mode agen bisa dimatikan di Pengaturan; AI lalu hanya menjawab dan Anda yang menyisipkan hasilnya.

## Biaya

Add-in, pemasangan, dan hosting gratis. Yang berbayar hanya pemakaian AI, dan itu memakai **API key, bukan langganan
chat**: Gemini Advanced, ChatGPT Plus, atau Claude Pro tidak menyertakan akses API.

- **Gemini** punya paket API gratis, tetapi kuotanya kecil dan dihitung per model per hari. Saat diuji (Oktober
  2026), Gemini 3.6 Flash dibatasi 20 permintaan per hari, dan satu tugas agen memakai 3 sampai 6 permintaan. Karena
  itu model bawaan add-in adalah Flash-Lite (`gemini-flash-lite-latest`). Batas akun Anda terlihat di
  <https://aistudio.google.com/rate-limit>. Untuk pemakaian rutin, aktifkan penagihan (bayar per pemakaian) lalu
  pilih Flash atau Pro. Di paket gratis, Google memakai isi permintaan untuk meningkatkan produknya, jadi hindari
  dokumen rahasia.
- **DeepSeek** prabayar: isi saldo, lalu dipotong per token.
- **Ollama** menjalankan model di komputer sendiri: gratis dan dokumen tidak keluar dari perangkat. Hasil mode agen
  sangat bergantung pada kemampuan modelnya memanggil alat.

## Menjalankan

Butuh Node.js, dan Microsoft 365 atau Office 2021 ke atas.

```bash
npm install
```

Cara termudah di Windows, tanpa terminal: klik dua kali

- **`Aktifkan AI Assistant.cmd`** — pilih Word, Excel, PowerPoint, atau ketiganya. Aplikasinya terbuka dengan panel
  AI Assistant, dimuat dari GitHub Pages (tanpa server lokal, perlu internet). Pilihannya diingat: lain kali cukup
  klik dua kali dan tunggu 3 detik (tekan **G** kalau ingin memilih aplikasi lain). Jendela skripnya menutup sendiri
  kalau semuanya lancar; kalau ada yang perlu dibaca, jendelanya tetap terbuka.
- **`Nonaktifkan AI Assistant.cmd`** — melepas add-in dari Office (dan mematikan server lokal kalau menyala).
  Aplikasi Office yang sedang terbuka tidak ditutup.

**Buka aplikasinya lewat skrip ini setiap kali.** Word, Excel, atau PowerPoint yang dibuka dengan cara biasa (menu
Start, klik dua kali dokumen, pintasan taskbar) tidak memuat add-in, jadi tombol AI Assistant tidak muncul. Add-in
pengembang hanya termuat pada aplikasi yang dimulai lewat skrip.

**Agar tombolnya selalu muncul, walau aplikasinya dibuka dengan cara biasa:** klik dua kali
**`Pasang Permanen.cmd`** sekali saja dan setujui permintaan administrator dari Windows. Skrip itu membuat satu
folder berbagi hanya-baca berisi `manifest.prod.xml` dan mendaftarkannya sebagai katalog tepercaya di Office. Karena
mengubah pengaturan keamanan Office, skrip ini sengaja dijalankan sendiri oleh pemilik komputer. Sesudahnya, di
tiap aplikasi sekali saja: **Insert → Add-ins → My Add-ins → tab Shared Folder → AI Assistant → Add**. Setelah itu
tombol AI Assistant ada di tab Home setiap kali. `Pasang Permanen.cmd -Periksa` hanya menampilkan rencananya tanpa
mengubah apa pun; `-Lepas` membatalkan semuanya. Skrip ini belum diuji sampai ke tampilan Word-nya.

**Dokumen kosong yang terbuka bersama add-in bersifat sementara.** Berkasnya ada di folder Temp dan dibuat ulang
setiap kali add-in diaktifkan. Simpan pekerjaan dengan Save As ke folder lain, atau buka berkas Anda sendiri lewat
File → Open. Kalau telanjur menyimpan di dokumen sementara itu, aktivasi berikutnya menyalinnya dulu ke folder
`Dokumen tersimpan` di dalam proyek.

Mode pengembangan (server lokal, untuk mencoba perubahan sebelum dikirim ke GitHub), lewat terminal:

```bash
npm start
```

`npm start` menyalakan server lokal di `http://localhost:3000` dan membuka Word dengan add-in terpasang. Untuk Excel
atau PowerPoint: `npm run start:excel` / `npm run start:powerpoint`; ketiganya sekaligus: `npm run start:all`.
Tidak perlu sertifikat di Windows.

Dua hal yang perlu diketahui:

- **Tutup dulu aplikasinya.** Add-in hanya termuat kalau Word (atau Excel/PowerPoint) dimulai lewat skrip atau
  perintah itu. Dokumen yang dibuka ke aplikasi yang sudah berjalan tidak memuatnya.
- **Server harus menyala** selama add-in dipakai (khusus `npm start`): jendela terminalnya jangan ditutup. Skrip
  `.cmd` tidak memakai server, jadi bagian ini tidak berlaku untuknya.

Setelah panel terbuka: tempel API key di **Pengaturan**, klik **Muat daftar model** kalau ingin memilih model, lalu
mulai memberi perintah. Untuk memakai add-in pada berkas Anda sendiri, buka berkasnya dari Word yang dimulai tadi
(File → Open); tombol **AI Assistant** ada di tab Home.

`npm stop` mematikan server (kalau berjalan di latar belakang) dan melepas add-in dari Office.

### Tanpa server lokal (GitHub Pages)

Supaya tidak perlu menyalakan server setiap kali:

1. Unggah folder proyek ini ke repositori publik di GitHub.
2. Di GitHub: **Settings → Pages → Deploy from a branch → `main`, folder `/docs`**.
3. `npm run manifest -- https://NAMA.github.io/REPO`
4. Tutup Word, lalu `npm run sideload` (`sideload:excel`, `sideload:powerpoint` untuk yang lain).

**Memperbarui add-in.** Ubah berkas di `docs/`, lalu jalankan:

```bash
npm run rilis -- "Pesan singkat tentang perubahannya"
```

Perintah ini memeriksa rahasia (API key, token, kunci privat, berkas `.env`), menampilkan berkas yang akan dikirim,
bertanya konfirmasi, lalu commit dan push. Kalau ada yang tampak seperti rahasia, tidak ada yang dikirim. Pages
membangun ulang dalam sekitar 1 menit; tutup panel AI Assistant lalu buka lagi untuk memuat versi baru (cache
browser sampai 10 menit). Tambahkan `--dry-run` untuk hanya melihat apa yang akan dikirim. Kalau `manifest.xml`
berubah, `manifest.prod.xml` dibuat ulang otomatis dan setiap komputer pengguna perlu `npm run sideload` lagi.

Cara resmi Microsoft tanpa perintah: taruh `manifest.prod.xml` di folder jaringan, daftarkan di **File → Options →
Trust Center → Trust Center Settings → Trusted Add-in Catalogs** (centang *Show in Menu*), mulai ulang Office, lalu
pilih add-in dari **Home → Add-ins → More Add-ins → Shared Folder**. Di Mac dan Office versi web, pasang lewat
**Add-ins → Upload My Add-in**; keduanya mewajibkan HTTPS (`npm run serve:https` untuk server lokal bersertifikat).

### Mencoba tanpa API key

```bash
npm run preview
```

Buka <http://localhost:3000/taskpane.html>, pilih penyedia **Kustom**, isi Base URL
`http://localhost:3000/mock/openai`, muat daftar model, lalu kirim pesan. Jawabannya dari AI tiruan di
`tools/mock-llm.js`, yang juga bisa memanggil alat kalau panel dibuka di dalam Office.

## Menambah penyedia AI

Kalau API-nya OpenAI-compatible, tambahkan satu entri di `PRESETS` pada `docs/providers.js`:

```js
mistral: { label: "Mistral", kind: "openai", baseUrl: "https://api.mistral.ai/v1", model: "", keyUrl: "https://console.mistral.ai/api-keys" },
```

Syaratnya, API tersebut mengizinkan panggilan dari browser (CORS). Untuk format lain, tulis adapter baru di berkas
yang sama: `stream()` yang menghasilkan teks dan panggilan alat, dan `listModels()`.

## Menambah alat

Alat tiap aplikasi ada di `docs/tools/`. Satu alat adalah satu objek:

```js
define("nama_alat", "write", "Penjelasan untuk model.", { address: text("Alamat rentang.") }, ["address"],
  (args) => `Label yang tampil di panel`,
  async (args, context) => { /* Office.js */ context.undo(() => /* cara membatalkan */); return { hasil: "…" }; });
```

`risk` menentukan perlakuannya: `read` langsung jalan, `write` jalan dan bisa dibatalkan, `danger` selalu minta izin.

## Privasi dan keamanan

- API key dan riwayat percakapan disimpan di perangkat ini saja. Di Windows, penyimpanan itu dipakai bersama oleh
  Word, Excel, dan PowerPoint, jadi API key cukup diisi sekali.
- Isi dokumen dikirim ke penyedia AI saat AI membacanya atau saat Anda menyertakannya sebagai konteks.
- Isi dokumen diperlakukan sebagai data, bukan perintah. Tetap berhati-hatilah dengan berkas dari pihak lain:
  teks tersembunyi bisa mencoba mengarahkan AI. Tinjau permintaan izin sebelum menyetujuinya.
- AI hanya bisa menyentuh dokumen yang sedang terbuka lewat alat di `docs/tools/`. Tidak ada alat untuk berkas lain,
  jaringan, atau sistem.

## Yang tidak ada dibanding Claude for Microsoft 365

- Lintas aplikasi di sini berupa alat, bukan percakapan bersama: percakapan tetap tampil hanya di panel tempat Anda
  mengetik, dan panel di aplikasi lain harus terbuka. Fitur ini terbukti di Windows saja.
- Outlook.
- Connectors data pihak ketiga.
- Grafik PowerPoint yang tertaut ke Excel: grafik di sini asli dan bisa diedit, tetapi datanya berdiri sendiri.
- Makro/VBA dan data table Excel.
- Kualitas: add-in resmi disetel khusus untuk model Claude; di sini hasilnya mengikuti model yang Anda pilih.

## Status pengujian

Diuji di Word, Excel, dan PowerPoint asli (Microsoft 365 untuk Windows, build 16.0.20430):

- Semua alat di ketiga aplikasi, termasuk pembatalan di Excel dan PowerPoint, serta terima/tolak tracked changes di
  Word (menolak mengembalikan dokumen persis seperti semula).
- Loop agen dari ujung ke ujung: langkah alat, permintaan izin, penolakan, pembatalan, coba-ulang saat kuota habis,
  riwayat yang pulih setelah panel dimuat ulang, skills, dan instruksi tetap.
- Ketiga format API (OpenAI-compatible, Gemini, Claude) terhadap AI tiruan yang, seperti aslinya, menolak permintaan
  bila giliran model tidak dikirim balik utuh.
- **Gemini sungguhan** (3.5 Flash-Lite dan 3.6 Flash, lewat API key): menulis kerangka di Word, menambah kolom rumus
  dan format bersyarat di Excel, dan dari Word mengambil data Excel lalu menuliskannya sebagai tabel. Tiap tugas
  selesai dalam 4 sampai 7 detik dengan 3 sampai 6 permintaan.
- Lintas aplikasi Word ↔ Excel: baca, tulis, permintaan izin dari jauh (ditolak dan diizinkan), pembatalan dari jauh,
  dan penolakan alat yang tidak bisa dibatalkan.
- `manifest.xml` lolos validator resmi (`npm run validate`).

Belum diuji:

- DeepSeek, Claude, dan penyedia lain dengan API key asli. Adapternya mengikuti dokumentasi terbaru masing-masing
  dan lolos terhadap AI tiruan, tetapi belum pernah dijalankan ke layanan aslinya.
- Tugas panjang dan rumit dengan model sungguhan; yang diuji baru tugas pendek.
- PowerPoint dengan model sungguhan dan dalam lintas aplikasi (alatnya sendiri sudah diuji).
- Mac, Office versi web, dan Office dalam bahasa selain Inggris (nama layout dan gaya bisa berbeda).
- Pemadatan otomatis percakapan panjang.

Yang perlu diketahui dari pengujian:

- Model yang kelebihan beban kadang menerima sambungan lalu diam. Add-in menyerah setelah 90 detik dengan saran
  memilih model lain. Saat diuji, `gemini-flash-latest` (Gemini 3.8 Flash) dan Gemini 3.7 Flash sedang begitu.
- Model kecil kadang memilih langkah yang boros atau menulis butir dengan tanda "•"; alat dan instruksi agen sudah
  disesuaikan untuk itu, tetapi model yang lebih besar tetap memberi hasil lebih rapi.

## Struktur

```
manifest.xml              manifest add-in (https://localhost:3000)
server.js                 server statis untuk docs/
docs/                     add-in itu sendiri; folder inilah yang di-host
  taskpane.html/.css/.js    tampilan, pengaturan, dan alur percakapan
  agent.js                  loop agen, izin, pemadatan
  providers.js              adapter Gemini, OpenAI-compatible, Claude
  tools/excel.js            alat Excel
  tools/word.js             alat Word
  tools/powerpoint.js       alat PowerPoint
  host.js                   baca pilihan dan sisip manual
  markdown.js               pengurai Markdown (HTML untuk panel, paragraf untuk Word)
  store.js                  riwayat percakapan (IndexedDB)
  peers.js                  lintas aplikasi: panel saling menjalankan alat lewat BroadcastChannel
  vendor/pptxgen.bundle.js  PptxGenJS, untuk slide bergrafik
tools/
  start.js                  npm start, dan "Aktifkan AI Assistant.cmd"
  stop.js                   npm stop, dan "Nonaktifkan AI Assistant.cmd"
  publish.js                npm run rilis: periksa rahasia, commit, dan push ke GitHub
  make-manifest.js          membuat manifest.http.xml / manifest.prod.xml
  mock-llm.js               AI tiruan
  dev-bridge*.js, dev-run.js  jembatan uji: menjalankan potongan JS di panel yang terbuka (npm run dev)
```
