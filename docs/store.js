// Riwayat percakapan, disimpan di IndexedDB perangkat ini (tidak dikirim ke mana pun): satu daftar untuk semua
// berkas dan aplikasi. Satu percakapan = satu catatan. `doc` menunjuk berkas tempat percakapan itu terakhir dipakai,
// supaya panel bisa memulihkannya saat berkas yang sama dibuka lagi; null = tidak menempel ke berkas mana pun.
// Hanya yang terbaru yang disimpan; yang paling lama tidak dibuka dibuang.

const DB = "office-ai-addin";
const STORE = "sessions";
const KEEP = 60;
const TITLE_MAX = 60;

export const newId = () => crypto.randomUUID();

// Judul diambil dari pesan pertama pengguna.
export function titleOf(history) {
  const first = history.find((entry) => entry.role === "user");
  const text = (first ? first.text ?? first.content : "").replace(/\s+/g, " ").trim();
  if (!text) return "Percakapan";
  return text.length > TITLE_MAX ? `${text.slice(0, TITLE_MAX)}…` : text;
}

// Nama berkas dari alamatnya, tanpa kode persen (%20).
export function fileName(url) {
  const name = url.split(/[\\/]/).pop();
  try {
    return decodeURIComponent(name);
  } catch {
    return name;
  }
}

// Versi 1 menyimpan satu percakapan per berkas dalam toko "chats", berkunci "aplikasi|alamat".
function fromV1({ key, updated, history }) {
  const split = key.indexOf("|");
  return {
    id: newId(),
    doc: key,
    host: key.slice(0, split),
    docName: fileName(key.slice(split + 1)),
    title: titleOf(history),
    created: updated,
    updated,
    history,
  };
}

let opened = null;

// Satu sambungan dipakai bersama, jadi permintaan diproses sesuai urutan dipanggil.
function open() {
  if (!opened) {
    opened = new Promise((resolve, reject) => {
      const request = indexedDB.open(DB, 2);
      request.onupgradeneeded = (event) => {
        const db = request.result;
        const sessions = db.createObjectStore(STORE, { keyPath: "id" });
        sessions.createIndex("updated", "updated");
        // [berkas, waktu]: catatan dengan doc null tidak masuk indeks ini
        sessions.createIndex("doc", ["doc", "updated"]);
        if (event.oldVersion < 1) return;
        request.transaction.objectStore("chats").getAll().onsuccess = (found) => {
          for (const saved of found.target.result) sessions.put(fromV1(saved));
          db.deleteObjectStore("chats");
        };
      };
      request.onsuccess = () => {
        const db = request.result;
        // panel lain sedang memasang versi yang lebih baru: lepaskan sambungan supaya tidak menghalanginya
        db.onversionchange = () => {
          db.close();
          opened = null;
        };
        resolve(db);
      };
      request.onerror = () => {
        opened = null;
        reject(request.error);
      };
    });
  }
  return opened;
}

async function transact(mode, work) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE, mode);
    const request = work(transaction.objectStore(STORE));
    transaction.oncomplete = () => resolve(request && request.result);
    transaction.onerror = transaction.onabort = () => reject(transaction.error);
  });
}

export async function loadChat(id) {
  return (await transact("readonly", (store) => store.get(id))) || null;
}

// Percakapan terbaru yang menempel ke berkas ini, atau null.
export async function latestChat(doc) {
  const found = await transact("readonly", (store) => store.index("doc").getAll(IDBKeyRange.bound([doc, 0], [doc, Infinity])));
  return found[found.length - 1] || null;
}

// Daftar untuk tampilan, terbaru dulu, tanpa isi percakapannya.
export async function listChats() {
  const all = await transact("readonly", (store) => store.index("updated").getAll());
  return all.reverse().map(({ history, ...meta }) => meta);
}

// Menyimpan (atau memperbarui) satu percakapan. Judul dan waktu dibuat tetap seperti pertama kali disimpan.
export function saveChat(id, fields) {
  return transact("readwrite", (store) => {
    const request = store.get(id);
    request.onsuccess = () => {
      const now = Date.now();
      const old = request.result;
      store.put({ created: now, ...old, ...fields, id, title: (old && old.title) || fields.title, updated: now });
      const count = store.count();
      count.onsuccess = () => {
        let extra = count.result - KEEP;
        if (extra <= 0) return;
        store.index("updated").openKeyCursor().onsuccess = (event) => {
          const cursor = event.target.result;
          if (!cursor || extra-- <= 0) return;
          store.delete(cursor.primaryKey);
          cursor.continue();
        };
      };
    };
    return request;
  });
}

// Melepas percakapan dari berkas ini, supaya "Chat baru" tidak dipulihkan lagi saat panel dibuka ulang.
export const detachDoc = (doc) =>
  transact("readwrite", (store) => {
    const request = store.index("doc").getAll(IDBKeyRange.bound([doc, 0], [doc, Infinity]));
    request.onsuccess = () => request.result.forEach((chat) => store.put({ ...chat, doc: null }));
    return request;
  });

export const deleteChat = (id) => transact("readwrite", (store) => store.delete(id));

// Menghapus semua percakapan, atau hanya milik satu aplikasi ("Word" | "Excel" | "PowerPoint").
export const clearChats = (host) =>
  transact("readwrite", (store) => {
    if (!host) return store.clear();
    const request = store.openCursor();
    request.onsuccess = () => {
      const cursor = request.result;
      if (!cursor) return;
      if (cursor.value.host === host) cursor.delete();
      cursor.continue();
    };
    return request;
  });
