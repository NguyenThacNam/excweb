/* Lưu trữ cục bộ trong trình duyệt (IndexedDB): bản sao lưu, lịch sử sửa, file mở gần đây. */
(function () {
  'use strict';
  const DB_NAME = 'tra-cuu-thong-so';
  const VERSION = 1;
  let dbPromise = null;

  function open() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const r = indexedDB.open(DB_NAME, VERSION);
        r.onupgradeneeded = () => {
          const db = r.result;
          if (!db.objectStoreNames.contains('backups')) db.createObjectStore('backups', { keyPath: 'id', autoIncrement: true });
          if (!db.objectStoreNames.contains('history')) db.createObjectStore('history', { keyPath: 'id', autoIncrement: true });
          if (!db.objectStoreNames.contains('recent')) db.createObjectStore('recent', { keyPath: 'key' });
        };
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    }
    return dbPromise;
  }
  function req(r) {
    return new Promise((resolve, reject) => {
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  async function store(name, mode) {
    const db = await open();
    return db.transaction(name, mode || 'readonly').objectStore(name);
  }
  async function put(name, obj) { return req((await store(name, 'readwrite')).put(obj)); }
  async function all(name) { return req((await store(name)).getAll()); }
  async function del(name, key) { return req((await store(name, 'readwrite')).delete(key)); }
  async function clear(name) { return req((await store(name, 'readwrite')).clear()); }

  // Giữ tối đa `keep` bản sao lưu cho mỗi tên file
  async function addBackup(fileName, data, keep) {
    await put('backups', { name: fileName, time: Date.now(), size: data.byteLength, data });
    const list = (await all('backups')).filter((b) => b.name === fileName).sort((a, b) => b.time - a.time);
    for (const old of list.slice(keep || 5)) await del('backups', old.id);
  }

  window.WXStore = { available: typeof indexedDB !== 'undefined', put, all, del, clear, addBackup };
})();
