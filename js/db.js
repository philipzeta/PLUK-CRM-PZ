// Simple IndexedDB-backed key/value store. Each tab's whole data object is
// stored as one record, keyed by tab id. Falls back to localStorage if
// IndexedDB is unavailable (e.g. some privacy modes).

const DB_NAME = 'pru-crm';
const DB_VERSION = 1;
const STORE = 'tabs';

let dbPromise = null;

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (!window.indexedDB) { resolve(null); return; }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null); // fall back to localStorage
  });
  return dbPromise;
}

export async function dbGet(id) {
  const db = await openDb();
  if (!db) {
    const raw = localStorage.getItem('crm:' + id);
    return raw ? JSON.parse(raw) : null;
  }
  return new Promise((resolve) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get(id);
    req.onsuccess = () => resolve(req.result ? req.result.data : null);
    req.onerror = () => resolve(null);
  });
}

export async function dbSet(id, data) {
  const db = await openDb();
  if (!db) {
    localStorage.setItem('crm:' + id, JSON.stringify(data));
    return;
  }
  return new Promise((resolve) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put({ id, data, savedAt: Date.now() });
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => resolve(false);
  });
}

export async function dbGetAllIds() {
  const db = await openDb();
  if (!db) {
    return Object.keys(localStorage).filter(k => k.startsWith('crm:')).map(k => k.slice(4));
  }
  return new Promise((resolve) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).getAllKeys();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => resolve([]);
  });
}

export async function dbClearAll() {
  const db = await openDb();
  if (!db) {
    Object.keys(localStorage).filter(k => k.startsWith('crm:')).forEach(k => localStorage.removeItem(k));
    return;
  }
  return new Promise((resolve) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).clear();
    tx.oncomplete = () => resolve(true);
  });
}
