// =====================================================
// IndexedDB File Store
// =====================================================
// Pure-frontend replacement for the v2.1.0 server-side session.
// Files are stored as ArrayBuffer keyed by fileId; the fileId
// matches the in-memory `S.files[i].id` so the file list can be
// restored after a page reload.

const DB_NAME = 'ticketchan';
const DB_VERSION = 1;
const STORE = 'files';

function openDb() {
  return new Promise(function (resolve, reject) {
    var req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = function () {
      var db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = function () { resolve(req.result); };
    req.onerror = function () { reject(req.error); };
  });
}

function tx(db, mode) {
  return db.transaction(STORE, mode).objectStore(STORE);
}

export async function putFile(id, name, mime, buffer) {
  var db = await openDb();
  return new Promise(function (resolve, reject) {
    var req = tx(db, 'readwrite').put({
      id: id,
      name: name,
      mime: mime || 'application/octet-stream',
      buffer: buffer,
      savedAt: Date.now()
    });
    req.onsuccess = function () { resolve(id); db.close(); };
    req.onerror = function () { reject(req.error); db.close(); };
  });
}

export async function getFile(id) {
  var db = await openDb();
  return new Promise(function (resolve, reject) {
    var req = tx(db, 'readonly').get(id);
    req.onsuccess = function () {
      resolve(req.result || null);
      db.close();
    };
    req.onerror = function () { reject(req.error); db.close(); };
  });
}

export async function listFiles() {
  var db = await openDb();
  return new Promise(function (resolve, reject) {
    var req = tx(db, 'readonly').getAll();
    req.onsuccess = function () {
      resolve(req.result || []);
      db.close();
    };
    req.onerror = function () { reject(req.error); db.close(); };
  });
}

export async function deleteFile(id) {
  var db = await openDb();
  return new Promise(function (resolve, reject) {
    var req = tx(db, 'readwrite').delete(id);
    req.onsuccess = function () { resolve(); db.close(); };
    req.onerror = function () { reject(req.error); db.close(); };
  });
}

export async function clearAll() {
  var db = await openDb();
  return new Promise(function (resolve, reject) {
    var req = tx(db, 'readwrite').clear();
    req.onsuccess = function () { resolve(); db.close(); };
    req.onerror = function () { reject(req.error); db.close(); };
  });
}

window.__idb = { putFile: putFile, getFile: getFile, listFiles: listFiles, deleteFile: deleteFile, clearAll: clearAll };
