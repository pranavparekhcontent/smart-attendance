// ============================================================
//  keystore.js — Multi-layer license key persistence
//  Writes to: localStorage + IndexedDB + Cookie (all 3)
//  Reads from: first layer that has a valid value
//  A browser cache wipe rarely clears all 3 simultaneously.
//  Do NOT edit between projects.
// ============================================================

const KeyStore = (() => {

  const STORE_KEY  = () => APP_CONFIG.LICENSE_STORAGE_KEY;   // key name
  const DB_NAME    = "appstart_db";
  const DB_STORE   = "keystore";
  const COOKIE_DAYS = 365;   // cookie lifetime in days

  // ── Layer 1: Cookie (Primary) ─────────────────────────────
  const CK = {
    save(name, val) {
      try {
        const exp = new Date();
        exp.setDate(exp.getDate() + COOKIE_DAYS);
        const isSecure = window.location.protocol === 'https:';
        document.cookie =
          `${name}=${encodeURIComponent(val)};expires=${exp.toUTCString()};path=/;SameSite=Strict${isSecure ? ';Secure' : ''}`;
        return true;
      } catch { return false; }
    },
    load(name) {
      try {
        const prefix = name + "=";
        for (const part of document.cookie.split(";")) {
          const c = part.trim();
          if (c.startsWith(prefix)) {
            return decodeURIComponent(c.slice(prefix.length)) || null;
          }
        }
        return null;
      } catch { return null; }
    },
    clear(name) {
      try {
        document.cookie =
          `${name}=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/;SameSite=Strict`;
      } catch {}
    },
  };

  // ── Layer 2: IndexedDB (Secondary) ──────────────────────────
  const IDB = {
    _db: null,
    _open() {
      if (this._db) return Promise.resolve(this._db);
      return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = e => e.target.result.createObjectStore(DB_STORE);
        req.onsuccess = e => { this._db = e.target.result; resolve(this._db); };
        req.onerror = () => reject(req.error);
      });
    },
    async save(name, val) {
      try {
        const db = await this._open();
        const tx = db.transaction(DB_STORE, "readwrite");
        tx.objectStore(DB_STORE).put(val, name);
        await new Promise((res, rej) => { tx.oncomplete = res; tx.onerror = rej; });
        return true;
      } catch { return false; }
    },
    async load(name) {
      try {
        const db = await this._open();
        const tx = db.transaction(DB_STORE, "readonly");
        const req = tx.objectStore(DB_STORE).get(name);
        return await new Promise((res, rej) => {
          req.onsuccess = () => res(req.result || null);
          req.onerror = () => rej(req.error);
        });
      } catch { return null; }
    },
    async clear(name) {
      try {
        const db = await this._open();
        const tx = db.transaction(DB_STORE, "readwrite");
        tx.objectStore(DB_STORE).delete(name);
      } catch {}
    }
  };

  // ── Public API ─────────────────────────────────────────────
  async function save(key) {
    CK.save(STORE_KEY(), key);
    await IDB.save(STORE_KEY(), key);
  }

  async function load() {
    const fromCK  = CK.load(STORE_KEY());
    const fromIDB = await IDB.load(STORE_KEY());
    const found   = fromCK || fromIDB || null;
    if (found) {
      if (!fromCK)  CK.save(STORE_KEY(), found);
      if (!fromIDB) IDB.save(STORE_KEY(), found);
    }
    return found;
  }

  async function clear() {
    CK.clear(STORE_KEY());
    await IDB.clear(STORE_KEY());
  }

  // Exposed for internal use (e.g. version caching)
  function setItem(name, val) { CK.save(name, val); IDB.save(name, val); }
  async function getItem(name) { return CK.load(name) || await IDB.load(name); }

  return { save, load, clear, setItem, getItem };
})();
