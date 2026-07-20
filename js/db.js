(function () {
  'use strict';
  let sqlDb = null, api = null, saveTimer = null, SQLRef = null, idbConn = null;
  const IDB_NAME = 'stocktake-pwa', IDB_STORE = 'kv', IDB_KEY = 'db-blob-v1';

  function idbOpen() {
    if (idbConn) return Promise.resolve(idbConn);
    return new Promise((res, rej) => {
      const r = indexedDB.open(IDB_NAME, 1);
      r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains(IDB_STORE)) r.result.createObjectStore(IDB_STORE); };
      r.onsuccess = () => { idbConn = r.result; res(idbConn); };
      r.onerror = () => rej(r.error);
    });
  }
  async function loadBlob() {
    try {
      const db = await idbOpen();
      return await new Promise((res, rej) => {
        const tx = db.transaction(IDB_STORE, 'readonly');
        const rq = tx.objectStore(IDB_STORE).get(IDB_KEY);
        rq.onsuccess = () => res(rq.result || null);
        rq.onerror = () => rej(rq.error);
      });
    } catch (e) { return null; }
  }
  async function saveBlob(bytes) {
    const db = await idbOpen();
    return new Promise((res, rej) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(bytes, IDB_KEY);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
  }
  // 防抖落盘：操作时 600ms 后写盘，兼顾性能与实时性
  // （600ms 比默认 400ms 更稳：连续改数量/输入时减少整库 export 次数，大库更顺）
  function scheduleSave() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
      try { if (sqlDb) await saveBlob(sqlDb.export()); }
      catch (e) { console.error('persist failed', e); }
    }, 600);
  }
  // 立即落盘：关页面/切后台时调用，避免防抖窗口内数据丢失
  function flush() {
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
    if (!sqlDb) return;
    try { saveBlob(sqlDb.export()); } catch (e) { console.error('flush failed', e); }
  }
  // 整库备份：把当前数据库序列化为 base64 包进 JSON
  function getBackupString() {
    if (!sqlDb) return null;
    const bytes = sqlDb.export();
    let bin = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    const b64 = (typeof btoa !== 'undefined') ? btoa(bin) : Buffer.from(bin, 'binary').toString('base64');
    return JSON.stringify({ app: 'stocktake-pwa', version: 1, ts: Date.now(), data: b64 });
  }
  // 从备份字符串恢复：替换内存库并立即落盘
  async function importBackupString(str) {
    const obj = JSON.parse(str);
    if (!obj || !obj.data) throw new Error('备份文件格式不正确');
    const bin = (typeof atob !== 'undefined') ? atob(obj.data) : Buffer.from(obj.data, 'base64').toString('binary');
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    sqlDb = new SQLRef.Database(bytes);
    api = StockDBCore.create(sqlDb, scheduleSave);
    api.init();
    flush();
    return true;
  }

  async function init() {
    SQLRef = await initSqlJs({ locateFile: f => 'vendor/' + f });
    const blob = await loadBlob();
    sqlDb = blob ? new SQLRef.Database(new Uint8Array(blob)) : new SQLRef.Database();
    api = StockDBCore.create(sqlDb, scheduleSave);
    api.init();
    scheduleSave();
    // 页面隐藏/关闭时立即落盘（消除 400ms 防抖丢数据风险）
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });
    window.addEventListener('pagehide', flush);
    return api;
  }

  window.AppDB = { init, get api() { return api; }, flush, getBackupString, importBackupString };
})();
