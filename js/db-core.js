(function (global) {
  'use strict';

  const SCHEMA = `
    CREATE TABLE IF NOT EXISTS materials (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      code TEXT NOT NULL,
      name TEXT NOT NULL,
      unit TEXT DEFAULT '',
      spec TEXT DEFAULT '',
      warehouse TEXT DEFAULT '',
      deleted INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now','localtime')),
      updated_at TEXT DEFAULT (datetime('now','localtime')),
      UNIQUE(code)
    );
    CREATE TABLE IF NOT EXISTS count_sheets (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      status TEXT DEFAULT 'open',
      deleted INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now','localtime')),
      updated_at TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE TABLE IF NOT EXISTS count_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sheet_id INTEGER NOT NULL,
      material_id INTEGER,
      code TEXT DEFAULT '',
      name TEXT DEFAULT '',
      unit TEXT DEFAULT '',
      spec TEXT DEFAULT '',
      qty REAL DEFAULT 0,
      seq INTEGER DEFAULT 0,
      warehouse TEXT DEFAULT '',
      remark TEXT DEFAULT '',
      FOREIGN KEY(sheet_id) REFERENCES count_sheets(id) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_lines_sheet ON count_lines(sheet_id);
  `;

  // Factory: wraps a sql.js Database instance with the app's data API.
  // onMutate is called after every mutating operation (used for persistence).
  function createStockDB(db, onMutate) {
    function mut() { if (onMutate) { try { onMutate(); } catch (e) { console.error(e); } } }
    function run(sql, p) { db.run(sql, p || []); }
    function all(sql, p) {
      const st = db.prepare(sql);
      if (p && p.length) st.bind(p);
      const rows = [];
      while (st.step()) rows.push(st.getAsObject());
      st.free();
      return rows;
    }
    function get1(sql, p) {
      const st = db.prepare(sql);
      if (p && p.length) st.bind(p);
      let r = null;
      if (st.step()) r = st.getAsObject();
      st.free();
      return r;
    }
    function lastId() {
      const r = db.exec("SELECT last_insert_rowid() AS id");
      return r.length ? r[0].values[0][0] : null;
    }

    return {
      init() {
        db.run(SCHEMA);
        // 迁移：兼容旧库自动补列（warehouse / deleted），已有数据默认空/未删
        try {
          const mig = (table, col, ddl) => {
            const info = db.exec("PRAGMA table_info(" + table + ")");
            const has = !!(info.length && info[0].values.some(rw => String(rw[1]).toLowerCase() === col));
            if (!has) db.run(ddl);
          };
          mig('materials', 'warehouse', "ALTER TABLE materials ADD COLUMN warehouse TEXT DEFAULT ''");
          mig('materials', 'deleted', "ALTER TABLE materials ADD COLUMN deleted INTEGER DEFAULT 0");
          mig('count_sheets', 'deleted', "ALTER TABLE count_sheets ADD COLUMN deleted INTEGER DEFAULT 0");
          mig('count_lines', 'warehouse', "ALTER TABLE count_lines ADD COLUMN warehouse TEXT DEFAULT ''");
          mig('count_lines', 'remark', "ALTER TABLE count_lines ADD COLUMN remark TEXT DEFAULT ''");
          // 一次性归一化：清理物料仓库字段可能残留的前后空格，保证精确筛选匹配
          try { db.run("UPDATE materials SET warehouse = TRIM(warehouse) WHERE warehouse IS NOT NULL AND TRIM(warehouse) <> warehouse"); } catch (_) {}
        } catch (e) { console.error('migrate failed', e); }
      },

      // ---- materials ----
      // 注意：基础条件必须用括号包裹 (deleted=0 OR deleted IS NULL)，
      // 否则与 AND 拼接后会被解析成 (deleted=0) OR (...)，导致筛选恒真、返回全部。
      listMaterials(term, warehouse) {
        const where = ['(deleted = 0 OR deleted IS NULL)'], params = [];
        if (term) {
          const t = '%' + term + '%';
          where.push('(code LIKE ? OR name LIKE ? OR spec LIKE ? OR warehouse LIKE ?)');
          params.push(t, t, t, t);
        }
        if (warehouse) { where.push('warehouse = ?'); params.push(warehouse); }
        const sql = 'SELECT * FROM materials WHERE ' + where.join(' AND ') + ' ORDER BY code';
        return all(sql, params);
      },
      getMaterial(id) { return get1("SELECT * FROM materials WHERE id=?", [id]); },
      getMaterialByCode(code) { return get1("SELECT * FROM materials WHERE code=?", [code]); },
      addMaterial(m) {
        run("INSERT INTO materials(code,name,unit,spec,warehouse,deleted) VALUES(?,?,?,?,?,0)", [m.code, m.name, m.unit || '', m.spec || '', (m.warehouse || '').trim()]);
        mut(); return lastId();
      },
      updateMaterial(id, m) {
        run("UPDATE materials SET code=?,name=?,unit=?,spec=?,warehouse=?,deleted=0,updated_at=datetime('now') WHERE id=?", [m.code, m.name, m.unit || '', m.spec || '', (m.warehouse || '').trim(), id]);
        mut();
      },
      deleteMaterial(id) { run("UPDATE materials SET deleted=1, updated_at=datetime('now') WHERE id=?", [id]); mut(); },
      upsertMaterial(m) {
        const ex = get1("SELECT id FROM materials WHERE code=?", [m.code]);
        const wh = (m.warehouse || '').trim();
        if (ex) {
          run("UPDATE materials SET name=?,unit=?,spec=?,warehouse=?,deleted=0,updated_at=datetime('now') WHERE id=?", [m.name, m.unit || '', m.spec || '', wh, ex.id]);
          mut(); return ex.id;
        }
        run("INSERT INTO materials(code,name,unit,spec,warehouse,deleted) VALUES(?,?,?,?,?,0)", [m.code, m.name, m.unit || '', m.spec || '', wh]);
        mut(); return lastId();
      },

      // ---- count sheets ----
      listSheets() {
        return all("SELECT s.*, (SELECT COUNT(*) FROM count_lines l WHERE l.sheet_id=s.id) AS line_count FROM count_sheets s WHERE s.deleted = 0 OR s.deleted IS NULL ORDER BY created_at DESC");
      },
      getSheet(id) { return get1("SELECT * FROM count_sheets WHERE id=?", [id]); },
      createSheet(title) { run("INSERT INTO count_sheets(title) VALUES(?)", [title]); mut(); return lastId(); },
      deleteSheet(id) { run("UPDATE count_sheets SET deleted=1 WHERE id=?", [id]); mut(); },

      // ---- count lines ----
      listLines(sheetId) {
        return all("SELECT * FROM count_lines WHERE sheet_id=? ORDER BY seq, id", [sheetId]);
      },
      getLine(id) { return get1("SELECT * FROM count_lines WHERE id=?", [id]); },
      addLine(sheetId, mat) {
        const seq = all("SELECT COALESCE(MAX(seq),0)+1 AS n FROM count_lines WHERE sheet_id=?", [sheetId])[0].n;
        run(
          "INSERT INTO count_lines(sheet_id,material_id,code,name,unit,spec,warehouse,remark,qty,seq) VALUES(?,?,?,?,?,?,?,?,?,?)",
          [sheetId, mat.id || null, mat.code || '', mat.name || '', mat.unit || '', mat.spec || '', mat.warehouse || '', mat.remark || '', 0, seq]
        );
        mut(); return lastId();
      },
      updateLineQty(lineId, qty) { run("UPDATE count_lines SET qty=MAX(0,?) WHERE id=?", [qty, lineId]); mut(); },
      incLine(lineId, delta) { run("UPDATE count_lines SET qty = MAX(0, qty + ?) WHERE id=?", [delta, lineId]); mut(); },
      removeLine(lineId) { run("DELETE FROM count_lines WHERE id=?", [lineId]); mut(); },
      updateLineRemark(lineId, remark) { run("UPDATE count_lines SET remark=? WHERE id=?", [remark || '', lineId]); mut(); },

      // ---- 回收站 ----
      listDeletedMaterials() {
        return all("SELECT * FROM materials WHERE deleted=1 ORDER BY code");
      },
      listDeletedSheets() {
        return all("SELECT s.*, (SELECT COUNT(*) FROM count_lines l WHERE l.sheet_id=s.id) AS line_count FROM count_sheets s WHERE s.deleted=1 ORDER BY created_at DESC");
      },
      restoreMaterial(id) { run("UPDATE materials SET deleted=0, updated_at=datetime('now') WHERE id=?", [id]); mut(); },
      restoreSheet(id) { run("UPDATE count_sheets SET deleted=0 WHERE id=?", [id]); mut(); },
      // 彻底删除：解除被引用行的 material_id 悬空，再删物料
      purgeMaterial(id) {
        run("UPDATE count_lines SET material_id=NULL WHERE material_id=?", [id]);
        run("DELETE FROM materials WHERE id=?", [id]); mut();
      },
      purgeSheet(id) {
        run("DELETE FROM count_lines WHERE sheet_id=?", [id]);
        run("DELETE FROM count_sheets WHERE id=?", [id]); mut();
      },
      emptyTrashMaterials() {
        // 先解绑引用本仓已删物料的明细行，避免 material_id 悬空
        run("UPDATE count_lines SET material_id=NULL WHERE material_id IN (SELECT id FROM materials WHERE deleted=1)");
        run("DELETE FROM materials WHERE deleted=1"); mut();
      },
      emptyTrashSheets() {
        run("DELETE FROM count_lines WHERE sheet_id IN (SELECT id FROM count_sheets WHERE deleted=1)");
        run("DELETE FROM count_sheets WHERE deleted=1"); mut();
      }
    };
  }

  const core = { create: createStockDB, SCHEMA };
  if (typeof module !== 'undefined' && module.exports) module.exports = core;
  global.StockDBCore = core;
})(typeof window !== 'undefined' ? window : globalThis);
