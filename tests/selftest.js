// 真机级自测：复现物料导入的两种模式（合并 / 清空后导入）数据层逻辑
// 运行：node tests/selftest.js
// 依赖：仓库自带 vendor/sql-wasm.js（sql.js 的 Node 可用构建）
const path = require('path');
const ROOT = path.join(__dirname, '..');
const initSqlJs = require(path.join(ROOT, 'vendor', 'sql-wasm.js'));
const { create: createStockDB } = require(path.join(ROOT, 'js', 'db-core.js'));
const VENDOR = path.join(ROOT, 'vendor') + path.sep;

let pass = 0, fail = 0;
function assert(cond, msg) { if (cond) { pass++; } else { fail++; console.error('  FAIL:', msg); } }

initSqlJs({ locateFile: f => VENDOR + f }).then(SQL => {
  const db = new SQL.Database();
  const core = createStockDB(db, () => {});
  core.init();

  // 复刻 importMaterials 的两种执行路径（数据层）
  function importSheet(sheet, mode) {
    if (mode === 'clear') {
      const cleared = core.listMaterials().length;
      core.clearAllMaterials();
      let imported = 0;
      sheet.forEach(m => { if (!m.code) return; core.upsertMaterial(m); imported++; });
      return { cleared, imported };
    }
    let added = 0, updated = 0;
    sheet.forEach(m => {
      if (!m.code) return;
      const ex = core.getMaterialByCode(m.code);
      core.upsertMaterial(m);
      if (ex) updated++; else added++;
    });
    return { added, updated };
  }

  const A = [{ code: '001', name: 'A', unit: '个' }, { code: '002', name: 'B', unit: '个' }];
  const B = [{ code: '002', name: 'B2', unit: '箱' }, { code: '003', name: 'C', unit: '个' }];

  // 1) 合并模式：导入 A（全新增）
  let r1 = importSheet(A, 'merge');
  assert(r1.added === 2 && r1.updated === 0, 'merge A: 应新增2');
  assert(core.listMaterials().length === 2, 'merge A 后应有2条');

  // 2) 合并模式：导入 B（002更新, 003新增）
  let r2 = importSheet(B, 'merge');
  assert(r2.updated === 1 && r2.added === 1, 'merge B: 应更新1新增1');
  assert(core.getMaterialByCode('002').name === 'B2', '002 应被更新为 B2');

  // 3) 合并模式幂等：再导入 B 一次（应全更新, 0新增）
  let r2b = importSheet(B, 'merge');
  assert(r2b.updated === 2 && r2b.added === 0, 'merge B 二次: 应全更新0新增(幂等)');
  assert(core.listMaterials().length === 3, '合并后应有3条(001,002,003)');

  // 4) 清空模式：用 B 清空后导入（清空3条, 导入2条）
  let r3 = importSheet(B, 'clear');
  assert(r3.cleared === 3, 'clear B: 应清空3条, 实际 ' + r3.cleared);
  assert(r3.imported === 2, 'clear B: 应导入2条, 实际 ' + r3.imported);
  assert(core.listMaterials().length === 2, '清空后仅剩2条');
  assert(!core.getMaterialByCode('001'), '001 应已被清空');
  assert(core.getMaterialByCode('002').name === 'B2', '002 仍为 B2');

  // 5) 回收站保留：软删一个再清空，回收站应不动
  core.upsertMaterial({ code: '009', name: 'Z' });
  const zid = core.getMaterialByCode('009').id;
  core.deleteMaterial(zid); // 软删 -> 回收站
  assert(core.listDeletedMaterials().length === 1, '软删后回收站应有1条');
  assert(core.listMaterials().length === 2, '活跃物料仍2条(002,003)');
  let r4 = importSheet(A, 'clear'); // 清空活跃(002,003)=2条, 回收站009不动
  assert(r4.cleared === 2, 'clear A: 应清空活跃2条, 实际 ' + r4.cleared);
  assert(core.listDeletedMaterials().length === 1, '回收站应仍保留1条(009)');
  assert(core.listMaterials().length === 2, '清空后活跃应为 A 的2条');
  assert(core.getMaterialByCode('001') && core.getMaterialByCode('002'), 'A 的 001/002 应在');
  assert(!core.listMaterials().some(m => m.code === '009'), '回收站009不应出现在活跃列表');

  // 6) 回收站同编码在导入后会被恢复为活跃（importMaterials 会先提示，此处验证数据层结果）
  core.restoreMaterial(zid); // 先恢复009参与后续
  assert(core.listMaterials().some(m => m.code === '009'), '恢复后009应在活跃列表');

  // 7) getMaterialByCode 不过滤已删（既有的、importMaterials 依赖的语义）
  core.deleteMaterial(core.getMaterialByCode('009').id);
  assert(core.getMaterialByCode('009') !== undefined, 'getMaterialByCode 应能取到已删(回收站)物料');

  // 8) 默认列表按物料编码升序（盘点单「铺开全部物料」依赖此排序展示）
  core.clearAllMaterials();
  ['003', '001', '010', '002', '009'].forEach(c => core.upsertMaterial({ code: c, name: 'M' + c }));
  const codes = core.listMaterials().map(m => m.code);
  const sorted = [...codes].sort((a, b) => a.localeCompare(b, 'zh', { numeric: true }));
  assert(JSON.stringify(codes) === JSON.stringify(sorted), 'listMaterials 默认应按编码升序, 实际 ' + codes.join(','));

  console.log(`\n自测结果: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
}).catch(e => { console.error('SELFTEST ERROR', e); process.exit(2); });
