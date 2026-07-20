(function () {
  'use strict';

  const ALIASES = {
    code: ['物料编码', '编码', '物料代码', '物料号', '货号', 'code'],
    name: ['物料名称', '名称', '品名', '物料', 'name'],
    unit: ['单位', '计量单位', 'unit'],
    spec: ['规格', '规格信息', '规格型号', '型号', 'spec'],
    warehouse: ['仓库', '库房', '仓位', '仓库名称', 'warehouse', 'wh']
  };

  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 8); i++) {
      const joined = (rows[i] || []).map(c => c == null ? '' : String(c)).join('|');
      if (/物料编码|编码|物料名称|名称|货号|code/i.test(joined)) return i;
    }
    return 0;
  }
  function matchHeader(headers) {
    const map = {};
    for (const field in ALIASES) {
      for (const alias of ALIASES[field]) {
        const idx = headers.findIndex(h => h && String(h).trim() !== '' && String(h).includes(alias));
        if (idx >= 0) { map[field] = idx; break; }
      }
    }
    return map;
  }

  async function parseFile(file) {
    await ensureXLSX();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const hIdx = findHeaderRow(rows);
    const headers = (rows[hIdx] || []).map(c => c == null ? '' : String(c).trim());
    const map = matchHeader(headers);
    if (map.code === undefined && map.name === undefined) {
      throw new Error('未识别到表头（需包含「物料编码 / 物料名称」等列）');
    }
    const out = []; let skipped = 0;
    for (let i = hIdx + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const code = map.code !== undefined ? String(r[map.code] ?? '').trim() : '';
      const name = map.name !== undefined ? String(r[map.name] ?? '').trim() : '';
      if (!code && !name) continue;
      if (!code || !name) { skipped++; continue; }
      const unit = map.unit !== undefined ? String(r[map.unit] ?? '').trim() : '';
      const spec = map.spec !== undefined ? String(r[map.spec] ?? '').trim() : '';
      const warehouse = map.warehouse !== undefined ? String(r[map.warehouse] ?? '').trim() : '';
      out.push({ code, name, unit, spec, warehouse });
    }
    return { items: out, skipped };
  }

  async function downloadTemplate() {
    await ensureXLSX();
    const aoa = [
      ['物料编码', '物料名称', '单位', '规格信息', '仓库'],
      ['A001', '示例物料-螺丝', '个', 'M8×20 不锈钢', '原料仓'],
      ['A002', '示例物料-垫圈', '个', 'Φ10 橡胶', '成品仓']
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 14 }, { wch: 22 }, { wch: 8 }, { wch: 24 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '物料档案');
    XLSX.writeFile(wb, '物料档案导入模板.xlsx');
  }

  window.ImportXLSX = { parseFile, downloadTemplate };
})();
