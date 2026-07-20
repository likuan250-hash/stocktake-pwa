(function () {
  'use strict';

  async function exportSheet(sheet, lines) {
    await ensureXLSX();
    const aoa = [];
    aoa.push([sheet.title || '盘点单']);
    aoa.push(['盘点日期', (sheet.created_at || '').slice(0, 10)]);
    aoa.push([]);
    aoa.push(['物料编码', '物料名称', '单位', '规格信息', '仓库', '盘点数量', '备注']);
    for (const l of lines) {
      aoa.push([l.code, l.name, l.unit, l.spec, l.warehouse || '', l.qty, l.remark || '']);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 14 }, { wch: 22 }, { wch: 8 }, { wch: 24 }, { wch: 12 }, { wch: 12 }, { wch: 28 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '盘点明细');
    const safe = (sheet.title || 'export').replace(/[\\/:*?"<>|]/g, '_');
    XLSX.writeFile(wb, `盘点单_${safe}.xlsx`);
  }

  window.ExportXLSX = { exportSheet };
})();
