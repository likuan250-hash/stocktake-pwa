(function () {
  'use strict';

  const view = document.getElementById('view');
  const topTitle = document.getElementById('topTitle');
  const themeBtn = document.getElementById('themeBtn');
  const modal = document.getElementById('modal');
  const fileInput = document.getElementById('fileInput');
  const backupInput = document.getElementById('backupInput');
  const db = () => window.AppDB.api;
  let currentSheetId = null;
  // 跨函数共享的「已勾选」状态：候选物料与明细行。必须位于 II FE 顶层作用域，
  // 否则 searchMaterialsToAdd / addSelectedLines / loadLines 等外层函数引用时会 ReferenceError。
  const selectedMatIds = new Set();
  const selectedLineIds = new Set();
  // 物料档案「批量删除」选择模式状态（独立于候选列表的 selectedMatIds，避免切视图冲突）
  let archiveSelectMode = false;
  const archiveSelIds = new Set();
  // 盘点单列表「批量删除」选择模式状态（独立于详情页的 selectedLineIds 等）
  let sheetSelectMode = false;
  const sheetSelIds = new Set();
  // 盘点单详情「已录入物料」多选：是否处于选择模式（勾选框默认隐藏，点「选择」才显示）
  let detailSelectMode = false;

  // xlsx（932KB）体积大，改为「按需懒加载」：首次用到导入/导出时才注入脚本，
  // 不再阻塞首屏启动。SW 已在 install 阶段预缓存该文件，故离线也可用。
  let _xlsxPromise = null;
  function ensureXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (_xlsxPromise) return _xlsxPromise;
    _xlsxPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'vendor/xlsx.full.min.js';
      s.onload = () => resolve(window.XLSX);
      s.onerror = () => reject(new Error('xlsx 组件加载失败'));
      document.body.appendChild(s);
    });
    return _xlsxPromise;
  }
  window.ensureXLSX = ensureXLSX;

  // 防抖：用于候选搜索输入，避免每个字符重建大列表
  function debounce(fn, ms) {
    let t = null;
    const d = function () { clearTimeout(t); t = setTimeout(() => fn(), ms); };
    d.cancel = () => clearTimeout(t);
    return d;
  }

  // ---------- utils ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  // 自定义 Toast（底部提示条，替代原生 alert）
  let _toastEl = null;
  function toast(msg) {
    if (!_toastEl) {
      _toastEl = document.createElement('div');
      _toastEl.id = 'toastContainer';
      document.body.appendChild(_toastEl);
    }
    const bar = document.createElement('div');
    bar.className = 'toast-bar';
    bar.textContent = msg;
    _toastEl.appendChild(bar);
    setTimeout(() => { if (bar.parentNode) bar.remove(); }, 2400);
  }

  // 自定义确认弹窗（大按钮易触控，替代原生 confirm）
  // 返回 Promise<boolean>，调用方用 await
  function askConfirm(msg, opts) {
    opts = opts || {};
    return new Promise(resolve => {
      const ov = document.createElement('div');
      ov.className = 'dialog-overlay';
      const okClass = opts.danger ? 'danger' : 'ok';
      const okText = opts.okText || (opts.danger ? '删除' : '确定');
      ov.innerHTML =
        '<div class="dialog-box">' +
          '<div class="dialog-msg">' + esc(msg) + '</div>' +
          '<div class="dialog-btns">' +
            '<button class="dialog-btn cancel">取消</button>' +
            '<button class="dialog-btn ' + okClass + '">' + okText + '</button>' +
          '</div>' +
        '</div>';
      ov.querySelector('.cancel').onclick = () => { ov.remove(); resolve(false); };
      ov.querySelector('.' + okClass).onclick = () => { ov.remove(); resolve(true); };
      // 点遮罩也关闭（算取消）
      ov.addEventListener('click', e => { if (e.target === ov) { ov.remove(); resolve(false); } });
      document.body.appendChild(ov);
    });
  }
  // 数量框算式求值：仅允许数字与 + - * / . ( )，安全求值，返回数字或 NaN
  function evalQtyExpr(raw) {
    const s = String(raw == null ? '' : raw).trim();
    if (s === '') return 0;
    if (/^-?\d*\.?\d+$/.test(s)) return parseFloat(s);
    if (!/^[0-9+\-*/().\s]+$/.test(s)) return NaN;
    try {
      const v = Function('"use strict";return (' + s + ')')();
      if (typeof v !== 'number' || !isFinite(v)) return NaN;
      return Math.round(v * 1e6) / 1e6;
    } catch (e) { return NaN; }
  }

  // ---------- theme ----------
  const THEMES = ['light', 'dark', 'system'];
  function applyTheme(t) {
    document.documentElement.removeAttribute('data-theme');
    if (t !== 'system') document.documentElement.setAttribute('data-theme', t);
    localStorage.setItem('theme', t);
    themeBtn.textContent = t === 'light' ? '☀️' : t === 'dark' ? '🌙' : '🌗';
  }
  themeBtn.addEventListener('click', () => {
    const cur = localStorage.getItem('theme') || 'system';
    const nx = THEMES[(THEMES.indexOf(cur) + 1) % 3];
    applyTheme(nx);
  });
  const dataBtn = document.getElementById('dataBtn');
  if (dataBtn) dataBtn.addEventListener('click', openDataPanel);

  // ---------- modal ----------
  function openModal(html) {
    modal.innerHTML = `<div class="modal-card">${html}</div>`;
    modal.classList.remove('hidden');
  }
  function closeModal() { modal.classList.add('hidden'); modal.innerHTML = ''; }
  modal.addEventListener('click', e => { if (e.target === modal) closeModal(); });

  // ============ 数据备份 / 恢复 ============
  function openDataPanel() {
    openModal(`
      <h3>数据备份与恢复</h3>
      <p class="muted" style="font-size:13px;margin:0 0 14px;line-height:1.6;">数据只保存在本机浏览器。建议每月导出一份备份；换手机或清缓存时，用备份文件即可完整恢复。</p>
      <div class="modal-actions" style="flex-direction:column;gap:10px;">
        <button class="btn primary" id="bk_export">导出备份（.json）</button>
        <button class="btn ghost" id="bk_import">从备份文件恢复…</button>
      </div>
      <div class="modal-actions">
        <button class="btn-close" id="bk_close">✕</button>
      </div>`);
    modal.querySelector('#bk_close').onclick = closeModal;
    modal.querySelector('#bk_export').onclick = downloadBackup;
    modal.querySelector('#bk_import').onclick = () => { closeModal(); if (backupInput) backupInput.click(); };
  }
  function downloadBackup() {
    const str = window.AppDB.getBackupString ? window.AppDB.getBackupString() : null;
    if (!str) { toast('暂无数据可备份'); return; }
    const blob = new Blob([str], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const d = new Date(), pad = n => String(n).padStart(2, '0');
    const name = '物料盘点_备份_' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '.json';
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('已导出数据库备份文件');
  }

  // ---------- navigation ----------
  function setActiveTab(v) {
    document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.view === v));
  }
  function goMaterials() { setActiveTab('materials'); renderMaterials(); }
  function goSheets() { setActiveTab('sheets'); renderSheets(); }
  function goDetail(id) { setActiveTab('sheets'); renderSheetDetail(id); }

  document.querySelectorAll('.tab').forEach(t => {
    t.addEventListener('click', () => {
      const v = t.dataset.view;
      if (v === 'materials') goMaterials();
      else if (v === 'sheets') goSheets();
      else if (v === 'trash') goTrash();
    });
  });

  // ============ 物料档案 ============
  function renderMaterials() {
    topTitle.textContent = '物料档案';
    view.innerHTML = `
      <div class="view-fixed">
        <div class="toolbar">
          <input id="mSearch" class="search" placeholder="搜索编码 / 名称 / 规格 / 仓库…">
          <div class="toolbtns">
            <button id="btnTpl" class="btn ghost">模板</button>
            <button id="btnImp" class="btn ghost">导入</button>
            <button id="btnAdd" class="btn primary">+ 新增</button>
          </div>
        </div>
        <div class="toolbar sub">
          <select id="mWh" class="search"><option value="">全部仓库</option></select>
          <button id="btnSelMode" class="btn ghost sm">选择</button>
        </div>
        <div class="count-bar"><span id="mCount" class="count"></span></div>
        <div id="archiveBulk" class="bulk-head hidden">
          <label class="bh-selall"><input type="checkbox" id="archiveSelAll"> 全选当前</label>
          <span class="sel-count" id="archiveCount">已选 0 项</span>
        </div>
      </div>
      <div id="mList" class="list view-scroll"></div>
      <button id="archiveBar" class="bulk-add batch-bar-fixed hidden">
        <span class="bd-count" id="archiveBarCount">已选 0 项</span>
        <span class="bd-action">批量删除</span>
      </button>`;
    const search = view.querySelector('#mSearch');
    const whSel = view.querySelector('#mWh');
    fillWarehouseOptions(whSel, '');
    search.addEventListener('input', () => loadMList(search.value.trim(), whSel.value));
    whSel.addEventListener('change', () => loadMList(search.value.trim(), whSel.value));
    view.querySelector('#btnTpl').onclick = () => ImportXLSX.downloadTemplate().catch(() => toast('模板生成失败'));
    view.querySelector('#btnImp').onclick = () => fileInput.click();
    view.querySelector('#btnAdd').onclick = () => openMaterialForm(null);
    view.querySelector('#btnSelMode').onclick = () => toggleArchiveSelect();
    const archiveSelAll = view.querySelector('#archiveSelAll');
    archiveSelAll.addEventListener('change', () => {
      view.querySelectorAll('#mList .card .line-sel').forEach(cb => {
        const mid = +cb.dataset.id;
        if (archiveSelAll.checked) archiveSelIds.add(mid); else archiveSelIds.delete(mid);
        cb.checked = archiveSelAll.checked;
        cb.closest('.card').classList.toggle('selected', archiveSelAll.checked);
      });
      updateArchiveBulk();
    });
    view.querySelector('#archiveBar').onclick = async () => {
      const ids = [...archiveSelIds];
      if (!ids.length) return;
      if (!(await askConfirm('确定批量删除选中的 ' + ids.length + ' 个物料？将移入回收站，可在回收站还原', { danger: true }))) return;
      ids.forEach(id => db().deleteMaterial(id));
      archiveSelIds.clear();
      archiveSelectMode = false;
      const wh = view.querySelector('#mWh');
      if (wh) fillWarehouseOptions(wh, wh.value);
      loadMList(view.querySelector('#mSearch').value.trim(), wh ? wh.value : '');
      updateArchiveBulk();
      toast('已移入回收站 ' + ids.length + ' 个');
    };
    loadMList('', '');
  }

  function fillWarehouseOptions(sel, current) {
    const whs = [...new Set(db().listMaterials().map(m => (m.warehouse || '').trim()).filter(Boolean))].sort();
    sel.innerHTML = '<option value="">全部仓库</option>' + whs.map(w => `<option value="${esc(w)}">${esc(w)}</option>`).join('');
    if (current) sel.value = current;
  }

  function loadMList(term, warehouse) {
    const rows = db().listMaterials(term, warehouse || '');
    const cnt = view.querySelector('#mCount');
    if (cnt) cnt.textContent = warehouse ? ('共 ' + rows.length + ' 个物料（已按仓库筛选）') : ('共 ' + db().listMaterials().length + ' 个物料');
    const el = view.querySelector('#mList');
    if (!rows.length) { el.innerHTML = '<div class="empty">暂无物料，点“新增”或“导入”</div>'; return; }
    el.innerHTML = rows.map(m => {
      if (archiveSelectMode) {
        const on = archiveSelIds.has(m.id);
        return `
      <div class="card row${on ? ' selected' : ''}" data-id="${m.id}">
        <input type="checkbox" class="line-sel mat-sel" data-id="${m.id}" ${on ? 'checked' : ''}>
        <div class="row-main">
          <div class="row-title">${esc(m.name)} <span class="muted">${esc(m.code)}</span></div>
          <div class="row-sub">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}</div>
        </div>
      </div>`;
      }
      return `
      <div class="card row" data-id="${m.id}">
        <div class="row-main">
          <div class="row-title">${esc(m.name)} <span class="muted">${esc(m.code)}</span></div>
          <div class="row-sub">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}</div>
        </div>
        <div class="row-actions">
          <button class="icon-btn" data-act="edit">✎</button>
          <button class="icon-btn danger" data-act="del">🗑</button>
        </div>
      </div>`;
    }).join('');
    el.querySelectorAll('.card.row').forEach(c => {
      const id = +c.dataset.id;
      if (archiveSelectMode) {
        const cb = c.querySelector('.line-sel');
        // 直接点复选框：原生切换 + change 更新集合（卡片点击此处早退，避免双触发）
        cb.addEventListener('change', () => {
          if (cb.checked) archiveSelIds.add(id); else archiveSelIds.delete(id);
          c.classList.toggle('selected', cb.checked);
          updateArchiveBulk();
        });
        // 点卡片其它区域：手动切换复选框（程序化改 checked 不触发 change），更新集合
        c.addEventListener('click', (e) => {
          if (e.target === cb) return;
          cb.checked = !cb.checked;
          if (cb.checked) archiveSelIds.add(id); else archiveSelIds.delete(id);
          c.classList.toggle('selected', cb.checked);
          updateArchiveBulk();
        });
      } else {
        c.querySelector('[data-act=edit]').onclick = () => openMaterialForm(id);
        c.querySelector('[data-act=del]').onclick = async () => {
          if (!(await askConfirm('删除该物料？将移入回收站，可在回收站还原'))) return;
          db().deleteMaterial(id);
          const wh = view.querySelector('#mWh');
          if (wh) fillWarehouseOptions(wh, wh.value);
          loadMList(view.querySelector('#mSearch').value.trim(), wh ? wh.value : '');
          toast('已移入回收站');
        };
      }
    });
    updateArchiveBulk();
  }

  // 切换物料档案「选择模式」
  function toggleArchiveSelect() {
    archiveSelectMode = !archiveSelectMode;
    if (!archiveSelectMode) archiveSelIds.clear();
    loadMList(view.querySelector('#mSearch').value.trim(), (view.querySelector('#mWh') || {}).value || '');
    updateArchiveBulk();
  }

  // 同步选择模式的头部（全选当前）与底部批量栏（计数、显隐、按钮文案）
  function updateArchiveBulk() {
    const bulk = view.querySelector('#archiveBulk');
    const bar = view.querySelector('#archiveBar');
    const btn = view.querySelector('#btnSelMode');
    if (btn) btn.textContent = archiveSelectMode ? '取消' : '选择';
    if (!bulk || !bar) return;
    bulk.classList.toggle('hidden', !archiveSelectMode);
    bar.classList.toggle('hidden', !archiveSelectMode);
    const n = archiveSelIds.size;
    const t = '已选 ' + n + ' 项';
    if (view.querySelector('#archiveCount')) view.querySelector('#archiveCount').textContent = t;
    if (view.querySelector('#archiveBarCount')) view.querySelector('#archiveBarCount').textContent = t;
    const cbs = [...view.querySelectorAll('#mList .card .line-sel')];
    const all = cbs.length > 0 && cbs.every(cb => cb.checked);
    const some = cbs.some(cb => cb.checked);
    const selAll = view.querySelector('#archiveSelAll');
    if (selAll) { selAll.checked = all; selAll.indeterminate = !all && some; }
  }

  function openMaterialForm(id) {
    const m = id ? db().getMaterial(id) : null;
    openModal(`
      <h3>${id ? '编辑' : '新增'}物料<button class="btn-close" id="m_close">✕</button></h3>
      <label>物料编码 *<input id="f_code" value="${esc(m ? m.code : '')}"></label>
      <label>物料名称 *<input id="f_name" value="${esc(m ? m.name : '')}"></label>
      <label>单位<input id="f_unit" value="${esc(m ? m.unit : '')}"></label>
      <label>规格信息<input id="f_spec" value="${esc(m ? m.spec : '')}"></label>
      <label>仓库<input id="f_wh" value="${esc(m ? m.warehouse : '')}" placeholder="如：原料仓 / 成品仓"></label>
      <div class="modal-actions">
        <button class="btn ghost" id="m_cancel">取消</button>
        <button class="btn primary" id="m_save">保存</button>
      </div>`);
    modal.querySelector('#m_cancel').onclick = closeModal;
    const mCloseBtn = modal.querySelector('#m_close');
    if (mCloseBtn) mCloseBtn.onclick = closeModal;
    modal.querySelector('#m_save').onclick = async () => {
      const code = modal.querySelector('#f_code').value.trim();
      const name = modal.querySelector('#f_name').value.trim();
      if (!code || !name) { toast('编码和名称必填'); return; }
      const data = {
        code, name,
        unit: modal.querySelector('#f_unit').value.trim(),
        spec: modal.querySelector('#f_spec').value.trim(),
        warehouse: modal.querySelector('#f_wh').value.trim()
      };
      // 先查重，再决定新增还是更新（避免重复调用）
      const dup = db().getMaterialByCode(code);
      if (dup && dup.id !== id) {
        if (!(await askConfirm('物料编码「' + code + '」已存在，是否覆盖更新已有物料？'))) return;
        db().updateMaterial(dup.id, data);
        toast('已更新已有物料');
      } else if (id) {
        db().updateMaterial(id, data);
      } else {
        db().addMaterial(data);
      }
      closeModal();
      const wh = view.querySelector('#mWh');
      if (wh) fillWarehouseOptions(wh, wh.value);
      loadMList(view.querySelector('#mSearch').value.trim(), wh ? wh.value : '');
    };
  }

  // ============ 盘点单 ============
  function renderSheets() {
    topTitle.textContent = '盘点单';
    view.innerHTML = `
      <div class="view-fixed">
        <div class="toolbar">
          <button id="btnSelModeSheets" class="btn ghost">选择</button>
          <div class="toolbtns right">
            <button id="btnNew" class="btn primary">+ 新建盘点单</button>
          </div>
        </div>
        <div id="sheetBulk" class="bulk-head hidden">
          <label class="bh-selall"><input type="checkbox" id="sheetSelAll"> 全选当前</label>
          <span class="sel-count" id="sheetCount">已选 0 项</span>
        </div>
      </div>
      <div id="sList" class="list view-scroll"></div>
      <button id="sheetBar" class="bulk-add batch-bar-fixed hidden">
        <span class="bd-count" id="sheetBarCount">已选 0 项</span>
        <span class="bd-action">批量删除</span>
      </button>`;
    view.querySelector('#btnNew').onclick = openNewSheet;
    view.querySelector('#btnSelModeSheets').onclick = () => toggleSheetSelect();
    const sheetSelAll = view.querySelector('#sheetSelAll');
    sheetSelAll.addEventListener('change', () => {
      view.querySelectorAll('#sList .card .line-sel').forEach(cb => {
        const sid = +cb.dataset.id;
        if (sheetSelAll.checked) sheetSelIds.add(sid); else sheetSelIds.delete(sid);
        cb.checked = sheetSelAll.checked;
        cb.closest('.card').classList.toggle('selected', sheetSelAll.checked);
      });
      updateSheetBulk();
    });
    view.querySelector('#sheetBar').onclick = async () => {
      const ids = [...sheetSelIds];
      if (!ids.length) return;
      if (!(await askConfirm('确定批量删除选中的 ' + ids.length + ' 个盘点单？将移入回收站，可在回收站还原', { danger: true }))) return;
      ids.forEach(id => db().deleteSheet(id));
      sheetSelIds.clear();
      sheetSelectMode = false;
      loadSheets();
      updateSheetBulk();
      toast('已移入回收站 ' + ids.length + ' 个');
    };
    loadSheets();
  }

  function loadSheets() {
    const rows = db().listSheets();
    const el = view.querySelector('#sList');
    if (!rows.length) { el.innerHTML = '<div class="empty">还没有盘点单，点“新建盘点单”开始</div>'; return; }
    el.innerHTML = rows.map(s => {
      if (sheetSelectMode) {
        const on = sheetSelIds.has(s.id);
        return `
      <div class="card row sheet${on ? ' selected' : ''}" data-id="${s.id}">
        <input type="checkbox" class="line-sel mat-sel" data-id="${s.id}" ${on ? 'checked' : ''}>
        <div class="row-main">
          <div class="row-title">${esc(s.title)}</div>
          <div class="row-sub">${esc((s.created_at || '').slice(0, 10))} · ${s.line_count} 项</div>
        </div>
      </div>`;
      }
      return `
      <div class="card row sheet" data-id="${s.id}">
        <div class="row-main">
          <div class="row-title">${esc(s.title)}</div>
          <div class="row-sub">${esc((s.created_at || '').slice(0, 10))} · ${s.line_count} 项</div>
        </div>
        <div class="row-actions">
          <button class="icon-btn danger" data-act="del">🗑</button>
        </div>
      </div>`;
    }).join('');
    el.querySelectorAll('.card.sheet').forEach(c => {
      const id = +c.dataset.id;
      if (sheetSelectMode) {
        const cb = c.querySelector('.line-sel');
        cb.addEventListener('change', () => {
          if (cb.checked) sheetSelIds.add(id); else sheetSelIds.delete(id);
          c.classList.toggle('selected', cb.checked);
          updateSheetBulk();
        });
        c.addEventListener('click', (e) => {
          if (e.target === cb) return;
          cb.checked = !cb.checked;
          if (cb.checked) sheetSelIds.add(id); else sheetSelIds.delete(id);
          c.classList.toggle('selected', cb.checked);
          updateSheetBulk();
        });
      } else {
        c.querySelector('.row-main').onclick = () => goDetail(id);
        c.querySelector('[data-act=del]').onclick = async (e) => {
          e.stopPropagation();
          if (!(await askConfirm('删除该盘点单？将移入回收站，可在回收站还原'))) return;
          db().deleteSheet(id); loadSheets(); toast('已移入回收站');
        };
      }
    });
    updateSheetBulk();
  }

  // 切换盘点单列表「选择模式」
  function toggleSheetSelect() {
    sheetSelectMode = !sheetSelectMode;
    if (!sheetSelectMode) sheetSelIds.clear();
    loadSheets();
    updateSheetBulk();
  }

  // 同步选择模式的头部（全选当前）与底部批量栏（计数、显隐、按钮文案）
  function updateSheetBulk() {
    const bulk = view.querySelector('#sheetBulk');
    const bar = view.querySelector('#sheetBar');
    const btn = view.querySelector('#btnSelModeSheets');
    if (btn) btn.textContent = sheetSelectMode ? '取消' : '选择';
    if (!bulk || !bar) return;
    bulk.classList.toggle('hidden', !sheetSelectMode);
    bar.classList.toggle('hidden', !sheetSelectMode);
    const n = sheetSelIds.size;
    const t = '已选 ' + n + ' 项';
    if (view.querySelector('#sheetCount')) view.querySelector('#sheetCount').textContent = t;
    if (view.querySelector('#sheetBarCount')) view.querySelector('#sheetBarCount').textContent = t;
    const cbs = [...view.querySelectorAll('#sList .card .line-sel')];
    const all = cbs.length > 0 && cbs.every(cb => cb.checked);
    const some = cbs.some(cb => cb.checked);
    const selAll = view.querySelector('#sheetSelAll');
    if (selAll) { selAll.checked = all; selAll.indeterminate = !all && some; }
  }

  function openNewSheet() {
    const def = '月末盘点 ' + new Date().toISOString().slice(0, 7);
    openModal(`
      <h3>新建盘点单<button class="btn-close" id="ns_close">✕</button></h3>
      <label>名称<input id="s_title" value="${esc(def)}"></label>
      <div class="modal-actions">
        <button class="btn ghost" id="s_cancel">取消</button>
        <button class="btn primary" id="s_ok">创建</button>
      </div>`);
    modal.querySelector('#s_cancel').onclick = closeModal;
    const nsClose = modal.querySelector('#ns_close');
    if (nsClose) nsClose.onclick = closeModal;
    modal.querySelector('#s_ok').onclick = () => {
      const t = modal.querySelector('#s_title').value.trim() || ('盘点单' + Date.now());
      const id = db().createSheet(t);
      closeModal();
      goDetail(id);
    };
  }

  // ============ 回收站 ============
  let trashType = 'materials';
  function goTrash() { setActiveTab('trash'); renderTrash(); }

  function renderTrash() {
    topTitle.textContent = '回收站';
    view.innerHTML = `
      <div class="view-fixed">
        <div class="seg">
          <button data-type="materials" class="${trashType === 'materials' ? 'active' : ''}">回收的物料</button>
          <button data-type="sheets" class="${trashType === 'sheets' ? 'active' : ''}">回收的盘点单</button>
        </div>
        <div class="toolbar" style="margin-top:10px;">
          <button id="btnEmpty" class="btn danger">清空回收站</button>
        </div>
      </div>
      <div id="trashList" class="list view-scroll"></div>`;
    view.querySelectorAll('.seg button').forEach(b => {
      b.onclick = () => { trashType = b.dataset.type; renderTrash(); };
    });
    view.querySelector('#btnEmpty').onclick = async () => {
      const label = trashType === 'materials' ? '回收的物料' : '回收的盘点单';
      if (!(await askConfirm('确定清空' + label + '？所有项目将彻底删除且不可恢复', { danger: true }))) return;
      if (trashType === 'materials') db().emptyTrashMaterials(); else db().emptyTrashSheets();
      renderTrash();
      toast('回收站已清空');
    };
    loadTrash();
  }

  function loadTrash() {
    const el = view.querySelector('#trashList');
    if (trashType === 'materials') {
      const rows = db().listDeletedMaterials();
      if (!rows.length) { el.innerHTML = '<div class="empty">回收站里没有物料</div>'; return; }
      el.innerHTML = rows.map(m => `
        <div class="card row" data-id="${m.id}">
          <div class="row-main">
            <div class="row-title">${esc(m.name)} <span class="muted">${esc(m.code)}</span></div>
            <div class="row-sub">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}</div>
          </div>
          <div class="row-actions">
            <button class="btn ghost sm" data-act="restore">还原</button>
            <button class="btn danger sm" data-act="purge">删除</button>
          </div>
        </div>`).join('');
      el.querySelectorAll('.card.row').forEach(c => {
        const id = +c.dataset.id;
        c.querySelector('[data-act=restore]').onclick = () => { db().restoreMaterial(id); renderTrash(); toast('已还原到物料档案'); };
        c.querySelector('[data-act=purge]').onclick = async () => {
          if (!(await askConfirm('彻底删除该物料？此操作不可恢复', { danger: true }))) return;
          db().purgeMaterial(id); renderTrash(); toast('已彻底删除');
        };
      });
    } else {
      const rows = db().listDeletedSheets();
      if (!rows.length) { el.innerHTML = '<div class="empty">回收站里没有盘点单</div>'; return; }
      el.innerHTML = rows.map(s => `
        <div class="card row" data-id="${s.id}">
          <div class="row-main">
            <div class="row-title">${esc(s.title)}</div>
            <div class="row-sub">${esc((s.created_at || '').slice(0, 10))} · ${s.line_count} 项</div>
          </div>
          <div class="row-actions">
            <button class="btn ghost sm" data-act="restore">还原</button>
            <button class="btn danger sm" data-act="purge">删除</button>
          </div>
        </div>`).join('');
      el.querySelectorAll('.card.row').forEach(c => {
        const id = +c.dataset.id;
        c.querySelector('[data-act=restore]').onclick = () => { db().restoreSheet(id); renderTrash(); toast('已还原到盘点单'); };
        c.querySelector('[data-act=purge]').onclick = async () => {
          if (!(await askConfirm('彻底删除该盘点单及其所有明细？此操作不可恢复', { danger: true }))) return;
          db().purgeSheet(id); renderTrash(); toast('已彻底删除');
        };
      });
    }
  }

  // ============ 盘点单明细 ============
  function renderSheetDetail(id) {
    const sheet = db().getSheet(id);
    if (!sheet) { goSheets(); return; }
    currentSheetId = id;
    // 多选状态集合（声明在 II FE 顶层作用域）：打开单据时重置勾选。
    selectedMatIds.clear();
    selectedLineIds.clear();
    topTitle.textContent = sheet.title;
    view.innerHTML = `
      <div class="view-fixed">
        <div class="detail-head">
          <button id="back" class="icon-btn">←</button>
          <div class="head-info">
            <div class="row-title">${esc(sheet.title)}</div>
            <div class="row-sub">${esc((sheet.created_at || '').slice(0, 10))}</div>
          </div>
        </div>
        <div class="detail-actions">
          <button id="btnExp" class="btn ghost">导出 Excel</button>
          <button id="btnDel" class="btn danger">删除盘点单</button>
        </div>
        <div class="toolbar">
          <div class="search-row">
            <input id="sSearch" class="search" placeholder="搜索添加物料…">
            <input id="lSearch" class="search" placeholder="本单查找…">
          </div>
        </div>
        <div class="toolbar sub">
          <select id="sWh" class="search"><option value="">全部仓库</option></select>
          <button id="btnSelModeDetail" class="btn ghost sm">选择</button>
        </div>
        <div id="sResults" class="results hidden"></div>
        <div class="count-bar"><span id="lCount" class="count"></span></div>
        <div id="lineBulk" class="bulk-line-bar hidden">
          <label class="lb-selall"><input type="checkbox" id="lineSelAll"> 全选</label>
          <span id="lineSelCount" class="sel-count">已选 0</span>
          <div class="lb-actions">
            <button class="btn ghost sm" id="lineBatchRemark">批量备注</button>
            <button class="btn danger sm" id="lineBatchDel">批量删除</button>
          </div>
        </div>
      </div>
      <div id="lList" class="list view-scroll"></div>`;
    const sSearch = view.querySelector('#sSearch');
    const sWh = view.querySelector('#sWh');
    const sResults = view.querySelector('#sResults');
    fillWarehouseOptions(sWh, '');
    view.querySelector('#back').onclick = goSheets;
    sSearch.addEventListener('input', debounce(() => searchMaterialsToAdd(sResults, id), 180));
    sWh.addEventListener('change', () => searchMaterialsToAdd(sResults, id));
    const lSearch = view.querySelector('#lSearch');
    lSearch.addEventListener('input', () => filterLines(lSearch.value));
    view.querySelector('#btnExp').onclick = async () => {
      try { await ensureXLSX(); } catch (e) { toast('组件加载失败，无法导出'); return; }
      const lines = db().listLines(id);
      ExportXLSX.exportSheet(sheet, lines);
    };
    view.querySelector('#btnDel').onclick = async () => {
      if (!(await askConfirm('删除该盘点单？将移入回收站，可在回收站还原'))) return;
      db().deleteSheet(id); goSheets(); toast('已移入回收站');
    };
    // 明细多选批量栏
    const lineBulk = view.querySelector('#lineBulk');
    const lineSelAll = view.querySelector('#lineSelAll');
    const lineSelCount = view.querySelector('#lineSelCount');
    lineSelAll.addEventListener('change', () => {
      view.querySelectorAll('#lList .card').forEach(c => {
        if (c.style.display === 'none') return; // 仅勾选当前过滤可见的行
        const cb = c.querySelector('.line-sel');
        if (!cb || cb.disabled) return;
        cb.checked = lineSelAll.checked;
        const lid = +cb.dataset.id;
        if (lineSelAll.checked) selectedLineIds.add(lid); else selectedLineIds.delete(lid);
      });
      updateLineBulk();
    });
    view.querySelector('#lineBatchDel').addEventListener('click', async () => {
      const ids = [...selectedLineIds];
      if (!ids.length) return;
      if (!(await askConfirm('确定批量删除选中的 ' + ids.length + ' 个物料行？此操作不可恢复', { danger: true }))) return;
      ids.forEach(lid => db().removeLine(lid));
      selectedLineIds.clear();
      lineSelAll.checked = false;
      loadLines(id);
      toast('已删除 ' + ids.length + ' 行');
    });
    view.querySelector('#lineBatchRemark').addEventListener('click', () => {
      const ids = [...selectedLineIds];
      if (!ids.length) return;
      openBatchRemarkModal(ids, id);
    });
    view.querySelector('#btnSelModeDetail').onclick = () => toggleDetailSelect();
    loadLines(id);
  }

  // 盘点单添加物料：候选列表支持多选（复选框 + 全选当前 + 底部「添加选中」）
  function searchMaterialsToAdd(resultsEl, sheetId) {
    const sEl = view.querySelector('#sSearch');
    if (!sEl) return; // 详情页已切走（如返回列表），pending 的防抖不应再渲染
    const term = (sEl.value || '').trim();
    const whEl = view.querySelector('#sWh');
    const wh = whEl ? (whEl.value || '').trim() : '';
    if (!term && !wh) { resultsEl.classList.add('hidden'); resultsEl.innerHTML = ''; selectedMatIds.clear(); return; }
    const lineByMat = {};
    db().listLines(sheetId).forEach(l => { if (l.material_id != null) lineByMat[l.material_id] = l.id; });
    let rows = db().listMaterials(term, wh);
    // 仓库筛选客户端兜底：只保留该仓库的物料
    if (wh) rows = rows.filter(m => (m.warehouse || '').trim() === wh);
    if (!rows.length) {
      resultsEl.innerHTML = `<div class="empty small">${wh ? '该仓库下没有可添加的物料' : '无匹配物料（可去"物料档案"先导入）'}</div>`;
      resultsEl.classList.remove('hidden'); return;
    }
    rows = rows.slice(0, 300);
    resultsEl.innerHTML = `
      <div class="bulk-head">
        <label class="bh-selall"><input type="checkbox" id="selAll"> 全选当前 (${rows.length})</label>
        <span class="sel-count" id="selCount">已选 0</span>
      </div>` + rows.map(m => {
        const exists = lineByMat[m.id] != null;
        const checked = selectedMatIds.has(m.id) && !exists;
        return `
        <div class="result-item${exists ? ' exists' : ''}" data-id="${m.id}" data-line="${exists ? lineByMat[m.id] : ''}">
          <input type="checkbox" class="sel" data-id="${m.id}" ${exists ? 'disabled' : ''} ${checked ? 'checked' : ''}>
          <span class="ri-main">${esc(m.name)} <span class="muted">${esc(m.code)}</span></span>
          <span class="ri-sub muted">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}${exists ? ' · <span class="added">已添加</span>' : ''}</span>
        </div>`;
      }).join('') + `
      <button id="btnAddSel" class="bulk-add">添加选中 (0)</button>`;
    resultsEl.classList.remove('hidden');
    const selAll = resultsEl.querySelector('#selAll');
    const selCount = resultsEl.querySelector('#selCount');
    const btnAdd = resultsEl.querySelector('#btnAddSel');
    function syncSel() {
      const boxes = resultsEl.querySelectorAll('.sel:not(:disabled)');
      const checkedBoxes = resultsEl.querySelectorAll('.sel:not(:disabled):checked');
      selCount.textContent = '已选 ' + checkedBoxes.length;
      btnAdd.textContent = '添加选中 (' + checkedBoxes.length + ')';
      selAll.checked = boxes.length > 0 && boxes.length === checkedBoxes.length;
      selAll.indeterminate = checkedBoxes.length > 0 && checkedBoxes.length < boxes.length;
    }
    resultsEl.querySelectorAll('.sel').forEach(cb => {
      cb.addEventListener('change', () => {
        const mid = +cb.dataset.id;
        if (cb.checked) selectedMatIds.add(mid); else selectedMatIds.delete(mid);
        syncSel();
      });
    });
    resultsEl.querySelectorAll('.result-item').forEach(it => {
      it.addEventListener('click', (e) => {
        if (e.target.classList.contains('sel')) return; // 复选框自行处理
        const mid = +it.dataset.id;
        const lineId = it.dataset.line ? +it.dataset.line : null;
        if (lineId != null) { locateLine(lineId, sheetId); return; } // 已添加：定位到本单行
        const cb = it.querySelector('.sel');
        if (cb && !cb.disabled) { cb.checked = !cb.checked; cb.dispatchEvent(new Event('change')); }
      });
    });
    selAll.addEventListener('change', () => {
      resultsEl.querySelectorAll('.sel:not(:disabled)').forEach(cb => {
        cb.checked = selAll.checked;
        const mid = +cb.dataset.id;
        if (selAll.checked) selectedMatIds.add(mid); else selectedMatIds.delete(mid);
      });
      syncSel();
    });
    btnAdd.addEventListener('click', () => addSelectedLines(sheetId, resultsEl));
    syncSel();
  }

  // 批量将勾选的候选物料加入盘点单
  function addSelectedLines(sheetId, resultsEl) {
    const ids = [...selectedMatIds];
    if (!ids.length) { toast('请先勾选要添加的物料'); return; }
    let added = 0;
    ids.forEach(mid => {
      const m = db().getMaterial(mid);
      if (m) { db().addLine(sheetId, m); added++; }
    });
    selectedMatIds.clear();
    const sInput = view.querySelector('#sSearch'); if (sInput) sInput.value = '';
    const lInput = view.querySelector('#lSearch'); if (lInput) lInput.value = '';
    resultsEl.classList.add('hidden'); resultsEl.innerHTML = '';
    loadLines(sheetId);
    const lastCard = view.querySelector('#lList .card:last-child');
    if (lastCard) {
      lastCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
      lastCard.classList.add('flash');
      setTimeout(() => lastCard.classList.remove('flash'), 1600);
    }
    toast('已添加 ' + added + ' 个物料');
  }

  // 已存在的物料：收起搜索、滚动并高亮定位到对应行，便于直接修改
  function locateLine(lineId, sheetId) {
    const sInput = view.querySelector('#sSearch');
    if (sInput) sInput.value = '';
    const resultsEl = view.querySelector('#sResults');
    if (resultsEl) { resultsEl.classList.add('hidden'); resultsEl.innerHTML = ''; }
    loadLines(sheetId);
    const card = view.querySelector(`#lList .card[data-id="${lineId}"]`);
    if (!card) return;
    card.style.display = ''; // 不受「本单查找」过滤影响，强制显示，否则 scrollIntoView 看不到
    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    card.classList.add('flash');
    setTimeout(() => card.classList.remove('flash'), 1600);
  }

  // 本单内查找：实时过滤已录入的物料行，便于在大量物料中快速定位到要改的那条
  function filterLines(term) {
    const el = view.querySelector('#lList');
    if (!el) return;
    term = (term || '').trim().toLowerCase();
    let visible = 0;
    el.querySelectorAll('.card').forEach(c => {
      const title = c.querySelector('.row-title') ? c.querySelector('.row-title').textContent.toLowerCase() : '';
      const sub = c.querySelector('.row-sub') ? c.querySelector('.row-sub').textContent.toLowerCase() : '';
      const match = !term || title.includes(term) || sub.includes(term);
      c.style.display = match ? '' : 'none';
      if (match) visible++;
    });
    let emptyEl = el.querySelector('.empty.filter');
    if (term && visible === 0) {
      if (!emptyEl) { emptyEl = document.createElement('div'); emptyEl.className = 'empty filter'; el.appendChild(emptyEl); }
      emptyEl.textContent = '本单中没有匹配“' + term + '”的物料';
    } else if (emptyEl) {
      emptyEl.remove();
    }
  }

  function loadLines(sheetId) {
    const lines = db().listLines(sheetId);
    const el = view.querySelector('#lList');
    if (!lines.length) { el.innerHTML = '<div class="empty">还没有物料，上方搜索“物料编码/名称”添加</div>'; return; }
    el.innerHTML = lines.map(l => `
      <div class="card${detailSelectMode ? ' selectmode' : ''}${selectedLineIds.has(l.id) ? ' selected' : ''}" data-id="${l.id}">
        <input type="checkbox" class="line-sel mat-sel" data-id="${l.id}" ${selectedLineIds.has(l.id) ? 'checked' : ''}>
        <div class="card-body">
        <div class="row">
          <div class="row-main">
            <div class="row-title">${esc(l.name)} <span class="muted">${esc(l.code)}</span>${l.remark ? ' <span class="remark-flag">已备注</span>' : ''}</div>
            <div class="row-sub">${esc(l.unit || '')} ${esc(l.spec || '')}${l.warehouse ? ' · ' + esc(l.warehouse) : ''}</div>
          </div>
          <div class="row-actions">
            <button class="icon-btn" data-act="remark" title="备注">📝</button>
            <button class="icon-btn danger" data-act="rm">🗑</button>
          </div>
        </div>
        <div class="stepper">
          <button class="step-btn" data-act="dec">−</button>
          <input class="qty" type="text" inputmode="text" value="${l.qty}" data-id="${l.id}" placeholder="可输入算式">
          <button class="step-btn" data-act="inc">＋</button>
        </div>
        <div class="ops" data-id="${l.id}">
          <button data-op="+">+</button>
          <button data-op="-">−</button>
          <button data-op="*">×</button>
          <button data-op="/">÷</button>
          <button data-op="=">=</button>
          <button data-op="c" class="op-c">C</button>
        </div>
        ${l.remark ? `<div class="remark-preview">${esc(l.remark)}</div>` : ''}
      </div>
      </div>`).join('');
    el.querySelectorAll('.card').forEach(c => {
      const lid = +c.dataset.id;
      const lineSel = c.querySelector('.line-sel');
      lineSel.addEventListener('change', () => {
        if (lineSel.checked) selectedLineIds.add(lid); else selectedLineIds.delete(lid);
        c.classList.toggle('selected', lineSel.checked);
        updateLineBulk();
      });
      if (detailSelectMode) {
        c.addEventListener('click', (e) => {
          if (e.target.closest('.line-sel')) return; // 点复选框本身由 change 处理，避免双触发
          lineSel.checked = !lineSel.checked;
          if (lineSel.checked) selectedLineIds.add(lid); else selectedLineIds.delete(lid);
          c.classList.toggle('selected', lineSel.checked);
          updateLineBulk();
        });
      }
      c.querySelector('[data-act=dec]').onclick = () => {
        const cur = db().getLine(lid).qty;
        if (cur <= 0) { toast('盘点数量已是最小 0'); return; }
        db().incLine(lid, -1);
        qty.value = db().getLine(lid).qty; // 原地更新，不重渲染整单
      };
      c.querySelector('[data-act=inc]').onclick = () => {
        db().incLine(lid, 1);
        qty.value = db().getLine(lid).qty; // 原地更新，不重渲染整单
      };
      c.querySelector('[data-act=rm]').onclick = async () => { if (!(await askConfirm('移除该物料行？'))) return; db().removeLine(lid); selectedLineIds.delete(lid); loadLines(sheetId); };
      c.querySelector('[data-act=remark]').onclick = () => openRemarkModal(lid, db().getLine(lid).remark || '', sheetId);
      const qty = c.querySelector('.qty');
      qty.addEventListener('focus', () => c.classList.add('editing'));
      qty.addEventListener('blur', () => {
        c.classList.remove('editing');
        if (qty.value.trim() === '') { db().updateLineQty(lid, 0); qty.value = 0; return; }
        const v = evalQtyExpr(qty.value);
        if (isNaN(v)) { toast('算式无法计算'); qty.value = l.qty; return; }
        if (v < 0) { toast('盘点数量不能为负数，已设为 0'); db().updateLineQty(lid, 0); qty.value = 0; return; }
        db().updateLineQty(lid, v);
        qty.value = v;
      });
      // 运算符小按钮（聚焦时浮出）；pointerdown 阻止失焦，保证连续输入
      c.querySelectorAll('.ops button').forEach(b => {
        b.addEventListener('pointerdown', e => {
          e.preventDefault();
          const op = b.dataset.op;
          if (op === 'c') { qty.value = ''; qty.focus(); return; }
          if (op === '=') { qty.blur(); return; }
          qty.value = qty.value + op;
          qty.focus();
          const len = qty.value.length;
          try { qty.setSelectionRange(len, len); } catch (_) {}
        });
      });
    });
    const lf = view.querySelector('#lSearch');
    if (lf) filterLines(lf.value);
    const lc = view.querySelector('#lCount');
    if (lc) lc.textContent = '本单已录入 ' + lines.length + ' 个物料';
    updateLineBulk();
  }

  // 盘点单行备注弹窗（多行输入，随盘点单导出）
  function openRemarkModal(lineId, current, sheetId) {
    openModal(`
      <h3>物料备注</h3>
      <p class="muted" style="font-size:13px;margin:0 0 10px;line-height:1.5;">记录该物料的异常情况，如破损、临期、数量不符等。备注会随盘点单一起导出。</p>
      <textarea id="rmkInput" class="remark-input" rows="4" placeholder="例如：外箱破损3个 / 临期商品 / 实盘比账面少2">${esc(current)}</textarea>
      <div class="modal-actions">
        <button class="btn ghost" id="rmkCancel">取消</button>
        <button class="btn primary" id="rmkSave">保存</button>
      </div>`);
    const ta = modal.querySelector('#rmkInput');
    ta.focus();
    modal.querySelector('#rmkCancel').onclick = closeModal;
    modal.querySelector('#rmkSave').onclick = () => {
      const v = ta.value.trim();
      db().updateLineRemark(lineId, v);
      closeModal();
      loadLines(sheetId);
      toast(v ? '备注已保存' : '备注已清空');
    };
  }

  // 切换盘点单详情「已录入物料」选择模式（勾选框默认隐藏，点「选择」才显示）
  function toggleDetailSelect() {
    detailSelectMode = !detailSelectMode;
    if (!detailSelectMode) selectedLineIds.clear();
    loadLines(currentSheetId);
    updateLineBulk();
  }

  // 明细多选：根据 selectedLineIds 同步批量栏显隐、计数、全选态（仅看可见行）
  function updateLineBulk() {
    const bar = view.querySelector('#lineBulk');
    const cnt = view.querySelector('#lineSelCount');
    const all = view.querySelector('#lineSelAll');
    const btn = view.querySelector('#btnSelModeDetail');
    if (btn) btn.textContent = detailSelectMode ? '取消' : '选择';
    if (!bar) return;
    const n = selectedLineIds.size;
    if (cnt) cnt.textContent = '已选 ' + n;
    bar.classList.toggle('hidden', !detailSelectMode);
    if (all) {
      const visBoxes = [...view.querySelectorAll('#lList .card')]
        .filter(c => c.style.display !== 'none')
        .map(c => c.querySelector('.line-sel'))
        .filter(cb => cb && !cb.disabled);
      const visChecked = visBoxes.filter(cb => cb.checked);
      all.checked = visBoxes.length > 0 && visBoxes.length === visChecked.length;
      all.indeterminate = visChecked.length > 0 && visChecked.length < visBoxes.length;
    }
  }

  // 批量给选中的明细行填同一备注
  function openBatchRemarkModal(ids, sheetId) {
    openModal(`
      <h3>批量备注（${ids.length} 行）</h3>
      <p class="muted" style="font-size:13px;margin:0 0 10px;line-height:1.5;">将为选中的 ${ids.length} 个物料行填写相同备注，覆盖各自原有备注。备注随盘点单一起导出。</p>
      <textarea id="rmkInput" class="remark-input" rows="4" placeholder="例如：整批临期 / 外箱破损"></textarea>
      <div class="modal-actions">
        <button class="btn ghost" id="rmkCancel">取消</button>
        <button class="btn primary" id="rmkSave">保存</button>
      </div>`);
    const ta = modal.querySelector('#rmkInput');
    ta.focus();
    modal.querySelector('#rmkCancel').onclick = closeModal;
    modal.querySelector('#rmkSave').onclick = () => {
      const v = ta.value.trim();
      ids.forEach(lid => db().updateLineRemark(lid, v));
      selectedLineIds.clear();
      closeModal();
      loadLines(sheetId);
      toast('已为 ' + ids.length + ' 行添加备注');
    };
  }

  // ============ 导入文件 ============
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      await ensureXLSX();
      const { items, skipped } = await ImportXLSX.parseFile(f);
      let n = 0;
      for (const it of items) { db().upsertMaterial(it); n++; }
      toast(`导入完成：成功 ${n} 条${skipped ? `，跳过 ${skipped} 条（缺编码或名称）` : ''}`);
      if (document.querySelector('#mList')) loadMList(view.querySelector('#mSearch').value.trim());
    } catch (err) {
      toast('导入失败：' + err.message);
    }
    fileInput.value = '';
  });

  // ============ 备份文件恢复 ============
  backupInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const text = await f.text();
      if (!(await askConfirm('恢复备份将覆盖当前所有数据，确定继续？'))) { backupInput.value = ''; return; }
      await window.AppDB.importBackupString(text);
      toast('备份已恢复');
      goMaterials();
    } catch (err) {
      toast('恢复失败：' + err.message);
    }
    backupInput.value = '';
  });

  // ============ 启动 ============
  // 轻量更新提示：新版本 SW 安装完成后，提示用户刷新（无需再手动杀后台）
  function showUpdateBanner() {
    if (document.getElementById('updateBar')) return;
    const bar = document.createElement('div');
    bar.id = 'updateBar';
    bar.className = 'update-bar';
    bar.innerHTML = '<span>发现新版本，建议刷新以使用最新功能</span><button id="updateReload">立即刷新</button>';
    document.body.appendChild(bar);
    bar.querySelector('#updateReload').onclick = () => location.reload();
  }

  async function boot() {
    applyTheme(localStorage.getItem('theme') || 'system');
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('sw.js').then(reg => {
        reg.addEventListener('updatefound', () => {
          const nw = reg.installing;
          if (!nw) return;
          nw.addEventListener('statechange', () => {
            // 仅在「已存在旧版本（有 controller）且新版本已安装」时提示，避免首装误弹
            if (nw.state === 'installed' && navigator.serviceWorker.controller) showUpdateBanner();
          });
        });
      }).catch(() => {});
    }
    try {
      await window.AppDB.init();
      goMaterials();
    } catch (err) {
      view.innerHTML = `<div class="empty">初始化失败：${esc(err.message)}</div>`;
      console.error(err);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
