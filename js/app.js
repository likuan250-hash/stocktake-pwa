(function () {
  'use strict';

  const view = document.getElementById('view');
  const topTitle = document.getElementById('topTitle');
  const themeBtn = document.getElementById('themeBtn');
  const modal = document.getElementById('modal');
  const fileInput = document.getElementById('fileInput');
  const backupInput = document.getElementById('backupInput');
  const db = () => window.AppDB.api;
  // 当前应用版本（发布时与 sw.js 的 CACHE 名 stocktake-pwa-<ver> 保持同步递增）
  const APP_VERSION = 'v33';
  const verBtn = document.getElementById('verBtn');
  let currentSheetId = null;
  // 跨函数共享的「已勾选」状态：候选物料与明细行。必须位于 II FE 顶层作用域，
  // 否则 searchMaterialsToAdd / addSelectedLines / loadLines 等外层函数引用时会 ReferenceError。
  const selectedMatIds = new Set();
  const selectedLineIds = new Set();
  // 盘点单「添加物料」候选列表滚动加载状态（默认铺开全部物料，边滚边加载）
  let addAll = [];        // 当前匹配的物料（已排序、已按仓库过滤）
  let addRendered = 0;    // 已渲染到 DOM 的条数
  const ADD_PAGE = 150;   // 每批渲染条数
  let resultsScrollHandler = null; // resultsEl 的 scroll 监听引用，便于重绑时解绑
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
          <button id="btnKd" class="btn ghost">金蝶导入</button>
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
    view.querySelector('#btnKd').onclick = openKingdeeImport;
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

  // ============ 从金蝶盘点单导入物料 ============
  // 静态同步库（window.KINGDEE_SHEETS，由 KingdeeMCP 按单号拉取生成）。
  // 物料档案页按单号选单 → 预览 → 确认 upsert 进物料主数据（加法功能，不影响现有导入/新增）。
  function openKingdeeImport() {
    const sheets = window.KINGDEE_SHEETS || [];
    if (!sheets.length) { toast('暂无可用的金蝶盘点单'); return; }
    let importMode = 'merge'; // 'merge' = 更新添加(默认,安全) | 'clear' = 清空后导入
    renderModeChoice();

    function refresh() {
      const wh = view.querySelector('#mWh');
      if (wh) fillWarehouseOptions(wh, wh.value);
      loadMList(view.querySelector('#mSearch').value.trim(), wh ? wh.value : '');
    }

    function renderModeChoice() {
      openModal(`
        <h3>从金蝶盘点单导入<button class="btn-close" id="kd_close">✕</button></h3>
        <p class="muted" style="font-size:13px;margin:0 0 12px;line-height:1.6;">请选择导入方式：</p>
        <div class="list" style="gap:10px;">
          <div class="card row kd-mode ${importMode === 'merge' ? 'selected' : ''}" data-mode="merge">
            <div class="row-main">
              <div class="row-title">更新添加（保留现有物料）</div>
              <div class="row-sub">按编码合并：已有则更新，没有则新增。不影响现有物料。</div>
            </div>
          </div>
          <div class="card row kd-mode danger ${importMode === 'clear' ? 'selected' : ''}" data-mode="clear">
            <div class="row-main">
              <div class="row-title">清空后导入（先删全部再导入所选单）</div>
              <div class="row-sub">会先删除全部现有物料（回收站保留不动），再导入所选单，结果仅含该单物料。</div>
            </div>
          </div>
        </div>
        <div class="modal-actions" style="margin-top:14px;">
          <button class="btn ghost" id="kd_cancel">取消</button>
          <button class="btn primary" id="kd_next">继续</button>
        </div>
      `);
      modal.querySelector('#kd_close').onclick = closeModal;
      modal.querySelector('#kd_cancel').onclick = closeModal;
      modal.querySelectorAll('.kd-mode').forEach(c => {
        c.onclick = () => {
          importMode = c.dataset.mode;
          modal.querySelectorAll('.kd-mode').forEach(x => x.classList.remove('selected'));
          c.classList.add('selected');
        };
      });
      modal.querySelector('#kd_next').onclick = renderSheetList;
    }

    function renderSheetList() {
      openModal(`
        <h3>从金蝶盘点单导入<button class="btn-close" id="kd_close">✕</button></h3>
        <p class="muted" style="font-size:13px;margin:0 0 12px;line-height:1.6;">${importMode === 'clear' ? '清空模式：导入后物料档案将仅含所选单。' : '更新添加模式：按编码合并进现有物料。'}选择一个金蝶「物料盘点作业表」。</p>
        <div id="kdList" class="list" style="max-height:58vh;overflow-y:auto;gap:8px;"></div>
      `);
      modal.querySelector('#kd_close').onclick = closeModal;
      const el = modal.querySelector('#kdList');
      const activeCodes = new Set(db().listMaterials().map(m => m.code));
      el.innerHTML = sheets.map((s, i) => {
        const newCnt = importMode === 'clear' ? s.materials.length : s.materials.filter(m => m.code && !activeCodes.has(m.code)).length;
        const updCnt = s.materials.length - newCnt;
        return `
        <div class="card row kd-sheet" data-i="${i}">
          <div class="row-main" data-i="${i}">
            <div class="row-title">${esc(s.billNo)} <span class="muted">${esc(s.org)}</span></div>
            <div class="row-sub">${esc(s.date)} · ${s.materials.length} 个物料（新增 ${newCnt} / 更新 ${updCnt}）</div>
          </div>
          <div class="row-actions"><button class="btn primary sm" data-i="${i}">选择</button></div>
        </div>`;
      }).join('');
      el.querySelectorAll('.kd-sheet').forEach(c => {
        const go = () => renderPreview(+c.dataset.i);
        c.querySelector('.row-main').onclick = go;
        c.querySelector('.btn').onclick = go;
      });
    }

    function renderPreview(i) {
      const s = sheets[i];
      const isClear = importMode === 'clear';
      openModal(`
        <h3>${esc(s.billNo)} 物料预览<button class="btn-close" id="kd_back">←</button></h3>
        <p class="muted" style="font-size:13px;margin:0 0 10px;line-height:1.6;">${esc(s.org)} · ${esc(s.date)} · 共 ${s.materials.length} 个物料。${isClear ? '确认后将<strong>先清空全部现有物料（回收站保留）</strong>，再导入本单。' : '确认后按编码 upsert 进物料档案。'}</p>
        <div id="kdPrev" class="list" style="max-height:52vh;overflow-y:auto;gap:8px;"></div>
        <div class="modal-actions" style="margin-top:12px;">
          <button class="btn ghost" id="kd_back2">返回</button>
          <button class="btn ${isClear ? 'danger' : 'primary'}" id="kd_confirm">${isClear ? '清空并导入 ' + s.materials.length + ' 个' : '确认导入 ' + s.materials.length + ' 个'}</button>
        </div>
      `);
      const prev = modal.querySelector('#kdPrev');
      prev.innerHTML = s.materials.map(m => {
        const isNew = isClear ? true : (m.code && !db().getMaterialByCode(m.code));
        return `
        <div class="card row">
          <div class="row-main">
            <div class="row-title">${esc(m.name)} <span class="muted">${esc(m.code)}</span> ${isNew ? '<span class="remark-flag">新增</span>' : '<span class="muted">更新</span>'}</div>
            <div class="row-sub">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}</div>
          </div>
        </div>`;
      }).join('');
      modal.querySelector('#kd_back').onclick = renderSheetList;
      modal.querySelector('#kd_back2').onclick = renderSheetList;
      modal.querySelector('#kd_confirm').onclick = async () => {
        await importMaterials(s.materials, importMode, { label: s.billNo, onDone: refresh });
      };
    }
  }

  // ============ 共享：物料导入（金蝶单 / XLSX 通用） ============
  // mode: 'merge'(更新添加,默认) | 'clear'(清空后导入)
  // 统一处理：回收站同编码提示、二次确认、清空模式整库快照+失败回滚兜底。
  async function importMaterials(items, mode, opts) {
    opts = opts || {};
    const onDone = typeof opts.onDone === 'function' ? opts.onDone : function () {};
    const valid = (items || []).filter(m => m.code);
    if (!valid.length) { toast('没有可导入的物料（缺少编码）'); return false; }

    // 回收站同编码提示：合并/清空两种模式 upsert 都会把同编码的回收站物料恢复为活跃
    const trashCodes = new Set(db().listDeletedMaterials().map(m => m.code));
    const overlap = valid.filter(m => trashCodes.has(m.code)).map(m => m.code);
    if (overlap.length) {
      const preview = overlap.slice(0, 12).join('、') + (overlap.length > 12 ? ' 等' : '');
      const ok = await askConfirm('以下 ' + overlap.length + ' 个编码在回收站中也存在，导入后将恢复为活跃物料：\n' + preview + '\n继续导入？');
      if (!ok) return false;
    }

    if (mode === 'clear') {
      const snapshot = window.AppDB.getBackupString();
      const cleared = db().listMaterials().length;
      const ok = await askConfirm('此操作不可恢复：将先删除全部现有物料（回收站保留不动），再导入 ' + valid.length + ' 个物料。确定清空并导入？');
      if (!ok) return false;
      try {
        db().clearAllMaterials();
        let imported = 0;
        valid.forEach(m => { db().upsertMaterial(m); imported++; });
        closeModal();
        toast('已清空 ' + cleared + ' 条、导入 ' + imported + ' 个');
        onDone();
        return true;
      } catch (e) {
        console.error('清空导入失败，正在回滚', e);
        try { await window.AppDB.importBackupString(snapshot); } catch (e2) { console.error('回滚失败', e2); }
        closeModal();
        toast('清空导入失败，已回滚到导入前状态');
        onDone();
        return false;
      }
    }

    // merge 模式
    const ok = await askConfirm('确认将 ' + valid.length + ' 个物料 upsert 进物料档案（按编码合并：已有则更新，没有则新增）？');
    if (!ok) return false;
    let added = 0, updated = 0;
    try {
      valid.forEach(m => {
        const ex = db().getMaterialByCode(m.code);
        db().upsertMaterial(m);
        if (ex) updated++; else added++;
      });
      closeModal();
      toast('已导入 ' + (added + updated) + ' 个（新增 ' + added + ' / 更新 ' + updated + '）');
      onDone();
      return true;
    } catch (e) {
      console.error('导入失败', e);
      closeModal();
      toast('导入失败：' + (e && e.message ? e.message : e));
      onDone();
      return false;
    }
  }

  // 通用「导入方式」选择弹窗（金蝶单与 XLSX 复用）
  function renderImportModeChoice(opts) {
    opts = opts || {};
    const count = opts.count || 0;
    const onNext = opts.onNext || function () {};
    let mode = 'merge';
    openModal(`
      <h3>导入物料<button class="btn-close" id="im_close">✕</button></h3>
      <p class="muted" style="font-size:13px;margin:0 0 12px;line-height:1.6;">请选择导入方式（共 ${count} 个有效物料）：</p>
      <div class="list" style="gap:10px;">
        <div class="card row kd-mode selected" data-mode="merge">
          <div class="row-main">
            <div class="row-title">更新添加（保留现有物料）</div>
            <div class="row-sub">按编码合并：已有则更新，没有则新增。不影响现有物料。</div>
          </div>
        </div>
        <div class="card row kd-mode danger" data-mode="clear">
          <div class="row-main">
            <div class="row-title">清空后导入（先删全部再导入）</div>
            <div class="row-sub">会先删除全部现有物料（回收站保留不动），再导入所选数据，结果仅含本次数据。</div>
          </div>
        </div>
      </div>
      <div class="modal-actions" style="margin-top:14px;">
        <button class="btn ghost" id="im_cancel">取消</button>
        <button class="btn primary" id="im_next">继续</button>
      </div>
    `);
    modal.querySelector('#im_close').onclick = closeModal;
    modal.querySelector('#im_cancel').onclick = closeModal;
    modal.querySelectorAll('.kd-mode').forEach(c => {
      c.onclick = () => {
        mode = c.dataset.mode;
        modal.querySelectorAll('.kd-mode').forEach(x => x.classList.remove('selected'));
        c.classList.add('selected');
      };
    });
    modal.querySelector('#im_next').onclick = () => onNext(mode);
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
    sSearch.addEventListener('focus', () => searchMaterialsToAdd(sResults, id)); // 点搜索框即铺开全部（空搜索展示全部）
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

  // 盘点单添加物料：候选列表支持多选（复选框 + 全选 + 底部「添加选中」）。
  // 默认（空搜索）即铺开全部物料（底层 ORDER BY code，编码升序），支持滚动加载（每批 ADD_PAGE 条）。
  function searchMaterialsToAdd(resultsEl, sheetId) {
    const sEl = view.querySelector('#sSearch');
    if (!sEl) return; // 详情页已切走（如返回列表），pending 的防抖不应再渲染
    // 清理上一次可能残留的滚动监听，避免重复绑定
    if (resultsScrollHandler) { resultsEl.removeEventListener('scroll', resultsScrollHandler); resultsScrollHandler = null; }
    const term = (sEl.value || '').trim();
    const whEl = view.querySelector('#sWh');
    const wh = whEl ? (whEl.value || '').trim() : '';
    const lineByMat = {};
    db().listLines(sheetId).forEach(l => { if (l.material_id != null) lineByMat[l.material_id] = l.id; });
    let rows = db().listMaterials(term, wh);
    // 仓库筛选客户端兜底：只保留该仓库的物料
    if (wh) rows = rows.filter(m => (m.warehouse || '').trim() === wh);
    addAll = rows;
    addRendered = 0;
    if (!addAll.length) {
      resultsEl.innerHTML = `<div class="empty small">${wh ? '该仓库下没有可添加的物料' : '无匹配物料（可去"物料档案"先导入）'}</div>`;
      resultsEl.classList.remove('hidden');
      return;
    }
    // 骨架：全选栏 + 列表容器 + 加载更多（兜底）+ 添加选中
    resultsEl.innerHTML = `
      <div class="bulk-head">
        <label class="bh-selall"><input type="checkbox" id="selAll"> 全选 (${addAll.length})</label>
        <span class="sel-count" id="selCount">已选 0</span>
      </div>
      <div id="matItems"></div>
      <button id="btnLoadMore" class="bulk-add more" style="display:none">加载更多</button>
      <button id="btnAddSel" class="bulk-add">添加选中 (0)</button>`;
    resultsEl.classList.remove('hidden');
    const matItems = resultsEl.querySelector('#matItems');
    const selAll = resultsEl.querySelector('#selAll');
    const selCount = resultsEl.querySelector('#selCount');
    const btnAdd = resultsEl.querySelector('#btnAddSel');
    const btnLoadMore = resultsEl.querySelector('#btnLoadMore');

    function makeItem(m) {
      const exists = lineByMat[m.id] != null;
      const checked = selectedMatIds.has(m.id) && !exists;
      const it = document.createElement('div');
      it.className = 'result-item' + (exists ? ' exists' : '');
      it.dataset.id = m.id;
      it.dataset.line = exists ? lineByMat[m.id] : '';
      it.innerHTML = `
        <input type="checkbox" class="sel" data-id="${m.id}" ${exists ? 'disabled' : ''} ${checked ? 'checked' : ''}>
        <span class="ri-main">${esc(m.name)} <span class="muted">${esc(m.code)}</span></span>
        <span class="ri-sub muted">${esc(m.unit || '')} ${esc(m.spec || '')}${m.warehouse ? ' · ' + esc(m.warehouse) : ''}${exists ? ' · <span class="added">已添加</span>' : ''}</span>`;
      it.querySelector('.sel').addEventListener('change', () => {
        if (it.querySelector('.sel').checked) selectedMatIds.add(m.id); else selectedMatIds.delete(m.id);
        syncSel();
      });
      it.addEventListener('click', (e) => {
        if (e.target.classList.contains('sel')) return; // 复选框自行处理
        const lineId = it.dataset.line ? +it.dataset.line : null;
        if (lineId != null) { locateLine(lineId, sheetId); return; } // 已添加：定位到本单行
        const cb = it.querySelector('.sel');
        if (cb && !cb.disabled) { cb.checked = !cb.checked; cb.dispatchEvent(new Event('change')); }
      });
      return it;
    }

    function appendBatch() {
      const batch = addAll.slice(addRendered, addRendered + ADD_PAGE);
      const frag = document.createDocumentFragment();
      batch.forEach(m => frag.appendChild(makeItem(m)));
      matItems.appendChild(frag);
      addRendered += batch.length;
      if (addRendered >= addAll.length) btnLoadMore.style.display = 'none';
      else { btnLoadMore.style.display = ''; btnLoadMore.textContent = '加载更多 (' + (addAll.length - addRendered) + ')'; }
    }

    function syncSel() {
      const boxes = resultsEl.querySelectorAll('.sel:not(:disabled)');
      const checkedBoxes = resultsEl.querySelectorAll('.sel:not(:disabled):checked');
      selCount.textContent = '已选 ' + checkedBoxes.length;
      btnAdd.textContent = '添加选中 (' + checkedBoxes.length + ')';
      selAll.checked = boxes.length > 0 && boxes.length === checkedBoxes.length;
      selAll.indeterminate = checkedBoxes.length > 0 && checkedBoxes.length < boxes.length;
    }

    appendBatch();
    syncSel();

    selAll.addEventListener('change', () => {
      // 全选作用于「全部匹配结果」（跨分页），不仅限于已渲染
      addAll.forEach(m => {
        if (lineByMat[m.id] != null) return; // 已添加的不参与
        if (selAll.checked) selectedMatIds.add(m.id); else selectedMatIds.delete(m.id);
      });
      resultsEl.querySelectorAll('.sel:not(:disabled)').forEach(cb => { cb.checked = selAll.checked; });
      syncSel();
    });
    btnAdd.addEventListener('click', () => addSelectedLines(sheetId, resultsEl));
    btnLoadMore.addEventListener('click', appendBatch);

    // 滚动到底自动加载下一批
    resultsScrollHandler = () => {
      if (addRendered >= addAll.length) return;
      if (resultsEl.scrollTop + resultsEl.clientHeight >= resultsEl.scrollHeight - 60) appendBatch();
    };
    resultsEl.addEventListener('scroll', resultsScrollHandler);
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
    let parsed;
    try {
      await ensureXLSX();
      parsed = await ImportXLSX.parseFile(f);
    } catch (err) {
      toast('导入失败：' + err.message);
      fileInput.value = '';
      return;
    }
    const { items, skipped } = parsed;
    fileInput.value = '';
    if (!items.length) { toast('未解析到有效物料（需含编码与名称列）'); return; }
    renderImportModeChoice({
      count: items.length,
      onNext: (mode) => importMaterials(items, mode, {
        label: 'xlsx',
        onDone: () => { if (document.querySelector('#mList')) loadMList(view.querySelector('#mSearch').value.trim()); }
      })
    });
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

  // ============ 版本查看 + 手动检测更新 ============
  function openVersionPanel() {
    openModal(`
      <h3>版本信息<button class="btn-close" id="vp_close">✕</button></h3>
      <div class="row" style="margin:6px 0 14px;">
        <div class="row-main">
          <div class="row-title">当前版本：${APP_VERSION}</div>
          <div class="row-sub">离线优先 PWA · 所有数据仅存本机浏览器</div>
        </div>
      </div>
      <p id="vpStatus" class="muted" style="font-size:13px;line-height:1.6;min-height:20px;margin:0 0 14px;">点击下方「检测更新」检查服务器是否有新版本。</p>
      <div class="modal-actions">
        <button class="btn ghost" id="vp_close2">关闭</button>
        <button class="btn primary" id="vp_check">检测更新</button>
      </div>
    `);
    modal.querySelector('#vp_close').onclick = closeModal;
    modal.querySelector('#vp_close2').onclick = closeModal;
    modal.querySelector('#vp_check').onclick = checkForUpdate;
  }

  async function checkForUpdate() {
    const status = modal.querySelector('#vpStatus');
    if (!('serviceWorker' in navigator)) {
      status.textContent = '当前环境不支持离线更新（Service Worker 不可用）。';
      return;
    }
    const reg = await navigator.serviceWorker.getRegistration();
    if (!reg) {
      status.textContent = '尚未注册 Service Worker，无法检测更新。';
      return;
    }
    const btn = modal.querySelector('#vp_check');
    btn.disabled = true; btn.textContent = '检测中…';
    status.textContent = '正在向服务器检查新版本…';
    const before = navigator.serviceWorker.controller;
    try { await reg.update(); } catch (e) { /* 忽略网络抖动，下面按 controller 是否变化判断 */ }
    // 等待新 SW 安装并激活（sw.js 内 skipWaiting 会立即接管），controllerchange 即代表有新版本生效
    await new Promise(res => {
      let done = false; const finish = () => { if (!done) { done = true; res(); } };
      navigator.serviceWorker.addEventListener('controllerchange', finish, { once: true });
      setTimeout(finish, 4000);
    });
    const after = navigator.serviceWorker.controller;
    const changed = after && after !== before;
    btn.disabled = false; btn.textContent = '检测更新';
    if (changed) {
      status.textContent = '发现新版本！点击下方按钮立即刷新以应用更新。';
      const acts = modal.querySelector('.modal-actions');
      acts.innerHTML = '<button class="btn ghost" id="vp_later">稍后</button><button class="btn primary" id="vp_reload">立即刷新</button>';
      modal.querySelector('#vp_later').onclick = closeModal;
      modal.querySelector('#vp_reload').onclick = () => location.reload();
    } else {
      status.textContent = '已是最新版本（' + APP_VERSION + '）。';
    }
  }

  async function boot() {
    applyTheme(localStorage.getItem('theme') || 'system');
    if (verBtn) { verBtn.textContent = APP_VERSION; verBtn.onclick = openVersionPanel; }
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
      }).catch(err => {
        console.error('Service Worker 注册失败（离线缓存将不可用）', err);
        toast('离线缓存不可用，本次使用需保持联网');
      });
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
