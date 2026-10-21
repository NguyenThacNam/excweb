/* Giao diện: mở file, tìm kiếm, sửa, chọn và xuất. */
(function () {
  'use strict';
  const C = window.WXCore;
  const S = window.WXStore;
  const $ = (sel) => document.querySelector(sel);

  const FS_OPEN = typeof window.showOpenFilePicker === 'function';
  const FS_DIR = typeof window.showDirectoryPicker === 'function';
  const FS_SAVE = typeof window.showSaveFilePicker === 'function';
  const XLSX_TYPES = [{ description: 'File Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }];
  const PAGE = 200;

  const state = {
    files: [],
    products: [],
    byPid: new Map(),
    query: '',
    filterFile: '',
    filtered: [],
    limit: PAGE,
    selected: new Map(), // pid -> true, giữ thứ tự chọn
    active: null,
    undo: [],
    lastClick: null,
  };
  let seq = 0;
  const uid = (p) => `${p}${++seq}`;

  // ================================================================== tiện ích
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function norm1(ch) {
    if (/\s/.test(ch)) return ' ';
    return ch.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  }
  // Tô sáng từ khoá, không phân biệt dấu
  function hl(text, tokens) {
    text = String(text == null ? '' : text);
    if (!tokens.length || !text) return esc(text);
    let n = '';
    const map = [];
    for (let i = 0; i < text.length; i++) {
      for (const ch of norm1(text[i])) { n += ch; map.push(i); }
    }
    const marked = new Uint8Array(text.length);
    for (const t of tokens) {
      let k = n.indexOf(t);
      while (k !== -1) {
        for (let j = k; j < k + t.length; j++) marked[map[j]] = 1;
        k = n.indexOf(t, k + t.length);
      }
    }
    let out = '', open = false;
    for (let i = 0; i < text.length; i++) {
      if (marked[i] && !open) { out += '<mark>'; open = true; }
      if (!marked[i] && open) { out += '</mark>'; open = false; }
      out += esc(text[i]);
    }
    return open ? out + '</mark>' : out;
  }
  function toast(msg, type, ms) {
    const el = document.createElement('div');
    el.className = 'toast' + (type ? ' ' + type : '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => el.remove(), ms || (type === 'error' ? 7000 : 3500));
    return el;
  }
  function fmtTime(t) {
    const d = new Date(t);
    const p = (x) => String(x).padStart(2, '0');
    return `${p(d.getHours())}:${p(d.getMinutes())} ${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
  }
  function stamp() {
    const d = new Date();
    const p = (x) => String(x).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  }
  function fmtSize(b) { return b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB'; }
  function download(data, name) {
    const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* bỏ qua */ } }
  async function safe(fn, fallback) { try { return await fn(); } catch (e) { return fallback; } }

  // ================================================================== mở file
  async function openFilesDialog() {
    if (!FS_OPEN) return $('#fileInput').click();
    let handles;
    try { handles = await window.showOpenFilePicker({ multiple: true, types: XLSX_TYPES, excludeAcceptAllOption: false }); }
    catch (e) { return; }
    await loadMany(handles.map((h) => ({ handle: h, name: h.name, path: h.name })));
    rememberRecent(handles.map((h) => ({ kind: 'file', handle: h })));
  }

  async function openFolderDialog() {
    if (!FS_DIR) return $('#folderInput').click();
    let dir;
    try { dir = await window.showDirectoryPicker({ mode: 'readwrite' }); }
    catch (e) { return; }
    await openDirectory(dir);
    rememberRecent([{ kind: 'dir', handle: dir }]);
  }

  async function openDirectory(dir) {
    const found = [];
    await walk(dir, dir.name, 0, found);
    if (!found.length) return toast(`Thư mục "${dir.name}" không có file .xlsx nào`, 'error');
    await loadMany(found);
  }
  async function walk(dir, path, depth, out) {
    for await (const entry of dir.values()) {
      if (entry.kind === 'directory') {
        if (depth < 3 && !entry.name.startsWith('.') && entry.name !== '_saoluu') await walk(entry, `${path}/${entry.name}`, depth + 1, out);
      } else if (/\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$')) {
        out.push({ handle: entry, name: entry.name, path: `${path}/${entry.name}` });
      } else if (/\.(xls|xlsm|xlsb)$/i.test(entry.name) && !entry.name.startsWith('~$')) {
        out.push({ unsupported: true, name: entry.name });
      }
    }
  }

  async function loadMany(list) {
    const unsupported = list.filter((x) => x.unsupported || !/\.xlsx$/i.test(x.name));
    const todo = list.filter((x) => !unsupported.includes(x));
    if (unsupported.length) {
      toast(`Bỏ qua ${unsupported.length} file không phải .xlsx (${unsupported.slice(0, 3).map((x) => x.name).join(', ')}${unsupported.length > 3 ? '…' : ''}). Mở bằng Excel rồi "Lưu thành" .xlsx để dùng được.`, 'error', 9000);
    }
    if (!todo.length) return;
    const progress = todo.length > 1 ? toast(`Đang đọc 0/${todo.length} file…`, '', 600000) : null;
    let ok = 0, n = 0;
    for (const item of todo) {
      n++;
      if (progress) progress.textContent = `Đang đọc ${n}/${todo.length}: ${item.name}`;
      try {
        if (await loadOne(item)) ok++;
      } catch (e) {
        console.error(e);
        toast(`Không đọc được "${item.name}": ${e.message || e}`, 'error');
      }
      await new Promise((r) => setTimeout(r, 0));
    }
    if (progress) progress.remove();
    rebuildIndex();
    renderAll();
    if (ok) toast(`Đã mở ${ok} file, ${state.products.length} sản phẩm`, 'ok');
  }

  async function loadOne(item) {
    let blob = item.file;
    if (!blob && item.handle) blob = await item.handle.getFile();
    for (const f of state.files) {
      if (item.handle && f.handle && await safe(() => f.handle.isSameEntry(item.handle), false)) return false;
      if (!item.handle && f.name === blob.name && f.size === blob.size && f.lastModified === blob.lastModified) return false;
    }
    const buffer = new Uint8Array(await blob.arrayBuffer());
    const prepared = await C.prepareXlsx(buffer, window.JSZip);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(prepared.buffer);
    const file = {
      id: uid('f'), name: item.name || blob.name, path: item.path || item.name || blob.name,
      handle: item.handle || null, wb, tables: [], changes: new Map(),
      origBuffer: buffer, backedUp: false, warnings: prepared.warnings,
      size: blob.size, lastModified: blob.lastModified,
    };
    wb.eachSheet((ws) => {
      if (ws.state && ws.state !== 'visible') return;
      file.tables.push(makeTable(file, ws));
    });
    state.files.push(file);
    return true;
  }

  // ================================================================== bảng & cấu hình cột
  function configKey(ws, auto) {
    const sig = auto && auto.signature;
    return 'wx-cot:' + (sig || `sheet:${C.normHeader(ws.name)}`);
  }
  function makeTable(file, ws) {
    const table = { id: uid('t'), file, ws, name: ws.name, t: null, status: 'none', products: [], key: '' };
    analyzeTable(table);
    return table;
  }
  function analyzeTable(table, override) {
    const auto = C.analyzeSheet(table.ws);
    table.key = configKey(table.ws, auto);
    const cfg = override || lsGet(table.key);
    let t = auto;
    if (cfg && cfg.ignore) { table.t = auto; table.status = 'off'; table.products = []; return; }
    if (cfg && cfg.headerRow) t = C.analyzeSheet(table.ws, { headerRow: cfg.headerRow, roles: cfg.roles || undefined });
    table.t = t;
    table.status = t ? 'ok' : 'none';
    table.products = t ? buildProducts(table) : [];
  }
  function buildProducts(table) {
    const t = table.t;
    const groupTitle = {};
    for (const r of t.rows) if (r.kind === 'group') groupTitle[r.row] = r.title;
    return t.rows.filter((r) => r.kind === 'product').map((r) => {
      const e = { pid: `${table.id}:${r.row}`, table, file: table.file, row: r.row, group: r.group, groupTitle: groupTitle[r.group] || '' };
      refreshEntry(e);
      return e;
    });
  }
  function refreshEntry(e) {
    const t = e.table.t;
    e.p = C.readProduct(e.table.ws, t, { row: e.row, group: e.group });
    const rt = (k) => C.roleText(t, e.p, k);
    const code = rt('code'), model = rt('model');
    e.code = code || model;
    e.model = code && model ? model : '';
    e.name = rt('name');
    const extra = C.extraSpecText(t, e.p);
    e.spec = rt('spec') || extra;
    e.unit = rt('unit');
    e.price = rt('price');
    e.brand = rt('brand');
    e.search = C.norm(Object.values(e.p.cells).join(' '));
    e.codeN = C.norm(code);
    e.modelN = C.norm(model);
    e.nameN = C.norm(e.name);
  }
  function rebuildIndex() {
    state.products = [];
    state.byPid = new Map();
    for (const f of state.files) for (const t of f.tables) for (const e of t.products) {
      state.products.push(e);
      state.byPid.set(e.pid, e);
    }
    for (const pid of [...state.selected.keys()]) if (!state.byPid.has(pid)) state.selected.delete(pid);
    if (state.active && !state.byPid.has(state.active)) state.active = null;
  }

  // ================================================================== tìm kiếm
  function runSearch() {
    const q = C.norm(state.query);
    const tokens = q ? q.split(' ').filter(Boolean) : [];
    let list = state.products;
    if (state.filterFile) list = list.filter((e) => e.file.id === state.filterFile);
    if (tokens.length) {
      const scored = [];
      for (const e of list) {
        if (!tokens.every((t) => e.search.includes(t))) continue;
        let s = 0;
        if (e.codeN === q || e.modelN === q) s += 100;
        else if ((e.codeN && e.codeN.startsWith(q)) || (e.modelN && e.modelN.startsWith(q))) s += 60;
        else if (tokens.some((t) => e.codeN.includes(t) || e.modelN.includes(t))) s += 30;
        if (e.nameN.includes(q)) s += 25;
        else if (tokens.every((t) => e.nameN.includes(t))) s += 12;
        scored.push({ e, s });
      }
      scored.sort((a, b) => b.s - a.s);
      list = scored.map((x) => x.e);
    }
    state.filtered = list;
    state.tokens = tokens;
  }

  // ================================================================== hiển thị
  function renderAll() {
    runSearch();
    renderFiles();
    renderResults();
    renderDrawer();
    renderCart();
    renderFilter();
  }

  function renderFilter() {
    const sel = $('#fileFilter');
    const cur = state.filterFile;
    if (cur && !state.files.some((f) => f.id === cur)) state.filterFile = '';
    sel.innerHTML = `<option value="">Tất cả file (${state.files.length})</option>` +
      state.files.map((f) => `<option value="${f.id}" ${f.id === state.filterFile ? 'selected' : ''}>${esc(f.name)}</option>`).join('');
    sel.hidden = state.files.length < 2;
  }

  function renderFiles() {
    const box = $('#fileList');
    $('#fileCount').textContent = state.files.length ? `(${state.files.length})` : '';
    if (!state.files.length) {
      box.innerHTML = '<p class="muted" style="padding:4px 6px;font-size:13px">Chưa mở file nào.</p>';
    } else {
      box.innerHTML = state.files.map((f) => {
        const dirty = f.changes.size;
        const sheets = f.tables.map((t) => {
          const info = t.status === 'ok' ? `${t.products.length} SP` : t.status === 'off' ? 'bỏ qua' : 'không nhận diện được';
          return `<div class="sheet ${t.status === 'ok' ? '' : t.status}">
            <span class="sname" title="${esc(t.name)}">${esc(t.name)} · <span class="muted">${info}</span></span>
            <button class="link" data-act="cols" data-tid="${t.id}">Cột</button>
          </div>`;
        }).join('');
        const warn = f.warnings.length
          ? `<div class="file-warn">File có ${esc(f.warnings.join(', '))}: có thể mất khi lưu đè. Bản gốc được tự sao lưu trước khi lưu.</div>` : '';
        return `<div class="file ${dirty ? 'dirty' : ''}" data-fid="${f.id}">
          <div class="file-head">
            <span class="fname" title="${esc(f.path)}">${esc(f.name)}</span>
            ${dirty ? `<span class="badge dirty">${dirty} ô chưa lưu</span>` : ''}
          </div>
          <div class="sheets">${sheets}</div>
          <div class="file-actions">
            <button class="btn small ${dirty ? 'primary' : ''}" data-act="save" ${dirty ? '' : 'disabled'}>Lưu</button>
            <button class="btn small ghost" data-act="close">Đóng</button>
          </div>${warn}
        </div>`;
      }).join('');
    }
    const anyDirty = state.files.some((f) => f.changes.size);
    $('#btnSaveAll').disabled = !anyDirty;
    $('#modeNote').innerHTML = FS_OPEN
      ? '<b>Lưu thẳng vào file gốc:</b> có (Edge/Chrome).'
      : '<b>Trình duyệt này không lưu đè được</b>: khi lưu sẽ tải file về. Nên dùng Edge hoặc Chrome.';
  }

  function isChanged(e) {
    for (const col of e.table.t.columns) if (e.file.changes.has(changeKey(e.table, e.row, col.col))) return true;
    return false;
  }

  function renderResults() {
    const has = state.files.length > 0;
    $('#empty').hidden = has;
    $('#resultsWrap').hidden = !has;
    const list = state.filtered;
    const shown = list.slice(0, state.limit);
    const tk = state.tokens || [];
    const multiFile = state.files.length > 1;
    $('#count').textContent = has ? (state.query || state.filterFile ? `${list.length} / ${state.products.length} sản phẩm` : `${state.products.length} sản phẩm`) : '';
    $('#tbody').innerHTML = shown.length ? shown.map((e) => {
      const sel = state.selected.has(e.pid);
      return `<tr data-pid="${e.pid}" class="${sel ? 'sel' : ''} ${state.active === e.pid ? 'active' : ''}">
        <td class="ck"><input type="checkbox" ${sel ? 'checked' : ''} aria-label="Chọn"></td>
        <td class="code"><b>${hl(e.code, tk)}</b>${e.model ? `<small>${hl(e.model, tk)}</small>` : ''}${isChanged(e) ? '<span class="dot-changed" title="Có ô đã sửa, chưa lưu"></span>' : ''}</td>
        <td class="name">${hl(e.name, tk)}${e.groupTitle ? `<small>${esc(e.groupTitle)}</small>` : ''}${e.brand ? `<small>${hl(e.brand, tk)}</small>` : ''}</td>
        <td class="spec"><div class="clamp">${hl(e.spec, tk)}</div></td>
        <td class="unit">${esc(e.unit)}</td>
        <td class="num">${esc(e.price)}</td>
        <td class="src" title="${esc(e.file.path)} › ${esc(e.table.name)} › dòng ${e.row}">${multiFile ? `<span class="fn">${esc(e.file.name)}</span>` : ''}<small>${esc(e.table.name)} · dòng ${e.row}</small></td>
      </tr>`;
    }).join('') : (has ? `<tr><td colspan="7" class="muted" style="text-align:center;padding:28px">Không tìm thấy sản phẩm nào khớp "${esc(state.query)}"</td></tr>` : '');
    const more = $('#btnMore');
    more.hidden = list.length <= state.limit;
    more.textContent = `Hiện thêm (còn ${list.length - state.limit})`;
    const ck = $('#ckAll');
    const selCount = list.filter((e) => state.selected.has(e.pid)).length;
    ck.checked = list.length > 0 && selCount === list.length;
    ck.indeterminate = selCount > 0 && selCount < list.length;
  }

  function renderCart() {
    const n = state.selected.size;
    $('#cartbar').hidden = n === 0;
    $('#cartCount').textContent = n;
    const files = new Set([...state.selected.keys()].map((pid) => state.byPid.get(pid).file.id));
    $('#cartFiles').textContent = files.size > 1 ? `· từ ${files.size} file` : '';
  }

  // ================================================================== ngăn chi tiết & sửa
  function changeKey(table, row, col) { return `${table.id}!${row}!${col}`; }

  function renderDrawer() {
    const d = $('#drawer');
    const e = state.active && state.byPid.get(state.active);
    if (!e) { d.hidden = true; d.innerHTML = ''; return; }
    d.hidden = false;
    const t = e.table.t;
    const fields = t.columns.map((col) => {
      const cell = e.table.ws.getCell(e.row, col.col);
      const key = changeKey(e.table, e.row, col.col);
      const ch = e.file.changes.get(key);
      const role = col.role ? `<span class="chip">${esc(C.ROLE_LABEL[col.role])}</span>` : '';
      const text = e.p.cells[col.col] || '';
      if (C.isSlave(cell)) return '';
      if (e.p.formulas[col.col]) {
        return `<div class="field"><label><span class="lbl">${esc(col.header)}</span>${role}</label>
          <div class="formula">${esc(text) || '&nbsp;'} <small>· công thức, Excel tự tính</small></div></div>`;
      }
      return `<div class="field ${ch ? 'changed' : ''}">
        <label for="fld-${col.col}"><span class="lbl">${esc(col.header)}</span>${role}
          ${ch ? `<button class="link" data-revert="${col.col}" title="Trả lại giá trị ban đầu">Hoàn tác ô này</button>` : ''}</label>
        <textarea id="fld-${col.col}" data-col="${col.col}" rows="1" spellcheck="false">${esc(text)}</textarea>
        ${ch ? `<div class="orig">Ban đầu: ${esc(C.valueText(ch.orig)) || '(trống)'}</div>` : ''}
      </div>`;
    }).join('');
    const sel = state.selected.has(e.pid);
    d.innerHTML = `
      <div class="drawer-head">
        <div style="flex:1;min-width:0">
          <div class="eyebrow">${esc(e.file.name)} › ${esc(e.table.name)} › dòng ${e.row}${e.groupTitle ? ' · ' + esc(e.groupTitle) : ''}</div>
          <h2>${esc(e.name || e.code || '(không tên)')}</h2>
        </div>
        <button class="icon-btn" data-act="close-drawer" title="Đóng (Esc)" aria-label="Đóng">×</button>
      </div>
      <div class="drawer-body">${fields}
        <p class="muted" style="font-size:12px">Sửa xong bấm ra ngoài ô là ghi nhận. Bấm <b>Lưu vào file</b> (hoặc <kbd>Ctrl</kbd>+<kbd>S</kbd>) để ghi vào file Excel.</p>
      </div>
      <div class="drawer-foot">
        <label><input type="checkbox" data-act="toggle-sel" ${sel ? 'checked' : ''}> Chọn để xuất</label>
        <button class="btn primary" data-act="save-active" ${e.file.changes.size ? '' : 'disabled'}>Lưu vào file</button>
      </div>`;
    d.querySelectorAll('textarea').forEach(autoGrow);
  }
  function patchDrawer(e) {
    const d = $('#drawer');
    for (const ta of d.querySelectorAll('textarea[data-col]')) {
      const col = Number(ta.dataset.col);
      const ch = e.file.changes.get(changeKey(e.table, e.row, col));
      const field = ta.closest('.field');
      field.classList.toggle('changed', !!ch);
      let orig = field.querySelector('.orig');
      let rv = field.querySelector('[data-revert]');
      if (ch) {
        if (!orig) { orig = document.createElement('div'); orig.className = 'orig'; field.appendChild(orig); }
        orig.textContent = 'Ban đầu: ' + (C.valueText(ch.orig) || '(trống)');
        if (!rv) {
          rv = document.createElement('button');
          rv.className = 'link'; rv.dataset.revert = col; rv.textContent = 'Hoàn tác ô này'; rv.title = 'Trả lại giá trị ban đầu';
          field.querySelector('label').appendChild(rv);
        }
      } else {
        if (orig) orig.remove();
        if (rv) rv.remove();
      }
      const shown = e.p.cells[col] || '';
      if (document.activeElement !== ta && ta.value !== shown) { ta.value = shown; autoGrow(ta); }
    }
    d.querySelector('h2').textContent = e.name || e.code || '(không tên)';
    const sb = d.querySelector('[data-act="save-active"]');
    if (sb) sb.disabled = !e.file.changes.size;
  }
  function autoGrow(ta) {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight + 2, 420) + 'px';
    ta.style.overflowY = ta.scrollHeight > 420 ? 'auto' : 'hidden';
  }

  function applyEdit(e, col, text, fromUndo) {
    const ws = e.table.ws;
    const key = changeKey(e.table, e.row, col);
    const cell = ws.getCell(e.row, col);
    const before = cell.value;
    const beforeCopy = before && typeof before === 'object' && !(before instanceof Date) ? JSON.parse(JSON.stringify(before)) : before;
    try { C.setCellText(ws, e.row, col, text); }
    catch (err) { toast(err.message, 'error'); return false; }
    if (!fromUndo) state.undo.push({ pid: e.pid, col, value: beforeCopy });
    trackChange(e, col, key, beforeCopy);
    return true;
  }
  function trackChange(e, col, key, before) {
    const now = e.table.ws.getCell(e.row, col).value;
    const ch = e.file.changes.get(key);
    if (!ch) e.file.changes.set(key, { orig: before, pid: e.pid, col, row: e.row, table: e.table });
    else if (C.valueText(ch.orig) === C.valueText(now)) e.file.changes.delete(key);
    refreshEntry(e);
  }

  function undo() {
    const u = state.undo.pop();
    if (!u) return toast('Không còn thao tác để hoàn tác');
    const e = state.byPid.get(u.pid);
    if (!e) return undo();
    const key = changeKey(e.table, e.row, u.col);
    const before = e.table.ws.getCell(e.row, u.col).value;
    C.restoreCellValue(e.table.ws, e.row, u.col, u.value);
    trackChange(e, u.col, key, before);
    state.active = e.pid;
    renderAll();
    toast(`Đã hoàn tác: ${e.name || e.code}`);
  }

  function revertCell(e, col) {
    const key = changeKey(e.table, e.row, col);
    const ch = e.file.changes.get(key);
    if (!ch) return;
    const before = e.table.ws.getCell(e.row, col).value;
    state.undo.push({ pid: e.pid, col, value: before });
    C.restoreCellValue(e.table.ws, e.row, col, ch.orig);
    e.file.changes.delete(key);
    refreshEntry(e);
    renderAll();
  }

  // ================================================================== lưu
  async function saveFile(file) {
    if (!file.changes.size) return true;
    let buf;
    try { buf = await file.wb.xlsx.writeBuffer(); }
    catch (e) { toast(`Lỗi tạo file "${file.name}": ${e.message}`, 'error'); return false; }
    if (!file.backedUp && S.available) {
      await safe(() => S.addBackup(file.name, file.origBuffer, 5));
      file.backedUp = true;
    }
    if (file.handle && file.handle.createWritable) {
      try {
        let perm = await file.handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') perm = await file.handle.requestPermission({ mode: 'readwrite' });
        if (perm !== 'granted') throw new Error('Bạn chưa cho phép sửa file');
        const w = await file.handle.createWritable();
        await w.write(buf);
        await w.close();
      } catch (e) {
        const locked = /lock|in use|NoModification|InvalidState|being used/i.test(String(e && (e.name + e.message)));
        toast(locked ? `Không lưu được "${file.name}": file đang mở trong Excel. Hãy đóng file trong Excel rồi lưu lại.`
          : `Không lưu được "${file.name}": ${e.message || e}`, 'error', 9000);
        return false;
      }
    } else {
      download(buf, file.name);
    }
    // ghi lịch sử
    const log = [...file.changes.values()].map((ch) => {
      const e = state.byPid.get(ch.pid);
      const col = ch.table.t.columns.find((c) => c.col === ch.col);
      return {
        time: Date.now(), file: file.name, sheet: ch.table.name, row: ch.row,
        product: e ? (e.name || e.code) : '', code: e ? e.code : '', column: col ? col.header : C.colLetter(ch.col),
        before: C.valueText(ch.orig), after: C.valueText(ch.table.ws.getCell(ch.row, ch.col).value),
      };
    });
    if (S.available) for (const l of log) await safe(() => S.put('history', l));
    file.changes.clear();
    state.undo = state.undo.filter((u) => { const e = state.byPid.get(u.pid); return e && e.file !== file; });
    toast(file.handle ? `Đã lưu "${file.name}"` : `Đã tải về "${file.name}" (trình duyệt không cho lưu đè)`, 'ok');
    return true;
  }
  async function saveAll() {
    for (const f of state.files.filter((x) => x.changes.size)) await saveFile(f);
    renderAll();
  }

  // ================================================================== chọn
  function toggleSelect(pid, on) {
    if (on == null) on = !state.selected.has(pid);
    if (on) state.selected.set(pid, true); else state.selected.delete(pid);
  }

  // ================================================================== hộp thoại
  function modal(title, body, buttons, opts) {
    const dlg = $('#modal');
    dlg.style.width = opts && opts.wide ? 'min(1000px, calc(100vw - 32px))' : '';
    dlg.innerHTML = `<div class="modal-inner">
      <div class="modal-head"><h3>${esc(title)}</h3><button class="icon-btn" data-close aria-label="Đóng">×</button></div>
      <div class="modal-body">${body}</div>
      <div class="modal-foot">${(buttons || []).map((b, i) => `<button class="btn ${b.cls || ''}" data-btn="${i}">${esc(b.label)}</button>`).join('')}</div>
    </div>`;
    dlg.querySelector('[data-close]').onclick = () => dlg.close();
    (buttons || []).forEach((b, i) => {
      dlg.querySelector(`[data-btn="${i}"]`).onclick = async () => {
        const r = b.onClick ? await b.onClick(dlg) : undefined;
        if (r !== false) dlg.close();
      };
    });
    if (!dlg.open) dlg.showModal();
    return dlg;
  }

  // ---------- cấu hình cột
  function openColumns(table) {
    const ws = table.ws;
    let headerRow = (table.t && table.t.headerRow) || 1;
    let roles = {};
    let preview = table.t;
    if (preview) for (const c of preview.columns) if (c.role) roles[c.col] = c.role;

    const draw = (dlg) => {
      const cols = preview ? preview.columns : [];
      const products = preview ? preview.rows.filter((r) => r.kind === 'product').slice(0, 3) : [];
      const sample = (col) => products.map((r) => C.valueText(ws.getCell(r.row, col).value)).filter(Boolean).join(' | ');
      const opts = (cur) => `<option value="">(không dùng)</option>` + C.ROLES.map((r) => `<option value="${r.key}" ${cur === r.key ? 'selected' : ''}>${esc(r.label)}</option>`).join('');
      dlg.querySelector('.modal-body').innerHTML = `
        <div class="note">Công cụ tự đoán dòng tiêu đề và ý nghĩa từng cột. Nếu đoán sai, chỉnh lại ở đây. Cấu hình được <b>nhớ</b> cho mọi file có cùng kiểu tiêu đề.</div>
        <div class="form-row" style="display:flex;gap:10px;align-items:end">
          <div style="width:160px"><label class="title" for="hdrRow">Dòng tiêu đề</label><input type="number" id="hdrRow" min="1" value="${headerRow}"></div>
          <button class="btn" id="btnReHdr">Đọc lại từ dòng này</button>
          <span class="muted" style="font-size:12.5px">${preview ? `Dữ liệu: dòng ${preview.dataStart}–${preview.lastDataRow}, ${preview.rows.filter((r) => r.kind === 'product').length} SP${preview.rows.some((r) => r.kind === 'group') ? `, ${preview.rows.filter((r) => r.kind === 'group').length} mục` : ''}` : 'Chưa tìm thấy bảng dữ liệu'}</span>
        </div>
        ${cols.length ? `<table class="grid"><thead><tr><th>Cột</th><th>Tiêu đề trong file</th><th style="width:200px">Ý nghĩa</th><th>Ví dụ dữ liệu</th></tr></thead><tbody>
          ${cols.map((c) => `<tr class="${roles[c.col] ? 'role-set' : ''}"><td>${C.colLetter(c.col)}</td><td>${esc(c.header)}</td>
            <td><select data-col="${c.col}">${opts(roles[c.col])}</select></td><td class="sample" title="${esc(sample(c.col))}">${esc(sample(c.col))}</td></tr>`).join('')}
        </tbody></table>
        <p class="muted" style="font-size:12.5px">Cột "(không dùng)" vẫn hiện khi xem chi tiết và vẫn được giữ khi xuất file. Cột nào cần tìm kiếm theo tên/mã thì nên đặt ý nghĩa.</p>` : ''}`;
      dlg.querySelector('#btnReHdr').onclick = () => {
        headerRow = Math.max(1, Number(dlg.querySelector('#hdrRow').value) || 1);
        preview = C.analyzeSheet(ws, { headerRow });
        roles = {};
        if (preview) for (const c of preview.columns) if (c.role) roles[c.col] = c.role;
        draw(dlg);
      };
      dlg.querySelectorAll('select[data-col]').forEach((s) => {
        s.onchange = () => {
          const col = Number(s.dataset.col);
          if (s.value) for (const k of Object.keys(roles)) if (roles[k] === s.value) delete roles[k];
          if (s.value) roles[col] = s.value; else delete roles[col];
          preview = C.analyzeSheet(ws, { headerRow, roles: Object.assign({}, roles) });
          draw(dlg);
        };
      });
    };

    const dlg = modal(`Cột của sheet "${table.name}" · ${table.file.name}`, '', [
      { label: table.status === 'off' ? 'Dùng lại sheet này' : 'Bỏ qua sheet này', cls: 'ghost left', onClick: () => {
        if (table.status === 'off') { lsSet(table.key, null); analyzeTable(table); }
        else { lsSet(table.key, { ignore: true }); analyzeTable(table, { ignore: true }); }
        afterTableChange();
      } },
      { label: 'Huỷ' },
      { label: 'Lưu cấu hình', cls: 'primary', onClick: () => {
        if (table.file.changes.size && [...table.file.changes.values()].some((c) => c.table === table)) {
          toast('Sheet này có ô đã sửa chưa lưu. Hãy lưu file trước khi đổi cấu hình cột.', 'error');
          return false;
        }
        const cfg = { headerRow, roles: Object.assign({}, roles) };
        lsSet(table.key, cfg);
        analyzeTable(table, cfg);
        afterTableChange();
        // áp dụng cho các sheet cùng kiểu đang mở
        for (const f of state.files) for (const t of f.tables) if (t !== table && t.key === table.key) analyzeTable(t, cfg);
        afterTableChange();
        toast('Đã lưu cấu hình cột', 'ok');
      } },
    ], { wide: true });
    draw(dlg);
  }
  function afterTableChange() {
    rebuildIndex();
    renderAll();
  }

  // ---------- giỏ
  function cartGroups() {
    const groups = new Map();
    for (const pid of state.selected.keys()) {
      const e = state.byPid.get(pid);
      if (!e) continue;
      if (!groups.has(e.table)) groups.set(e.table, []);
      groups.get(e.table).push(e);
    }
    return groups;
  }
  function openCart() {
    const draw = (dlg) => {
      const groups = cartGroups();
      dlg.querySelector('.modal-body').innerHTML = groups.size ? [...groups].map(([t, list]) => `
        <div class="cart-group"><h4>${esc(t.file.name)} › ${esc(t.name)} <span class="muted">(${list.length})</span></h4>
        <table class="grid"><tbody>${list.sort((a, b) => a.row - b.row).map((e) => `<tr>
          <td style="width:140px"><b>${esc(e.code)}</b></td><td>${esc(e.name)}</td><td class="muted" style="width:80px">dòng ${e.row}</td>
          <td style="width:60px;text-align:right"><button class="link" data-rm="${e.pid}">Bỏ</button></td></tr>`).join('')}</tbody></table></div>`).join('')
        : '<p class="muted">Chưa chọn sản phẩm nào.</p>';
      dlg.querySelectorAll('[data-rm]').forEach((b) => { b.onclick = () => { state.selected.delete(b.dataset.rm); renderResults(); renderCart(); draw(dlg); }; });
    };
    const dlg = modal(`Sản phẩm đã chọn (${state.selected.size})`, '', [
      { label: 'Đóng' },
      { label: 'Xuất file Excel', cls: 'primary', onClick: () => { setTimeout(openExport, 0); } },
    ]);
    draw(dlg);
  }

  // ---------- xuất
  function openExport() {
    const groups = cartGroups();
    if (!groups.size) return toast('Hãy tích chọn sản phẩm trước');
    const tables = [...groups.keys()];
    const single = tables.length === 1;
    let template = tables.slice().sort((a, b) => groups.get(b).length - groups.get(a).length)[0];
    const allTables = [];
    for (const f of state.files) for (const t of f.tables) if (t.status === 'ok') allTables.push(t);
    const total = state.selected.size;

    const body = () => {
      const t = template;
      const hasGroups = t.t.rows.some((r) => r.kind === 'group');
      const sheetCount = t.file.wb.worksheets.length;
      const base = t.file.name.replace(/\.xlsx$/i, '');
      return `
        ${single
          ? `<div class="note">Xuất <b>${total}</b> sản phẩm từ <b>${esc(t.file.name)}</b> › ${esc(t.name)}. File mới giữ nguyên định dạng, phần đầu và phần cuối (tổng cộng, chữ ký…) của file gốc. Công thức tổng được tự tính lại.</div>`
          : `<div class="note">Các sản phẩm đến từ <b>${tables.length}</b> sheet/file khác nhau. Chọn một file làm <b>mẫu định dạng</b>. Sản phẩm từ file khác được ghép theo cột tương ứng (Mã, Tên, Thông số…). Thông tin mà mẫu không có cột sẽ được ghép vào ô Thông số kỹ thuật.</div>
             <div class="form-row"><label class="title" for="tplSel">File mẫu định dạng</label>
             <select id="tplSel">${allTables.map((x) => `<option value="${x.id}" ${x === t ? 'selected' : ''}>${esc(x.file.name)} › ${esc(x.name)}${groups.has(x) ? ` (${groups.get(x).length} SP đã chọn)` : ''}</option>`).join('')}</select></div>`}
        <div class="form-row">
          ${t.t.roleCol.stt ? '<label class="check"><input type="checkbox" id="optRenumber" checked> Đánh lại số thứ tự (STT)</label>' : ''}
          ${single && hasGroups ? '<label class="check"><input type="checkbox" id="optGroups" checked> Giữ dòng tiêu đề mục (I, II, III…) của các sản phẩm được chọn</label>' : ''}
          ${sheetCount > 1 ? `<label class="check"><input type="checkbox" id="optOnlySheet" checked> Chỉ giữ sheet "${esc(t.name)}" (xoá ${sheetCount - 1} sheet khác trong file xuất)</label>` : ''}
          <label class="check"><input type="checkbox" id="optOrder" checked> Sắp xếp theo thứ tự trong file gốc (bỏ chọn = theo thứ tự đã tích)</label>
        </div>
        <div class="form-row"><label class="title" for="outName">Tên file xuất</label>
          <input type="text" id="outName" value="${esc(base)}_chon-${total}sp_${stamp()}.xlsx"></div>
        ${t.file.warnings.length ? `<div class="note warn">File mẫu có ${esc(t.file.warnings.join(', '))}: phần này có thể không giữ được trong file xuất.</div>` : ''}`;
    };
    const dlg = modal('Xuất file Excel', body(), [
      { label: 'Huỷ' },
      { label: 'Xuất file', cls: 'primary', onClick: async (d) => {
        const opts = {
          renumber: !!(d.querySelector('#optRenumber') || {}).checked,
          groups: !!(d.querySelector('#optGroups') || {}).checked,
          onlySheet: !!(d.querySelector('#optOnlySheet') || {}).checked,
          keepOrder: !!(d.querySelector('#optOrder') || {}).checked,
          name: (d.querySelector('#outName').value.trim() || 'xuat.xlsx').replace(/[\\/:*?"<>|]/g, '-'),
        };
        if (!/\.xlsx$/i.test(opts.name)) opts.name += '.xlsx';
        const btn = d.querySelector('.modal-foot .primary');
        btn.disabled = true; btn.textContent = 'Đang tạo file…';
        try { await doExport(template, groups, opts); }
        catch (e) { console.error(e); toast(`Lỗi xuất file: ${e.message || e}`, 'error'); btn.disabled = false; btn.textContent = 'Xuất file'; return false; }
      } },
    ]);
    const bind = () => {
      const sel = dlg.querySelector('#tplSel');
      if (sel) sel.onchange = () => { template = allTables.find((x) => x.id === sel.value); dlg.querySelector('.modal-body').innerHTML = body(); bind(); };
    };
    bind();
  }

  async function doExport(template, groups, opts) {
    const file = template.file;
    const buf = await file.wb.xlsx.writeBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet(template.name);
    const cfg = lsGet(template.key);
    const t = C.analyzeSheet(ws, cfg && cfg.headerRow ? { headerRow: cfg.headerRow, roles: cfg.roles } : undefined);
    if (!t) throw new Error('Không đọc lại được bảng của file mẫu');

    const order = [...state.selected.keys()];
    const own = (groups.get(template) || []).slice();
    if (opts.keepOrder) own.sort((a, b) => a.row - b.row);
    else own.sort((a, b) => order.indexOf(a.pid) - order.indexOf(b.pid));
    const items = [];
    let lastGroup = null;
    for (const e of own) {
      if (opts.groups && e.group && e.group !== lastGroup) { items.push({ type: 'group', row: e.group }); lastGroup = e.group; }
      items.push({ type: 'row', row: e.row });
    }
    for (const [table, list] of groups) {
      if (table === template) continue;
      const sorted = list.slice();
      if (opts.keepOrder) sorted.sort((a, b) => a.row - b.row);
      else sorted.sort((a, b) => order.indexOf(a.pid) - order.indexOf(b.pid));
      for (const e of sorted) items.push(C.makeMapItem(table.t, e.p, t));
    }
    C.buildExport(wb, t, items, { renumber: opts.renumber });
    if (opts.onlySheet) for (const other of wb.worksheets.slice()) if (other.name !== template.name) wb.removeWorksheet(other.id);
    const out = await wb.xlsx.writeBuffer();

    if (FS_SAVE) {
      let h;
      try { h = await window.showSaveFilePicker({ suggestedName: opts.name, types: XLSX_TYPES }); }
      catch (e) { if (e && e.name === 'AbortError') return; throw e; }
      const w = await h.createWritable();
      await w.write(out);
      await w.close();
      toast(`Đã xuất "${h.name}" (${items.filter((i) => i.type !== 'group').length} SP)`, 'ok');
    } else {
      download(out, opts.name);
      toast(`Đã tải về "${opts.name}"`, 'ok');
    }
  }

  // ---------- lịch sử, sao lưu, hướng dẫn
  async function openHistory() {
    const pending = [];
    for (const f of state.files) for (const ch of f.changes.values()) {
      const e = state.byPid.get(ch.pid);
      const col = ch.table.t.columns.find((c) => c.col === ch.col);
      pending.push({ file: f.name, product: e ? (e.name || e.code) : '', column: col ? col.header : '', before: C.valueText(ch.orig), after: C.valueText(ch.table.ws.getCell(ch.row, ch.col).value), row: ch.row });
    }
    const saved = S.available ? (await safe(() => S.all('history'), [])).sort((a, b) => b.time - a.time).slice(0, 300) : [];
    const rowHtml = (h, time) => `<tr><td class="muted" style="white-space:nowrap">${time}</td><td>${esc(h.file)}<div class="sub">dòng ${h.row}</div></td><td>${esc(h.product)}<div class="sub">${esc(h.column)}</div></td>
      <td><span class="old">${esc(h.before) || '(trống)'}</span><br><span class="new">${esc(h.after) || '(trống)'}</span></td></tr>`;
    const table = (rows) => `<table class="grid"><thead><tr><th>Thời gian</th><th>File</th><th>Sản phẩm / Cột</th><th>Cũ → Mới</th></tr></thead><tbody>${rows}</tbody></table>`;
    modal('Lịch sử sửa', `
      ${pending.length ? `<h4 style="margin:6px 0">Chưa lưu (${pending.length})</h4>${table(pending.map((h) => rowHtml(h, 'chưa lưu')).join(''))}` : ''}
      <h4 style="margin:14px 0 6px">Đã lưu ${saved.length ? `(${saved.length} gần nhất)` : ''}</h4>
      ${saved.length ? table(saved.map((h) => rowHtml(h, fmtTime(h.time))).join('')) : '<p class="muted">Chưa có.</p>'}`,
    [
      ...(saved.length ? [{ label: 'Xoá lịch sử', cls: 'ghost left danger', onClick: async () => { if (confirm('Xoá toàn bộ lịch sử sửa đã lưu?')) await safe(() => S.clear('history')); } }] : []),
      { label: 'Đóng' },
    ], { wide: true });
  }

  async function openBackups() {
    const list = S.available ? (await safe(() => S.all('backups'), [])).sort((a, b) => b.time - a.time) : [];
    const dlg = modal('Bản sao lưu', `
      <div class="note">Trước lần lưu đầu tiên của mỗi file, công cụ tự cất bản gốc ở đây (trong trình duyệt của máy này, tối đa 5 bản mỗi file). Lỡ sửa nhầm thì tải bản gốc về.</div>
      ${list.length ? `<table class="grid" style="margin-top:12px"><thead><tr><th>File</th><th>Thời điểm sao lưu</th><th>Dung lượng</th><th></th></tr></thead><tbody>
        ${list.map((b) => `<tr><td>${esc(b.name)}</td><td>${fmtTime(b.time)}</td><td>${fmtSize(b.size)}</td><td style="text-align:right"><button class="link" data-dl="${b.id}">Tải về</button></td></tr>`).join('')}
      </tbody></table>` : '<p class="muted" style="margin-top:12px">Chưa có bản sao lưu nào.</p>'}`, [{ label: 'Đóng' }], { wide: true });
    dlg.querySelectorAll('[data-dl]').forEach((b) => {
      b.onclick = () => {
        const item = list.find((x) => String(x.id) === b.dataset.dl);
        if (item) download(item.data, item.name.replace(/\.xlsx$/i, '') + `_saoluu_${stamp()}.xlsx`);
      };
    });
  }

  function openHelp() {
    modal('Hướng dẫn nhanh', `<div class="help">
      <h4>1. Mở file</h4>
      <p>Bấm <b>Mở file Excel</b> (chọn được nhiều file) hoặc <b>Mở thư mục</b> (đọc mọi file .xlsx trong thư mục và thư mục con). Có thể kéo thả file vào cửa sổ. Lần sau bấm <b>Mở lại lần trước</b>.</p>
      <h4>2. Tìm kiếm</h4>
      <p>Gõ tên, mã, model hoặc thông số vào ô tìm kiếm, không cần dấu. Ví dụ <i>nguon 12v</i>, <i>NG-5V</i>, <i>schneider</i>. Tìm trên tất cả các file đang mở.</p>
      <h4>3. Sửa thông số</h4>
      <p>Bấm vào một dòng để mở ngăn chi tiết bên phải. Sửa ô, bấm ra ngoài là ghi nhận (ô đổi màu vàng). Bấm <b>Lưu vào file</b> hoặc <kbd>Ctrl</kbd>+<kbd>S</kbd> để ghi vào file Excel gốc. Trước lần lưu đầu, bản gốc được tự sao lưu (nút <b>Bản sao lưu</b>).</p>
      <p>Lưu ý: hãy <b>đóng file trong Excel</b> trước khi lưu, nếu không Windows sẽ khoá file.</p>
      <h4>4. Chọn và xuất</h4>
      <p>Tích ô vuông đầu dòng để chọn (giữ <kbd>Shift</kbd> để chọn cả đoạn). Bấm <b>Xuất file Excel</b>. File mới giữ nguyên định dạng file gốc, STT và công thức tổng được tự chỉnh. Chọn từ nhiều file thì chọn một file làm mẫu, sản phẩm được ghép theo cột tương ứng.</p>
      <h4>5. Cột bị nhận sai?</h4>
      <p>Bấm <b>Cột</b> cạnh tên sheet ở cột trái để chọn lại dòng tiêu đề và ý nghĩa từng cột. Cấu hình được nhớ cho các file cùng kiểu.</p>
      <h4>Phím tắt</h4>
      <p><kbd>/</kbd> tìm kiếm · <kbd>Ctrl</kbd>+<kbd>S</kbd> lưu tất cả · <kbd>Ctrl</kbd>+<kbd>Z</kbd> hoàn tác (khi không gõ trong ô) · <kbd>Esc</kbd> đóng ngăn chi tiết</p>
    </div>`, [{ label: 'Đã hiểu', cls: 'primary' }]);
  }

  // ---------- mở lại lần trước
  async function rememberRecent(items) {
    if (!S.available) return;
    for (const it of items) await safe(() => S.put('recent', { key: `${it.kind}:${it.handle.name}`, kind: it.kind, handle: it.handle, name: it.handle.name, time: Date.now() }));
    const all = (await safe(() => S.all('recent'), [])).sort((a, b) => b.time - a.time);
    for (const old of all.slice(20)) await safe(() => S.del('recent', old.key));
    updateReopen();
  }
  async function updateReopen() {
    if (!S.available || !FS_OPEN) return;
    const all = await safe(() => S.all('recent'), []);
    const btn = $('#btnReopen');
    btn.hidden = !all.length || state.files.length > 0;
    btn.title = all.map((r) => r.name).join('\n');
  }
  async function reopen() {
    const all = (await safe(() => S.all('recent'), [])).sort((a, b) => b.time - a.time);
    const files = [];
    for (const r of all) {
      try {
        let perm = await r.handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') perm = await r.handle.requestPermission({ mode: 'readwrite' });
        if (perm !== 'granted') continue;
        if (r.kind === 'dir') await walk(r.handle, r.handle.name, 0, files);
        else files.push({ handle: r.handle, name: r.handle.name, path: r.handle.name });
      } catch (e) {
        await safe(() => S.del('recent', r.key));
      }
    }
    if (files.length) await loadMany(files);
    else toast('Không mở lại được (file đã bị xoá/di chuyển hoặc chưa cấp quyền)', 'error');
    updateReopen();
  }

  // ================================================================== sự kiện
  function bind() {
    $('#btnOpenFiles').onclick = openFilesDialog;
    $('#btnOpenFolder').onclick = openFolderDialog;
    $('#btnReopen').onclick = reopen;
    $('#btnSaveAll').onclick = saveAll;
    $('#btnHistory').onclick = openHistory;
    $('#btnBackups').onclick = openBackups;
    $('#btnHelp').onclick = openHelp;
    $('#btnCartView').onclick = openCart;
    $('#btnCartClear').onclick = () => { state.selected.clear(); renderResults(); renderCart(); renderDrawer(); };
    $('#btnExport').onclick = openExport;
    $('#btnMore').onclick = () => { state.limit += PAGE; renderResults(); };

    $('#fileInput').onchange = async (ev) => {
      const files = [...ev.target.files];
      ev.target.value = '';
      await loadMany(files.map((f) => ({ file: f, name: f.name, path: f.name })));
    };
    $('#folderInput').onchange = async (ev) => {
      const files = [...ev.target.files].filter((f) => !f.name.startsWith('~$') && !/(^|\/)_saoluu\//.test(f.webkitRelativePath));
      ev.target.value = '';
      await loadMany(files.map((f) => ({ file: f, name: f.name, path: f.webkitRelativePath || f.name })));
    };

    let timer;
    $('#q').addEventListener('input', (ev) => {
      clearTimeout(timer);
      timer = setTimeout(() => { state.query = ev.target.value; state.limit = PAGE; runSearch(); renderResults(); }, 90);
    });
    $('#q').addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' && state.filtered.length) { state.active = state.filtered[0].pid; renderResults(); renderDrawer(); }
    });
    $('#fileFilter').onchange = (ev) => { state.filterFile = ev.target.value; state.limit = PAGE; runSearch(); renderResults(); };

    $('#ckAll').onchange = (ev) => {
      for (const e of state.filtered) toggleSelect(e.pid, ev.target.checked);
      renderResults(); renderCart(); renderDrawer();
    };

    $('#tbody').addEventListener('click', (ev) => {
      const tr = ev.target.closest('tr[data-pid]');
      if (!tr) return;
      const pid = tr.dataset.pid;
      if (ev.target.matches('input[type=checkbox]') || ev.target.closest('td.ck')) {
        const idx = state.filtered.findIndex((e) => e.pid === pid);
        const on = !state.selected.has(pid);
        if (ev.shiftKey && state.lastClick != null) {
          const a = Math.min(idx, state.lastClick), b = Math.max(idx, state.lastClick);
          for (let i = a; i <= b; i++) toggleSelect(state.filtered[i].pid, on);
        } else toggleSelect(pid, on);
        state.lastClick = idx;
        renderResults(); renderCart(); renderDrawer();
        return;
      }
      state.active = state.active === pid ? null : pid;
      renderResults(); renderDrawer();
    });

    $('#fileList').addEventListener('click', async (ev) => {
      const btn = ev.target.closest('[data-act]');
      if (!btn) return;
      const fileEl = ev.target.closest('[data-fid]');
      const file = fileEl && state.files.find((f) => f.id === fileEl.dataset.fid);
      if (btn.dataset.act === 'cols') {
        for (const f of state.files) for (const t of f.tables) if (t.id === btn.dataset.tid) openColumns(t);
      } else if (btn.dataset.act === 'save' && file) {
        await saveFile(file); renderAll();
      } else if (btn.dataset.act === 'close' && file) {
        if (file.changes.size && !confirm(`"${file.name}" có ${file.changes.size} ô đã sửa chưa lưu. Vẫn đóng và bỏ các thay đổi?`)) return;
        state.files = state.files.filter((f) => f !== file);
        state.undo = state.undo.filter((u) => { const e = state.byPid.get(u.pid); return e && e.file !== file; });
        rebuildIndex(); renderAll(); updateReopen();
      }
    });

    const drawer = $('#drawer');
    drawer.addEventListener('input', (ev) => { if (ev.target.matches('textarea')) autoGrow(ev.target); });
    drawer.addEventListener('change', (ev) => {
      if (!ev.target.matches('textarea[data-col]')) return;
      const e = state.byPid.get(state.active);
      if (!e) return;
      const col = Number(ev.target.dataset.col);
      if ((e.p.cells[col] || '') === ev.target.value) return;
      // không vẽ lại cả ngăn chi tiết: tránh mất cú bấm vào nút "Lưu" ngay sau khi rời ô
      if (applyEdit(e, col, ev.target.value)) { renderResults(); renderFiles(); patchDrawer(e); }
    });
    drawer.addEventListener('keydown', (ev) => {
      if (ev.target.matches('textarea') && ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) { ev.preventDefault(); ev.target.blur(); }
    });
    drawer.addEventListener('click', async (ev) => {
      const e = state.byPid.get(state.active);
      const rv = ev.target.closest('[data-revert]');
      if (rv && e) return revertCell(e, Number(rv.dataset.revert));
      const btn = ev.target.closest('[data-act]');
      if (!btn) return;
      if (btn.dataset.act === 'close-drawer') { state.active = null; renderResults(); renderDrawer(); }
      if (btn.dataset.act === 'toggle-sel' && e) { toggleSelect(e.pid, btn.checked); renderResults(); renderCart(); }
      if (btn.dataset.act === 'save-active' && e) {
        if (document.activeElement && document.activeElement.matches('textarea')) document.activeElement.blur();
        await saveFile(e.file); renderAll();
      }
    });

    document.addEventListener('keydown', (ev) => {
      const typing = ev.target.matches('input, textarea, select');
      if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 's') {
        ev.preventDefault();
        if (document.activeElement && document.activeElement.matches('textarea')) document.activeElement.blur();
        setTimeout(saveAll, 0);
      } else if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 'z' && !typing) {
        ev.preventDefault(); undo();
      } else if (ev.key === '/' && !typing) {
        ev.preventDefault(); $('#q').focus(); $('#q').select();
      } else if (ev.key === 'Escape' && !$('#modal').open && state.active) {
        state.active = null; renderResults(); renderDrawer();
      }
    });

    window.addEventListener('beforeunload', (ev) => {
      if (state.files.some((f) => f.changes.size)) { ev.preventDefault(); ev.returnValue = ''; }
    });

    // kéo thả
    let dragDepth = 0;
    document.addEventListener('dragenter', (ev) => { if (ev.dataTransfer && [...ev.dataTransfer.types].includes('Files')) { dragDepth++; $('#dropOverlay').hidden = false; } });
    document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('#dropOverlay').hidden = true; } });
    document.addEventListener('dragover', (ev) => ev.preventDefault());
    document.addEventListener('drop', async (ev) => {
      ev.preventDefault();
      dragDepth = 0; $('#dropOverlay').hidden = true;
      const items = [...(ev.dataTransfer.items || [])].filter((i) => i.kind === 'file');
      const list = [];
      const remember = [];
      if (items.length && typeof items[0].getAsFileSystemHandle === 'function') {
        const handles = await Promise.all(items.map((i) => i.getAsFileSystemHandle()));
        for (const h of handles) {
          if (!h) continue;
          if (h.kind === 'directory') { await walk(h, h.name, 0, list); remember.push({ kind: 'dir', handle: h }); }
          else { list.push({ handle: h, name: h.name, path: h.name }); remember.push({ kind: 'file', handle: h }); }
        }
      } else {
        for (const f of ev.dataTransfer.files) list.push({ file: f, name: f.name, path: f.name });
      }
      await loadMany(list);
      if (remember.length) rememberRecent(remember);
    });
  }

  bind();
  renderAll();
  updateReopen();
  window.__wx = state; // để kiểm thử
})();
