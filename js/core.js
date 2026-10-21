/*
 * Lõi xử lý Excel: nhận diện bảng sản phẩm, sửa ô, xuất file giữ định dạng.
 * Chạy được cả trên trình duyệt (window.WXCore) lẫn Node (module.exports) để kiểm thử.
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------------ chữ
  function norm(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .toLowerCase().replace(/\s+/g, ' ').trim();
  }
  function normHeader(s) {
    return norm(s).replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function valueText(v) {
    if (v == null) return '';
    if (v instanceof Date) {
      const d = v.getUTCDate(), m = v.getUTCMonth() + 1, y = v.getUTCFullYear();
      return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
    }
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map((t) => t.text).join('');
      if ('formula' in v || 'sharedFormula' in v) return valueText(v.result);
      if ('text' in v) return String(v.text);
      if ('error' in v) return String(v.error);
      return '';
    }
    return String(v);
  }
  function primitive(v) {
    if (v == null) return null;
    if (typeof v === 'object' && !(v instanceof Date)) {
      if ('formula' in v || 'sharedFormula' in v) return primitive(v.result);
      return valueText(v);
    }
    return v;
  }
  function isFormulaValue(v) {
    return v != null && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v);
  }
  function formatNumber(n) {
    if (typeof n !== 'number' || !isFinite(n)) return String(n);
    const r = Math.round(n * 100) / 100;
    return r.toLocaleString('vi-VN');
  }
  function cellDisplay(cell) {
    const p = primitive(cell.value);
    if (typeof p === 'number') {
      const fmt = cell.numFmt || '';
      if (/[#0]/.test(fmt) || Math.abs(p) >= 10000) return formatNumber(p);
      return String(p);
    }
    return valueText(cell.value);
  }
  // "1.250.000" | "1,250,000" | "1250000" | "12,5" -> số; còn lại null
  function parseNumber(text) {
    const t = String(text).trim().replace(/\s/g, '');
    if (!/^-?[\d.,]+$/.test(t)) return null;
    let s = t;
    if (/^-?\d{1,3}([.]\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
    const n = Number(s);
    return isFinite(n) ? n : null;
  }

  // ------------------------------------------------------------------ vai trò cột
  const ROLES = [
    { key: 'stt', label: 'STT', kw: ['stt', 'tt', 'so tt', 'so thu tu', 'no'] },
    { key: 'code', label: 'Mã SP', kw: ['ma', 'ma sp', 'ma hang', 'ma san pham', 'ma vt', 'ma vat tu', 'ma thiet bi', 'ma hang hoa', 'code', 'part no', 'part number', 'sku'] },
    { key: 'model', label: 'Model / Ký mã hiệu', kw: ['model', 'ky ma hieu', 'ky hieu', 'ma hieu', 'kieu', 'type'] },
    { key: 'name', label: 'Tên thiết bị', kw: ['ten', 'ten thiet bi', 'ten hang', 'ten hang hoa', 'ten san pham', 'ten vat tu', 'danh muc', 'danh muc thiet bi', 'danh muc hang hoa', 'noi dung', 'hang muc', 'description', 'item', 'product name'] },
    { key: 'spec', label: 'Thông số kỹ thuật', kw: ['thong so', 'thong so ky thuat', 'tskt', 'yeu cau ky thuat', 'dac tinh ky thuat', 'dac tinh', 'cau hinh', 'mo ta', 'mo ta ky thuat', 'quy cach', 'specification', 'specifications', 'technical'] },
    { key: 'unit', label: 'ĐVT', kw: ['dvt', 'don vi', 'don vi tinh', 'dv', 'unit', 'uom'] },
    { key: 'qty', label: 'Số lượng', kw: ['so luong', 'sl', 'khoi luong', 'kl', 'qty', 'quantity'] },
    { key: 'price', label: 'Đơn giá', kw: ['don gia', 'gia', 'gia ban', 'gia ban le', 'gia niem yet', 'unit price', 'price'] },
    { key: 'amount', label: 'Thành tiền', kw: ['thanh tien', 'tong tien', 'gia tri', 'amount', 'total'] },
    { key: 'brand', label: 'Hãng', kw: ['hang', 'hang sx', 'hang san xuat', 'nha san xuat', 'thuong hieu', 'brand', 'manufacturer'] },
    { key: 'origin', label: 'Xuất xứ', kw: ['xuat xu', 'nuoc sx', 'nuoc san xuat', 'origin', 'country'] },
    { key: 'note', label: 'Ghi chú', kw: ['ghi chu', 'note', 'notes', 'remark'] },
  ];
  const ROLE_LABEL = Object.fromEntries(ROLES.map((r) => [r.key, r.label]));

  function guessRole(headerText) {
    const h = normHeader(headerText);
    if (!h || h.length > 60) return null;
    let best = null, bestScore = 0;
    const padded = ' ' + h + ' ';
    for (const role of ROLES) {
      for (const kw of role.kw) {
        let score = 0;
        if (h === kw) score = 100 + kw.length;
        else if (padded.includes(' ' + kw + ' ')) score = kw.length;
        if (score > bestScore) { bestScore = score; best = role.key; }
      }
    }
    return best;
  }

  const FOOTER_KW = ['tong cong', 'cong tien hang', 'cong', 'tong', 'tong so', 'thue', 'vat', 'bang chu', 'so tien bang chu',
    'nguoi lap', 'nguoi lap bieu', 'dai dien', 'giam doc', 'ky ten', 'ke toan', 'thu truong', 'xac nhan'];
  function isFooterText(t) {
    const n = normHeader(t);
    return FOOTER_KW.some((k) => n === k || n.startsWith(k + ' '));
  }

  // ------------------------------------------------------------------ đọc ô
  function isSlave(cell) {
    return cell.isMerged && cell.master && cell.master.address !== cell.address;
  }
  function rowTexts(ws, r, maxCol) {
    const out = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      out[c] = isSlave(cell) ? '' : valueText(cell.value).trim();
    }
    return out;
  }
  function mergeSpan(ws, r, c) {
    const cell = ws.getCell(r, c);
    if (!cell.isMerged) return 1;
    const m = cell.master;
    let n = 0;
    for (let cc = 1; cc <= ws.columnCount; cc++) {
      const x = ws.getCell(r, cc);
      if (x.isMerged && x.master.address === m.address) n++;
    }
    return n;
  }

  function isNumberingRow(texts) {
    const vals = texts.filter((t) => t);
    if (vals.length < 3) return false;
    return vals.every((t) => /^\(?[0-9]+\)?$/.test(t) || /^\([a-z0-9]+\)$/i.test(t));
  }

  // ------------------------------------------------------------------ nhận diện bảng
  function detectHeader(ws) {
    const maxCol = Math.min(ws.columnCount || 0, 60);
    const lastRow = Math.min(ws.rowCount || 0, 40);
    let best = null;
    for (let r = 1; r <= lastRow; r++) {
      const texts = rowTexts(ws, r, maxCol);
      const roles = new Set();
      let filled = 0;
      for (let c = 1; c <= maxCol; c++) {
        if (!texts[c]) continue;
        filled++;
        const role = guessRole(texts[c]);
        if (role) roles.add(role);
      }
      if (roles.size < 2 || filled < 2) continue;
      let score = roles.size + (roles.has('name') ? 2 : 0) + (roles.has('code') || roles.has('model') ? 1 : 0);
      if (!best || score > best.score) best = { row: r, score };
    }
    return best ? best.row : null;
  }

  /**
   * Phân tích một sheet. opts.headerRow / opts.roles (col -> role) để ép theo cấu hình người dùng.
   * Trả về null nếu sheet không giống bảng sản phẩm.
   */
  function analyzeSheet(ws, opts) {
    opts = opts || {};
    const autoHeader = detectHeader(ws);
    const headerRow = opts.headerRow || autoHeader;
    if (!headerRow) return null;
    let maxCol = Math.min(ws.columnCount || 0, 60);

    const head = rowTexts(ws, headerRow, maxCol);
    // dòng tiêu đề phụ (tiêu đề 2 tầng)
    let dataStart = headerRow + 1;
    const next = rowTexts(ws, headerRow + 1, maxCol);
    const headerTexts = head.slice();
    if (!isNumberingRow(next)) {
      let sub = false;
      for (let c = 1; c <= maxCol; c++) {
        if (!next[c]) continue;
        const top = ws.getCell(headerRow, c);
        if (top.isMerged && mergeSpan(ws, headerRow, c) > 1 && !/^[\d.,]+$/.test(next[c]) && next[c].length < 40) sub = true;
      }
      if (sub) {
        for (let c = 1; c <= maxCol; c++) {
          const top = ws.getCell(headerRow, c);
          const topText = valueText((top.isMerged ? top.master : top).value).trim();
          headerTexts[c] = next[c] ? (topText ? topText + ' - ' + next[c] : next[c]) : topText;
        }
        dataStart++;
      }
    }
    if (isNumberingRow(rowTexts(ws, dataStart, maxCol))) dataStart++;

    while (maxCol > 1 && !headerTexts[maxCol]) maxCol--;

    const autoRoles = {};
    const used = new Set();
    for (let c = 1; c <= maxCol; c++) {
      if (!headerTexts[c]) continue;
      let role = guessRole(headerTexts[c]);
      if (!role && headerTexts[c].includes(' - ')) role = guessRole(headerTexts[c].split(' - ').pop());
      if (role && !used.has(role)) { autoRoles[c] = role; used.add(role); }
    }
    const roles = opts.roles || autoRoles;
    const columns = [];
    for (let c = 1; c <= maxCol; c++) {
      if (!headerTexts[c] && !roles[c]) continue;
      columns.push({ col: c, header: headerTexts[c] || `Cột ${colLetter(c)}`, role: roles[c] || null });
    }
    const roleCol = {};
    for (const col of columns) if (col.role && !roleCol[col.role]) roleCol[col.role] = col.col;

    // phân loại dòng
    const rows = [];
    const lastRow = ws.rowCount || 0;
    let emptyRun = 0, footerStart = null, currentGroup = null, lastDataRow = dataStart - 1;
    const keyCols = ['code', 'name', 'model'].map((k) => roleCol[k]).filter(Boolean);
    for (let r = dataStart; r <= lastRow; r++) {
      const texts = rowTexts(ws, r, maxCol);
      const filled = [];
      for (let c = 1; c <= maxCol; c++) if (texts[c]) filled.push(c);
      if (!filled.length) {
        if (++emptyRun >= 30) break;
        continue;
      }
      emptyRun = 0;
      const hasKey = keyCols.length ? keyCols.some((c) => texts[c]) : filled.length >= 2;
      const firstText = texts[filled[0]];
      if (!hasKey || filled.length === 1) {
        if (isFooterText(firstText) || filled.some((c) => isFooterText(texts[c]))) { footerStart = r; break; }
      }
      const span = mergeSpan(ws, r, filled[0]);
      const roman = /^([IVXLC]+|[A-Z])[.)]\s/.test(firstText);
      let isGroup = false;
      if (filled.length === 1) isGroup = !hasKey || span >= 2 || roman;
      else if (roleCol.stt && /^([IVXLC]+|[A-Z])\.?$/.test(texts[roleCol.stt] || '')) {
        // "I | HỆ THỐNG ĐIỆN" : STT là số La Mã/chữ cái, không có mã, số lượng, đơn giá
        isGroup = !['code', 'qty', 'price', 'unit'].some((k) => roleCol[k] && texts[roleCol[k]]);
      }
      if (isGroup) {
        currentGroup = r;
        const title = filled.map((c) => texts[c]).join(' ');
        rows.push({ row: r, kind: 'group', title });
        lastDataRow = r;
        continue;
      }
      if (!hasKey) continue;
      rows.push({ row: r, kind: 'product', group: currentGroup });
      lastDataRow = r;
    }
    if (footerStart == null) footerStart = lastDataRow + 1;
    // dòng trống giữa dữ liệu và chân trang thuộc về chân trang
    if (footerStart > lastDataRow + 1) footerStart = lastDataRow + 1;

    const products = rows.filter((x) => x.kind === 'product');
    if (!products.length) return null;

    // STT có đánh lại từ 1 sau mỗi mục không?
    let sttRestarts = false;
    if (roleCol.stt) {
      let seenGroup = false, count = 0;
      for (const x of rows) {
        if (x.kind === 'group') { seenGroup = count > 0 || seenGroup; continue; }
        count++;
        const n = Number(valueText(ws.getCell(x.row, roleCol.stt).value));
        if (seenGroup && n === 1) sttRestarts = true;
      }
    }

    return {
      sheetName: ws.name,
      headerRow, autoHeaderRow: autoHeader, dataStart, lastDataRow, footerStart,
      maxCol, columns, roleCol, rows, sttRestarts,
      signature: sheetSignature(ws, autoHeader),
    };
  }

  function sheetSignature(ws, headerRow) {
    if (!headerRow) return '';
    const maxCol = Math.min(ws.columnCount || 0, 60);
    return rowTexts(ws, headerRow, maxCol).filter(Boolean).map(normHeader).join('|');
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colNumber(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  // ------------------------------------------------------------------ sản phẩm
  function readProduct(ws, table, rowInfo) {
    const cells = {};
    const raw = {};
    const formulas = {};
    for (const col of table.columns) {
      const cell = ws.getCell(rowInfo.row, col.col);
      if (isSlave(cell)) { cells[col.col] = ''; continue; }
      cells[col.col] = cellDisplay(cell).trim();
      raw[col.col] = primitive(cell.value);
      if (isFormulaValue(cell.value)) formulas[col.col] = true;
    }
    return { row: rowInfo.row, group: rowInfo.group, cells, raw, formulas };
  }

  function roleText(table, product, role) {
    const c = table.roleCol[role];
    return c ? product.cells[c] || '' : '';
  }

  // Các cột không có vai trò rõ ràng: coi như thông số phụ ("Điện áp: 5V; Dòng: 2A")
  function extraSpecText(table, product) {
    return table.columns
      .filter((c) => !c.role && product.cells[c.col])
      .map((c) => `${c.header}: ${product.cells[c.col]}`)
      .join('; ');
  }

  // ------------------------------------------------------------------ sửa ô
  function setCellText(ws, row, col, text) {
    const cell = ws.getCell(row, col);
    if (isFormulaValue(cell.value)) throw new Error('Ô này là công thức, không sửa trực tiếp được');
    const old = cell.value;
    let v = text;
    if (text === '') v = null;
    else if (typeof primitive(old) === 'number' || old == null) {
      const n = parseNumber(text);
      if (n != null && (typeof primitive(old) === 'number' || /^-?\d+([.,]\d+)?$/.test(text.trim()))) v = n;
    }
    cell.value = v;
    if (typeof v === 'string' && v.includes('\n')) {
      cell.alignment = Object.assign({}, cell.alignment || {}, { wrapText: true });
    }
    growRowHeight(ws, row);
    return old;
  }

  function restoreCellValue(ws, row, col, value) {
    ws.getCell(row, col).value = value;
  }

  // Ước lượng chiều cao dòng để chữ dài/xuống dòng không bị che (chỉ tăng, không giảm)
  function growRowHeight(ws, r) {
    const row = ws.getRow(r);
    let lines = 1;
    row.eachCell((cell, c) => {
      if (isSlave(cell)) return;
      const t = valueText(cell.value);
      if (!t) return;
      let width = ws.getColumn(c).width || 9;
      if (cell.isMerged) width = Math.max(width, width * mergeSpan(ws, r, c));
      const wrap = cell.alignment && cell.alignment.wrapText;
      const perLine = Math.max(4, Math.floor(width * 1.15));
      const n = t.split('\n').reduce((s, part) => s + (wrap ? Math.max(1, Math.ceil(part.length / perLine)) : 1), 0);
      lines = Math.max(lines, n);
    });
    if (lines <= 1) return;
    const need = lines * 15 + 4;
    if (!row.height || row.height < need) row.height = need;
  }

  // ------------------------------------------------------------------ công thức
  const REF_RE = /(?<![A-Za-z0-9_.!$'\]])(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})(?::(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6}))?(?![A-Za-z0-9_(])/g;

  /**
   * Dịch tham chiếu dòng trong công thức sau khi sắp xếp lại dòng.
   * mapRow(old) -> new | null (null = dòng đã bị bỏ).
   * ctx.dataFirst..ctx.dataLast: vùng dữ liệu cũ; ctx.newFirst..ctx.newLast: vùng dữ liệu mới.
   */
  function translateFormula(f, mapRow, ctx) {
    const parts = f.split(/("(?:[^"]|"")*")/);
    for (let i = 0; i < parts.length; i += 2) {
      parts[i] = parts[i].replace(REF_RE, (m, a1, c1, a2, r1, b1, c2, b2, r2) => {
        if (colNumber(c1) > 16384) return m;
        r1 = Number(r1);
        if (!c2) {
          const n = mapRow(r1);
          if (n == null) return '#REF!';
          return `${a1}${c1}${a2}${n}`;
        }
        r2 = Number(r2);
        const lo = Math.min(r1, r2), hi = Math.max(r1, r2);
        let min = Infinity, max = -Infinity;
        if (ctx && lo <= ctx.dataFirst && hi >= ctx.dataLast && ctx.newLast >= ctx.newFirst) {
          min = ctx.newFirst; max = ctx.newLast;
        }
        if (hi - lo <= 20000) {
          for (let r = lo; r <= hi; r++) {
            const n = mapRow(r);
            if (n == null) continue;
            if (n < min) min = n;
            if (n > max) max = n;
          }
        }
        if (min === Infinity) return '#REF!';
        return `${a1}${c1}${a2}${min}:${b1}${c2}${b2}${max}`;
      });
    }
    return parts.join('');
  }

  // ------------------------------------------------------------------ xuất file
  function cloneValue(v) {
    if (v == null || typeof v !== 'object') return v;
    if (v instanceof Date) return new Date(v.getTime());
    return JSON.parse(JSON.stringify(v));
  }
  function cloneStyle(s) {
    return s ? JSON.parse(JSON.stringify(s)) : {};
  }

  function snapshotRow(ws, r) {
    const row = ws.getRow(r);
    const cells = [];
    row.eachCell({ includeEmpty: true }, (cell, c) => {
      let value = null, formula = null, result;
      if (!isSlave(cell)) {
        if (isFormulaValue(cell.value)) {
          formula = cell.formula || (cell.value && cell.value.formula) || null;
          result = cell.value.result;
        } else value = cloneValue(cell.value);
      }
      cells.push({ col: c, value, formula, result, style: cloneStyle(cell.style) });
    });
    return { row: r, height: row.height, hidden: row.hidden, cells };
  }

  function listMerges(ws) {
    const out = [];
    const merges = ws._merges || {};
    for (const key of Object.keys(merges)) {
      const m = merges[key];
      const model = m.model || m;
      out.push({ top: model.top, left: model.left, bottom: model.bottom, right: model.right });
    }
    return out;
  }

  /**
   * Tạo workbook xuất ra trên workbook mẫu `wb` (đã nạp bằng ExcelJS, sẽ bị sửa trực tiếp).
   * table: kết quả analyzeSheet của sheet mẫu.
   * items: mảng theo thứ tự xuất:
   *   { type: 'row', row }                 // chép nguyên dòng của sheet mẫu
   *   { type: 'group', row }               // dòng tiêu đề mục của sheet mẫu
   *   { type: 'map', values, extras, specExtra } // SP từ file khác, ghép theo vai trò cột
   *       values: { role: giá trị }, extras: { headerChuẩnHóa: giá trị }, specExtra: chuỗi thông số phụ
   * options: { renumber: true }
   */
  function buildExport(wb, table, items, options) {
    options = Object.assign({ renumber: true }, options || {});
    const ws = wb.getWorksheet(table.sheetName);
    const dataStart = table.dataStart;
    const oldEnd = Math.max(ws.rowCount || 0, table.footerStart);
    const products = table.rows.filter((x) => x.kind === 'product');
    const firstProductRow = products[0].row;
    const ctxOld = { dataFirst: firstProductRow, dataLast: table.lastDataRow };

    // 1. chụp lại toàn bộ vùng dữ liệu + chân trang
    const snaps = {};
    for (let r = dataStart; r <= oldEnd; r++) snaps[r] = snapshotRow(ws, r);
    const merges = listMerges(ws).filter((m) => m.top >= dataStart);
    for (const m of merges) {
      try { ws.unMergeCells(m.top, m.left, m.bottom, m.right); } catch (e) { /* bỏ qua */ }
    }
    const media = (ws._media || []).filter((m) => m && m.range && m.range.tl);
    const mediaRows = media.map((m) => ({ m, row: Math.floor(m.range.tl.nativeRow) + 1 }));

    // 2. xoá trắng vùng cũ
    for (let r = dataStart; r <= oldEnd; r++) {
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell) => { cell.value = null; cell.style = {}; });
      row.height = undefined;
      row.hidden = false;
    }

    // 3. sắp đặt dòng mới
    const plan = []; // { snap, kind, item }
    const templateRow = snaps[firstProductRow];
    for (const it of items) {
      if (it.type === 'row' || it.type === 'group') plan.push({ snap: snaps[it.row], kind: it.type, item: it });
      else plan.push({ snap: templateRow, kind: 'map', item: it });
    }
    const dataCount = plan.length;
    const footerRows = [];
    for (let r = table.footerStart; r <= oldEnd; r++) footerRows.push(snaps[r]);

    const rowMap = new Map();
    for (let r = 1; r < dataStart; r++) rowMap.set(r, r);
    plan.forEach((p, i) => { if (p.kind !== 'map') rowMap.set(p.snap.row, dataStart + i); });
    footerRows.forEach((s, i) => rowMap.set(s.row, dataStart + dataCount + i));
    const ctx = Object.assign({}, ctxOld, { newFirst: dataStart, newLast: dataStart + dataCount - 1 });
    // mọi thứ nằm dưới dòng cuối (hình con dấu, ghi chú lẻ...) dịch theo phần chân trang
    const shiftAfter = (dataStart + dataCount + footerRows.length - 1) - oldEnd;
    const baseMap = (r) => (rowMap.has(r) ? rowMap.get(r) : (r > oldEnd ? r + shiftAfter : null));

    const sttCol = table.roleCol.stt;
    let stt = 0;
    const writeRow = (snap, newR, kind, item) => {
      const row = ws.getRow(newR);
      if (snap.height) row.height = snap.height;
      if (snap.hidden) row.hidden = true;
      const localMap = kind === 'map'
        ? (r) => (r === snap.row ? newR : baseMap(r))
        : baseMap;
      for (const c of snap.cells) {
        const cell = row.getCell(c.col);
        cell.style = cloneStyle(c.style);
        const colInfo = table.columns.find((x) => x.col === c.col);
        const role = colInfo && colInfo.role;
        if (c.formula) {
          if (kind === 'map' && role && item.values[role] != null && role !== 'amount') {
            cell.value = item.values[role];
          } else {
            cell.value = { formula: translateFormula(c.formula, localMap, ctx) };
          }
          continue;
        }
        if (kind === 'map') {
          let v = null;
          if (role && item.values[role] != null && item.values[role] !== '') v = item.values[role];
          else if (colInfo && item.extras && item.extras[normHeader(colInfo.header)] != null) v = item.extras[normHeader(colInfo.header)];
          if (role === 'spec' && item.specExtra) v = v ? `${v}\n${item.specExtra}` : item.specExtra;
          cell.value = v;
        } else {
          cell.value = cloneValue(c.value);
        }
      }
      if (kind === 'map') {
        // cột có trong mẫu nhưng không có ô ở dòng mẫu
        for (const colInfo of table.columns) {
          if (snap.cells.some((c) => c.col === colInfo.col)) continue;
          const role = colInfo.role;
          if (role && item.values[role] != null) row.getCell(colInfo.col).value = item.values[role];
        }
        if (table.roleCol.spec) growRowHeight(ws, newR);
      }
      if (options.renumber && sttCol && (kind === 'row' || kind === 'map')) {
        const cur = row.getCell(sttCol);
        const old = primitive(cur.value);
        if (kind === 'map' || old == null || old === '' || /^\d+$/.test(String(old))) {
          stt++;
          cur.value = stt;
        }
      }
      if (kind === 'group' && options.renumber && table.sttRestarts) stt = 0;
    };

    plan.forEach((p, i) => writeRow(p.snap, dataStart + i, p.kind, p.item));
    footerRows.forEach((s, i) => writeRow(s, dataStart + dataCount + i, 'footer', null));

    // 4. gộp ô lại theo vị trí mới
    const addMerge = (top, left, bottom, right) => {
      try { ws.mergeCells(top, left, bottom, right); } catch (e) { /* trùng vùng gộp thì bỏ qua */ }
    };
    for (const m of merges) {
      const height = m.bottom - m.top;
      const newTop = rowMap.get(m.top);
      if (newTop != null) {
        if (height === 0 || rowMap.get(m.bottom) === newTop + height) addMerge(newTop, m.left, newTop + height, m.right);
      }
      if (height === 0 && m.top === firstProductRow) {
        plan.forEach((p, i) => { if (p.kind === 'map') addMerge(dataStart + i, m.left, dataStart + i, m.right); });
      }
    }
    // gộp ô làm mất định dạng ô phụ (viền) -> áp lại
    const reapply = (snap, newR) => {
      for (const c of snap.cells) ws.getCell(newR, c.col).style = cloneStyle(c.style);
    };
    plan.forEach((p, i) => reapply(p.snap, dataStart + i));
    footerRows.forEach((s, i) => reapply(s, dataStart + dataCount + i));

    // 5. công thức ở phần đầu trang trỏ xuống vùng dữ liệu
    for (let r = 1; r < dataStart; r++) {
      ws.getRow(r).eachCell((cell) => {
        if (isFormulaValue(cell.value) && cell.formula) {
          cell.value = { formula: translateFormula(cell.formula, baseMap, ctx) };
        }
      });
    }

    // 6. hình ảnh (logo, con dấu, ảnh sản phẩm)
    const keep = [];
    for (const { m, row } of mediaRows) {
      if (row < dataStart) { keep.push(m); continue; }
      const nr = baseMap(row);
      if (nr == null) continue;
      const delta = nr - row;
      m.range.tl.nativeRow += delta;
      if (m.range.br) m.range.br.nativeRow += delta;
      keep.push(m);
    }
    if (ws._media) ws._media = ws._media.filter((m) => !m.range || !m.range.tl || keep.includes(m));

    // 7. bộ lọc tự động
    if (typeof ws.autoFilter === 'string') {
      ws.autoFilter = ws.autoFilter.replace(/(\d+)$/, String(Math.max(dataStart, dataStart + dataCount - 1)));
    } else if (ws.autoFilter && ws.autoFilter.to) {
      const to = ws.autoFilter.to;
      if (typeof to === 'string') ws.autoFilter.to = to.replace(/\d+$/, String(dataStart + dataCount - 1));
    }

    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return { dataRows: dataCount, lastRow: dataStart + dataCount + footerRows.length - 1 };
  }

  /**
   * Chuyển một SP của bảng nguồn thành mục 'map' để ghép vào bảng mẫu có cột khác.
   * Thông tin mà mẫu không có cột tương ứng sẽ được ghép vào ô thông số kỹ thuật.
   */
  function makeMapItem(srcTable, product, tplTable) {
    const values = {}, extras = {};
    for (const col of srcTable.columns) {
      const v = product.raw[col.col];
      if (v == null || v === '') continue;
      if (col.role && values[col.role] == null) values[col.role] = v;
      extras[normHeader(col.header)] = v;
    }
    const tplRoles = new Set(tplTable.columns.map((c) => c.role).filter(Boolean));
    const tplHeaders = new Set(tplTable.columns.map((c) => normHeader(c.header)));
    // mã <-> ký mã hiệu thay thế cho nhau khi mẫu chỉ có một trong hai
    if (!tplRoles.has('model') && values.model != null && values.code == null) { values.code = values.model; delete values.model; }
    if (!tplRoles.has('code') && values.code != null && values.model == null) { values.model = values.code; delete values.code; }
    const extra = [];
    if (tplRoles.has('spec')) {
      for (const col of srcTable.columns) {
        const text = product.cells[col.col];
        if (!text || tplHeaders.has(normHeader(col.header))) continue;
        if (col.role && (tplRoles.has(col.role) || ['stt', 'amount', 'qty', 'price', 'code', 'model'].includes(col.role))) continue;
        if (col.role === 'spec' || col.role === 'name') continue;
        extra.push(`${col.header}: ${text}`);
      }
    }
    return { type: 'map', values, extras, specExtra: extra.join('; ') };
  }

  // ------------------------------------------------------------------ sửa lỗi cấu trúc file
  /**
   * Một số phần mềm (openpyxl, WPS, phần mềm kế toán...) ghi đường dẫn tuyệt đối trong *.rels
   * (Target="/xl/drawings/drawing1.xml") khiến ExcelJS không đọc được. Đổi về đường dẫn tương đối.
   * Trả lại buffer gốc nếu không cần sửa.
   */
  async function normalizeXlsx(buffer, JSZip) {
    return (await prepareXlsx(buffer, JSZip)).buffer;
  }

  /** Như normalizeXlsx, kèm danh sách cảnh báo về những thứ có thể mất khi lưu lại. */
  async function prepareXlsx(buffer, JSZip) {
    const zip = await JSZip.loadAsync(buffer);
    const names = Object.keys(zip.files);
    const warnings = [];
    if (names.some((n) => n.startsWith('xl/charts/'))) warnings.push('biểu đồ');
    if (names.some((n) => n.startsWith('xl/pivotTables/'))) warnings.push('bảng Pivot');
    if (names.some((n) => /vbaProject\.bin$/.test(n))) warnings.push('macro VBA');
    if (names.some((n) => n.startsWith('xl/slicers/'))) warnings.push('slicer');
    for (const n of names.filter((x) => /^xl\/drawings\/drawing\d+\.xml$/.test(x))) {
      if (/<(xdr:)?sp>|<(xdr:)?sp\s/.test(await zip.file(n).async('string'))) { warnings.push('hình vẽ/khung chữ'); break; }
    }
    let changed = false;
    const relFiles = Object.keys(zip.files).filter((n) => n.endsWith('.rels'));
    for (const name of relFiles) {
      const xml = await zip.file(name).async('string');
      if (!/Target="\//.test(xml)) continue;
      // xl/worksheets/_rels/sheet1.xml.rels -> thư mục nguồn: xl/worksheets
      const srcDir = name.replace(/_rels\/[^/]+$/, '').replace(/\/$/, '');
      const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
        if (/TargetMode="External"/.test(tag)) return tag;
        return tag.replace(/Target="\/([^"]*)"/, (m, abs) => `Target="${relativePath(srcDir, abs)}"`);
      });
      if (fixed !== xml) { zip.file(name, fixed); changed = true; }
    }
    // Hình vẽ ghi bằng namespace mặc định (<wsDr xmlns="...">) thay vì tiền tố xdr: -> ExcelJS bỏ qua và lỗi
    const SD = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    for (const name of Object.keys(zip.files).filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))) {
      const xml = await zip.file(name).async('string');
      if (!new RegExp(`<wsDr[^>]*xmlns="${SD}"`).test(xml)) continue;
      const fixed = xml
        .replace(`xmlns="${SD}"`, `xmlns:xdr="${SD}"`)
        .replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s/>])/g, (m, slash, tag) => `<${slash}xdr:${tag}`);
      zip.file(name, fixed);
      changed = true;
    }
    if (!changed) return { buffer, warnings };
    return { buffer: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), warnings };
  }
  function relativePath(fromDir, target) {
    const a = fromDir ? fromDir.split('/') : [];
    const b = target.split('/');
    let i = 0;
    while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    return [...Array(a.length - i).fill('..'), ...b.slice(i)].join('/');
  }

  const api = {
    normalizeXlsx, prepareXlsx,
    norm, normHeader, valueText, primitive, isFormulaValue, cellDisplay, parseNumber, formatNumber,
    ROLES, ROLE_LABEL, guessRole, detectHeader, analyzeSheet, readProduct, roleText, extraSpecText,
    setCellText, restoreCellValue, translateFormula, buildExport, makeMapItem, colLetter, colNumber, isSlave,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.WXCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
