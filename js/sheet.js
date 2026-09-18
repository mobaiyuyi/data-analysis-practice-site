/* mobai训练 · 表格组件
 * 目标是「手感像 Excel」，不是功能像 Excel。三个反射动作必须成立：
 *   1) 写公式时用鼠标点 / 拖单元格 → 引用自动插进公式，F4 切 $ 的四种写法
 *   2) 选区右下角的填充柄：拖动填充（等差 / 日期 / 文本编号 / 公式平移），双击一拉到底
 *   3) 一眼认得出的外壳：功能区（js/ribbon.js）+ 名称框 + fx + 行号列标 + 状态栏 + 工作表标签
 *
 * 编辑用「定位在单元格上的浮层 input」，不用 contenteditable——
 * IME（拼音输入）在 contenteditable 里的表现远比原生 input 难控制。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;
  var E = Mobai.Errors;

  var MIN_COL_WIDTH = 48;
  var DEFAULT_COL_WIDTH = 96;
  var ROW_HEIGHT = 28;
  var MIN_COLS = 12;
  var MIN_ROWS = 26;
  var MAX_HISTORY = 60;

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = text;
    return n;
  }

  /* ---------- 组件 ---------- */

  /**
   * @param {HTMLElement} container
   * @param {{workbook: Mobai.Workbook, onChange: Function, editable: boolean, note: {title,text}}} opts
   */
  function mount(container, opts) {
    opts = opts || {};
    var wb = opts.workbook;
    if (!wb) throw new Error('Mobai.Sheet.mount 需要一个 workbook');
    var onChange = opts.onChange || function () {};
    var editable = opts.editable !== false;

    var history = [];
    var redoStack = [];
    var widths = {};
    var clipboard = null;
    var tds = [];               // tds[row][col] 只记「锚点格」（合并区只有锚点有 td）
    var rowTds = [];            // rowTds[row] = 这一行真正渲染出来的 td
    var colHeads = [];
    var rowHeads = [];
    var dims = { cols: MIN_COLS, rows: MIN_ROWS };
    var selection = { anchor: { col: 0, row: 1 }, focus: { col: 0, row: 1 } };
    var extras = [];            // Ctrl 加选的不连续区域
    var edit = null;            // { mode:'cell'|'bar', input, col, row }
    var refPoint = null;        // 公式点选引用：{ input, at, len, anchor, focus, whole, abs }
    var pointDrag = false;
    var fillDrag = null;
    var dragSel = false;
    var suppressBlur = false;
    var menu = null;

    /* --- 骨架 --- */

    var ribbonSlot = el('div', 'ribbon-slot');
    var bar = el('div', 'sheet-bar');
    var addrBox = el('input', 'sheet-addr');
    addrBox.setAttribute('aria-label', '名称框：输入 A1 或 A1:C10 回车跳过去');
    addrBox.spellcheck = false;
    var fx = el('span', 'sheet-fx', 'fx');
    var formulaBar = el('input', 'sheet-formula');
    formulaBar.setAttribute('aria-label', '公式栏');
    formulaBar.spellcheck = false;
    var barSep = el('span', 'sheet-bar-sep', '=');
    bar.appendChild(addrBox);
    bar.appendChild(barSep);
    bar.appendChild(fx);
    bar.appendChild(formulaBar);

    var tip = el('div', 'sheet-tip');

    var body = el('div', 'sheet-body');
    var grid = el('div', 'sheet-grid');
    grid.tabIndex = 0;
    var table = el('table', 'sheet-table');
    var colgroup = el('colgroup');
    var thead = el('thead');
    var tbody = el('tbody');
    table.appendChild(colgroup);
    table.appendChild(thead);
    table.appendChild(tbody);
    grid.appendChild(table);

    var note = el('div', 'sheet-note');
    note.style.display = 'none';
    body.appendChild(grid);
    body.appendChild(note);

    var handle = el('div', 'sheet-fill-handle');
    handle.title = '拖动填充：数字按等差、日期按天、公式按相对引用平移；双击填到相邻数据末行';
    handle.setAttribute('role', 'button');
    handle.setAttribute('aria-label', '填充柄：按住拖动向下或向右填充，双击填到底');
    grid.appendChild(handle);
    var preview = el('div', 'sheet-fill-preview');
    preview.style.display = 'none';
    grid.appendChild(preview);
    var bubble = el('div', 'sheet-fill-bubble');
    bubble.style.display = 'none';
    grid.appendChild(bubble);
    var refBox = el('div', 'sheet-ref-box');
    refBox.style.display = 'none';
    grid.appendChild(refBox);
    var refTip = el('div', 'sheet-ref-bubble');
    refTip.style.display = 'none';
    grid.appendChild(refTip);

    var editor = el('input', 'sheet-editor');
    editor.spellcheck = false;
    editor.style.display = 'none';
    grid.appendChild(editor);

    var status = el('div', 'sheet-status');
    var statusLeft = el('span', 'sheet-status-left', '就绪');
    var statusRight = el('span', 'sheet-status-right');
    status.appendChild(statusLeft);
    status.appendChild(statusRight);

    var tabs = el('div', 'sheet-tabs');
    var tabGrid = el('button', 'sheet-tab is-active', 'Sheet1');
    tabGrid.type = 'button';
    tabs.appendChild(tabGrid);
    var tabNote = null;
    if (opts.note) {
      tabNote = el('button', 'sheet-tab', opts.note.tab || '说明');
      tabNote.type = 'button';
      tabs.appendChild(tabNote);
    }

    container.appendChild(ribbonSlot);
    container.appendChild(bar);
    container.appendChild(tip);
    container.appendChild(body);
    container.appendChild(status);
    container.appendChild(tabs);
    container.classList.add('sheet');

    if (opts.note) {
      var h = el('h3', null, opts.note.title || '这一题在练什么');
      note.appendChild(h);
      var paras = String(opts.note.text || '').split('\n\n');
      paras.forEach(function (p) { note.appendChild(el('p', null, p)); });
    }

    /* --- 选区 --- */

    function selRange() {
      return U.normRange(selection.anchor, selection.focus);
    }

    function allRanges() {
      var out = [];
      for (var i = 0; i < extras.length; i++) out.push(extras[i]);
      out.push(selRange());
      return out;
    }

    function isMulti() {
      var r = selRange();
      return r.start.col !== r.end.col || r.start.row !== r.end.row;
    }

    function addrLabel() {
      var r = selRange();
      return U.rangeEqual(r, U.normRange(r.start, r.start)) ? U.addrOf(r.start.col, r.start.row) : U.rangeToString(r);
    }

    /** 合并区里的任意一格 → 整个合并区 */
    function expandToMerge(cell) {
      var m = wb.mergeOf(cell.col, cell.row);
      if (!m) return { start: { col: cell.col, row: cell.row }, end: { col: cell.col, row: cell.row } };
      return m;
    }

    /* --- 渲染 --- */

    function colWidth(col) {
      var w = widths[col];
      if (typeof w === 'number') return w;
      var preset = wb.workbookOpts.widths || {};
      var fromOpts = preset[U.colToName(col)];
      return typeof fromOpts === 'number' ? fromOpts : DEFAULT_COL_WIDTH;
    }

    function applyInlineStyle(td, st, col, row) {
      td.style.background = '';
      td.style.color = '';
      td.style.textAlign = '';
      td.style.verticalAlign = '';
      td.style.fontFamily = '';
      td.style.fontSize = '';
      if (!st) return;
      if (st.bg) td.style.background = st.bg;
      if (st.fg) td.style.color = st.fg;
      if (st.alignH) td.style.textAlign = st.alignH;
      if (st.alignV) td.style.verticalAlign = st.alignV;
      if (st.fontName) td.style.fontFamily = st.fontName;
      if (st.fontSize) td.style.fontSize = st.fontSize + 'px';
    }

    function paintCell(td, col, row) {
      var entry = wb.entry(col, row);
      var v = entry ? wb.getCell(col, row) : null;
      var cls = 'sheet-cell';
      td.textContent = entry ? wb.display(col, row) : '';
      if (!entry || entry.kind === 'empty') cls += ' is-empty';
      if (wb.isReadonly(col, row)) cls += ' is-ro';
      var st = wb.styleOf(col, row);
      if (st) {
        if (st.bold) cls += ' is-bold';
        if (st.italic) cls += ' is-italic';
        if (st.underline) cls += ' is-underline';
        if (st.wrap) cls += ' is-wrap';
      }
      if (U.addrInRanges(col, row, wb.bold)) cls += ' is-bold';
      var isDate = wb.isDateCell(col, row);
      var numeric = typeof v === 'number';
      if (Mobai.Format && entry && entry.format && Mobai.Format.isNumericFormat(entry.format)) numeric = true;
      if (numeric && !isDate) cls += ' is-num';
      if (isDate) cls += ' is-date';
      if (E.isError(v)) cls += ' is-err';
      td.className = cls;
      applyInlineStyle(td, st, col, row);
    }

    function render() {
      dims = wb.dims(MIN_COLS, MIN_ROWS);

      colgroup.innerHTML = '';
      colgroup.appendChild(el('col', 'sheet-colno')).style.width = '42px';
      for (var c = 0; c < dims.cols; c++) {
        var colEl = el('col');
        colEl.style.width = colWidth(c) + 'px';
        colgroup.appendChild(colEl);
      }

      thead.innerHTML = '';
      var htr = el('tr');
      var corner = el('th', 'sheet-corner', '');
      corner.dataset.corner = '1';
      htr.appendChild(corner);
      colHeads = [];
      for (var c2 = 0; c2 < dims.cols; c2++) {
        var th = el('th', 'sheet-colhead', U.colToName(c2));
        th.dataset.colhead = String(c2);
        var rs = el('span', 'sheet-col-resize');
        rs.dataset.col = String(c2);
        th.appendChild(rs);
        htr.appendChild(th);
        colHeads[c2] = th;
      }
      thead.appendChild(htr);

      tbody.innerHTML = '';
      tds = [];
      rowTds = [];
      rowHeads = [];
      for (var row = 1; row <= dims.rows; row++) {
        var tr = el('tr');
        var rh = el('th', 'sheet-rowhead', String(row));
        rh.dataset.rowhead = String(row);
        tr.appendChild(rh);
        rowHeads[row] = rh;
        var line = [];
        var lineTds = [];
        for (var c3 = 0; c3 < dims.cols; c3++) {
          var m = wb.mergeOf(c3, row);
          if (m && (m.start.col !== c3 || m.start.row !== row)) continue;   // 被合并盖住的格子不渲染
          var td = el('td');
          td.dataset.col = String(c3);
          td.dataset.row = String(row);
          var c1 = c3;
          var r1 = row;
          if (m) {
            c1 = m.end.col;
            r1 = m.end.row;
            td.colSpan = c1 - c3 + 1;
            if (r1 > row) td.rowSpan = r1 - row + 1;
          }
          td._c0 = c3;
          td._r0 = row;
          td._c1 = c1;
          td._r1 = r1;
          paintCell(td, c3, row);
          tr.appendChild(td);
          line[c3] = td;
          lineTds.push(td);
        }
        tbody.appendChild(tr);
        tds[row] = line;
        rowTds[row] = lineTds;
      }

      paintSelection();
    }

    function tdAt(col, row) {
      if (tds[row] && tds[row][col]) return tds[row][col];
      var m = wb.mergeOf(col, row);
      if (m && tds[m.start.row] && tds[m.start.row][m.start.col]) return tds[m.start.row][m.start.col];
      return null;
    }

    function scrollIntoView(col, row) {
      var td = tdAt(col, row);
      if (td && td.scrollIntoView) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }

    function tdRect(td) {
      return {
        start: { col: td._c0, row: td._r0 },
        end: { col: td._c1, row: td._r1 }
      };
    }

    function paintSelection() {
      var ranges = allRanges();
      for (var row = 1; row < rowTds.length; row++) {
        var line = rowTds[row];
        if (!line) continue;
        for (var i = 0; i < line.length; i++) {
          var td = line[i];
          var rect = tdRect(td);
          var inSel = false;
          for (var k = 0; k < ranges.length; k++) {
            if (U.rangeOverlap(rect, ranges[k])) { inSel = true; break; }
          }
          var cls = td.className.replace(/\s*is-(sel|active|ro-sel)\b/g, '');
          if (inSel) {
            cls += ' is-sel';
            if (wb.isReadonly(td._c0, td._r0)) cls += ' is-ro-sel';
          }
          if (selection.focus.col >= td._c0 && selection.focus.col <= td._c1 &&
            selection.focus.row >= td._r0 && selection.focus.row <= td._r1) cls += ' is-active';
          td.className = cls;
        }
      }
      refreshBar();
      markHeaders();
      positionHandle();
      updateStatus();
      if (edit && edit.mode === 'cell') positionEditor();
      if (api.onState) api.onState(stateOfSelection());
    }

    /** 只要动了值就把整张表重绘一遍。
     *  只重绘当前格是不够的：别处的公式可能正引用着这一格，
     *  改了 C2 而 C11 的 =SUM(C2:C9) 还挂着旧结果，是最容易被当成「没生效」的那种坑。 */
    function paintAll() {
      var trs = grid.querySelectorAll('.sheet-table tbody tr');
      for (var i = 0; i < trs.length; i++) {
        var cells = trs[i].children;
        for (var j = 0; j < cells.length; j++) {
          var td = cells[j];
          if (typeof td._c0 !== 'number') continue;
          paintCell(td, td._c0, td._r0);
        }
      }
      paintSelection();
    }

    /* 当前格所在的行号 / 列标点亮，一眼知道自己在哪一行哪一列 */
    function markHeaders() {
      var hotCol = -1;
      var hotRow = -1;
      if (selection.focus) {
        hotCol = selection.focus.col;
        hotRow = selection.focus.row;
      }
      for (var c = 0; c < colHeads.length; c++) {
        if (!colHeads[c]) continue;
        if (c === hotCol) colHeads[c].classList.add('is-hot');
        else colHeads[c].classList.remove('is-hot');
      }
      for (var r = 1; r < rowHeads.length; r++) {
        if (!rowHeads[r]) continue;
        if (r === hotRow) rowHeads[r].classList.add('is-hot');
        else rowHeads[r].classList.remove('is-hot');
      }
    }

    function refreshBar() {
      addrBox.value = addrLabel();
      formulaBar.value = activeEntryText();
    }

    function activeEntryText() {
      var f = selection.focus;
      return wb.rawText(f.col, f.row);
    }

    function updateStatus() {
      var r = selRange();
      var cells = U.expandRange(r);
      var nums = [];
      var count = 0;
      for (var i = 0; i < cells.length; i++) {
        var v = wb.getCell(cells[i].col, cells[i].row);
        if (v === null || v === undefined || E.isError(v)) continue;
        count++;
        if (typeof v === 'number') nums.push(v);
      }
      var left = isMulti() ? '选中 ' + U.rangeToString(r) + '（' + count + ' 格有内容）' : '就绪';
      statusLeft.textContent = left;
      if (!nums.length) {
        statusRight.textContent = '';
        return;
      }
      var sum = nums.reduce(function (a, b) { return a + b; }, 0);
      statusRight.textContent =
        '求和 ' + U.formatNumber(Math.round(sum * 1e10) / 1e10) +
        '　平均值 ' + U.formatNumber(Math.round((sum / nums.length) * 1e4) / 1e4) +
        '　数值个数 ' + nums.length;
    }

    function setTip(text, kind) {
      tip.textContent = text || '';
      tip.className = 'sheet-tip' + (kind ? ' is-' + kind : '');
    }

    /* --- 撤销 / 重做 --- */

    function pushUndo() {
      history.push(wb.snapshot());
      if (history.length > MAX_HISTORY) history.shift();
      redoStack.length = 0;
    }

    function undo() {
      if (!history.length) {
        setTip('没有可以撤销的操作了。');
        return;
      }
      cancelEdit();
      redoStack.push(wb.snapshot());
      wb.restore(history.pop());
      render();
      setTip('已撤销上一步。');
      onChange('undo');
    }

    function redo() {
      if (!redoStack.length) {
        setTip('没有可以重做的操作了。');
        return;
      }
      cancelEdit();
      history.push(wb.snapshot());
      wb.restore(redoStack.pop());
      render();
      setTip('已重做。');
      onChange('redo');
    }

    /* --- 编辑 --- */

    function canPoint(input) {
      return editable && input && input.value.charAt(0) === '=';
    }

    function startEditCell(initialText) {
      if (!editable) return;
      var col = selection.focus.col;
      var row = selection.focus.row;
      if (wb.isReadonly(col, row)) {
        setTip(U.addrOf(col, row) + ' 是参考数据，不能改。要练的地方在别的区域。', 'warn');
        return;
      }
      editor.style.display = 'block';
      editor.value = initialText !== undefined ? initialText : wb.rawText(col, row);
      edit = { mode: 'cell', input: editor, col: col, row: row };
      positionEditor();
      editor.focus();
      if (initialText === undefined) editor.select();
      else editor.setSelectionRange(editor.value.length, editor.value.length);
      setTip('正在编辑 ' + U.addrOf(col, row) + '：写 = 开头的公式时，可以直接用鼠标点单元格把引用点进来。');
    }

    function positionEditor() {
      if (!edit || edit.mode !== 'cell') return;
      var td = tdAt(edit.col, edit.row);
      if (!td) return;
      editor.style.left = td.offsetLeft + 'px';
      editor.style.top = td.offsetTop + 'px';
      editor.style.width = Math.max(td.offsetWidth, colWidth(edit.col)) + 'px';
      editor.style.height = td.offsetHeight + 'px';
    }

    function commitEdit(move, keepFocus) {
      if (!edit) return;
      var e = edit;
      var input = e.input;
      var text = input.value;
      edit = null;
      refPoint = null;
      hideRefBox();
      if (input === editor) {
        editor.style.display = 'none';
        editor.value = '';
      }
      var before = wb.rawText(e.col, e.row);
      if (text !== before) {
        pushUndo();
        wb.setInput(e.col, e.row, text);
        paintAll();
        onChange('edit');
      }
      setTip('');
      if (move) moveFocus(move);
      else if (!keepFocus) grid.focus();
      refreshBar();
    }

    function cancelEdit() {
      if (!edit) return;
      if (edit.input === editor) {
        editor.style.display = 'none';
        editor.value = '';
      }
      edit = null;
      refPoint = null;
      hideRefBox();
      grid.focus();
    }

    function commitPending() {
      if (edit) commitEdit(null, true);
    }

    /* --- 移动选中 --- */

    function moveFocus(delta) {
      if (!delta) { grid.focus(); return; }
      var col = Math.min(Math.max(selection.focus.col + (delta.col || 0), 0), dims.cols - 1);
      var row = Math.min(Math.max(selection.focus.row + (delta.row || 0), 1), dims.rows);
      selection.anchor = { col: col, row: row };
      selection.focus = { col: col, row: row };
      extras = [];
      paintSelection();
      scrollIntoView(col, row);
      grid.focus();
    }

    function extendFocus(delta) {
      selection.focus = {
        col: Math.min(Math.max(selection.focus.col + (delta.col || 0), 0), dims.cols - 1),
        row: Math.min(Math.max(selection.focus.row + (delta.row || 0), 1), dims.rows)
      };
      paintSelection();
      scrollIntoView(selection.focus.col, selection.focus.row);
    }

    function selectAll() {
      selection.anchor = { col: 0, row: 1 };
      selection.focus = { col: dims.cols - 1, row: dims.rows };
      extras = [];
      paintSelection();
    }

    function setSelection(range, keepExtras) {
      selection.anchor = { col: range.start.col, row: range.start.row };
      selection.focus = { col: range.end.col, row: range.end.row };
      if (!keepExtras) extras = [];
      paintSelection();
    }

    /* --- 清空 / 填充 / 复制粘贴 --- */

    function editableCellsIn(ranges) {
      var out = [];
      var seen = Object.create(null);
      ranges.forEach(function (r) {
        U.expandRange(r).forEach(function (c) {
          var key = U.addrOf(c.col, c.row);
          if (seen[key]) return;
          if (wb.isReadonly(c.col, c.row)) return;
          seen[key] = 1;
          out.push(c);
        });
      });
      return out;
    }

    function clearSelection() {
      var targets = editableCellsIn(allRanges());
      if (!targets.length) {
        setTip('选中的都是参考数据，不能清空。', 'warn');
        return;
      }
      pushUndo();
      targets.forEach(function (c) { wb.clearCell(c.col, c.row); });
      render();
      setTip('已清空 ' + targets.length + ' 格内容（格式留着，跟 Excel 按 Delete 一样）。');
      onChange('clear');
    }

    function rawOf(col, row) {
      return wb.rawText(col, row);
    }

    /** 一次填充的公共部分：把 plan 的结果写进目标格 */
    function writeFill(axis, src, count) {
      var written = 0;
      var lastPlan = null;
      if (axis === 'row') {
        for (var col = src.start.col; col <= src.end.col; col++) {
          var sources = [];
          var dates = [];
          for (var row = src.start.row; row <= src.end.row; row++) {
            sources.push(rawOf(col, row));
            dates.push(wb.isDateCell(col, row));
          }
          var plan = Mobai.Fill.planFill(sources, count, { dates: dates, axis: 'row' });
          lastPlan = plan;
          for (var j = 0; j < count; j++) {
            var trow = src.end.row + 1 + j;
            if (trow > dims.rows) break;
            if (wb.isReadonly(col, trow)) continue;
            var val = plan.values[j];
            if (val === '' || val === undefined) continue;
            wb.setInput(col, trow, val);
            written++;
          }
        }
      } else {
        for (var r2 = src.start.row; r2 <= src.end.row; r2++) {
          var srcs = [];
          var ds = [];
          for (var c2 = src.start.col; c2 <= src.end.col; c2++) {
            srcs.push(rawOf(c2, r2));
            ds.push(wb.isDateCell(c2, r2));
          }
          var plan2 = Mobai.Fill.planFill(srcs, count, { dates: ds, axis: 'col' });
          lastPlan = plan2;
          for (var j2 = 0; j2 < count; j2++) {
            var tcol = src.end.col + 1 + j2;
            if (tcol > dims.cols - 1) break;
            if (wb.isReadonly(tcol, r2)) continue;
            var val2 = plan2.values[j2];
            if (val2 === '' || val2 === undefined) continue;
            wb.setInput(tcol, r2, val2);
            written++;
          }
        }
      }
      return { written: written, label: Mobai.Fill.previewText(lastPlan) };
    }

    function extendSelection(axis, src, count) {
      if (axis === 'row') {
        selection.anchor = { col: src.start.col, row: src.start.row };
        selection.focus = { col: src.end.col, row: src.end.row + count };
      } else {
        selection.anchor = { col: src.start.col, row: src.start.row };
        selection.focus = { col: src.end.col + count, row: src.end.row };
      }
    }

    function fillBy(axis, count) {
      if (count <= 0) return;
      var src = selRange();
      pushUndo();
      extras = [];
      var res = writeFill(axis, src, count);
      extendSelection(axis, src, count);
      render();
      setTip(res.written
        ? '已' + (axis === 'row' ? '向下' : '向右') + '填充 ' + res.written + ' 格（' + res.label + '）。'
        : '要填的地方在参考区里，跳过了。', res.written ? '' : 'warn');
      onChange('fill');
    }

    /** Ctrl+D：向下填充 */
    function fillDown() {
      var r = selRange();
      var single = r.start.row === r.end.row;
      if (single && r.start.row <= 1) {
        setTip('第一行没有上一行可以填下来。', 'warn');
        return;
      }
      var srcRow = single ? r.start.row - 1 : r.start.row;
      var fromRow = single ? r.start.row : r.start.row + 1;
      if (fromRow > r.end.row) {
        setTip('向下填充会把选中区域的第一行复制到下面几行。', 'warn');
        return;
      }
      var src = {
        start: { col: r.start.col, row: srcRow },
        end: { col: r.end.col, row: srcRow }
      };
      pushUndo();
      var res = writeFill('row', src, r.end.row - srcRow);
      render();
      setTip(res.written ? '已向下填充 ' + res.written + ' 格（' + res.label + '）。' : '没找到可填的内容。',
        res.written ? '' : 'warn');
      onChange('fill');
    }

    function selectionRaw() {
      var r = selRange();
      var out = [];
      for (var row = r.start.row; row <= r.end.row; row++) {
        var line = [];
        for (var col = r.start.col; col <= r.end.col; col++) line.push(rawOf(col, row));
        out.push(line);
      }
      return { raw: out, anchor: { col: r.start.col, row: r.start.row } };
    }

    function toTSV(matrix) {
      return matrix.map(function (line) { return line.join('\t'); }).join('\n');
    }

    function copySelection(cut) {
      var s = selectionRaw();
      var text = toTSV(s.raw);
      clipboard = { text: text, raw: s.raw, anchor: s.anchor };
      if (cut) {
        var targets = editableCellsIn([selRange()]);
        pushUndo();
        targets.forEach(function (c) { wb.clearCell(c.col, c.row); });
        render();
        setTip('已剪切 ' + s.raw.length + ' 行。');
        onChange('cut');
      } else {
        setTip('已复制 ' + s.raw.length + ' 行。');
      }
      return text;
    }

    /** 粘贴：优先走内部剪贴板（能平移相对引用），否则按纯文本写入 */
    function pasteTC(text) {
      if (!text) return;
      var fromInternal = clipboard && clipboard.text === text;
      var matrix = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map(function (line) {
        return line.split('\t');
      });
      var start = { col: selection.focus.col, row: selection.focus.row };
      var dCol = fromInternal ? start.col - clipboard.anchor.col : 0;
      var dRow = fromInternal ? start.row - clipboard.anchor.row : 0;
      var blocked = 0;
      var written = 0;

      pushUndo();
      for (var i = 0; i < matrix.length; i++) {
        for (var j = 0; j < matrix[i].length; j++) {
          var col = start.col + j;
          var row = start.row + i;
          if (col >= dims.cols || row > dims.rows) continue;
          if (wb.isReadonly(col, row)) { blocked++; continue; }
          var value = matrix[i][j];
          if (fromInternal && i < clipboard.raw.length && j < clipboard.raw[i].length) {
            var raw = clipboard.raw[i][j];
            value = raw.charAt(0) === '=' ? Mobai.Fill.shiftFormulaText(raw, dCol, dRow) : raw;
          }
          wb.setInput(col, row, value);
          written++;
        }
      }
      render();
      selection.anchor = { col: start.col, row: start.row };
      selection.focus = {
        col: Math.min(start.col + matrix[0].length - 1, dims.cols - 1),
        row: Math.min(start.row + matrix.length - 1, dims.rows)
      };
      extras = [];
      paintSelection();
      setTip(blocked ? '已粘贴 ' + written + ' 格，' + blocked + ' 格落在参考区，跳过了。' : '已粘贴 ' + written + ' 格。');
      onChange('paste');
    }

    /* --- 填充柄 --- */

    function positionHandle() {
      var r = selRange();
      var td = tdAt(r.end.col, r.end.row);
      if (!td || !editable || extras.length) {
        handle.style.display = 'none';
        return;
      }
      handle.style.display = 'block';
      // handle 是 14×14 的透明命中区，中间画 7px 的小方块，视觉跟 Excel 一样但好点得多
      handle.style.left = (td.offsetLeft + td.offsetWidth - 7) + 'px';
      handle.style.top = (td.offsetTop + td.offsetHeight - 7) + 'px';
    }

    function rectBox(r) {
      var a = tdAt(r.start.col, r.start.row);
      var b = tdAt(r.end.col, r.end.row);
      if (!a || !b) return null;
      var left = Math.min(a.offsetLeft, b.offsetLeft);
      var top = Math.min(a.offsetTop, b.offsetTop);
      var right = Math.max(a.offsetLeft + a.offsetWidth, b.offsetLeft + b.offsetWidth);
      var bottom = Math.max(a.offsetTop + a.offsetHeight, b.offsetTop + b.offsetHeight);
      return { left: left, top: top, width: right - left, height: bottom - top };
    }

    function showPreview(r, label) {
      var box = rectBox(r);
      if (!box) return;
      preview.style.display = 'block';
      preview.style.left = box.left + 'px';
      preview.style.top = box.top + 'px';
      preview.style.width = box.width + 'px';
      preview.style.height = box.height + 'px';
      bubble.style.display = 'block';
      bubble.textContent = label;
      bubble.style.left = box.left + 'px';
      bubble.style.top = Math.max(0, box.top - 24) + 'px';
    }

    function hidePreview() {
      preview.style.display = 'none';
      bubble.style.display = 'none';
    }

    function fillPreviewPlan(axis, count) {
      var src = selRange();
      var sources = [];
      var dates = [];
      if (axis === 'row') {
        for (var row = src.start.row; row <= src.end.row; row++) {
          sources.push(rawOf(src.start.col, row));
          dates.push(wb.isDateCell(src.start.col, row));
        }
      } else {
        for (var col = src.start.col; col <= src.end.col; col++) {
          sources.push(rawOf(col, src.start.row));
          dates.push(wb.isDateCell(col, src.start.row));
        }
      }
      return Mobai.Fill.planFill(sources, count, { dates: dates, axis: axis });
    }

    function fillTargetRange(axis, count) {
      var src = selRange();
      if (axis === 'row') {
        return {
          start: { col: src.start.col, row: src.end.row + 1 },
          end: { col: src.end.col, row: src.end.row + count }
        };
      }
      return {
        start: { col: src.end.col + 1, row: src.start.row },
        end: { col: src.end.col + count, row: src.end.row }
      };
    }

    /** 双击填充柄：填到相邻列/行数据的末行 */
    function autoFillEnd() {
      var src = selRange();
      var cols = [src.start.col - 1, src.end.col + 1];
      var end = src.end.row;
      for (var i = 0; i < cols.length && end === src.end.row; i++) {
        var c = cols[i];
        if (c < 0 || c > dims.cols - 1) continue;
        var last = src.start.row - 1;
        for (var row = src.start.row; row <= dims.rows; row++) {
          var e = wb.entry(c, row);
          var has = e && (e.kind !== 'empty' || e.style) && wb.rawText(c, row) !== '';
          if (!has) break;
          last = row;
        }
        if (last >= src.start.row) end = last;
      }
      if (end <= src.end.row) {
        setTip('双击填充要看向相邻列的连续数据，这次旁边没有可参照的长度。用鼠标拖填充柄也一样。', 'warn');
        return;
      }
      fillBy('row', end - src.end.row);
    }

    function startFillDrag() {
      var r = selRange();
      fillDrag = {
        src: { start: { col: r.start.col, row: r.start.row }, end: { col: r.end.col, row: r.end.row } },
        axis: 'row',
        count: 0
      };
    }

    function updateFillDrag(cell) {
      if (!fillDrag || !cell) return;
      var src = fillDrag.src;
      var down = cell.row - src.end.row;
      var right = cell.col - src.end.col;
      var axis;
      var count;
      if (Math.abs(down) >= Math.abs(right)) {
        axis = 'row';
        count = Math.max(0, down);
      } else {
        axis = 'col';
        count = Math.max(0, right);
      }
      fillDrag.axis = axis;
      fillDrag.count = count;
      if (!count) {
        hidePreview();
        return;
      }
      var plan = fillPreviewPlan(axis, count);
      showPreview(fillTargetRange(axis, count), Mobai.Fill.previewText(plan));
    }

    function endFillDrag() {
      if (!fillDrag) return;
      var fd = fillDrag;
      fillDrag = null;
      hidePreview();
      if (fd.count > 0) fillBy(fd.axis, fd.count);
      else positionHandle();
    }

    /* --- 公式点选引用 --- */

    function hideRefBox() {
      refBox.style.display = 'none';
      refTip.style.display = 'none';
    }

    function refRect() {
      if (!refPoint) return null;
      var a = refPoint.anchor;
      var b = refPoint.focus || a;
      if (!a) return null;
      return U.normRange(a, b);
    }

    var ABS_STEPS = [
      { colAbs: false, rowAbs: false },
      { colAbs: true, rowAbs: true },
      { colAbs: false, rowAbs: true },
      { colAbs: true, rowAbs: false }
    ];

    function refTextOf() {
      var r = refRect();
      if (!r) return '';
      var a = ABS_STEPS[refPoint.abs % 4];
      var s = (a.colAbs ? '$' : '') + U.colToName(r.start.col) + (a.rowAbs ? '$' : '') + r.start.row;
      var e = (a.colAbs ? '$' : '') + U.colToName(r.end.col) + (a.rowAbs ? '$' : '') + r.end.row;
      return s === e ? s : s + ':' + e;
    }

    function updatePointRef() {
      if (!refPoint) return;
      var text = refTextOf();
      var input = refPoint.input;
      var v = input.value;
      var at = Math.min(refPoint.at, v.length);
      var rest = v.slice(at + refPoint.len);
      input.value = v.slice(0, at) + text + rest;
      refPoint.len = text.length;
      try { input.setSelectionRange(at + text.length, at + text.length); } catch (ex) { /* 忽略 */ }
      drawRefBox();
    }

    function drawRefBox() {
      var r = refRect();
      if (!r) { hideRefBox(); return; }
      var box = rectBox(r);
      if (!box) { hideRefBox(); return; }
      refBox.style.display = 'block';
      refBox.style.left = box.left + 'px';
      refBox.style.top = box.top + 'px';
      refBox.style.width = box.width + 'px';
      refBox.style.height = box.height + 'px';
      refTip.style.display = 'block';
      var rows = r.end.row - r.start.row + 1;
      var cols = r.end.col - r.start.col + 1;
      refTip.textContent = cols + 'C × ' + rows + 'R';
      refTip.style.left = (box.left + box.width - 46) + 'px';
      refTip.style.top = (box.top + box.height + 2) + 'px';
    }

    function startPoint(hit) {
      var input = edit && edit.input;
      if (!input) return;
      if (!refPoint || refPoint.input !== input) {
        refPoint = {
          input: input,
          at: input.selectionStart === null ? input.value.length : input.selectionStart,
          len: 0,
          anchor: null,
          focus: null,
          whole: null,
          abs: 0
        };
      }
      refPoint.abs = 0;
      var used = wb.usedRange();
      if (hit && hit.axis === 'row') {
        refPoint.whole = 'row';
        refPoint.anchor = { col: Math.max(0, used.minCol), row: hit.index };
        refPoint.focus = { col: Math.max(used.maxCol, Math.max(0, used.minCol)), row: hit.index };
      } else if (hit && hit.axis === 'col') {
        refPoint.whole = 'col';
        refPoint.anchor = { col: hit.index, row: 1 };
        refPoint.focus = { col: hit.index, row: Math.max(1, used.maxRow) };
      } else if (hit) {
        refPoint.whole = null;
        var m = wb.mergeOf(hit.col, hit.row);
        var r = m || { start: { col: hit.col, row: hit.row }, end: { col: hit.col, row: hit.row } };
        refPoint.anchor = { col: r.start.col, row: r.start.row };
        refPoint.focus = { col: r.end.col, row: r.end.row };
      } else {
        refPoint.anchor = { col: selection.focus.col, row: selection.focus.row };
        refPoint.focus = null;
      }
      updatePointRef();
      pointDrag = true;
    }

    function updatePointFocus(hit) {
      if (!refPoint || !hit || refPoint.whole) return;
      if (hit.axis) return;
      var m = wb.mergeOf(hit.col, hit.row);
      var e = m ? { col: m.end.col, row: m.end.row } : hit;
      refPoint.focus = e;
      updatePointRef();
    }

    function cycleRefAbs() {
      if (!refPoint) return false;
      refPoint.abs = (refPoint.abs + 1) % 4;
      updatePointRef();
      return true;
    }

    /* --- 右键菜单 --- */

    function closeMenu() {
      if (!menu) return;
      if (menu.parentNode) menu.parentNode.removeChild(menu);
      menu = null;
    }

    function openMenu(x, y) {
      closeMenu();
      commitPending();
      var items = [
        { label: '剪切', run: function () { copySelection(true); } },
        { label: '复制', run: function () { copySelection(false); } },
        { label: '粘贴', run: function () { pasteFromMenu(); } },
        { sep: true },
        { label: '插入行', run: function () { insertRows(1); } },
        { label: '插入列', run: function () { insertCols(1); } },
        { label: '删除行', run: function () { deleteRows(); } },
        { label: '删除列', run: function () { deleteCols(); } },
        { sep: true },
        { label: '清除内容', run: clearSelection },
        { label: '清除格式', run: function () { clearFormat(); } },
        { sep: true },
        {
          label: wb.mergeOf(selection.focus.col, selection.focus.row) ? '取消合并' : '合并居中',
          run: function () {
            if (wb.mergeOf(selection.focus.col, selection.focus.row)) unmerge();
            else mergeCenter();
          }
        }
      ];
      menu = el('div', 'sheet-menu');
      items.forEach(function (it) {
        if (it.sep) {
          menu.appendChild(el('div', 'sheet-menu-sep'));
          return;
        }
        var btn = el('button', 'sheet-menu-item', it.label);
        btn.type = 'button';
        btn.addEventListener('click', function () {
          closeMenu();
          it.run();
        });
        menu.appendChild(btn);
      });
      document.body.appendChild(menu);
      var w = menu.offsetWidth;
      var hgt = menu.offsetHeight;
      menu.style.left = Math.max(4, Math.min(x, window.innerWidth - w - 6)) + 'px';
      menu.style.top = Math.max(4, Math.min(y, window.innerHeight - hgt - 6)) + 'px';
    }

    function pasteFromMenu() {
      if (clipboard) pasteTC(clipboard.text);
      else setTip('浏览器不让网页直接读系统剪贴板，用 Ctrl+V 粘贴吧。', 'warn');
    }

    /* --- 命令（功能区调用） --- */

    function act(kind, label, fn) {
      commitPending();
      pushUndo();
      var note = fn();
      render();
      setTip(typeof note === 'string' ? note : label);
      if (kind !== 'style' && kind !== 'format') onChange(kind);
      else onChange(kind);
    }

    function selectionOrCell() {
      return selRange();
    }

    function applyStyle(patch, label) {
      act('style', label, function () {
        wb.setStyle(selectionOrCell(), patch);
        return label;
      });
    }

    function toggleStyle(key) {
      var st = wb.styleOf(selection.focus.col, selection.focus.row) || {};
      var on = !st[key];
      var patch = {};
      patch[key] = on ? true : null;
      applyStyle(patch, (on ? '已加上' : '已去掉') + ({ bold: '加粗', italic: '倾斜', underline: '下划线' }[key] || key));
    }

    function setFormat(code) {
      act('format', '已设置数字格式：' + code, function () {
        wb.setFormat(selectionOrCell(), code);
        return '已设置数字格式：' + code + '（只影响显示，判分还是看原始值）。';
      });
    }

    function stepDecimals(delta) {
      var cur = wb.formatOf(selection.focus.col, selection.focus.row) || 'General';
      if (!Mobai.Format || Mobai.Format.isGeneral(cur) || Mobai.Format.isDate(cur)) {
        cur = '0.00';
      }
      var next = Mobai.Format.withDecimals(cur, delta);
      setFormat(next);
    }

    function setAlign(h, v) {
      var patch = {};
      if (h) patch.alignH = h;
      if (v) patch.alignV = v;
      applyStyle(patch, '已设置对齐');
    }

    function toggleWrap() {
      var st = wb.styleOf(selection.focus.col, selection.focus.row) || {};
      applyStyle({ wrap: st.wrap ? null : true }, st.wrap ? '已取消自动换行' : '已自动换行');
    }

    function mergeCenter() {
      var r = selRange();
      if (U.rangeCount(r) < 2) {
        setTip('合并至少要选两格。', 'warn');
        return;
      }
      act('merge', '已合并居中', function () {
        wb.addMerge(r);
        wb.setStyle(r, { alignH: 'center', alignV: 'middle' });
        return '已合并 ' + U.rangeToString(r) + ' 并居中。';
      });
    }

    function unmerge() {
      act('merge', '已取消合并', function () {
        var f = selection.focus;
        var m = wb.mergeOf(f.col, f.row);
        if (m) wb.removeMergeAt(f.col, f.row);
        return '已取消合并。';
      });
    }

    function clearFormat() {
      act('style', '已清除格式', function () {
        var targets = editableCellsIn([selectionOrCell()]);
        targets.forEach(function (c) { wb.clearFormat(c.col, c.row); });
        return '已清除 ' + targets.length + ' 格的格式（内容还在）。';
      });
    }

    function insertRows(n) {
      var at = selRange().start.row;
      act('insert', '已插入行', function () {
        wb.insertLines('row', at, n || 1);
        return '已在第 ' + at + ' 行前插入 ' + (n || 1) + ' 行，公式引用跟着改了。';
      });
    }

    function insertCols(n) {
      var at = selRange().start.col;
      act('insert', '已插入列', function () {
        wb.insertLines('col', at, n || 1);
        return '已在 ' + U.colToName(at) + ' 列前插入 ' + (n || 1) + ' 列。';
      });
    }

    function deleteRows() {
      var r = selRange();
      var n = r.end.row - r.start.row + 1;
      act('delete', '已删除行', function () {
        wb.deleteLines('row', r.start.row, n);
        return '已删除第 ' + r.start.row + ' 行起的 ' + n + ' 行；指向它们的引用变成了 #REF!。';
      });
    }

    function deleteCols() {
      var r = selRange();
      var n = r.end.col - r.start.col + 1;
      act('delete', '已删除列', function () {
        wb.deleteLines('col', r.start.col, n);
        return '已删除 ' + U.colToName(r.start.col) + ' 列起的 ' + n + ' 列。';
      });
    }

    /* --- 事件 --- */

    function hitFromNode(node) {
      while (node && node !== grid) {
        if (node.classList && node.classList.contains('sheet-col-resize')) return 'resize';
        if (node.dataset) {
          if (node.dataset.corner !== undefined) return { axis: 'all' };
          if (node.dataset.rowhead !== undefined) return { axis: 'row', index: parseInt(node.dataset.rowhead, 10) };
          if (node.dataset.colhead !== undefined) return { axis: 'col', index: parseInt(node.dataset.colhead, 10) };
          if (node.dataset.col !== undefined && node.dataset.row !== undefined) {
            return { col: parseInt(node.dataset.col, 10), row: parseInt(node.dataset.row, 10) };
          }
        }
        node = node.parentNode;
      }
      return null;
    }

    function hitFromEvent(e) {
      return hitFromNode(e.target);
    }

    function hitFromPoint(x, y) {
      return hitFromNode(document.elementFromPoint(x, y));
    }

    function cellOf(hit) {
      if (!hit || hit.axis) return null;
      var m = wb.mergeOf(hit.col, hit.row);
      if (m) return { col: m.start.col, row: m.start.row };
      return hit;
    }

    grid.addEventListener('mousedown', function (e) {
      if (e.button === 2) return;
      // 在编辑框内部点：是想挪光标，不能当成「点选引用」，否则会平白插入一个当前格的引用
      if (editor.contains(e.target)) return;
      var hit = hitFromEvent(e);
      if (hit === 'resize') return;

      // 编辑态 + 公式：点格子 = 把引用点进公式
      if (edit && edit.input && canPoint(edit.input)) {
        e.preventDefault();
        var cell = cellOf(hit);
        startPoint(hit && hit.axis ? hit : cell);
        return;
      }
      if (edit) commitEdit(null, true);

      e.preventDefault();
      closeMenu();
      if (!hit) return;

      if (hit.axis === 'all') { selectAll(); grid.focus(); return; }
      if (hit.axis === 'row') {
        var fullRow = { start: { col: 0, row: hit.index }, end: { col: dims.cols - 1, row: hit.index } };
        if (e.shiftKey) selection.focus = { col: dims.cols - 1, row: hit.index };
        else setSelection(fullRow);
        paintSelection();
        scrollIntoView(0, hit.index);
        grid.focus();
        return;
      }
      if (hit.axis === 'col') {
        var fullCol = { start: { col: hit.index, row: 1 }, end: { col: hit.index, row: dims.rows } };
        if (e.shiftKey) selection.focus = { col: hit.index, row: dims.rows };
        else setSelection(fullCol);
        paintSelection();
        scrollIntoView(hit.index, 1);
        grid.focus();
        return;
      }

      var c = cellOf(hit);
      var mr = expandToMerge({ col: c.col, row: c.row });
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
        extras.push(selRange());
        setSelection(mr, true);
      } else if (e.shiftKey) {
        selection.focus = { col: mr.end.col, row: mr.end.row };
        paintSelection();
      } else {
        setSelection(mr);
      }
      dragSel = true;
      grid.focus();
      setTip(U.addrOf(c.col, c.row) + (wb.isReadonly(c.col, c.row) ? ' 是参考数据，只能看不能改。' : ''));
    });

    grid.addEventListener('dblclick', function (e) {
      if (!editable) return;
      var hit = hitFromEvent(e);
      var c = cellOf(hit);
      if (!c) return;
      e.preventDefault();
      setSelection(expandToMerge(c));
      startEditCell();
    });

    grid.addEventListener('contextmenu', function (e) {
      var hit = hitFromEvent(e);
      var c = cellOf(hit);
      if (!c) return;
      e.preventDefault();
      closeMenu();
      if (!(c.col >= selRange().start.col && c.col <= selRange().end.col &&
        c.row >= selRange().start.row && c.row <= selRange().end.row)) {
        setSelection(expandToMerge(c));
      }
      openMenu(e.clientX, e.clientY);
    });

    handle.addEventListener('mousedown', function (e) {
      if (e.button !== 0) return;
      if (e.detail >= 2) return;
      e.preventDefault();
      e.stopPropagation();
      startFillDrag();
      document.body.classList.add('is-fill-dragging');
    });

    handle.addEventListener('dblclick', function (e) {
      e.preventDefault();
      e.stopPropagation();
      autoFillEnd();
    });

    window.addEventListener('mousemove', function (e) {
      if (fillDrag) {
        updateFillDrag(cellOf(hitFromPoint(e.clientX, e.clientY)));
        return;
      }
      if (pointDrag && refPoint) {
        updatePointFocus(hitFromPoint(e.clientX, e.clientY));
        return;
      }
      if (dragSel) {
        var c = cellOf(hitFromPoint(e.clientX, e.clientY));
        if (!c) return;
        var m = wb.mergeOf(selection.anchor.col, selection.anchor.row);
        if (m) selection.anchor = { col: m.start.col, row: m.start.row };
        selection.focus = { col: c.col, row: c.row };
        paintSelection();
      }
    });

    window.addEventListener('mouseup', function (e) {
      if (fillDrag) {
        document.body.classList.remove('is-fill-dragging');
        endFillDrag();
        return;
      }
      if (pointDrag) {
        pointDrag = false;
        if (refPoint && refPoint.input) {
          refPoint.input.focus();
          try { refPoint.input.setSelectionRange(refPoint.at + refPoint.len, refPoint.at + refPoint.len); } catch (ex) { /* 忽略 */ }
        }
        return;
      }
      if (dragSel) {
        dragSel = false;
        var r = selRange();
        if (extras.length) setTip('按住 Ctrl 加选了 ' + (extras.length + 1) + ' 块区域。');
        else if (isMulti()) setTip('选中 ' + U.rangeToString(r) + '，右下角的小方块可以拖动填充。');
        else setTip('');
      }
    });

    /* --- 键盘 --- */

    function editorKeydown(e) {
      // 提交之后 edit 已经变成 null，同一次按键会再冒泡到 grid 上被当成「移动一格」，
      // 结果是回车跳两格。这里直接把事件截住。
      e.stopPropagation();
      var ctrl = e.ctrlKey || e.metaKey;
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit({ row: e.shiftKey ? -1 : 1 });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit({ col: e.shiftKey ? -1 : 1 });
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
        return;
      }
      if (e.key === 'F4') {
        if (cycleRefAbs()) e.preventDefault();
        return;
      }
      if (ctrl && (e.key === 'a' || e.key === 'A')) {
        e.preventDefault();
        var input = edit.input;
        input.setSelectionRange(0, input.value.length);
        return;
      }
      if (ctrl && (e.key === 'z' || e.key === 'Z')) {
        e.preventDefault();
        undo();
      }
    }

    function bindEditInput(input, isCell) {
      input.addEventListener('keydown', editorKeydown);
      input.addEventListener('input', function () { refPoint = null; });
      input.addEventListener('blur', function () {
        if (suppressBlur) return;
        if (!edit || edit.input !== input) return;
        commitEdit(null, true);
      });
    }

    bindEditInput(editor, true);

    formulaBar.addEventListener('focus', function () {
      if (edit) return;
      if (!editable) return;
      var f = selection.focus;
      if (wb.isReadonly(f.col, f.row)) {
        setTip('这一格是参考数据，不能改。', 'warn');
        suppressBlur = true;
        formulaBar.blur();
        suppressBlur = false;
        return;
      }
      edit = { mode: 'bar', input: formulaBar, col: f.col, row: f.row };
      setTip('正在编辑 ' + U.addrOf(f.col, f.row) + '：写公式时可以直接点单元格把引用点进来。');
    });
    bindEditInput(formulaBar, false);

    grid.addEventListener('keydown', function (e) {
      if (edit && e.target === editor) return;   // 交给编辑器
      var key = e.key;
      var ctrl = e.ctrlKey || e.metaKey;

      if (ctrl && (key === 'z' || key === 'Z')) {
        e.preventDefault();
        if (e.shiftKey) redo(); else undo();
        return;
      }
      if (ctrl && (key === 'y' || key === 'Y')) { e.preventDefault(); redo(); return; }
      if (ctrl && (key === 'd' || key === 'D')) { e.preventDefault(); fillDown(); return; }
      if (ctrl && (key === 'a' || key === 'A')) { e.preventDefault(); selectAll(); return; }
      if (ctrl && (key === 'b' || key === 'B')) { e.preventDefault(); toggleStyle('bold'); return; }
      if (ctrl && (key === 'i' || key === 'I')) { e.preventDefault(); toggleStyle('italic'); return; }
      if (ctrl && (key === 'u' || key === 'U')) { e.preventDefault(); toggleStyle('underline'); return; }
      if (ctrl && (key === 'c' || key === 'C')) { copySelection(false); return; }
      if (ctrl && (key === 'x' || key === 'X')) { copySelection(true); return; }
      if (ctrl && (key === 'v' || key === 'V')) { e.preventDefault(); return; }  // 交给 paste 事件

      if (key === 'ArrowUp') { e.preventDefault(); e.shiftKey ? extendFocus({ row: -1 }) : moveFocus({ row: -1 }); return; }
      if (key === 'ArrowDown') { e.preventDefault(); e.shiftKey ? extendFocus({ row: 1 }) : moveFocus({ row: 1 }); return; }
      if (key === 'ArrowLeft') { e.preventDefault(); e.shiftKey ? extendFocus({ col: -1 }) : moveFocus({ col: -1 }); return; }
      if (key === 'ArrowRight') { e.preventDefault(); e.shiftKey ? extendFocus({ col: 1 }) : moveFocus({ col: 1 }); return; }
      if (key === 'Tab') { e.preventDefault(); moveFocus({ col: e.shiftKey ? -1 : 1 }); return; }
      if (key === 'Enter') {
        e.preventDefault();
        moveFocus({ row: e.shiftKey ? -1 : 1 });
        return;
      }
      if (key === 'F2') { e.preventDefault(); startEditCell(); return; }
      if (key === 'F4') { e.preventDefault(); return; }
      if (key === 'Escape') { cancelEdit(); closeMenu(); return; }
      if (key === 'Delete' || key === 'Backspace') { e.preventDefault(); clearSelection(); return; }
      if (key === 'Home') { e.preventDefault(); moveFocus({ col: -selection.focus.col }); return; }
      if (key === 'End') { e.preventDefault(); moveFocus({ col: dims.cols - 1 - selection.focus.col }); return; }

      // 直接打字：进入编辑态，第一个字符带进去
      if (!ctrl && !e.altKey && key.length === 1) {
        e.preventDefault();
        startEditCell(key);
      }
    });

    grid.addEventListener('copy', function (e) {
      var text = copySelection(false);
      if (e.clipboardData) {
        e.clipboardData.setData('text/plain', text);
        e.preventDefault();
      }
    });

    grid.addEventListener('cut', function (e) {
      var text = copySelection(true);
      if (e.clipboardData) {
        e.clipboardData.setData('text/plain', text);
        e.preventDefault();
      }
    });

    grid.addEventListener('paste', function (e) {
      e.preventDefault();
      var text = '';
      if (e.clipboardData) text = e.clipboardData.getData('text/plain');
      if (!text && clipboard) text = clipboard.text;
      pasteTC(text);
    });

    addrBox.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      var parsed = U.rangeOf(addrBox.value);
      if (!parsed) {
        addrBox.value = addrLabel();
        setTip('名称框里写 A1 或 A1:C10 这样的区域。', 'warn');
        return;
      }
      commitPending();
      setSelection(parsed);
      scrollIntoView(parsed.start.col, parsed.start.row);
      grid.focus();
    });

    addrBox.addEventListener('blur', function () { addrBox.value = addrLabel(); });

    // 列宽拖拽 + 双击自适应
    var resizing = null;
    thead.addEventListener('mousedown', function (e) {
      var rs = e.target;
      if (!rs.classList || !rs.classList.contains('sheet-col-resize')) return;
      e.preventDefault();
      e.stopPropagation();
      var col = parseInt(rs.dataset.col, 10);
      resizing = { col: col, startX: e.clientX, startW: colWidth(col), moved: false };
      document.body.classList.add('is-col-resizing');
    });
    window.addEventListener('mousemove', function (e) {
      if (!resizing) return;
      resizing.moved = true;
      widths[resizing.col] = Math.max(MIN_COL_WIDTH, resizing.startW + (e.clientX - resizing.startX));
      var cols = colgroup.children;
      if (cols[resizing.col + 1]) cols[resizing.col + 1].style.width = widths[resizing.col] + 'px';
      if (edit && edit.mode === 'cell') positionEditor();
    });
    window.addEventListener('mouseup', function () {
      if (!resizing) return;
      if (!resizing.moved) autoFitWidth(resizing.col);
      resizing = null;
      document.body.classList.remove('is-col-resizing');
    });

    function autoFitWidth(col) {
      var w = MIN_COL_WIDTH;
      for (var row = 1; row <= dims.rows; row++) {
        var text = wb.display(col, row);
        if (!text) continue;
        var len = 0;
        for (var i = 0; i < text.length; i++) len += text.charCodeAt(i) > 255 ? 2 : 1;
        w = Math.max(w, Math.min(320, len * 7.4 + 18));
      }
      widths[col] = Math.round(w);
      var cols = colgroup.children;
      if (cols[col + 1]) cols[col + 1].style.width = widths[col] + 'px';
    }

    // 工作表标签
    tabGrid.addEventListener('click', function () {
      grid.style.display = 'block';
      note.style.display = 'none';
      tabGrid.classList.add('is-active');
      if (tabNote) tabNote.classList.remove('is-active');
    });
    if (tabNote) {
      tabNote.addEventListener('click', function () {
        commitPending();
        closeMenu();
        grid.style.display = 'none';
        note.style.display = 'block';
        tabNote.classList.add('is-active');
        tabGrid.classList.remove('is-active');
      });
    }

    // 点空白处关掉右键菜单
    document.addEventListener('mousedown', function (e) {
      if (!menu) return;
      if (menu.contains(e.target)) return;
      closeMenu();
    });
    window.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') closeMenu();
    });
    window.addEventListener('resize', function () { closeMenu(); });

    /* --- 对外 --- */

    function stateOfSelection() {
      var f = selection.focus;
      var st = wb.styleOf(f.col, f.row) || {};
      return {
        addr: addrLabel(),
        bold: !!st.bold,
        italic: !!st.italic,
        underline: !!st.underline,
        wrap: !!st.wrap,
        alignH: st.alignH || '',
        alignV: st.alignV || '',
        fontName: st.fontName || '',
        fontSize: st.fontSize || 0,
        format: wb.formatOf(f.col, f.row) || 'General',
        merged: !!wb.mergeOf(f.col, f.row),
        readonlyCell: wb.isReadonly(f.col, f.row),
        canUndo: history.length > 0,
        canRedo: redoStack.length > 0,
        multi: isMulti()
      };
    }

    var api = {
      render: render,
      refresh: render,
      undo: undo,
      redo: redo,
      focus: function () { grid.focus(); },
      selection: selRange,
      setTip: setTip,
      selectAddr: function (str) {
        var parsed = U.rangeOf(str);
        if (!parsed) return false;
        setSelection(parsed);
        scrollIntoView(parsed.start.col, parsed.start.row);
        return true;
      },
      setEditable: function (v) { editable = !!v; },
      state: stateOfSelection,
      onState: null,
      commands: {
        undo: undo,
        redo: redo,
        cut: function () { copySelection(true); },
        copy: function () { copySelection(false); },
        paste: pasteTC,
        pasteInternal: pasteFromMenu,
        toggleBold: function () { toggleStyle('bold'); },
        toggleItalic: function () { toggleStyle('italic'); },
        toggleUnderline: function () { toggleStyle('underline'); },
        setStyle: applyStyle,
        setFontName: function (name) { applyStyle({ fontName: name || null }, '已改字体'); },
        setFontSize: function (size) { applyStyle({ fontSize: size ? Number(size) : null }, '已改字号'); },
        setFontColor: function (color) { applyStyle({ fg: color }, '已改字体颜色'); },
        setFillColor: function (color) { applyStyle({ bg: color }, '已改填充色'); },
        setAlign: setAlign,
        toggleWrap: toggleWrap,
        mergeCenter: mergeCenter,
        unmerge: unmerge,
        isMerged: function () { return !!wb.mergeOf(selection.focus.col, selection.focus.row); },
        setFormat: setFormat,
        stepDecimals: stepDecimals,
        clearFormat: clearFormat,
        clearContents: clearSelection,
        insertRows: insertRows,
        insertCols: insertCols,
        deleteRows: deleteRows,
        deleteCols: deleteCols,
        focusNameBox: function () { addrBox.focus(); addrBox.select(); },
        selectAll: selectAll,
        fillDown: fillDown,
        freeze: function () { /* 第二批 */ }
      },
      workbook: wb
    };

    render();
    setTip('直接打字或按 F2 编辑；写公式时点单元格就能把引用点进来（F4 切 $），' +
      '选区右下角的小方块可以拖动填充。');

    return api;
  }

  Mobai.Sheet = {
    mount: mount,
    shiftFormulaText: function (text, dCol, dRow) {
      return Mobai.Fill.shiftFormulaText(text, dCol, dRow);
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);