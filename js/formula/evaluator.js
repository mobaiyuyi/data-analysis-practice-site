/* 研数通途 · 求值器与工作簿
 * 求值是「单元格 → 值」的函数：每次编辑后全量重算（表很小，不需要依赖图）。
 * 循环引用在求值栈上检测。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;
  var E = Mobai.Errors;
  var C = E.CODES;
  var F = Mobai.Functions;

  var MAX_RANGE_CELLS = 20000;

  /* ---------- 值比较 ---------- */

  function coerceForCompare(v, other) {
    if (v !== null && v !== undefined) return v;
    if (typeof other === 'number') return 0;
    if (typeof other === 'boolean') return false;
    return '';
  }

  function typeRank(v) {
    if (typeof v === 'number') return 1;
    if (typeof v === 'string') return 2;
    if (typeof v === 'boolean') return 3;
    return 0;
  }

  /** Excel 的大小规则：数字 < 文本 < FALSE < TRUE */
  function compareValues(a, b) {
    a = coerceForCompare(a, b);
    b = coerceForCompare(b, a);
    var ra = typeRank(a);
    var rb = typeRank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 1) return a === b ? 0 : (a < b ? -1 : 1);
    if (ra === 2) {
      var la = a.toLowerCase();
      var lb = b.toLowerCase();
      return la === lb ? 0 : (la < lb ? -1 : 1);
    }
    return a === b ? 0 : (a ? 1 : -1);
  }

  /* ---------- 求值 ---------- */

  function evalBinary(node, ctx) {
    var l = evalNode(node.left, ctx);
    if (E.isError(l)) return l;
    var r = evalNode(node.right, ctx);
    if (E.isError(r)) return r;
    var op = node.op;

    if (op === '&') {
      var lt = F.toText(l);
      if (E.isError(lt)) return lt;
      var rt = F.toText(r);
      if (E.isError(rt)) return rt;
      return lt + rt;
    }

    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      var c = compareValues(l, r);
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        default: return c >= 0;
      }
    }

    var ln = F.toNumber(l, {});
    if (E.isError(ln)) return ln;
    var rn = F.toNumber(r, {});
    if (E.isError(rn)) return rn;
    switch (op) {
      case '+': return ln + rn;
      case '-': return ln - rn;
      case '*': return ln * rn;
      case '/':
        if (rn === 0) return E.err(C.DIV0, {});
        return ln / rn;
      case '^': {
        var p = Math.pow(ln, rn);
        if (!isFinite(p) || isNaN(p)) return E.err(C.NUM, {});
        return p;
      }
      default:
        return E.err(C.VALUE, {});
    }
  }

  function evalNode(node, ctx) {
    if (!node) return null;
    switch (node.type) {
      case 'Number':
        return node.value;
      case 'String':
        return node.value;
      case 'Bool':
        return node.value;
      case 'Error':
        return E.err(node.value, {});
      case 'Name':
        return E.err(C.NAME, { name: node.name });
      case 'Ref':
        return ctx.getCell(node.ref.col, node.ref.row);
      case 'Range':
        // 区域出现在标量语境（比如 =A1:A5+1）
        return E.err(C.VALUE, { range: node.raw });
      case 'Binary':
        return evalBinary(node, ctx);
      case 'Unary': {
        var v = evalNode(node.operand, ctx);
        if (E.isError(v)) return v;
        var n = F.toNumber(v, {});
        if (E.isError(n)) return n;
        return node.op === '-' ? -n : n;
      }
      case 'Percent': {
        var pv = evalNode(node.operand, ctx);
        if (E.isError(pv)) return pv;
        var pn = F.toNumber(pv, {});
        if (E.isError(pn)) return pn;
        return pn / 100;
      }
      case 'Func':
        return F.call(node, ctx);
      default:
        return E.err(C.VALUE, {});
    }
  }

  /* ---------- 单元格内容 ---------- */

  var STYLE_KEYS = ['bold', 'italic', 'underline', 'alignH', 'alignV', 'wrap', 'bg', 'fg', 'fontName', 'fontSize'];

  /** 从题库写法里挑出样式字段；align 是 alignH 的简写 */
  function pickStyle(raw) {
    var st = null;
    for (var i = 0; i < STYLE_KEYS.length; i++) {
      var k = STYLE_KEYS[i];
      if (raw[k] !== undefined) { st = st || {}; st[k] = raw[k]; }
    }
    if (raw.align !== undefined) { st = st || {}; st.alignH = raw.align; }
    return st;
  }

  function normalizeEntry(raw) {
    if (raw && typeof raw === 'object') {
      var fmt = raw.format || raw.fmt;
      var st = pickStyle(raw);
      if (raw.d !== undefined) {
        // {d:'2026-03-09'} —— 题目里写日期用这个，省得手算序列号
        var s = U.serialFromISO(raw.d);
        var dEntry = { kind: 'number', value: s === null ? 0 : s, format: fmt || 'date' };
        if (st) dEntry.style = st;
        return dEntry;
      }
      var out;
      if (raw.f !== undefined) out = { kind: 'formula', formula: String(raw.f) };
      else if (raw.v !== undefined) out = normalizeEntry(raw.v);
      else out = { kind: 'text', value: '' };
      if (fmt) out.format = fmt;
      if (st) out.style = st;
      return out;
    }
    if (typeof raw === 'number') return { kind: 'number', value: raw };
    if (typeof raw === 'boolean') return { kind: 'bool', value: raw };
    var s = String(raw);
    if (s.charAt(0) === '=') return { kind: 'formula', formula: s };
    if (s.charAt(0) === "'") return { kind: 'text', value: s.slice(1) };
    var t = s.trim();
    if (t !== '' && /^-?\d+(\.\d+)?$/.test(t)) return { kind: 'number', value: parseFloat(t) };
    return { kind: 'text', value: s };
  }

  /* ---------- 工作簿 ---------- */

  /**
   * @param {Object} cells 初始单元格，如 { A1: '日期', B2: 1200, C2: '=A2*B2', D2: {v: 45000, format:'date'} }
   * @param {Object} opts  { today:'2026-03-16', dates:['A2:A9'], readonly:['A1:F1'] }
   */
  function Workbook(cells, opts) {
    opts = opts || {};
    this.cells = Object.create(null);
    this.workbookOpts = opts;
    this.readonly = opts.readonly || [];
    this.dates = opts.dates || [];
    this.bold = opts.bold || [];
    this.referenceTables = opts.referenceTables || [];
    this.merges = (opts.merges || []).map(function (m) {
      return typeof m === 'string' ? U.rangeOf(m) : U.normRange(m.start, m.end);
    }).filter(Boolean);
    this.todaySerial = opts.today ? U.serialFromISO(opts.today) : U.nowSerial();
    this._cache = Object.create(null);
    this._stack = [];
    if (cells) {
      var self = this;
      Object.keys(cells).forEach(function (addr) {
        var a = U.parseAddr(addr);
        if (a) self.cells[addr] = normalizeEntry(cells[addr]);
      });
    }
    this.recalc();
  }

  Workbook.prototype.recalc = function () {
    this._cache = Object.create(null);
  };

  Workbook.prototype.entry = function (col, row) {
    return this.cells[U.addrOf(col, row)] || null;
  };

  Workbook.prototype.isReadonly = function (col, row) {
    return U.addrInRanges(col, row, this.readonly);
  };

  Workbook.prototype.isDateCell = function (col, row) {
    var e = this.entry(col, row);
    var Fm = Mobai.Format;
    if (e && e.format && (Fm ? Fm.isDate(e.format) : e.format === 'date')) return true;
    return U.addrInRanges(col, row, this.dates);
  };

  /** 求值：单元格的计算结果（空单元格返回 null） */
  Workbook.prototype.getCell = function (col, row) {
    var addr = U.addrOf(col, row);
    if (addr in this._cache) return this._cache[addr];
    var entry = this.cells[addr];
    if (!entry) return null;
    var v;
    if (entry.kind === 'formula') {
      if (this._stack.indexOf(addr) >= 0) {
        v = E.err(C.REF, { circular: true, addr: addr });
      } else {
        var parsed = Mobai.Parser.parseFormula(entry.formula);
        if (!parsed.ok) {
          v = E.err(C.NAME, { syntax: parsed.message });
        } else {
          this._stack.push(addr);
          try {
            v = evalNode(parsed.ast, this);
          } catch (ex) {
            v = E.err(C.VALUE, {});
          } finally {
            this._stack.pop();
          }
        }
      }
    } else {
      v = entry.value;
    }
    if (v === undefined) v = null;
    this._cache[addr] = v;
    return v;
  };

  /** 显示文本：数字右对齐、日期按格式串、错误显示错误码 */
  Workbook.prototype.display = function (col, row) {
    var entry = this.entry(col, row);
    if (!entry) return '';
    var v = this.getCell(col, row);
    if (E.isError(v)) return v.code;
    if (v === null) return '';
    var Fm = Mobai.Format;
    if (Fm && entry.format && !Fm.isGeneral(entry.format)) return Fm.render(v, entry.format);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      if (this.isDateCell(col, row)) return U.isoFromSerial(v);
      return U.formatNumber(v);
    }
    return String(v);
  };

  /** 输入框里应该显示的内容：公式显示公式本身，其余显示可再编辑的文本 */
  Workbook.prototype.rawText = function (col, row) {
    var entry = this.entry(col, row);
    if (!entry) return '';
    if (entry.kind === 'formula') return entry.formula;
    if (entry.kind === 'empty') return '';
    if (entry.kind === 'bool') return entry.value ? 'TRUE' : 'FALSE';
    if (entry.kind === 'number') {
      if (this.isDateCell(col, row)) return U.isoFromSerial(entry.value);
      return U.formatNumber(entry.value);
    }
    // 文本型数字（"'128"）要把撇号显示出来，否则用户在编辑框里看到的是 128，
    // 改的时候「删了重打 128」跟原值一模一样，会以为改了其实没改。Excel 也是这样。
    if (/^-?\d+(\.\d+)?$/.test(entry.value)) return "'" + entry.value;
    return entry.value;
  };

  Workbook.prototype.formulaText = function (col, row) {
    var entry = this.entry(col, row);
    return entry && entry.kind === 'formula' ? entry.formula : '';
  };

  Workbook.prototype.setRaw = function (addr, raw) {
    var a = U.parseAddr(addr);
    if (!a) return false;
    if (raw === null || raw === undefined || raw === '') delete this.cells[addr];
    else this.cells[addr] = normalizeEntry(raw);
    this.recalc();
    return true;
  };

  /** 用户输入 → 单元格内容（日期区域里输入 2026-03-01 会自动存成序列号）
   * 样式与数字格式跟着格子走：改内容不该把「加粗 / 两位小数」弄丢。 */
  Workbook.prototype.setInput = function (col, row, text) {
    var addr = U.addrOf(col, row);
    var s = String(text == null ? '' : text);
    var prev = this.cells[addr];
    var keep = null;
    if (prev && (prev.style || prev.format)) keep = { style: prev.style, format: prev.format };
    if (s.trim() === '') {
      this.cells[addr] = emptyLike(keep);
      this.recalc();
      return;
    }
    var entry;
    if (s.charAt(0) === "'") {
      entry = { kind: 'text', value: s.slice(1) };
    } else if (this.isDateCell(col, row) && U.looksLikeISO(s)) {
      entry = { kind: 'number', value: U.serialFromISO(s), format: 'date' };
    } else {
      entry = normalizeEntry(s);
    }
    this.cells[addr] = applyKeep(entry, keep);
    this.recalc();
  };

  /** 清内容不清格式（Excel 里按 Delete 就是这个行为） */
  Workbook.prototype.clearCell = function (col, row) {
    var addr = U.addrOf(col, row);
    this.cells[addr] = emptyLike(keepOf(this.cells[addr]));
    this.recalc();
  };

  /** 连格式一起清掉 */
  Workbook.prototype.clearFormat = function (col, row) {
    var addr = U.addrOf(col, row);
    var e = this.cells[addr];
    if (!e) return;
    delete e.style;
    delete e.format;
    if (e.kind === 'empty') delete this.cells[addr];
    this.recalc();
  };

  function keepOf(entry) {
    if (!entry || (!entry.style && !entry.format)) return null;
    return { style: entry.style, format: entry.format };
  }

  function emptyLike(keep) {
    return applyKeep({ kind: 'empty' }, keep);
  }

  function applyKeep(entry, keep) {
    if (keep) {
      if (keep.style) entry.style = keep.style;
      if (keep.format) entry.format = keep.format;
    }
    return entry;
  }

  /* ---------- 样式 / 数字格式 ---------- */

  /** target 可以是单元格 {col,row}、区域 {start,end}，或 'C2:C11' / 'C2' 这样的字符串 */
  function eachCellIn(target, fn) {
    if (target && target.start && target.end) {
      var cells = U.expandRange(target);
      for (var i = 0; i < cells.length; i++) fn(cells[i].col, cells[i].row);
      return;
    }
    if (typeof target === 'string') {
      var r = U.rangeOf(target);
      if (r) return eachCellIn(r, fn);
      return;
    }
    if (target && typeof target.col === 'number') fn(target.col, target.row);
  }

  Workbook.prototype.styleOf = function (col, row) {
    var e = this.entry(col, row);
    return (e && e.style) || null;
  };

  /** 给区域打样式补丁；patch 里值为 null/false 表示去掉这个属性 */
  Workbook.prototype.setStyle = function (target, patch) {
    var self = this;
    eachCellIn(target, function (col, row) {
      if (self.isReadonly(col, row)) return;
      var addr = U.addrOf(col, row);
      var e = self.cells[addr] || { kind: 'empty' };
      var st = {};
      var old = e.style || {};
      Object.keys(old).forEach(function (k) { st[k] = old[k]; });
      Object.keys(patch).forEach(function (k) {
        if (patch[k] === null || patch[k] === undefined || patch[k] === false) delete st[k];
        else st[k] = patch[k];
      });
      if (Object.keys(st).length) e.style = st; else delete e.style;
      if (e.kind === 'empty' && !e.style && !e.format) delete self.cells[addr];
      else self.cells[addr] = e;
    });
    this.recalc();
  };

  Workbook.prototype.formatOf = function (col, row) {
    var e = this.entry(col, row);
    return (e && e.format) || null;
  };

  Workbook.prototype.setFormat = function (target, code) {
    var self = this;
    eachCellIn(target, function (col, row) {
      if (self.isReadonly(col, row)) return;
      var addr = U.addrOf(col, row);
      var e = self.cells[addr] || { kind: 'empty' };
      if (!code || code === 'General') delete e.format;
      else e.format = code;
      if (e.kind === 'empty' && !e.style && !e.format) delete self.cells[addr];
      else self.cells[addr] = e;
    });
    this.recalc();
  };

  /* ---------- 合并单元格 ---------- */

  Workbook.prototype.mergeOf = function (col, row) {
    for (var i = 0; i < this.merges.length; i++) {
      var m = this.merges[i];
      if (col >= m.start.col && col <= m.end.col && row >= m.start.row && row <= m.end.row) return m;
    }
    return null;
  };

  Workbook.prototype.addMerge = function (range) {
    var r = U.normRange(range.start, range.end);
    if (U.rangeCount(r) < 2) return false;
    this.merges = this.merges.filter(function (m) { return !U.rangeOverlap(m, r); });
    this.merges.push(r);
    return true;
  };

  Workbook.prototype.removeMergeAt = function (col, row) {
    var before = this.merges.length;
    this.merges = this.merges.filter(function (m) {
      return !(col >= m.start.col && col <= m.end.col && row >= m.start.row && row <= m.end.row);
    });
    return this.merges.length !== before;
  };

  /* ---------- 插入 / 删除行列 ----------
   * 格子跟着挪，公式里的引用也一起改写（$ 锁不住「整行搬家」这件事，
   * 跟 Excel 一致：插一行，$A$1 也会跟着往下走）；指向被删区域的引用变成 #REF!。
   */

  function shiftRangeList(list, axis, at, n) {
    return (list || []).map(function (item) {
      var r = typeof item === 'string' ? U.rangeOf(item) : item;
      if (!r) return item;
      var out = { start: { col: r.start.col, row: r.start.row }, end: { col: r.end.col, row: r.end.row } };
      if (axis === 'row') {
        if (out.start.row >= at) out.start.row += n;
        if (out.end.row >= at) out.end.row += n;
      } else {
        if (out.start.col >= at) out.start.col += n;
        if (out.end.col >= at) out.end.col += n;
      }
      return U.rangeToString(out);
    });
  }

  Workbook.prototype.insertLines = function (axis, at, n) {
    var self = this;
    var isRow = axis === 'row';
    n = n || 1;
    var out = Object.create(null);
    Object.keys(this.cells).forEach(function (addr) {
      var a = U.parseAddr(addr);
      if (!a) return;
      var e = self.cells[addr];
      var col = a.col;
      var row = a.row;
      if (isRow) { if (row >= at) row += n; } else if (col >= at) col += n;
      out[U.addrOf(col, row)] = mapRefs(e, function (ref) {
        if (isRow) return ref.row >= at ? cloneRef(ref, ref.col, ref.row + n) : ref;
        return ref.col >= at ? cloneRef(ref, ref.col + n, ref.row) : ref;
      });
    });
    this.cells = out;
    this.merges = this.merges.map(function (m) {
      var r = {
        start: { col: m.start.col, row: m.start.row },
        end: { col: m.end.col, row: m.end.row }
      };
      if (isRow) {
        if (r.start.row >= at) r.start.row += n;
        if (r.end.row >= at) r.end.row += n;
      } else {
        if (r.start.col >= at) r.start.col += n;
        if (r.end.col >= at) r.end.col += n;
      }
      return r;
    });
    this.readonly = shiftRangeList(this.readonly, axis, at, n);
    this.dates = shiftRangeList(this.dates, axis, at, n);
    this.bold = shiftRangeList(this.bold, axis, at, n);
    this.recalc();
  };

  Workbook.prototype.deleteLines = function (axis, at, n) {
    var self = this;
    var isRow = axis === 'row';
    n = n || 1;
    var out = Object.create(null);
    Object.keys(this.cells).forEach(function (addr) {
      var a = U.parseAddr(addr);
      if (!a) return;
      var e = self.cells[addr];
      var col = a.col;
      var row = a.row;
      var pos = isRow ? row : col;
      if (pos >= at && pos < at + n) return;             // 这一行/列被删掉了
      if (pos >= at + n) { if (isRow) row -= n; else col -= n; }
      out[U.addrOf(col, row)] = mapRefs(e, function (ref) {
        var p = isRow ? ref.row : ref.col;
        if (p >= at && p < at + n) return null;          // 指向被删区域 → #REF!
        if (p >= at + n) return isRow ? cloneRef(ref, ref.col, ref.row - n) : cloneRef(ref, ref.col - n, ref.row);
        return ref;
      });
    });
    this.cells = out;
    var kept = [];
    this.merges.forEach(function (m) {
      var s = isRow ? m.start.row : m.start.col;
      var e2 = isRow ? m.end.row : m.end.col;
      var ns = s >= at + n ? s - n : (s >= at ? at : s);
      var ne = e2 >= at + n ? e2 - n : (e2 >= at ? at - 1 : e2);
      if (ne < ns) return;
      var r = {
        start: { col: m.start.col, row: m.start.row },
        end: { col: m.end.col, row: m.end.row }
      };
      if (isRow) { r.start.row = ns; r.end.row = ne; } else { r.start.col = ns; r.end.col = ne; }
      kept.push(r);
    });
    this.merges = kept;
    this.readonly = shiftRangeList(this.readonly, axis, at, -n);
    this.dates = shiftRangeList(this.dates, axis, at, -n);
    this.bold = shiftRangeList(this.bold, axis, at, -n);
    this.recalc();
  };

  function cloneRef(ref, col, row) {
    return { col: col, row: row, colAbs: ref.colAbs, rowAbs: ref.rowAbs };
  }

  function mapRefs(entry, fn) {
    if (!entry || entry.kind !== 'formula') return entry;
    var Fill = Mobai.Fill;
    if (!Fill) return entry;
    var text = Fill.transformRefs(entry.formula, fn);
    if (text === entry.formula) return entry;
    return applyKeep({ kind: 'formula', formula: text }, keepOf(entry));
  }

  /** 有内容的最大行/列 */
  Workbook.prototype.usedRange = function () {
    var minCol = Infinity;
    var maxCol = -Infinity;
    var minRow = Infinity;
    var maxRow = -Infinity;
    var keys = Object.keys(this.cells);
    for (var i = 0; i < keys.length; i++) {
      var a = U.parseAddr(keys[i]);
      if (!a) continue;
      if (a.col < minCol) minCol = a.col;
      if (a.col > maxCol) maxCol = a.col;
      if (a.row < minRow) minRow = a.row;
      if (a.row > maxRow) maxRow = a.row;
    }
    if (!keys.length) return { minCol: 0, maxCol: 0, minRow: 1, maxRow: 1 };
    return { minCol: minCol, maxCol: maxCol, minRow: minRow, maxRow: maxRow };
  };

  /** 把区域裁剪到有数据的范围内（多留 3 行 1 列），并限制总格数 */
  Workbook.prototype.clampRange = function (r) {
    var u = this.usedRange();
    var out = {
      start: {
        col: Math.max(r.start.col, u.minCol),
        row: Math.max(r.start.row, u.minRow)
      },
      end: {
        col: Math.min(r.end.col, u.maxCol + 1),
        row: Math.min(r.end.row, u.maxRow + 3)
      }
    };
    if (out.start.col > out.end.col || out.start.row > out.end.row) {
      out = { start: { col: r.start.col, row: r.start.row }, end: { col: r.start.col, row: r.start.row } };
    }
    if (U.rangeCount(out) > MAX_RANGE_CELLS) {
      out.end.row = out.start.row + Math.max(1, Math.floor(MAX_RANGE_CELLS / (out.end.col - out.start.col + 1))) - 1;
    }
    return out;
  };

  /** 表格渲染范围（含空行空列，供新填的数据用） */
  Workbook.prototype.dims = function (minCols, minRows) {
    var u = this.usedRange();
    return {
      cols: Math.max(minCols || 0, u.maxCol + 2),
      rows: Math.max(minRows || 0, u.maxRow + 3)
    };
  };

  Workbook.prototype.snapshot = function () {
    return {
      cells: JSON.parse(JSON.stringify(this.cells)),
      merges: JSON.parse(JSON.stringify(this.merges))
    };
  };

  Workbook.prototype.restore = function (snap) {
    if (!snap || typeof snap !== 'object') return;
    // 兼容只有单元格表的旧快照
    var cells = snap.cells || snap;
    this.cells = JSON.parse(JSON.stringify(cells));
    this.merges = snap.merges ? JSON.parse(JSON.stringify(snap.merges)) : [];
    this.recalc();
  };

  /** 判分与校验用：把指定区域算成值数组 */
  Workbook.prototype.valuesOf = function (rangeStr) {
    var r = U.rangeOf(rangeStr);
    if (!r) return null;
    var self = this;
    return U.expandRange(r).map(function (c) {
      return self.getCell(c.col, c.row);
    });
  };

  Mobai.Evaluator = {
    evalNode: evalNode,
    compareValues: compareValues,
    normalizeEntry: normalizeEntry
  };
  Mobai.Workbook = Workbook;
})(typeof window !== 'undefined' ? window : globalThis);