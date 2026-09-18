/* mobai训练 · 函数库
 * 函数清单按 10 道练习倒推，不多做。
 * 每个函数收到的是「惰性参数」数组：[{node, value()}]，这样 IF / IFERROR / IFS
 * 才能只求值需要的分支。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;
  var E = Mobai.Errors;
  var C = E.CODES;

  var registry = {};

  function def(name, min, max, fn) {
    registry[name] = { name: name, min: min, max: max, fn: fn };
  }

  /* ---------- 参数与值处理 ---------- */

  function toNumber(v, info) {
    if (E.isError(v)) return v;
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    var s = String(v).trim();
    if (s === '') return 0;
    if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s)) return parseFloat(s);
    return E.err(C.VALUE, info || {});
  }

  function toText(v) {
    if (E.isError(v)) return v;
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return U.formatNumber(v);
    return String(v);
  }

  function toBool(v) {
    if (E.isError(v)) return v;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === null || v === undefined || v === '') return false;
    var s = String(v).trim().toUpperCase();
    if (s === 'FALSE') return false;
    return true;
  }

  /** 转成日期序列号：数字直接用，ISO 文本解析，其余报错 */
  function toSerial(v) {
    if (E.isError(v)) return v;
    if (typeof v === 'number') return Math.round(v);
    var iso = U.serialFromISO(v);
    if (iso !== null) return iso;
    return E.err(C.VALUE, {});
  }

  function displayOf(v) {
    if (v === null || v === undefined) return '（空）';
    if (E.isError(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  function firstError(list) {
    for (var i = 0; i < list.length; i++) {
      if (E.isError(list[i])) return list[i];
    }
    return null;
  }

  /** 区域 → 二维表（会被工作簿裁剪到有数据范围，避免 A:A 这种整列引用爆掉） */
  function matrixOf(arg, ctx) {
    var node = arg.node;
    if (node.type === 'Range') {
      var r = U.normRange(node.start, node.end);
      if (ctx && ctx.clampRange) r = ctx.clampRange(r);
      if (!r) return { error: E.err(C.VALUE, { range: node.raw }) };
      var rows = [];
      for (var row = r.start.row; row <= r.end.row; row++) {
        var line = [];
        for (var col = r.start.col; col <= r.end.col; col++) line.push(ctx.getCell(col, row));
        rows.push(line);
      }
      return { rows: rows };
    }
    var v = arg.value();
    if (E.isError(v)) return { error: v };
    return { rows: [[v]] };
  }

  /** 区域 → 一维值（行优先），标量 → 单元素 */
  function flatValues(arg, ctx) {
    if (arg.node.type !== 'Range') return [arg.value()];
    var m = matrixOf(arg, ctx);
    if (m.error) return [m.error];
    return U.flatten(m.rows);
  }

  /** 一维化：单行或单列 */
  function vectorOf(arg, ctx) {
    var m = matrixOf(arg, ctx);
    if (m.error) return { error: m.error };
    var rows = m.rows;
    if (rows.length === 1) return { values: rows[0], horizontal: true };
    if (rows[0].length === 1) {
      return { values: rows.map(function (r) { return r[0]; }), horizontal: false };
    }
    return { error: E.err(C.VALUE, { fn: 'MATCH' }) };
  }

  function scanNumbers(args, ctx) {
    var nums = [];
    var err = null;
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (a.node.type === 'Range') {
        var vals = flatValues(a, ctx);
        var e1 = firstError(vals);
        if (e1) { err = err || e1; continue; }
        for (var j = 0; j < vals.length; j++) {
          if (typeof vals[j] === 'number') nums.push(vals[j]);
        }
      } else {
        var v = a.value();
        if (E.isError(v)) { err = err || v; continue; }
        if (v === null || v === undefined || v === '') continue;
        // 直接写进参数里的逻辑值与文本，Excel 会尝试转换
        var n = toNumber(v, { fn: '聚合函数' });
        if (E.isError(n)) { err = err || n; continue; }
        nums.push(n);
      }
    }
    return { nums: nums, error: err };
  }

  function aggResult(nums, fn) {
    // 空区域（含「一片全是被当成文本的数字」）：Excel 的 SUM / MAX / MIN 都给 0，
    // 只有 AVERAGE 才报 #DIV/0!，而它在自己那边已经拦住了。
    if (!nums.length) return 0;
    var out = nums[0];
    for (var i = 1; i < nums.length; i++) {
      if (fn === 'SUM') out += nums[i];
      else if (fn === 'MAX') out = Math.max(out, nums[i]);
      else out = Math.min(out, nums[i]);
    }
    return out;
  }

  /* ---------- 条件匹配 ---------- */

  function wildcardRegex(pattern) {
    var hasWild = /[*?]/.test(pattern.replace(/~[*?]/g, ''));
    if (!hasWild) return null;
    var re = '';
    for (var i = 0; i < pattern.length; i++) {
      var ch = pattern[i];
      if (ch === '~' && i + 1 < pattern.length) {
        re += pattern[i + 1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        i++;
      } else if (ch === '*') re += '.*';
      else if (ch === '?') re += '.';
      else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }

  function textOf(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return U.formatNumber(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  /** 把条件（数字 / ">100" / "茶*" / 单元格）编译成判定函数 */
  function makeCriteria(crit) {
    if (typeof crit === 'number') {
      return function (v) {
        return typeof v === 'number' && U.almostEqual(v, crit);
      };
    }
    if (typeof crit === 'boolean') {
      return function (v) {
        return v === crit;
      };
    }
    var s = String(crit);
    var m = /^(<=|>=|<>|=|<|>)\s*([\s\S]*)$/.exec(s);
    var op = '=';
    var target = s;
    if (m) {
      op = m[1];
      target = m[2];
    }
    if (target === '') {
      if (op === '<>') {
        return function (v) {
          return !(v === null || v === undefined || v === '');
        };
      }
      return function (v) {
        return v === null || v === undefined || v === '';
      };
    }
    var numTarget = /^-?\d+(\.\d+)?$/.test(target.trim()) ? parseFloat(target) : null;
    if (numTarget !== null) {
      return function (v) {
        if (typeof v !== 'number') return false;
        switch (op) {
          case '=': return U.almostEqual(v, numTarget);
          case '<>': return !U.almostEqual(v, numTarget);
          case '<': return v < numTarget;
          case '>': return v > numTarget;
          case '<=': return v <= numTarget;
          default: return v >= numTarget;
        }
      };
    }
    var re = wildcardRegex(target);
    var low = target.toLowerCase();
    return function (v) {
      var t = textOf(v);
      var hit = re ? re.test(t) : t.toLowerCase() === low;
      switch (op) {
        case '=': return hit;
        case '<>': return !hit;
        case '<': return t.toLowerCase() < low;
        case '>': return t.toLowerCase() > low;
        case '<=': return t.toLowerCase() <= low || hit;
        default: return t.toLowerCase() >= low || hit;
      }
    };
  }

  /** 精确匹配（查找函数用）：不做数字与文本的互相转换——这正是第 6/7 题的坑 */
  function looseEqual(a, b) {
    if (a === null || a === undefined) return b === null || b === undefined || b === '';
    if (typeof a === 'number' && typeof b === 'number') return U.almostEqual(a, b);
    if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
    if (typeof a === 'string' && typeof b === 'string') {
      return a.toLowerCase() === b.toLowerCase();
    }
    return false;
  }

  /* ---------- 统计 ---------- */

  def('SUM', 1, null, function (args, ctx) {
    var s = scanNumbers(args, ctx);
    if (s.error) return s.error;
    return aggResult(s.nums, 'SUM');
  });

  def('AVERAGE', 1, null, function (args, ctx) {
    var s = scanNumbers(args, ctx);
    if (s.error) return s.error;
    if (!s.nums.length) return E.err(C.DIV0, {});
    var t = 0;
    for (var i = 0; i < s.nums.length; i++) t += s.nums[i];
    return t / s.nums.length;
  });

  def('MAX', 1, null, function (args, ctx) {
    var s = scanNumbers(args, ctx);
    if (s.error) return s.error;
    return aggResult(s.nums, 'MAX');
  });

  def('MIN', 1, null, function (args, ctx) {
    var s = scanNumbers(args, ctx);
    if (s.error) return s.error;
    return aggResult(s.nums, 'MIN');
  });

  def('ROUND', 2, 2, function (args) {
    var x = toNumber(args[0].value(), { fn: 'ROUND' });
    if (E.isError(x)) return x;
    var d = toNumber(args[1].value(), { fn: 'ROUND' });
    if (E.isError(d)) return d;
    d = Math.round(d);
    var f = Math.pow(10, d);
    return Math.round(x * f) / f;
  });

  def('ABS', 1, 1, function (args) {
    var x = toNumber(args[0].value(), { fn: 'ABS' });
    if (E.isError(x)) return x;
    return Math.abs(x);
  });

  /* ---------- 条件聚合 ---------- */

  def('COUNTIF', 2, 2, function (args, ctx) {
    var vals = flatValues(args[0], ctx);
    var e = firstError(vals);
    if (e) return e;
    var crit = args[1].value();
    if (E.isError(crit)) return crit;
    var test = makeCriteria(crit);
    var n = 0;
    for (var i = 0; i < vals.length; i++) if (test(vals[i])) n++;
    return n;
  });

  def('COUNTIFS', 2, null, function (args, ctx) {
    if (args.length % 2 !== 0) return E.err(C.VALUE, { arity: true, fn: 'COUNTIFS' });
    var groups = [];
    for (var i = 0; i < args.length; i += 2) {
      var vals = flatValues(args[i], ctx);
      var e = firstError(vals);
      if (e) return e;
      var crit = args[i + 1].value();
      if (E.isError(crit)) return crit;
      groups.push({ vals: vals, test: makeCriteria(crit) });
    }
    var len = groups[0].vals.length;
    for (var g = 1; g < groups.length; g++) {
      if (groups[g].vals.length !== len) return E.err(C.VALUE, { fn: 'COUNTIFS' });
    }
    var n = 0;
    for (var k = 0; k < len; k++) {
      var ok = true;
      for (var h = 0; h < groups.length; h++) {
        if (!groups[h].test(groups[h].vals[k])) { ok = false; break; }
      }
      if (ok) n++;
    }
    return n;
  });

  function conditionalSum(args, ctx, fn) {
    var pairsStart = 1;
    var sumVals = flatValues(args[0], ctx);
    var e0 = firstError(sumVals);
    if (e0) return e0;
    var groups = [];
    for (var i = pairsStart; i < args.length; i += 2) {
      var vals = flatValues(args[i], ctx);
      var e = firstError(vals);
      if (e) return e;
      var crit = args[i + 1].value();
      if (E.isError(crit)) return crit;
      groups.push({ vals: vals, test: makeCriteria(crit) });
    }
    for (var g = 0; g < groups.length; g++) {
      if (groups[g].vals.length !== sumVals.length) return E.err(C.VALUE, { fn: fn });
    }
    var total = 0;
    var count = 0;
    for (var k = 0; k < sumVals.length; k++) {
      var ok = true;
      for (var h = 0; h < groups.length; h++) {
        if (!groups[h].test(groups[h].vals[k])) { ok = false; break; }
      }
      if (ok && typeof sumVals[k] === 'number') {
        total += sumVals[k];
        count++;
      }
    }
    if (fn === 'AVERAGEIFS') {
      if (!count) return E.err(C.DIV0, {});
      return total / count;
    }
    return total;
  }

  // SUMIF / AVERAGEIF 的参数顺序是 (条件区域, 条件, 求和区域)，与 SUMIFS 相反
  def('SUMIF', 3, 3, function (args, ctx) {
    return conditionalSum([args[2], args[0], args[1]], ctx, 'SUMIF');
  });

  def('SUMIFS', 3, null, function (args, ctx) {
    if (args.length % 2 === 0) return E.err(C.VALUE, { arity: true, fn: 'SUMIFS' });
    return conditionalSum(args, ctx, 'SUMIFS');
  });

  def('AVERAGEIF', 3, 3, function (args, ctx) {
    return conditionalSum([args[2], args[0], args[1]], ctx, 'AVERAGEIF');
  });

  def('AVERAGEIFS', 3, null, function (args, ctx) {
    if (args.length % 2 === 0) return E.err(C.VALUE, { arity: true, fn: 'AVERAGEIFS' });
    return conditionalSum(args, ctx, 'AVERAGEIFS');
  });

  /* ---------- 逻辑 ---------- */

  def('IF', 2, 3, function (args) {
    var c = args[0].value();
    if (E.isError(c)) return c;
    var cond = toBool(c);
    if (E.isError(cond)) return cond;
    if (cond) return args[1].value();
    if (args.length >= 3) return args[2].value();
    return false;
  });

  def('IFS', 2, null, function (args) {
    if (args.length % 2 !== 0) return E.err(C.VALUE, { arity: true, fn: 'IFS' });
    for (var i = 0; i < args.length; i += 2) {
      var c = args[i].value();
      if (E.isError(c)) return c;
      var cond = toBool(c);
      if (E.isError(cond)) return cond;
      if (cond) return args[i + 1].value();
    }
    return E.err(C.NA, { fn: 'IFS' });
  });

  function logicAggregate(args, ctx, mode) {
    var values = [];
    for (var i = 0; i < args.length; i++) {
      var vals = flatValues(args[i], ctx);
      for (var j = 0; j < vals.length; j++) {
        var v = vals[j];
        if (E.isError(v)) return v;
        if (typeof v === 'boolean') values.push(v);
        else if (typeof v === 'number') values.push(v !== 0);
        // 文本与空单元格忽略
      }
    }
    if (!values.length) return E.err(C.VALUE, {});
    if (mode === 'AND') {
      for (var a = 0; a < values.length; a++) if (!values[a]) return false;
      return true;
    }
    for (var o = 0; o < values.length; o++) if (values[o]) return true;
    return false;
  }

  def('AND', 1, null, function (args, ctx) {
    return logicAggregate(args, ctx, 'AND');
  });

  def('OR', 1, null, function (args, ctx) {
    return logicAggregate(args, ctx, 'OR');
  });

  def('NOT', 1, 1, function (args) {
    var b = toBool(args[0].value());
    if (E.isError(b)) return b;
    return !b;
  });

  def('IFERROR', 2, 2, function (args) {
    var v = args[0].value();
    if (E.isError(v)) {
      if (args[0].node.type === 'Range') return E.err(C.VALUE, { range: args[0].node.raw });
      return args[1].value();
    }
    return v;
  });

  /* ---------- 查找 ---------- */

  def('VLOOKUP', 3, 4, function (args, ctx) {
    var lookup = args[0].value();
    if (E.isError(lookup)) return lookup;
    var m = matrixOf(args[1], ctx);
    if (m.error) return m.error;
    var colNum = toNumber(args[2].value(), { fn: 'VLOOKUP' });
    if (E.isError(colNum)) return colNum;
    colNum = Math.round(colNum);
    var approx = true;
    if (args.length >= 4) {
      var rv = args[3].value();
      if (E.isError(rv)) return rv;
      approx = toBool(rv);
      if (E.isError(approx)) return approx;
    }
    var width = m.rows[0].length;
    if (colNum < 1 || colNum > width) return E.err(C.REF, { fn: 'VLOOKUP' });
    var lastHit = -1;
    for (var i = 0; i < m.rows.length; i++) {
      if (looseEqual(m.rows[i][0], lookup)) return m.rows[i][colNum - 1];
      if (approx && typeof m.rows[i][0] === 'number' && typeof lookup === 'number' && m.rows[i][0] <= lookup) {
        lastHit = i;
      }
    }
    if (approx && lastHit >= 0) return m.rows[lastHit][colNum - 1];
    return E.err(C.NA, { fn: 'VLOOKUP', lookupValue: displayOf(lookup) });
  });

  def('XLOOKUP', 3, 6, function (args, ctx) {
    var lookup = args[0].value();
    if (E.isError(lookup)) return lookup;
    var keys = vectorOf(args[1], ctx);
    if (keys.error) return keys.error;
    var rets = vectorOf(args[2], ctx);
    if (rets.error) return rets.error;
    if (keys.values.length !== rets.values.length) {
      return E.err(C.VALUE, { fn: 'XLOOKUP' });
    }
    var ifNotFound = E.err(C.NA, { fn: 'XLOOKUP', lookupValue: displayOf(lookup) });
    if (args.length >= 4) ifNotFound = args[3].value();
    var matchMode = 0;
    if (args.length >= 5) {
      var mm = toNumber(args[4].value(), { fn: 'XLOOKUP' });
      if (E.isError(mm)) return mm;
      matchMode = Math.round(mm);
    }
    var searchMode = 1;
    if (args.length >= 6) {
      var sm = toNumber(args[5].value(), { fn: 'XLOOKUP' });
      if (E.isError(sm)) return sm;
      searchMode = sm < 0 ? -1 : 1;
    }
    var test;
    if (matchMode === 2) {
      test = makeCriteria(String(displayOf(lookup)));
    } else {
      test = function (v) { return looseEqual(v, lookup); };
    }
    var order = [];
    var n = keys.values.length;
    for (var i = 0; i < n; i++) order.push(searchMode < 0 ? n - 1 - i : i);

    if (matchMode === 0 || matchMode === 2) {
      for (var a = 0; a < order.length; a++) {
        if (test(keys.values[order[a]])) return rets.values[order[a]];
      }
      return ifNotFound;
    }
    // 近似匹配：要求按键升序
    var best = -1;
    for (var b = 0; b < n; b++) {
      var kv = keys.values[b];
      if (typeof kv !== 'number' || typeof lookup !== 'number') continue;
      if (matchMode === -1 && kv <= lookup) best = b;
      if (matchMode === 1 && kv >= lookup) { best = b; break; }
    }
    if (best >= 0) return rets.values[best];
    return ifNotFound;
  });

  def('INDEX', 2, 3, function (args, ctx) {
    var m = matrixOf(args[0], ctx);
    if (m.error) return m.error;
    var r = toNumber(args[1].value(), { fn: 'INDEX' });
    if (E.isError(r)) return r;
    r = Math.round(r);
    var c = 1;
    if (args.length >= 3) {
      var cv = toNumber(args[2].value(), { fn: 'INDEX' });
      if (E.isError(cv)) return cv;
      c = Math.round(cv);
    }
    if (m.rows.length === 1 && args.length < 3) return m.rows[0][r - 1];
    if (m.rows[0].length === 1 && args.length < 3) return m.rows[r - 1][0];
    if (r < 1 || r > m.rows.length || c < 1 || c > m.rows[0].length) return E.err(C.REF, { fn: 'INDEX' });
    return m.rows[r - 1][c - 1];
  });

  def('MATCH', 2, 3, function (args, ctx) {
    var lookup = args[0].value();
    if (E.isError(lookup)) return lookup;
    var vec = vectorOf(args[1], ctx);
    if (vec.error) return vec.error;
    var mode = 1;
    if (args.length >= 3) {
      var mv = toNumber(args[2].value(), { fn: 'MATCH' });
      if (E.isError(mv)) return mv;
      mode = Math.round(mv);
    }
    var vals = vec.values;
    if (mode === 0) {
      for (var i = 0; i < vals.length; i++) if (looseEqual(vals[i], lookup)) return i + 1;
      return E.err(C.NA, { fn: 'MATCH', lookupValue: displayOf(lookup) });
    }
    var best = -1;
    for (var j = 0; j < vals.length; j++) {
      if (typeof vals[j] !== 'number' || typeof lookup !== 'number') continue;
      if (mode === 1 && vals[j] <= lookup) best = j;
      if (mode === -1 && vals[j] >= lookup) best = j;
    }
    if (best < 0) return E.err(C.NA, { fn: 'MATCH', lookupValue: displayOf(lookup) });
    return best + 1;
  });

  /* ---------- 文本 ---------- */

  def('LEFT', 1, 2, function (args) {
    var t = toText(args[0].value());
    if (E.isError(t)) return t;
    var n = 1;
    if (args.length >= 2) {
      var nv = toNumber(args[1].value(), { fn: 'LEFT' });
      if (E.isError(nv)) return nv;
      n = Math.round(nv);
    }
    if (n < 0) return E.err(C.VALUE, { fn: 'LEFT' });
    return t.slice(0, n);
  });

  def('RIGHT', 1, 2, function (args) {
    var t = toText(args[0].value());
    if (E.isError(t)) return t;
    var n = 1;
    if (args.length >= 2) {
      var nv = toNumber(args[1].value(), { fn: 'RIGHT' });
      if (E.isError(nv)) return nv;
      n = Math.round(nv);
    }
    if (n < 0) return E.err(C.VALUE, { fn: 'RIGHT' });
    return n === 0 ? '' : t.slice(-n);
  });

  def('MID', 3, 3, function (args) {
    var t = toText(args[0].value());
    if (E.isError(t)) return t;
    var s = toNumber(args[1].value(), { fn: 'MID' });
    if (E.isError(s)) return s;
    var n = toNumber(args[2].value(), { fn: 'MID' });
    if (E.isError(n)) return n;
    s = Math.round(s);
    n = Math.round(n);
    if (s < 1 || n < 0) return E.err(C.VALUE, { fn: 'MID' });
    return t.substr(s - 1, n);
  });

  def('LEN', 1, 1, function (args) {
    var t = toText(args[0].value());
    if (E.isError(t)) return t;
    return t.length;
  });

  def('FIND', 2, 3, function (args) {
    var needle = toText(args[0].value());
    if (E.isError(needle)) return needle;
    var hay = toText(args[1].value());
    if (E.isError(hay)) return hay;
    var start = 1;
    if (args.length >= 3) {
      var sv = toNumber(args[2].value(), { fn: 'FIND' });
      if (E.isError(sv)) return sv;
      start = Math.round(sv);
    }
    var idx = hay.indexOf(needle, Math.max(0, start - 1));
    if (idx < 0) return E.err(C.VALUE, { fn: 'FIND' });
    return idx + 1;
  });

  def('TRIM', 1, 1, function (args) {
    var t = toText(args[0].value());
    if (E.isError(t)) return t;
    return t.replace(/\s+/g, ' ').replace(/^ | $/g, '');
  });

  def('CONCAT', 1, null, function (args, ctx) {
    return concatAll(args, ctx);
  });

  def('CONCATENATE', 1, null, function (args, ctx) {
    return concatAll(args, ctx);
  });

  function concatAll(args, ctx) {
    var out = '';
    for (var i = 0; i < args.length; i++) {
      var vals = flatValues(args[i], ctx);
      for (var j = 0; j < vals.length; j++) {
        if (E.isError(vals[j])) return vals[j];
        out += textOf(vals[j]);
      }
    }
    return out;
  }

  def('TEXT', 2, 2, function (args) {
    var v = args[0].value();
    if (E.isError(v)) return v;
    var code = toText(args[1].value());
    if (E.isError(code)) return code;
    return formatWithCode(v, code);
  });

  function formatWithCode(value, code) {
    var isDateCode = /[yd]/i.test(code) || /[^a-z]m{1,4}[^a-z]/i.test(code);
    if (isDateCode) {
      var serial = toSerial(value);
      if (E.isError(serial)) return serial;
      var p = U.partsFromSerial(serial);
      var out = code;
      out = out.replace(/yyyy/gi, String(p.y));
      out = out.replace(/yy/gi, String(p.y).slice(-2));
      out = out.replace(/mmmm/gi, p.m + '月');
      out = out.replace(/mm/g, U.pad2(p.m));
      out = out.replace(/m/gi, String(p.m));
      out = out.replace(/dd/gi, U.pad2(p.d));
      out = out.replace(/d/gi, String(p.d));
      out = out.replace(/aaaa/g, '');
      return out;
    }
    var n = toNumber(value, { fn: 'TEXT' });
    if (E.isError(n)) return n;
    var pct = code.indexOf('%') >= 0;
    var comma = code.indexOf(',') >= 0;
    var dot = code.lastIndexOf('.');
    var decimals = dot >= 0 ? code.slice(dot + 1).replace(/[^0#]/g, '').length : 0;
    var x = pct ? n * 100 : n;
    var s = Math.abs(x).toFixed(decimals);
    if (comma) {
      var parts = s.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      s = parts.join('.');
    }
    if (x < 0) s = '-' + s;
    return s + (pct ? '%' : '');
  }

  /* ---------- 日期 ---------- */

  def('TODAY', 0, 0, function (args, ctx) {
    return ctx.todaySerial;
  });

  def('DATE', 3, 3, function (args) {
    var y = toNumber(args[0].value(), { fn: 'DATE' });
    if (E.isError(y)) return y;
    var m = toNumber(args[1].value(), { fn: 'DATE' });
    if (E.isError(m)) return m;
    var d = toNumber(args[2].value(), { fn: 'DATE' });
    if (E.isError(d)) return d;
    return U.serialFromYMD(Math.round(y), Math.round(m), Math.round(d));
  });

  def('YEAR', 1, 1, function (args) {
    return datePart(args, 'y', 'YEAR');
  });
  def('MONTH', 1, 1, function (args) {
    return datePart(args, 'm', 'MONTH');
  });
  def('DAY', 1, 1, function (args) {
    return datePart(args, 'd', 'DAY');
  });

  function datePart(args, part, fn) {
    var serial = toSerial(args[0].value());
    if (E.isError(serial)) return serial;
    var p = U.partsFromSerial(serial);
    return p[part];
  }

  def('WEEKDAY', 1, 2, function (args) {
    var serial = toSerial(args[0].value());
    if (E.isError(serial)) return serial;
    var type = 1;
    if (args.length >= 2) {
      var tv = toNumber(args[1].value(), { fn: 'WEEKDAY' });
      if (E.isError(tv)) return tv;
      type = Math.round(tv);
    }
    var dow = U.partsFromSerial(serial).dow; // 0=周日
    if (type === 2) return dow === 0 ? 7 : dow;
    if (type === 3) return dow === 0 ? 6 : dow - 1;
    return dow + 1;
  });

  def('DATEDIF', 3, 3, function (args) {
    var a = toSerial(args[0].value());
    if (E.isError(a)) return a;
    var b = toSerial(args[1].value());
    if (E.isError(b)) return b;
    var unit = toText(args[2].value()).toUpperCase();
    if (b < a) return E.err(C.NUM, { fn: 'DATEDIF' });
    var pa = U.partsFromSerial(a);
    var pb = U.partsFromSerial(b);
    switch (unit) {
      case 'D':
        return b - a;
      case 'M':
        return (pb.y - pa.y) * 12 + (pb.m - pa.m) - (pb.d < pa.d ? 1 : 0);
      case 'Y':
        return pb.y - pa.y - (pb.m < pa.m || (pb.m === pa.m && pb.d < pa.d) ? 1 : 0);
      case 'MD': {
        var d = pb.d - pa.d;
        if (d < 0) {
          var prev = U.partsFromSerial(U.serialFromYMD(pb.y, pb.m, 1) - 1);
          d += prev.d;
        }
        return d;
      }
      case 'YM':
        return ((pb.m - pa.m) + 12 - (pb.d < pa.d ? 1 : 0)) % 12;
      case 'YD': {
        var anchor = U.serialFromYMD(pb.y, pa.m, pa.d);
        if (anchor > b) anchor = U.serialFromYMD(pb.y - 1, pa.m, pa.d);
        return b - anchor;
      }
      default:
        return E.err(C.NUM, { fn: 'DATEDIF' });
    }
  });

  def('EOMONTH', 2, 2, function (args) {
    var start = toSerial(args[0].value());
    if (E.isError(start)) return start;
    var months = toNumber(args[1].value(), { fn: 'EOMONTH' });
    if (E.isError(months)) return months;
    var p = U.partsFromSerial(start);
    var total = p.y * 12 + (p.m - 1) + Math.round(months);
    var y = Math.floor(total / 12);
    var m = (total % 12 + 12) % 12 + 1;
    var lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return U.serialFromYMD(y, m, lastDay);
  });

  def('NETWORKDAYS', 2, 3, function (args, ctx) {
    var a = toSerial(args[0].value());
    if (E.isError(a)) return a;
    var b = toSerial(args[1].value());
    if (E.isError(b)) return b;
    var holidays = [];
    if (args.length >= 3) {
      var hv = flatValues(args[2], ctx);
      var e = firstError(hv);
      if (e) return e;
      for (var i = 0; i < hv.length; i++) {
        var s = toSerial(hv[i]);
        if (!E.isError(s)) holidays.push(s);
      }
    }
    var lo = Math.min(a, b);
    var hi = Math.max(a, b);
    var n = 0;
    for (var d = lo; d <= hi; d++) {
      var dow = U.partsFromSerial(d).dow;
      if (dow === 0 || dow === 6) continue;
      if (holidays.indexOf(d) >= 0) continue;
      n++;
    }
    return a > b ? -n : n;
  });

  /* ---------- 判断 ---------- */

  def('ISNUMBER', 1, 1, function (args) {
    var v = args[0].value();
    if (E.isError(v)) return false;
    return typeof v === 'number';
  });

  def('ISTEXT', 1, 1, function (args) {
    var v = args[0].value();
    if (E.isError(v)) return false;
    return typeof v === 'string' && v !== '';
  });

  def('ISBLANK', 1, 1, function (args) {
    var v = args[0].value();
    if (E.isError(v)) return false;
    return v === null || v === undefined;
  });

  /* ---------- 调用入口 ---------- */

  function call(node, ctx) {
    var f = registry[node.name];
    if (!f) return E.err(C.NAME, { name: node.name });
    if (node.args.length < f.min || (f.max !== null && node.args.length > f.max)) {
      return E.err(C.VALUE, { arity: true, fn: node.name });
    }
    var args = node.args.map(function (n) {
      var done = false;
      var cache;
      return {
        node: n,
        value: function () {
          if (!done) {
            cache = Mobai.Evaluator.evalNode(n, ctx);
            done = true;
          }
          return cache;
        }
      };
    });
    var out = f.fn(args, ctx, node);
    return out === undefined ? null : out;
  }

  function names() {
    return Object.keys(registry).sort();
  }

  Mobai.Functions = {
    registry: registry,
    call: call,
    names: names,
    // 供其它模块复用的值处理
    toNumber: toNumber,
    toText: toText,
    toBool: toBool,
    toSerial: toSerial,
    textOf: textOf,
    displayOf: displayOf,
    looseEqual: looseEqual,
    makeCriteria: makeCriteria
  };
})(typeof window !== 'undefined' ? window : globalThis);