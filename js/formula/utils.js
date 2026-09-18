/* mobai训练 · 公式引擎通用工具
 * 地址、区域、数值比较、日期序列号。
 * 行列约定：列 0 起（A=0），行 1 起（与 Excel 一致）。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});

  /* ---------- 列名与地址 ---------- */

  function colToName(col) {
    var s = '';
    var n = col + 1;
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function nameToCol(name) {
    var s = 0;
    for (var i = 0; i < name.length; i++) {
      s = s * 26 + (name.toUpperCase().charCodeAt(i) - 64);
    }
    return s - 1;
  }

  function addrOf(col, row) {
    return colToName(col) + row;
  }

  var REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;

  /** 'A1' / '$A$1' / 'A$1' / '$A1' → {col,row,colAbs,rowAbs}，不是引用则返回 null */
  function parseRefToken(tok) {
    var m = REF_RE.exec(String(tok).trim());
    if (!m) return null;
    var col = nameToCol(m[2]);
    if (col > 16383) return null;
    return {
      col: col,
      row: parseInt(m[4], 10),
      colAbs: m[1] === '$',
      rowAbs: m[3] === '$'
    };
  }

  /** 'A1' → {col,row}，不关心绝对引用标志 */
  function parseAddr(addr) {
    var r = parseRefToken(addr);
    return r ? { col: r.col, row: r.row } : null;
  }

  function refToString(ref) {
    return (ref.colAbs ? '$' : '') + colToName(ref.col) + (ref.rowAbs ? '$' : '') + ref.row;
  }

  /* ---------- 区域 ---------- */

  function normRange(a, b) {
    return {
      start: { col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) },
      end: { col: Math.max(a.col, b.col), row: Math.max(a.row, b.row) }
    };
  }

  /** 'A1:B10' 或 'A1' → 规范化的区域对象；非法返回 null */
  function rangeOf(str) {
    var parts = String(str).trim().split(':');
    if (parts.length > 2) return null;
    var a = parseRefToken(parts[0]);
    if (!a) return null;
    var b = parts.length === 2 ? parseRefToken(parts[1]) : a;
    if (!b) return null;
    return normRange(a, b);
  }

  function rangeToString(r) {
    var s = colToName(r.start.col) + r.start.row;
    var e = colToName(r.end.col) + r.end.row;
    return s === e ? s : s + ':' + e;
  }

  /** 行优先展开为 [{col,row}, ...] */
  function expandRange(r) {
    var out = [];
    for (var row = r.start.row; row <= r.end.row; row++) {
      for (var col = r.start.col; col <= r.end.col; col++) {
        out.push({ col: col, row: row });
      }
    }
    return out;
  }

  function rangeCount(r) {
    return (r.end.col - r.start.col + 1) * (r.end.row - r.start.row + 1);
  }

  /** outer 是否完全覆盖 inner */
  function rangeContains(outer, inner) {
    return outer.start.col <= inner.start.col && outer.end.col >= inner.end.col &&
      outer.start.row <= inner.start.row && outer.end.row >= inner.end.row;
  }

  function rangeOverlap(a, b) {
    return !(a.end.col < b.start.col || b.end.col < a.start.col ||
      a.end.row < b.start.row || b.end.row < a.start.row);
  }

  function rangeEqual(a, b) {
    return a.start.col === b.start.col && a.start.row === b.start.row &&
      a.end.col === b.end.col && a.end.row === b.end.row;
  }

  /** 该地址是否落在某组区域字符串内（如 ['A1:H1', 'J2:L4']） */
  function addrInRanges(col, row, rangeStrings) {
    if (!rangeStrings) return false;
    for (var i = 0; i < rangeStrings.length; i++) {
      var r = typeof rangeStrings[i] === 'string' ? rangeOf(rangeStrings[i]) : rangeStrings[i];
      if (!r) continue;
      if (col >= r.start.col && col <= r.end.col && row >= r.start.row && row <= r.end.row) return true;
    }
    return false;
  }

  function flatten(matrix) {
    var out = [];
    for (var i = 0; i < matrix.length; i++) {
      for (var j = 0; j < matrix[i].length; j++) out.push(matrix[i][j]);
    }
    return out;
  }

  /* ---------- 数值 ---------- */

  var TOL = 1e-9;

  function almostEqual(a, b) {
    if (a === b) return true;
    if (typeof a !== 'number' || typeof b !== 'number') return false;
    if (!isFinite(a) || !isFinite(b)) return false;
    var d = Math.abs(a - b);
    return d <= TOL * Math.max(1, Math.abs(a), Math.abs(b));
  }

  /** 显示用数字：去掉浮点尾巴，不显示科学计数 */
  function formatNumber(n) {
    if (typeof n !== 'number' || !isFinite(n)) return String(n);
    if (Number.isInteger(n)) return String(n);
    var r = Math.round(n * 1e10) / 1e10;
    if (Number.isInteger(r)) return String(r);
    return String(r);
  }

  /* ---------- 日期 ----------
   * 用 Excel 序列号表示日期：1899-12-30 为 0。忽略 1900 闰年 bug。
   */

  var EPOCH = Date.UTC(1899, 11, 30);
  var DAY_MS = 86400000;

  function serialFromYMD(y, m, d) {
    return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS);
  }

  function partsFromSerial(serial) {
    var dt = new Date(EPOCH + Math.round(serial) * DAY_MS);
    return {
      y: dt.getUTCFullYear(),
      m: dt.getUTCMonth() + 1,
      d: dt.getUTCDate(),
      dow: dt.getUTCDay() // 0 = 周日
    };
  }

  function pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  function isoFromSerial(serial) {
    var p = partsFromSerial(serial);
    return p.y + '-' + pad2(p.m) + '-' + pad2(p.d);
  }

  function serialFromISO(iso) {
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(iso).trim());
    return m ? serialFromYMD(+m[1], +m[2], +m[3]) : null;
  }

  function looksLikeISO(text) {
    return /^\d{4}-\d{1,2}-\d{1,2}$/.test(String(text).trim());
  }

  function nowSerial() {
    var d = new Date();
    return serialFromYMD(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  /** 本地日期字符串 YYYY-MM-DD */
  function localDateStr(date) {
    var d = date || new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  Mobai.Utils = {
    colToName: colToName,
    nameToCol: nameToCol,
    addrOf: addrOf,
    parseRefToken: parseRefToken,
    parseAddr: parseAddr,
    refToString: refToString,
    normRange: normRange,
    rangeOf: rangeOf,
    rangeToString: rangeToString,
    expandRange: expandRange,
    rangeCount: rangeCount,
    rangeContains: rangeContains,
    rangeOverlap: rangeOverlap,
    rangeEqual: rangeEqual,
    addrInRanges: addrInRanges,
    flatten: flatten,
    TOL: TOL,
    almostEqual: almostEqual,
    formatNumber: formatNumber,
    serialFromYMD: serialFromYMD,
    partsFromSerial: partsFromSerial,
    isoFromSerial: isoFromSerial,
    serialFromISO: serialFromISO,
    looksLikeISO: looksLikeISO,
    nowSerial: nowSerial,
    localDateStr: localDateStr,
    pad2: pad2
  };
})(typeof window !== 'undefined' ? window : globalThis);