/* 研数通途 · 填充与引用改写
 * 三件事，都是纯函数（tools/verify.js 直接测）：
 *   shiftFormulaText  公式里的相对引用按 (dCol,dRow) 平移
 *   transformRefs     把公式里每个引用交给回调重写（插入/删除行列时用）
 *   planFill          从一串源格内容推导出后续 n 格该填什么（等差 / 日期 / 文本编号 / 复制）
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;

  var IDENT_CH = /[A-Za-z_$0-9]/;

  /**
   * 遍历公式里的引用 token。字符串字面量里的 A1 不动，函数名不动。
   * @param {string} text 以 = 开头的公式
   * @param {Function} fn (ref, raw) → 新的引用对象；返回 null 表示这一处失效（写成 #REF!）
   */
  function transformRefs(text, fn) {
    if (typeof text !== 'string' || text.charAt(0) !== '=') return text;
    var out = '=';
    var body = text.slice(1);
    var i = 0;
    while (i < body.length) {
      var ch = body[i];
      if (ch === '"') {
        // 字符串整体搬过去，里面出现的 A1 不能被当成引用
        out += ch;
        i++;
        while (i < body.length) {
          out += body[i];
          if (body[i] === '"') {
            if (body[i + 1] === '"') { out += body[i + 1]; i += 2; continue; }
            i++;
            break;
          }
          i++;
        }
        continue;
      }
      if (/[A-Za-z_$]/.test(ch)) {
        var j = i;
        while (j < body.length && IDENT_CH.test(body[j])) j++;
        var tok = body.slice(i, j);
        var ref = U.parseRefToken(tok);
        if (ref) {
          var next = fn(ref, tok);
          out += next ? U.refToString(next) : '#REF!';
        } else {
          out += tok;
        }
        i = j;
        continue;
      }
      out += ch;
      i++;
    }
    return out;
  }

  /** 相对引用平移（复制粘贴、拖动填充、Ctrl+D 都用它） */
  function shiftFormulaText(text, dCol, dRow) {
    if (!text || text.charAt(0) !== '=' || (!dCol && !dRow)) return text;
    return transformRefs(text, function (ref) {
      return {
        col: ref.col + (ref.colAbs ? 0 : dCol),
        row: ref.row + (ref.rowAbs ? 0 : dRow),
        colAbs: ref.colAbs,
        rowAbs: ref.rowAbs
      };
    });
  }

  /* ---------- 填充序列 ---------- */

  function isFormula(s) {
    return typeof s === 'string' && s.charAt(0) === '=';
  }

  function asNumber(s) {
    if (typeof s === 'number') return s;
    var t = String(s == null ? '' : s).trim();
    if (t === '' || !/^-?\d+(\.\d+)?$/.test(t)) return null;
    return parseFloat(t);
  }

  /** 文本末尾的编号：「第3周」→ {prefix:'第', digits:'3', suffix:'周'}；没有编号返回 null */
  function trailingNumber(s) {
    var m = /^(.*?)(\d+)$/.exec(String(s));
    if (!m) return null;
    return { prefix: m[1], digits: m[2], value: parseInt(m[2], 10) };
  }

  /**
   * 从源格内容推导后续 n 格。
   * @param {string[]} sources 源格原始内容（自上而下 / 自左而右）
   * @param {number} count 要新填几格
   * @param {{dates:boolean[], axis:'row'|'col'}} opts dates 与 sources 一一对应，标记这一格是不是日期
   * @returns {{kind:'copy'|'series'|'date'|'text'|'formula', values:string[]}}
   */
  function planFill(sources, count, opts) {
    opts = opts || {};
    var h = sources.length;
    var n = Math.max(0, count | 0);
    var empty = { kind: 'copy', values: [] };
    for (var e = 0; e < n; e++) empty.values.push('');
    if (!h || !n) return empty;

    var filled = sources.filter(function (s) { return s !== '' && s !== null && s !== undefined; });
    if (!filled.length) return empty;

    var axisIsRow = opts.axis !== 'col';

    // 1) 含公式：按相对引用平移（跟 Excel 一样，预览显示「复制」）
    if (sources.some(isFormula)) {
      var fOut = [];
      for (var j = 0; j < n; j++) {
        var i = j % h;
        var src = sources[i];
        if (!src) { fOut.push(''); continue; }
        var d = h + j - i;                  // 这一格相对源格走了几行/几列
        fOut.push(isFormula(src)
          ? shiftFormulaText(src, axisIsRow ? 0 : d, axisIsRow ? d : 0)
          : src);
      }
      return { kind: 'formula', values: fOut };
    }

    var dates = opts.dates || [];
    var allDateSlot = filled.length > 0 && sources.every(function (s, k) {
      return s === '' || s === null || s === undefined || (dates[k] && asNumber(s) !== null);
    }) && sources.some(function (s, k) { return s !== '' && dates[k] && asNumber(s) !== null; });

    // 2) 日期序列：单格按天 +1，多格沿用日差
    if (allDateSlot) {
      var dNums = filled.map(asNumber);
      var dStep = dNums.length > 1 ? (dNums[dNums.length - 1] - dNums[0]) / (dNums.length - 1) : 1;
      var dLast = dNums[dNums.length - 1];
      var dOut = [];
      for (var dj = 0; dj < n; dj++) dOut.push(U.isoFromSerial(dLast + dStep * (dj + 1)));
      return { kind: 'date', values: dOut };
    }

    // 3) 数字序列：单格按 1 递增，多格沿用等差步长
    var nums = filled.map(asNumber);
    if (nums.every(function (v) { return v !== null; })) {
      var step = nums.length > 1 ? (nums[nums.length - 1] - nums[0]) / (nums.length - 1) : 1;
      if (step === 0) return { kind: 'copy', values: repeatCopy(sources, n) };
      var last = nums[nums.length - 1];
      var nOut = [];
      for (var nj = 0; nj < n; nj++) nOut.push(U.formatNumber(last + step * (nj + 1)));
      return { kind: 'series', values: nOut };
    }

    // 4) 文本末尾带编号：「第1周」「A1」
    var tn = filled.map(trailingNumber);
    if (tn.every(function (t) { return t && t.prefix === tn[0].prefix; })) {
      var tNums = tn.map(function (t) { return t.value; });
      var tStep = tNums.length > 1 ? (tNums[tNums.length - 1] - tNums[0]) / (tNums.length - 1) : 1;
      if (tStep !== 0) {
        var padLen = tn[tn.length - 1].digits.length;
        var tLast = tNums[tNums.length - 1];
        var tOut = [];
        for (var tj = 0; tj < n; tj++) {
          var num = String(tLast + tStep * (tj + 1));
          while (num.length < padLen) num = '0' + num;
          tOut.push(tn[0].prefix + num);
        }
        return { kind: 'text', values: tOut };
      }
    }

    // 5) 其它：原样复制
    return { kind: 'copy', values: repeatCopy(sources, n) };
  }

  function repeatCopy(sources, n) {
    var out = [];
    for (var j = 0; j < n; j++) out.push(sources[j % sources.length]);
    return out;
  }

  /** 拖动填充时气泡里写什么（Excel 里「填充序列 97」「复制」） */
  function previewText(plan) {
    if (!plan || !plan.values.length) return '复制';
    var last = plan.values[plan.values.length - 1];
    if (plan.kind === 'series') return '填充序列 ' + last;
    if (plan.kind === 'date') return '填充日期 ' + last;
    if (plan.kind === 'text') return '填充序列 ' + last;
    return '复制';
  }

  Mobai.Fill = {
    shiftFormulaText: shiftFormulaText,
    transformRefs: transformRefs,
    planFill: planFill,
    previewText: previewText,
    trailingNumber: trailingNumber
  };
})(typeof window !== 'undefined' ? window : globalThis);