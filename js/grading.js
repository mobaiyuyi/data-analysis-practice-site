/* mobai训练 · 判分
 * 三类检查项，无权重，全部通过才算这道题通过。
 *   cell_values        算对了吗
 *   formula_functions  用对工具了吗
 *   formula_references 能维护吗（绝对引用、是否引用了指定区域、是否填满）
 *
 * 判分只看 AST，不做字符串匹配：「VLOOKUP ( A2 , B:C , 2 , 0 )」与
 * 「VLOOKUP(A2,B:C,2,0)」在这里是同一个公式。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;
  var E = Mobai.Errors;
  var P = Mobai.Parser;

  /* ---------- 展示用 ---------- */

  function showValue(v) {
    if (v === null || v === undefined) return '（空）';
    if (E.isError(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return U.formatNumber(v);
    return '"' + v + '"';
  }

  function errorNote(v, fallback) {
    if (!E.isError(v)) return fallback;
    return v.code + ' —— ' + E.explain(v);
  }

  function joinCells(list, max) {
    var cap = max || 4;
    if (list.length <= cap) return list.join('、');
    return list.slice(0, cap).join('、') + ' 等 ' + list.length + ' 处';
  }

  /* ---------- cell_values ---------- */

  /** 期望值写法：数字 / 文本 / TRUE / null（必须是空的）/ '#N/A'（必须出错）/ '2026-03-01'（日期） */
  function sameAsExpected(got, expected) {
    if (typeof expected === 'string' && expected.charAt(0) === '#') {
      return E.isError(got) && got.code === expected.toUpperCase();
    }
    var iso = typeof expected === 'string' && U.looksLikeISO(expected) ? U.serialFromISO(expected) : null;
    if (iso !== null) {
      if (typeof got === 'number' && U.almostEqual(got, iso)) return true;
      // 日期也可能被写成了文本
      return typeof got === 'string' && got.trim() === expected;
    }
    if (expected === null || expected === undefined) {
      if (got === null || got === undefined) return true;
      return typeof got === 'string' && got.trim() === '';
    }
    if (typeof expected === 'number') {
      return typeof got === 'number' && U.almostEqual(got, expected);
    }
    if (typeof expected === 'boolean') return got === expected;
    // 文本：忽略首尾空白与大小写（Excel 比较本来就忽略大小写）
    if (typeof got === 'string') {
      return got.trim().toLowerCase() === String(expected).trim().toLowerCase();
    }
    return false;
  }

  function checkCellValues(check, wb) {
    var r = U.rangeOf(check.target);
    if (!r) return { passed: false, message: '检查项写错了：target「' + check.target + '」不是合法区域' };
    var cells = U.expandRange(r);
    // expected 可以写成一维数组，也可以写成与区域同形的二维数组；统一按行优先摊平
    var expected = [].concat(check.expected == null ? [] : check.expected);
    if (expected.length && Array.isArray(expected[0])) expected = U.flatten(expected);
    if (expected.length !== cells.length) {
      return {
        passed: false,
        message: '检查项写错了：' + check.target + ' 有 ' + cells.length + ' 格，但给了 ' + expected.length + ' 个期望值'
      };
    }

    var bad = [];
    for (var i = 0; i < cells.length; i++) {
      var addr = U.addrOf(cells[i].col, cells[i].row);
      var got = wb.getCell(cells[i].col, cells[i].row);
      if (sameAsExpected(got, expected[i])) continue;
      var line = addr + ' 得到 ' + showValue(got) + '，应为 ' + showValue(expected[i]);
      if (E.isError(got)) line += '（' + E.explain(got) + '）';
      bad.push(line);
    }
    if (!bad.length) return { passed: true, message: check.target + ' 共 ' + cells.length + ' 格，全部算对了。' };
    return {
      passed: false,
      message: joinCells(bad, 3) + (bad.length > 3 ? '（共 ' + bad.length + ' 格不对）' : '')
    };
  }

  /* ---------- 公式解析小工具 ---------- */

  /** 拿某格的公式 AST 与函数名；不是公式或解析失败时带出原因 */
  function formulaInfoOf(wb, col, row) {
    var text = wb.formulaText(col, row);
    if (!text) {
      return { missing: true, reason: '不是公式（' + (wb.rawText(col, row) || '空单元格') + '）' };
    }
    var parsed = P.parseFormula(text);
    if (!parsed.ok) {
      return { missing: true, reason: '公式看不懂：' + parsed.message, text: text };
    }
    return {
      missing: false,
      text: text,
      ast: parsed.ast,
      functions: P.collectFunctions(parsed.ast),
      refs: P.collectRefs(parsed.ast)
    };
  }

  /* ---------- formula_functions ---------- */

  function checkFormulaFunctions(check, wb) {
    var r = U.rangeOf(check.target);
    if (!r) return { passed: false, message: '检查项写错了：target「' + check.target + '」不是合法区域' };
    var must = (check.must || []).map(function (s) { return String(s).toUpperCase(); });
    var mustNot = (check.mustNot || []).map(function (s) { return String(s).toUpperCase(); });
    var mustAny = (check.mustAny || []).map(function (s) { return String(s).toUpperCase(); });
    var cells = U.expandRange(r);

    var noFormula = [];
    var missing = [];   // 'F2 里少了 SUM'
    var banned = [];    // 'F2 里不该出现 VLOOKUP'

    for (var i = 0; i < cells.length; i++) {
      var col = cells[i].col;
      var row = cells[i].row;
      var addr = U.addrOf(col, row);
      var info = formulaInfoOf(wb, col, row);
      if (info.missing) {
        noFormula.push(addr + '（' + info.reason + '）');
        continue;
      }
      for (var a = 0; a < must.length; a++) {
        if (info.functions.indexOf(must[a]) < 0) missing.push(addr + ' 少用了 ' + must[a] + '（现在写的是 ' + info.text + '）');
      }
      if (mustAny.length) {
        var hitAny = mustAny.some(function (n) { return info.functions.indexOf(n) >= 0; });
        if (!hitAny) {
          missing.push(addr + ' 需要用到 ' + mustAny.join(' 或 ') + '（现在写的是 ' + info.text + '）');
        }
      }
      for (var b = 0; b < mustNot.length; b++) {
        if (info.functions.indexOf(mustNot[b]) >= 0) banned.push(addr + ' 不该出现 ' + mustNot[b]);
      }
    }

    var parts = [];
    if (noFormula.length) parts.push(joinCells(noFormula, 3) + ' 不是公式');
    if (missing.length) parts.push(joinCells(missing, 3));
    if (banned.length) parts.push(joinCells(banned, 3));
    if (!parts.length) {
      var used = must.length ? '用到了 ' + must.join('、') + '。' : '公式结构符合要求。';
      return { passed: true, message: check.target + ' 全部是公式，且 ' + used };
    }
    return { passed: false, message: parts.join('；') + '。' };
  }

  /* ---------- formula_references ---------- */

  function checkFormulaReferences(check, wb) {
    var r = U.rangeOf(check.target);
    if (!r) return { passed: false, message: '检查项写错了：target「' + check.target + '」不是合法区域' };
    var cells = U.expandRange(r);
    var needFilled = check.filled !== false;
    var requiredRanges = (check.mustReference || []).map(function (s) {
      return { text: s, range: U.rangeOf(s) };
    });
    for (var k = 0; k < requiredRanges.length; k++) {
      if (!requiredRanges[k].range) {
        return { passed: false, message: '检查项写错了：mustReference「' + requiredRanges[k].text + '」不是合法区域' };
      }
    }

    var notFormula = [];
    var notCovered = [];   // 必须引用某区域却没引用
    var notAbsolute = [];  // 该锁住的引用没锁

    for (var i = 0; i < cells.length; i++) {
      var col = cells[i].col;
      var row = cells[i].row;
      var addr = U.addrOf(col, row);
      var info = formulaInfoOf(wb, col, row);
      if (info.missing) {
        if (needFilled) notFormula.push(addr + '（' + info.reason + '）');
        continue;
      }

      // 1) 是否引用了指定区域
      for (var a = 0; a < requiredRanges.length; a++) {
        var need = requiredRanges[a];
        var hit = false;
        for (var b = 0; b < info.refs.length; b++) {
          var refRange = U.normRange(info.refs[b].start, info.refs[b].end);
          if (U.rangeContains(refRange, need.range)) { hit = true; break; }
        }
        if (!hit) {
          notCovered.push(addr + ' 没有引用到 ' + need.text + '（现在写的是 ' + info.text + '）');
        }
      }

      // 2) 绝对引用：只看落在 mustReference 区域内的引用；没给 mustReference 就要求全部引用都锁住
      if (check.requireAbsolute) {
        for (var c = 0; c < info.refs.length; c++) {
          var ref = info.refs[c];
          var refRange2 = U.normRange(ref.start, ref.end);
          var inScope = true;
          if (requiredRanges.length) {
            inScope = requiredRanges.some(function (need2) {
              return U.rangeOverlap(refRange2, need2.range) || U.rangeContains(refRange2, need2.range);
            });
          }
          if (!inScope) continue;
          if (!ref.fullyAbsolute) {
            notAbsolute.push(addr + ' 里的 ' + ref.raw + ' 没有锁定（应写成带 $ 的形式）');
          }
        }
      }
    }

    var parts = [];
    if (notFormula.length) parts.push(joinCells(notFormula, 3) + ' 还不是公式');
    if (notCovered.length) parts.push(joinCells(notCovered, 3));
    if (notAbsolute.length) parts.push(joinCells(notAbsolute, 3));
    if (!parts.length) {
      var desc = [];
      if (requiredRanges.length) desc.push('引用了 ' + requiredRanges.map(function (n) { return n.text; }).join('、'));
      if (check.requireAbsolute) desc.push('引用都用 $ 锁住了');
      return { passed: true, message: check.target + ' 的引用结构正确' + (desc.length ? '：' + desc.join('，') : '') + '。' };
    }
    return { passed: false, message: parts.join('；') + '。' };
  }

  /* ---------- 对外 ---------- */

  var RUNNERS = {
    cell_values: checkCellValues,
    formula_functions: checkFormulaFunctions,
    formula_references: checkFormulaReferences
  };

  /**
   * 跑一套检查项。
   * @param {Array} checks 题目里的 checks
   * @param {Mobai.Workbook} wb
   * @returns {Array<{index,type,target,hint,passed,message}>}
   */
  function run(checks, wb) {
    var out = [];
    (checks || []).forEach(function (check, i) {
      var runner = RUNNERS[check.type];
      var result;
      if (!runner) {
        result = { passed: false, message: '检查项写错了：不知道「' + check.type + '」是什么检查类型' };
      } else {
        try {
          result = runner(check, wb);
        } catch (ex) {
          result = { passed: false, message: '判分时出错：' + (ex && ex.message ? ex.message : ex) };
        }
      }
      out.push({
        index: i,
        type: check.type,
        target: check.target,
        hint: check.hint || '',
        passed: !!result.passed,
        message: result.message
      });
    });
    return out;
  }

  function allPassed(results) {
    if (!results || !results.length) return false;
    for (var i = 0; i < results.length; i++) {
      if (!results[i].passed) return false;
    }
    return true;
  }

  /** 未通过的检查项类型（进度的 failedChecks 用） */
  function failedTypes(results) {
    var out = [];
    (results || []).forEach(function (r) {
      if (!r.passed && out.indexOf(r.type) < 0) out.push(r.type);
    });
    return out;
  }

  var TYPE_LABEL = {
    cell_values: '算对了吗',
    formula_functions: '用对工具了吗',
    formula_references: '能维护吗'
  };

  Mobai.Grading = {
    run: run,
    allPassed: allPassed,
    failedTypes: failedTypes,
    TYPE_LABEL: TYPE_LABEL,
    sameAsExpected: sameAsExpected,
    showValue: showValue,
    errorNote: errorNote
  };
})(typeof window !== 'undefined' ? window : globalThis);