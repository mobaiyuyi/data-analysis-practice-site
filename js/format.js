/* mobai训练 · 数字格式
 * Excel 格式串的一小撮子集：General / 0 / 0.00 / #,##0 / 0% / ¥#,##0.00 / yyyy-mm-dd / @
 * 全是纯函数：render(值, 格式串) → 显示文本。
 *
 * 一条铁律：格式只影响「显示」，判分永远看原始值。所以这里永远不碰 workbook 的数据。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;

  /* ---------- 预设（功能区下拉里列的就是这些） ---------- */

  var PRESETS = [
    { code: 'General', label: '常规' },
    { code: '0', label: '数值' },
    { code: '0.00', label: '数值（两位小数）' },
    { code: '#,##0', label: '千分位' },
    { code: '#,##0.00', label: '千分位（两位小数）' },
    { code: '0%', label: '百分比' },
    { code: '0.00%', label: '百分比（两位小数）' },
    { code: '¥#,##0.00', label: '货币（¥）' },
    { code: '@', label: '文本' },
    { code: 'yyyy-mm-dd', label: '日期（2026-03-09）' }
  ];

  var DATE_STYLES = {
    date: 'yyyy-mm-dd',
    'yyyy-mm-dd': 'yyyy-mm-dd',
    'yyyy/m/d': 'yyyy/m/d',
    'm月d日': 'm月d日'
  };

  /* ---------- 解析 ---------- */

  /**
   * 格式串 → 描述对象
   * @returns {{kind:'general'|'text'|'date'|'number'|'percent'|'currency', decimals:number, thousands:boolean, symbol:string, style:string}}
   */
  function parse(code) {
    if (code === null || code === undefined) return { kind: 'general', decimals: 0, thousands: false, symbol: '', style: '' };
    var c = String(code).trim();
    if (c === '' || /^general$/i.test(c)) return { kind: 'general', decimals: 0, thousands: false, symbol: '', style: '' };
    if (c === '@') return { kind: 'text', decimals: 0, thousands: false, symbol: '', style: '' };
    if (DATE_STYLES[c]) return { kind: 'date', decimals: 0, thousands: false, symbol: '', style: DATE_STYLES[c] };

    var m = /^([^\d#]*)(#,##)?(0)(\.(0+))?(%?)$/.exec(c);
    if (m) {
      var percent = m[6] === '%';
      return {
        kind: percent ? 'percent' : (m[1] ? 'currency' : 'number'),
        decimals: m[4] ? m[4].length - 1 : 0,
        thousands: !!m[2],
        symbol: m[1] || '',
        style: ''
      };
    }
    return { kind: 'general', decimals: 0, thousands: false, symbol: '', style: '' };
  }

  /** 描述对象 → 格式串（增减小数位之后要重新拼回去） */
  function build(d) {
    if (!d || d.kind === 'general') return 'General';
    if (d.kind === 'text') return '@';
    if (d.kind === 'date') return d.style || 'date';
    var zeros = '0';
    if (d.decimals > 0) zeros += '.' + new Array(d.decimals + 1).join('0');
    return (d.symbol || '') + (d.thousands ? '#,##' : '') + zeros + (d.kind === 'percent' ? '%' : '');
  }

  function isDate(code) {
    return parse(code).kind === 'date';
  }

  function isGeneral(code) {
    return parse(code).kind === 'general';
  }

  function decimalsOf(code) {
    return parse(code).decimals;
  }

  /** 增减小数位；日期/文本/常规这三种不动 */
  function withDecimals(code, delta) {
    var d = parse(code);
    if (d.kind === 'date' || d.kind === 'text' || d.kind === 'general') return code;
    d.decimals = Math.max(0, Math.min(8, d.decimals + delta));
    return build(d);
  }

  /** 该格式是不是「千分位/货币这类数字格式」（用于判断要不要右对齐） */
  function isNumericFormat(code) {
    var k = parse(code).kind;
    return k === 'number' || k === 'percent' || k === 'currency';
  }

  /* ---------- 渲染 ---------- */

  function addThousands(s) {
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }

  function renderDate(serial, style) {
    var p = U.partsFromSerial(serial);
    if (style === 'yyyy/m/d') return p.y + '/' + p.m + '/' + p.d;
    if (style === 'm月d日') return p.m + '月' + p.d + '日';
    return U.isoFromSerial(serial);
  }

  /**
   * 把值按格式串渲染成显示文本。
   * @param {*} value 原始值（数字 / 日期序列号 / 文本 / 布尔）
   * @param {string} code 格式串
   */
  function render(value, code) {
    var d = parse(code);

    if (d.kind === 'general') {
      if (value === null || value === undefined) return '';
      if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
      if (typeof value === 'number') return U.formatNumber(value);
      return String(value);
    }

    if (d.kind === 'text') {
      if (value === null || value === undefined) return '';
      if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
      return String(value);
    }

    if (d.kind === 'date') {
      if (typeof value !== 'number' || !isFinite(value)) return value === null || value === undefined ? '' : String(value);
      return renderDate(value, d.style);
    }

    var n = typeof value === 'number' ? value : parseFloat(value);
    if (typeof n !== 'number' || isNaN(n) || !isFinite(n)) {
      return value === null || value === undefined ? '' : String(value);
    }
    if (d.kind === 'percent') n = n * 100;
    var negative = n < 0;
    var s = addThousands(Math.abs(n).toFixed(d.decimals));
    return (negative ? '-' : '') + d.symbol + s + (d.kind === 'percent' ? '%' : '');
  }

  Mobai.Format = {
    PRESETS: PRESETS,
    parse: parse,
    build: build,
    render: render,
    isDate: isDate,
    isGeneral: isGeneral,
    isNumericFormat: isNumericFormat,
    decimalsOf: decimalsOf,
    withDecimals: withDecimals
  };
})(typeof window !== 'undefined' ? window : globalThis);