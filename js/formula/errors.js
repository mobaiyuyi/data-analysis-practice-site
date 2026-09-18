/* mobai训练 · 错误值
 * 引擎只产生 6 个错误码；每个错误码配一份中文成因解释（独立的教学资产，
 * 与判分提示分开维护）。判分未通过时展示的是解释，而不是错误码本身。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});

  var CODES = {
    DIV0: '#DIV/0!',
    VALUE: '#VALUE!',
    NA: '#N/A',
    NAME: '#NAME?',
    REF: '#REF!',
    NUM: '#NUM!'
  };

  function FormulaError(code, info) {
    this.code = code;
    this.info = info || {};
  }

  FormulaError.prototype.toString = function () {
    return this.code;
  };

  function isError(v) {
    return !!v && v instanceof FormulaError;
  }

  function err(code, info) {
    return new FormulaError(code, info);
  }

  /* 中文成因解释 */
  function explain(value) {
    if (!isError(value)) return '';
    var info = value.info || {};
    switch (value.code) {
      case CODES.DIV0:
        return '除数为 0 或为空。先确认分母有没有值，再用 IF 或 IFERROR 兜住这种情况。';
      case CODES.VALUE:
        if (info.arity) {
          return '「' + (info.fn || '这个函数') + '」的参数个数不对，检查一下逗号有没有多写或少写。';
        }
        if (info.range) {
          return '一整片区域（' + info.range + '）不能当单个值参与运算，需要先把它聚合起来——比如外面套 SUM 或先选出某一个单元格。';
        }
        return '参数类型不对：多半是拿文本去做加减乘除，或者区域和单个值混在一起算。';
      case CODES.NA:
        var s = '查找函数没找到目标值。最常见的原因是查找值不在所选区域的第一列，或者两边格式不一致（一边是文本型数字）。';
        if (info.lookupValue !== undefined && info.lookupValue !== null && info.lookupValue !== '') {
          s = '没有找到「' + info.lookupValue + '」。' + s;
        }
        if (info.fn) s += '（出错的函数：' + info.fn + '）';
        return s;
      case CODES.NAME:
        if (info.syntax) {
          return '公式语法有问题：' + info.syntax;
        }
        if (info.name) {
          return '「' + info.name + '」不是引擎认识的名字：函数名拼错了，或者用了本引擎还没实现的函数。';
        }
        return '函数名拼写错误，或者用了引擎不支持的写法。';
      case CODES.REF:
        if (info.circular) {
          return '循环引用：' + (info.addr ? info.addr + ' 的公式' : '这个公式') + '直接或间接引用了自己。把链条上的某一环改成引用别的单元格。';
        }
        return '引用失效：引用的单元格不在表内，或指向了已经被清空的位置。';
      case CODES.NUM:
        return '数值不合法或超出范围（例如把负数当日期、开平方得到虚数）。检查一下参与计算的数字。';
      default:
        return '未知错误。';
    }
  }

  Mobai.Errors = {
    CODES: CODES,
    FormulaError: FormulaError,
    isError: isError,
    err: err,
    explain: explain,
    list: [CODES.DIV0, CODES.VALUE, CODES.NA, CODES.NAME, CODES.REF, CODES.NUM]
  };
})(typeof window !== 'undefined' ? window : globalThis);