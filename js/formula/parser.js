/* mobai训练 · 公式解析器
 * 先解析成 AST，再求值——判分要回答「用了哪个函数、引用了哪个区域、有没有绝对
 * 引用」，只有 AST 能可靠回答，所以引擎里不允许出现字符串匹配判分。
 *
 * 文法（由低到高，与 Excel 一致）：
 *   expr           → comparison
 *   comparison     → concat ( (= | <> | < | > | <= | >=) concat )*
 *   concat         → additive ( & additive )*
 *   additive       → multiplicative ( (+ | -) multiplicative )*
 *   multiplicative → power ( (* | /) power )*
 *   power          → signed ( ^ power )?
 *   signed         → (- | +) signed | postfix
 *   postfix        → primary ( % )*
 *   primary        → 数字 | 字符串 | 错误值 | TRUE | FALSE | 函数(args) | 引用 | 区域 | ( expr )
 *
 * 注意 signed 在 power 之内：Excel 里 =-2^2 得到 4，不是 -4。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;
  var E = Mobai.Errors;

  function ParseError(message) {
    this.message = message;
    this.isParseError = true;
  }
  ParseError.prototype.toString = function () {
    return this.message;
  };

  var OP_TOKENS = ['<=', '>=', '<>', '<', '>', '=', '+', '-', '*', '/', '^', '&', '%', '(', ')', ',', ':'];
  var COMPARE_OPS = ['=', '<>', '<', '>', '<=', '>='];

  /* ---------- 词法 ---------- */

  function tokenize(src) {
    var tokens = [];
    var i = 0;
    while (i < src.length) {
      var ch = src[i];
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '\u3000') {
        i++;
        continue;
      }

      // 错误值字面量：#DIV/0! / #N/A / ...
      if (ch === '#') {
        var hit = null;
        for (var k = 0; k < E.list.length; k++) {
          if (src.slice(i, i + E.list[k].length).toUpperCase() === E.list[k]) {
            hit = E.list[k];
            break;
          }
        }
        if (!hit) throw new ParseError('「#」开头的内容不是已知的错误值');
        tokens.push({ t: 'error', v: hit });
        i += hit.length;
        continue;
      }

      // 字符串，内部的 "" 表示一个引号
      if (ch === '"') {
        var buf = '';
        i++;
        for (;;) {
          if (i >= src.length) throw new ParseError('字符串缺少结尾的引号');
          if (src[i] === '"') {
            if (src[i + 1] === '"') {
              buf += '"';
              i += 2;
              continue;
            }
            i++;
            break;
          }
          buf += src[i++];
        }
        tokens.push({ t: 'string', v: buf });
        continue;
      }

      // 数字
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) {
        var mn = /^\d*\.?\d+(?:[eE][+-]?\d+)?/.exec(src.slice(i));
        tokens.push({ t: 'number', v: parseFloat(mn[0]) });
        i += mn[0].length;
        continue;
      }

      // 标识符：函数名、单元格引用、TRUE/FALSE
      if (/[A-Za-z_$]/.test(ch)) {
        var mi = /^[A-Za-z_$][A-Za-z0-9_$.]*/.exec(src.slice(i));
        tokens.push({ t: 'ident', v: mi[0] });
        i += mi[0].length;
        continue;
      }

      var op = null;
      for (var j = 0; j < OP_TOKENS.length; j++) {
        if (src.substr(i, OP_TOKENS[j].length) === OP_TOKENS[j]) {
          op = OP_TOKENS[j];
          break;
        }
      }
      if (op) {
        tokens.push({ t: 'op', v: op });
        i += op.length;
        continue;
      }

      throw new ParseError('出现了无法识别的字符「' + ch + '」');
    }
    tokens.push({ t: 'eof' });
    return tokens;
  }

  /* ---------- 语法 ---------- */

  function Parser(tokens) {
    this.ts = tokens;
    this.i = 0;
  }

  Parser.prototype.peek = function (k) {
    return this.ts[this.i + (k || 0)] || this.ts[this.ts.length - 1];
  };
  Parser.prototype.next = function () {
    return this.ts[this.i++];
  };
  Parser.prototype.isOp = function (v) {
    var t = this.peek();
    return t.t === 'op' && t.v === v;
  };
  Parser.prototype.eat = function (v) {
    if (this.isOp(v)) {
      this.i++;
      return true;
    }
    return false;
  };
  Parser.prototype.expect = function (v) {
    if (!this.eat(v)) throw new ParseError('这里应该是「' + v + '」');
  };

  Parser.prototype.parse = function () {
    var node = this.expr();
    if (this.peek().t !== 'eof') throw new ParseError('公式末尾还有多余的内容');
    return node;
  };

  Parser.prototype.expr = function () {
    return this.comparison();
  };

  Parser.prototype.comparison = function () {
    var left = this.concat();
    while (this.peek().t === 'op' && COMPARE_OPS.indexOf(this.peek().v) >= 0) {
      var op = this.next().v;
      left = { type: 'Binary', op: op, left: left, right: this.concat() };
    }
    return left;
  };

  Parser.prototype.concat = function () {
    var left = this.additive();
    while (this.isOp('&')) {
      this.next();
      left = { type: 'Binary', op: '&', left: left, right: this.additive() };
    }
    return left;
  };

  Parser.prototype.additive = function () {
    var left = this.multiplicative();
    while (this.isOp('+') || this.isOp('-')) {
      var op = this.next().v;
      left = { type: 'Binary', op: op, left: left, right: this.multiplicative() };
    }
    return left;
  };

  Parser.prototype.multiplicative = function () {
    var left = this.power();
    while (this.isOp('*') || this.isOp('/')) {
      var op = this.next().v;
      left = { type: 'Binary', op: op, left: left, right: this.power() };
    }
    return left;
  };

  Parser.prototype.power = function () {
    var base = this.signed();
    if (this.isOp('^')) {
      this.next();
      return { type: 'Binary', op: '^', left: base, right: this.power() };
    }
    return base;
  };

  Parser.prototype.signed = function () {
    if (this.isOp('-')) {
      this.next();
      return { type: 'Unary', op: '-', operand: this.signed() };
    }
    if (this.isOp('+')) {
      this.next();
      return this.signed();
    }
    return this.postfix();
  };

  Parser.prototype.postfix = function () {
    var node = this.primary();
    while (this.isOp('%')) {
      this.next();
      node = { type: 'Percent', operand: node };
    }
    return node;
  };

  Parser.prototype.primary = function () {
    var t = this.peek();

    if (t.t === 'number') {
      this.next();
      return { type: 'Number', value: t.v };
    }
    if (t.t === 'string') {
      this.next();
      return { type: 'String', value: t.v };
    }
    if (t.t === 'error') {
      this.next();
      return { type: 'Error', value: t.v };
    }
    if (this.isOp('(')) {
      this.next();
      var inner = this.expr();
      this.expect(')');
      return inner;
    }
    if (t.t === 'ident') {
      // 函数调用
      if (this.peek(1).t === 'op' && this.peek(1).v === '(') {
        this.next();
        this.next();
        var args = [];
        if (!this.isOp(')')) {
          args.push(this.expr());
          while (this.eat(',')) args.push(this.expr());
        }
        this.expect(')');
        return { type: 'Func', name: t.v.toUpperCase(), args: args, raw: t.v };
      }
      // 单元格引用 / 区域
      var ref = U.parseRefToken(t.v);
      if (ref) {
        this.next();
        if (this.isOp(':')) {
          this.next();
          var t2 = this.peek();
          if (t2.t !== 'ident') throw new ParseError('区域写法不完整，冒号后面应该是单元格地址');
          var ref2 = U.parseRefToken(t2.v);
          if (!ref2) throw new ParseError('「' + t2.v + '」不是合法的单元格地址');
          this.next();
          return { type: 'Range', start: ref, end: ref2, raw: t.v + ':' + t2.v };
        }
        return { type: 'Ref', ref: ref, raw: t.v };
      }
      // 常量与未知名字
      this.next();
      var up = t.v.toUpperCase();
      if (up === 'TRUE') return { type: 'Bool', value: true };
      if (up === 'FALSE') return { type: 'Bool', value: false };
      return { type: 'Name', name: t.v };
    }

    throw new ParseError('公式在这里不完整');
  };

  /* ---------- 对外接口 ---------- */

  /**
   * 解析公式文本（可带前导 =）。
   * @returns {{ok:true, ast:Object, text:string} | {ok:false, message:string}}
   */
  function parseFormula(src) {
    var text = String(src == null ? '' : src);
    if (text.charAt(0) === '=') text = text.slice(1);
    if (!text.trim()) return { ok: false, message: '公式是空的' };
    try {
      var ast = new Parser(tokenize(text)).parse();
      return { ok: true, ast: ast, text: text };
    } catch (ex) {
      if (ex && ex.isParseError) return { ok: false, message: ex.message };
      throw ex;
    }
  }

  /** 公式里用到的函数名（大写、去重） */
  function collectFunctions(ast) {
    var names = [];
    (function walk(node) {
      if (!node || typeof node !== 'object') return;
      if (node.type === 'Func') {
        if (names.indexOf(node.name) < 0) names.push(node.name);
        node.args.forEach(walk);
        return;
      }
      if (node.type === 'Binary') {
        walk(node.left);
        walk(node.right);
        return;
      }
      if (node.type === 'Unary' || node.type === 'Percent') {
        walk(node.operand);
      }
    })(ast);
    return names;
  }

  /**
   * 公式里的引用，按出现顺序去重返回。
   * { kind:'ref'|'range', start:{col,row}, end:{col,row}, fullyAbsolute:boolean, raw:string }
   */
  function collectRefs(ast) {
    var out = [];

    function push(entry) {
      for (var i = 0; i < out.length; i++) {
        if (out[i].kind === entry.kind && out[i].raw === entry.raw) return;
      }
      out.push(entry);
    }

    (function walk(node) {
      if (!node || typeof node !== 'object') return;
      if (node.type === 'Ref') {
        push({
          kind: 'ref',
          raw: node.raw,
          start: { col: node.ref.col, row: node.ref.row },
          end: { col: node.ref.col, row: node.ref.row },
          fullyAbsolute: node.ref.colAbs && node.ref.rowAbs
        });
        return;
      }
      if (node.type === 'Range') {
        push({
          kind: 'range',
          raw: node.raw,
          start: { col: node.start.col, row: node.start.row },
          end: { col: node.end.col, row: node.end.row },
          fullyAbsolute: node.start.colAbs && node.start.rowAbs && node.end.colAbs && node.end.rowAbs
        });
        return;
      }
      if (node.type === 'Func') {
        node.args.forEach(walk);
        return;
      }
      if (node.type === 'Binary') {
        walk(node.left);
        walk(node.right);
        return;
      }
      if (node.type === 'Unary' || node.type === 'Percent') {
        walk(node.operand);
      }
    })(ast);

    return out;
  }

  Mobai.Parser = {
    ParseError: ParseError,
    tokenize: tokenize,
    parseFormula: parseFormula,
    collectFunctions: collectFunctions,
    collectRefs: collectRefs
  };
})(typeof window !== 'undefined' ? window : globalThis);