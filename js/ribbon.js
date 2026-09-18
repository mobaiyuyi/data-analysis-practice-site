/* mobai训练 · 功能区
 * Excel 顶部那一排标签与按钮。只有「开始」有真按钮，其它标签点开写明「这个练习用不到」——
 * 不为凑数量放一堆灰按钮，那只会让人以为自己漏学了什么。
 *
 * 功能区不碰表格内部状态，所有动作都通过 sheet.commands 走。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});

  var TABS = ['文件', '开始', '插入', '公式', '数据', '审阅', '视图'];

  var TAB_NOTES = {
    '文件': '这里没有文件要保存——练习进度存在你自己的浏览器里。要备份或换电脑用，去模块页点「导出进度」。',
    '插入': '图表、图片、透视表在这里都用不到。这一批题目练的是函数与引用，不是插图。',
    '公式': '函数库、名称管理器在这里用不到。公式直接在下面的 fx 栏里写，也可以选中单元格后按 = 打头直接敲；' +
      '写公式时用鼠标点单元格就能把引用点进来，这比记住函数列表有用得多。',
    '数据': '排序、筛选、分列在这里用不到。要练的是「看到一张脏表，知道该上什么手法」，不是熟悉按钮位置。',
    '审阅': '批注、保护工作表在这里用不到。',
    '视图': '冻结窗格、网格线开关在这里用不到。'
  };

  var FONT_NAMES = [
    { value: '', label: '默认字体' },
    { value: '"PingFang SC","Microsoft YaHei",sans-serif', label: '黑体（无衬线）' },
    { value: '"Songti SC",SimSun,serif', label: '宋体（衬线）' },
    { value: 'Consolas,"Courier New",monospace', label: '等宽' }
  ];

  var FONT_SIZES = [11, 12, 14, 18, 24];

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = text;
    return n;
  }

  /* ---------- 小零件 ---------- */

  function btn(label, title, onClick, cls) {
    var b = el('button', 'rb-btn' + (cls ? ' ' + cls : ''), label);
    b.type = 'button';
    b.title = title || label;
    // 按下时不抢焦点：否则正在编辑的格子会先被提交、公式栏会闪一下
    b.addEventListener('mousedown', function (e) { e.preventDefault(); });
    b.addEventListener('click', function () { onClick(); });
    return b;
  }

  function select(options, title, onChange) {
    var s = el('select', 'rb-select');
    s.title = title || '';
    options.forEach(function (o) {
      var opt = el('option', null, o.label);
      opt.value = o.value;
      s.appendChild(opt);
    });
    s.addEventListener('mousedown', function (e) { e.stopPropagation(); });
    s.addEventListener('change', function () { onChange(s.value); });
    return s;
  }

  function colorInput(title, onChange) {
    var c = el('input', 'rb-color');
    c.type = 'color';
    c.title = title;
    c.value = title.indexOf('填充') >= 0 ? '#fff2cc' : '#c2364b';
    c.addEventListener('input', function () { onChange(c.value); });
    return c;
  }

  function group(label, items) {
    var g = el('div', 'rb-group');
    var row = el('div', 'rb-group-items');
    items.forEach(function (n) {
      if (n === '-') row.appendChild(el('span', 'rb-sep'));
      else row.appendChild(n);
    });
    g.appendChild(row);
    g.appendChild(el('div', 'rb-group-label', label));
    return g;
  }

  /* ---------- 挂载 ---------- */

  /**
   * @param {HTMLElement} container 功能区容器（表格组件里的 .ribbon-slot）
   * @param {Object} sheet Mobai.Sheet.mount 返回的实例
   */
  function mount(container, sheet) {
    var C = sheet.commands;
    var current = '开始';

    var wrap = el('div', 'ribbon');

    /* 标签条 */
    var tabStrip = el('div', 'rb-tabs');
    var tabButtons = {};
    TABS.forEach(function (name) {
      var b = el('button', 'rb-tab', name);
      b.type = 'button';
      b.addEventListener('mousedown', function (e) { e.preventDefault(); });
      b.addEventListener('click', function () { show(name); });
      tabButtons[name] = b;
      tabStrip.appendChild(b);
    });
    var tabHint = el('span', 'rb-tab-hint', '写公式时点单元格即可把引用点进来');
    tabStrip.appendChild(tabHint);
    wrap.appendChild(tabStrip);

    /* 「开始」面板 */
    var body = el('div', 'rb-body');

    var undoBtn = btn('撤销', '撤销（Ctrl+Z）', function () { C.undo(); }, 'rb-icon');
    var redoBtn = btn('重做', '重做（Ctrl+Y）', function () { C.redo(); }, 'rb-icon');
    body.appendChild(group('历史', [undoBtn, redoBtn]));

    body.appendChild(group('剪贴板', [
      btn('剪切', '剪切（Ctrl+X）', function () { C.cut(); }),
      btn('复制', '复制（Ctrl+C）', function () { C.copy(); }),
      btn('粘贴', '粘贴（Ctrl+V）。网页读不到系统剪贴板时，会粘内部剪贴板里最后一次复制的内容。',
        function () { C.pasteInternal(); })
    ]));

    var fontName = select(FONT_NAMES, '字体', function (v) { C.setFontName(v); });
    var fontSize = select(FONT_SIZES.map(function (n) {
      return { value: n, label: String(n) };
    }), '字号', function (v) { C.setFontSize(v); });
    var boldBtn = btn('加粗', '加粗（Ctrl+B）', function () { C.toggleBold(); }, 'rb-b');
    var italicBtn = btn('倾斜', '倾斜（Ctrl+I）', function () { C.toggleItalic(); }, 'rb-i');
    var underBtn = btn('下划线', '下划线（Ctrl+U）', function () { C.toggleUnderline(); }, 'rb-u');
    var fontColor = colorInput('字体颜色', function (c) { C.setFontColor(c); });
    var fillColor = colorInput('填充颜色', function (c) { C.setFillColor(c); });
    body.appendChild(group('字体', [
      fontName, fontSize, '-', boldBtn, italicBtn, underBtn, '-',
      labeled('A', fontColor, '字体颜色'),
      labeled('底色', fillColor, '填充颜色')
    ]));

    var alignL = btn('左', '左对齐', function () { C.setAlign('left', null); });
    var alignC = btn('中', '水平居中', function () { C.setAlign('center', null); });
    var alignR = btn('右', '右对齐', function () { C.setAlign('right', null); });
    var alignT = btn('上', '顶端对齐', function () { C.setAlign(null, 'top'); });
    var alignM = btn('中', '垂直居中', function () { C.setAlign(null, 'middle'); });
    var alignB = btn('下', '底端对齐', function () { C.setAlign(null, 'bottom'); });
    var wrapBtn = btn('自动换行', '自动换行', function () { C.toggleWrap(); });
    var mergeBtn = btn('合并居中', '合并居中 / 取消合并', function () {
      if (C.isMerged()) C.unmerge(); else C.mergeCenter();
    });
    body.appendChild(group('对齐方式', [
      alignL, alignC, alignR, '-', alignT, alignM, alignB, '-', wrapBtn, mergeBtn
    ]));

    var fmtSel = select(Mobai.Format.PRESETS.map(function (p) {
      return { value: p.code, label: p.label };
    }), '数字格式（只改显示，判分永远看原始值）', function (v) { C.setFormat(v); });
    body.appendChild(group('数字', [
      fmtSel,
      btn('增加小数位', '增加小数位', function () { C.stepDecimals(1); }),
      btn('减少小数位', '减少小数位', function () { C.stepDecimals(-1); })
    ]));

    body.appendChild(group('单元格', [
      btn('插入行', '在选区上方插入一行', function () { C.insertRows(1); }),
      btn('插入列', '在选区左侧插入一列', function () { C.insertCols(1); }),
      btn('删除行', '删除选中行', function () { C.deleteRows(); }),
      btn('删除列', '删除选中列', function () { C.deleteCols(); }),
      '-',
      btn('清除内容', '清除内容（Delete）', function () { C.clearContents(); }),
      btn('清除格式', '只清格式，内容留着', function () { C.clearFormat(); })
    ]));

    var fillBtn = btn('向下填充', '向下填充（Ctrl+D）', function () { C.fillDown(); });
    var nameBtn = btn('名称框', '跳到名称框，输入 A1 或 A1:C10 回车', function () { C.focusNameBox(); });
    var allBtn = btn('全选', '全选（Ctrl+A）', function () { C.selectAll(); });
    var soon1 = btn('排序和筛选', '第二批再说', function () { sheet.setTip('排序、筛选属于「功能」，等三个反射动作稳了再做。'); }, 'rb-disabled');
    var soon2 = btn('查找', '第二批再说', function () { sheet.setTip('查找替换属于「功能」，等三个反射动作稳了再做。'); }, 'rb-disabled');
    body.appendChild(group('编辑', [fillBtn, nameBtn, allBtn, '-', soon1, soon2]));

    wrap.appendChild(body);

    /* 其它标签的说明面板 */
    var notePanel = el('div', 'rb-note');
    wrap.appendChild(notePanel);

    container.appendChild(wrap);

    function labeled(text, node, title) {
      var box = el('span', 'rb-labeled');
      box.title = title || text;
      box.appendChild(el('span', 'rb-label-text', text));
      box.appendChild(node);
      return box;
    }

    function show(name) {
      current = name;
      TABS.forEach(function (t) { tabButtons[t].classList.toggle('is-active', t === name); });
      var isHome = name === '开始';
      body.style.display = isHome ? 'flex' : 'none';
      notePanel.style.display = isHome ? 'none' : 'block';
      if (!isHome) notePanel.textContent = (TAB_NOTES[name] || '这个练习用不到。');
      sync();
    }

    /** 按当前选区状态刷新按钮的高亮 / 下拉的值 */
    function sync(state) {
      var st = state || sheet.state();
      [['bold', boldBtn], ['italic', italicBtn], ['underline', underBtn], ['wrap', wrapBtn]].forEach(function (pair) {
        pair[1].classList.toggle('is-on', !!st[pair[0]]);
      });
      alignL.classList.toggle('is-on', st.alignH === 'left');
      alignC.classList.toggle('is-on', st.alignH === 'center');
      alignR.classList.toggle('is-on', st.alignH === 'right');
      alignT.classList.toggle('is-on', st.alignV === 'top');
      alignM.classList.toggle('is-on', st.alignV === 'middle');
      alignB.classList.toggle('is-on', st.alignV === 'bottom');
      mergeBtn.textContent = st.merged ? '取消合并' : '合并居中';
      undoBtn.disabled = !st.canUndo;
      redoBtn.disabled = !st.canRedo;
      fontName.value = FONT_NAMES.some(function (f) { return f.value === st.fontName; }) ? st.fontName : '';
      fontSize.value = st.fontSize ? String(st.fontSize) : '';
      var presets = Mobai.Format.PRESETS.map(function (p) { return p.code; });
      fmtSel.value = presets.indexOf(st.format) >= 0 ? st.format : 'General';
    }

    sheet.onState = sync;
    show('开始');

    return { sync: sync, show: show, tab: function () { return current; } };
  }

  Mobai.Ribbon = { mount: mount };
})(typeof window !== 'undefined' ? window : globalThis);