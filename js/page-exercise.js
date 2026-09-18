/* mobai训练 · 练习页
 * 左栏是题面与判分，右栏是表格。表格一改，上一次的判分结论就作废（清空），
 * 免得看着旧的绿字以为已经过了。
 * 手机（<1024px）不挂表格，只留题面、知识卡与进度——编辑要靠键盘，手机给不了。
 */
(function () {
  'use strict';

  var E = Mobai.Exercises;
  var P = Mobai.Progress;
  var G = Mobai.Grading;
  var Cards = Mobai.Cards;

  function byId(id) {
    return document.getElementById(id);
  }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined) n.textContent = text;
    return n;
  }

  /* ---------- 选题 ---------- */

  function queryId() {
    try {
      var m = /[?&]id=([^&]+)/.exec(window.location.search);
      return m ? decodeURIComponent(m[1]) : null;
    } catch (ex) {
      return null;
    }
  }

  var ex = E.get(queryId()) || E.first;
  var index = E.all.indexOf(ex);
  var isMobile = window.matchMedia('(max-width: 1023px)').matches;

  /* ---------- 题面 ---------- */

  function difficultyDots(n) {
    var out = '';
    for (var i = 1; i <= 5; i++) out += i <= n ? '●' : '○';
    return out;
  }

  function renderMeta() {
    var rec = P.get(ex.id);
    var meta = byId('ex-meta');
    meta.textContent =
      '第 ' + (index + 1) + ' 题 / 共 ' + E.count + ' 题 · 考点：' + ex.topicLabel +
      ' · ' + ex.stage + ' · 难度 ' + difficultyDots(ex.difficulty) + ' · 约 ' + ex.minutes + ' 分钟';
    if (rec && rec.passed) meta.appendChild(el('span', 'chip is-done', '已通过'));
    else if (rec && rec.attempts) meta.appendChild(el('span', 'chip is-doing', '进行中'));
    else meta.appendChild(el('span', 'chip', '未开始'));
  }

  function renderHead() {
    document.title = ex.title + ' · mobai训练';
    byId('ex-title').textContent = ex.title;
    renderMeta();
    byId('ex-background').textContent = ex.background;
    byId('ex-goal').textContent = ex.goal;

    var steps = byId('ex-steps');
    steps.textContent = '';
    if (ex.steps && ex.steps.length) {
      ex.steps.forEach(function (s) { steps.appendChild(el('li', null, s)); });
    } else {
      steps.style.display = 'none';
    }
  }

  function renderCard() {
    var card = ex.card ? Cards.get(ex.card) : null;
    if (!card) {
      byId('ex-card').style.display = 'none';
      return;
    }
    byId('ex-card-title').textContent = '知识点卡：' + card.title;
    var body = byId('ex-card-body');
    body.textContent = '';
    body.appendChild(el('p', null, card.concept));
    body.appendChild(el('h4', null, '例子'));
    body.appendChild(el('pre', null, card.example));
    body.appendChild(el('h4', null, '容易踩的坑'));
    var ul = el('ul');
    card.pitfalls.forEach(function (t) { ul.appendChild(el('li', null, t)); });
    body.appendChild(ul);
  }

  /* ---------- 表格 ---------- */

  var wb = new Mobai.Workbook(ex.workbook.cells, ex.workbook);
  var snapshot = wb.snapshot();
  var sheet = null;
  var ribbon = null;

  function mountSheet() {
    var host = byId('sheet-host');
    sheet = Mobai.Sheet.mount(host, {
      workbook: wb,
      // 只有真的改了数据才把上一次的判分结论清掉；改字体、改数字格式不影响对错，留着
      onChange: function (kind) {
        if (kind === 'style' || kind === 'format') return;
        clearResults();
      },
      note: {
        tab: '说明',
        title: '这一题在练什么',
        text: ex.goal + '\n\n' +
          '判分分三类：算对了吗（单元格的值）、用对工具了吗（函数）、能维护吗（引用结构）。' +
          '三类全过才算过，没有权重，所以不用琢磨怎么拿分。\n\n' +
          '灰色格是别人给你的原始数据，不能改；白色格才是你该动的地方。\n\n' +
          '手感：写公式时可以直接用鼠标点单元格把引用点进来（拖一下就是区域，F4 切 $ 的写法）；' +
          '选区右下角的小方块是填充柄，往下拖就按等差 / 日期 / 公式规律填，双击填到相邻数据的末行；' +
          '右键有剪切、插入行列、清除格式。日期格按 ' + ex.workbook.today + ' 算「今天」。'
      }
    });
    sheet.selectAddr(ex.checks && ex.checks.length ? ex.checks[0].target : 'A1');

    ribbon = Mobai.Ribbon.mount(host.querySelector('.ribbon-slot'), sheet);
  }

  /* ---------- 判分 ---------- */

  function clearResults() {
    byId('check-list').textContent = '';
    byId('result').textContent = '';
  }

  function renderResults(results, passed) {
    var list = byId('check-list');
    list.textContent = '';
    results.forEach(function (r, i) {
      var item = el('div', 'check-item ' + (r.passed ? 'is-pass' : 'is-fail'));
      var top = el('div', 'check-top');
      top.appendChild(el('span', 'check-label',
        (i + 1) + '. ' + (G.TYPE_LABEL[r.type] || r.type) + '（' + (r.passed ? '通过' : '没过') + '）'));
      top.appendChild(el('span', 'check-target', r.target));
      item.appendChild(top);
      item.appendChild(el('div', 'check-msg', r.message));
      if (!r.passed && r.hint) item.appendChild(el('div', 'check-hint', r.hint));
      list.appendChild(item);
    });

    var box = byId('result');
    box.textContent = '';
    var banner = el('div', 'result-banner ' + (passed ? 'is-pass' : 'is-fail'));
    if (passed) {
      banner.textContent = '全部通过。这道题要你记住的是：' + ex.goal;
    } else {
      var bad = results.filter(function (r) { return !r.passed; }).length;
      banner.textContent =
        '还有 ' + bad + ' 项没过。别一格一格乱试——红色的说明里写了这一格现在算出来是什么、' +
        '应该是什么，先看它，再回头看知识点卡。';
    }
    box.appendChild(banner);

    var next = E.all[index + 1];
    if (passed && next) {
      var a = el('a', 'btn btn-primary next-link', '下一题：' + next.title);
      a.href = 'exercise.html?id=' + next.id;
      box.appendChild(a);
    }
  }

  function submit() {
    var results = G.run(ex.checks, wb);
    var passed = G.allPassed(results);
    P.recordAttempt(ex.id, passed, G.failedTypes(results));
    renderResults(results, passed);
    renderMeta();
  }

  function reset() {
    wb.restore(snapshot);
    if (sheet) sheet.render();
    clearResults();
  }

  /* ---------- 导航 ---------- */

  function bindNav() {
    var prev = byId('btn-prev');
    var next = byId('btn-next');
    prev.disabled = index <= 0;
    next.disabled = index >= E.count - 1;
    prev.addEventListener('click', function () {
      if (index > 0) window.location.href = 'exercise.html?id=' + E.all[index - 1].id;
    });
    next.addEventListener('click', function () {
      if (index < E.count - 1) window.location.href = 'exercise.html?id=' + E.all[index + 1].id;
    });
    byId('btn-continue').addEventListener('click', function () {
      var last = P.lastPracticeId();
      window.location.href = 'exercise.html?id=' + (last || E.first.id);
    });
  }

  /* ---------- 启动 ---------- */

  renderHead();
  renderCard();
  bindNav();

  if (isMobile) {
    byId('panel-submit').style.display = 'none';
  } else {
    mountSheet();
    byId('btn-submit').addEventListener('click', submit);
    byId('btn-reset').addEventListener('click', reset);
  }
})();