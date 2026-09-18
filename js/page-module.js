/* mobai训练 · 模块页（Excel）
 * 三件事：把 5 阶段的考点清单画出来、把进度条与文案填上、接上导出/导入。
 * 知识卡折叠在每行下面，卡住时点开就行，不跳页。
 */
(function () {
  'use strict';

  var E = Mobai.Exercises;
  var P = Mobai.Progress;
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

  /* ---------- 顶部进度 ---------- */

  function renderProgress() {
    var s = P.stats();
    var total = E.count;
    byId('progress-inner').style.width = Math.round(s.passed / total * 100) + '%';
    var parts = ['已通过 ' + s.passed + ' / ' + total + ' 题'];
    if (s.streakDays) parts.push('连续练习 ' + s.streakDays + ' 天');
    if (s.attempts) parts.push('累计提交 ' + s.attempts + ' 次');
    if (s.pendingChecks) parts.push('还有 ' + s.pendingChecks + ' 类问题没解决');
    byId('progress-note').textContent = parts.join(' · ') + '（记录只存在这台电脑的浏览器里）';
  }

  /* ---------- 考点清单 ---------- */

  function cardBlock(cardId) {
    var box = el('div', 'topic-card');
    var card = cardId ? Cards.get(cardId) : null;
    if (!card) {
      box.appendChild(el('p', null, '这道题的知识卡还没写。'));
      return box;
    }
    box.appendChild(el('h4', null, card.title));
    box.appendChild(el('p', null, card.concept));
    box.appendChild(el('p', null, card.example));
    var ul = el('ul');
    card.pitfalls.forEach(function (t) { ul.appendChild(el('li', null, t)); });
    box.appendChild(ul);
    return box;
  }

  function statusChip(rec) {
    if (rec && rec.passed) return el('span', 'chip is-done', '已通过');
    if (rec && rec.attempts) return el('span', 'chip is-doing', '进行中');
    return el('span', 'chip', '未开始');
  }

  function topicRow(ex) {
    var rec = P.get(ex.id);
    var row = el('div', 'topic-row');

    var name = el('div', 'topic-name', ex.topicLabel);
    name.appendChild(el('span', null, '第 ' + (E.all.indexOf(ex) + 1) + ' 题 · ' + ex.title));
    row.appendChild(name);

    row.appendChild(statusChip(rec));
    row.appendChild(el('span', 'topic-attempts', rec && rec.attempts ? '提交 ' + rec.attempts + ' 次' : ''));

    var actions = el('div', 'nav-buttons');
    var cardBtn = el('button', 'btn btn-ghost', '看知识卡');
    cardBtn.type = 'button';
    var start = el('a', 'btn btn-primary', rec && rec.passed ? '再练一次' : '开始练习');
    start.href = 'exercise.html?id=' + ex.id;
    actions.appendChild(cardBtn);
    actions.appendChild(start);
    row.appendChild(actions);

    var box = cardBlock(ex.card);
    row.appendChild(box);
    cardBtn.addEventListener('click', function () {
      var open = box.classList.toggle('is-open');
      cardBtn.textContent = open ? '收起知识卡' : '看知识卡';
    });

    return row;
  }

  function renderStages() {
    var host = byId('stages');
    host.textContent = '';
    E.byStage().forEach(function (group) {
      var block = el('section', 'stage-block glass');
      block.appendChild(el('h2', null, group.stage + '（' + group.items.length + ' 题）'));
      if (!group.items.length) {
        block.appendChild(el('p', 'stage-empty', '这个阶段的题还没写。'));
        host.appendChild(block);
        return;
      }
      group.items.forEach(function (ex) { block.appendChild(topicRow(ex)); });
      host.appendChild(block);
    });
  }

  /* ---------- 导出 / 导入 ---------- */

  function bindIO() {
    var text = byId('io-text');
    var msg = byId('io-msg');

    byId('btn-export').addEventListener('click', function () {
      text.value = P.exportText();
      text.select();
      msg.textContent = '已生成，全选复制走就行。';
    });

    byId('btn-import').addEventListener('click', function () {
      if (!text.value.trim()) {
        msg.textContent = '先把要导入的 JSON 粘到上面的框里。';
        return;
      }
      var r = P.importText(text.value);
      msg.textContent = r.message;
      if (r.ok) {
        renderProgress();
        renderStages();
      }
    });
  }

  /* ---------- 导航 ---------- */

  function continueUrl() {
    var last = null;
    try {
      last = P.lastPracticeId();
    } catch (ex) {
      last = null;
    }
    return 'exercise.html?id=' + (last || E.first.id);
  }

  function bindNav() {
    byId('btn-continue').addEventListener('click', function () {
      window.location.href = continueUrl();
    });
    byId('btn-home').addEventListener('click', function () {
      window.location.href = 'index.html';
    });
  }

  renderProgress();
  renderStages();
  bindIO();
  bindNav();
})();