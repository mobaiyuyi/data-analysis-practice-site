/* 研数通途 · 首页
 * 只做两件事：把 4 个指标填上，把按钮指到该去的地方。
 */
(function () {
  'use strict';

  var FIRST_EXERCISE = 'excel-01';   // 题库第一题，懒得为了一个字符串把题库加载进来

  function byId(id) {
    return document.getElementById(id);
  }

  function go(url) {
    window.location.href = url;
  }

  function continueUrl() {
    var last = null;
    try {
      last = Mobai.Progress.lastPracticeId();
    } catch (ex) {
      last = null;
    }
    return 'exercise.html?id=' + (last || FIRST_EXERCISE);
  }

  function renderStats() {
    var s;
    try {
      s = Mobai.Progress.stats();
    } catch (ex) {
      s = { passed: 0, streakDays: 0, pendingChecks: 0, attempts: 0 };
    }
    byId('stat-passed').textContent = String(s.passed);
    byId('stat-streak').textContent = String(s.streakDays);
    byId('stat-pending').textContent = String(s.pendingChecks);
    byId('stat-attempts').textContent = String(s.attempts);
  }

  renderStats();

  ['btn-continue', 'btn-continue2'].forEach(function (id) {
    var el = byId(id);
    if (el) el.addEventListener('click', function () { go(continueUrl()); });
  });

  ['btn-tree', 'btn-tree2'].forEach(function (id) {
    var el = byId(id);
    if (el) el.addEventListener('click', function () { go('excel.html'); });
  });

  var start = byId('btn-start');
  if (start) {
    start.addEventListener('click', function () { go('exercise.html?id=' + FIRST_EXERCISE); });
  }
})();