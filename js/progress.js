/* mobai训练 · 进度
 * 只存在本机浏览器里（localStorage），没有后端。
 * localStorage 不可用（隐私模式、file:// 某些环境）时降级到内存，页面仍然能用，
 * 只是刷新后会丢。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});
  var U = Mobai.Utils;

  var KEY = 'mobai.progress.v1';

  function emptyData() {
    return { version: 1, exercises: {}, streak: { lastActiveDate: '', days: 0 } };
  }

  var memory = null;   // localStorage 不可用时的兜底

  function storage() {
    try {
      var s = root.localStorage;
      if (!s) return null;
      var probe = '__mobai_probe__';
      s.setItem(probe, '1');
      s.removeItem(probe);
      return s;
    } catch (ex) {
      return null;
    }
  }

  function load() {
    var s = storage();
    if (!s) return memory || (memory = emptyData());
    try {
      var raw = s.getItem(KEY);
      if (!raw) return emptyData();
      var data = JSON.parse(raw);
      if (!data || typeof data !== 'object' || !data.exercises) return emptyData();
      data.version = 1;
      data.streak = data.streak || { lastActiveDate: '', days: 0 };
      return data;
    } catch (ex) {
      return emptyData();
    }
  }

  function save(data) {
    var s = storage();
    if (!s) { memory = data; return false; }
    try {
      s.setItem(KEY, JSON.stringify(data));
      return true;
    } catch (ex) {
      memory = data;
      return false;
    }
  }

  function get(id) {
    var data = load();
    return data.exercises[id] || null;
  }

  function yesterday() {
    var d = new Date();
    d.setDate(d.getDate() - 1);
    return U.localDateStr(d);
  }

  function bumpStreak(data) {
    var today = U.localDateStr();
    var st = data.streak || (data.streak = { lastActiveDate: '', days: 0 });
    if (st.lastActiveDate === today) return;
    st.days = st.lastActiveDate === yesterday() ? (st.days || 0) + 1 : 1;
    st.lastActiveDate = today;
  }

  /**
   * 记录一次提交。
   * @param {string} id 练习标识
   * @param {boolean} passed 是否全部检查项通过
   * @param {string[]} failedChecks 未通过的检查项类型
   */
  function recordAttempt(id, passed, failedChecks) {
    var data = load();
    var now = new Date();
    var stamp = now.getFullYear() + '-' + U.pad2(now.getMonth() + 1) + '-' + U.pad2(now.getDate()) +
      'T' + U.pad2(now.getHours()) + ':' + U.pad2(now.getMinutes());
    var rec = data.exercises[id] || {
      attempts: 0, passed: false, firstPassedAt: '', lastAttemptAt: '', failedChecks: []
    };
    rec.attempts = (rec.attempts || 0) + 1;
    rec.lastAttemptAt = stamp;
    rec.failedChecks = failedChecks || [];
    if (passed && !rec.passed) {
      rec.passed = true;
      rec.firstPassedAt = stamp;
    }
    data.exercises[id] = rec;
    bumpStreak(data);
    save(data);
    return rec;
  }

  /** 换设备迁移或交给 AI 看时用 */
  function exportText() {
    return JSON.stringify(load(), null, 2);
  }

  function importText(text) {
    var data;
    try {
      data = JSON.parse(String(text));
    } catch (ex) {
      return { ok: false, message: '这段文本不是合法的 JSON，检查一下有没有漏复制或多了内容。' };
    }
    if (!data || typeof data !== 'object' || !data.exercises || typeof data.exercises !== 'object') {
      return { ok: false, message: 'JSON 里没有找到 exercises 字段，看起来不是本站导出的进度。' };
    }
    var ids = Object.keys(data.exercises);
    var normalized = emptyData();
    ids.forEach(function (id) {
      var rec = data.exercises[id] || {};
      normalized.exercises[id] = {
        attempts: rec.attempts || 0,
        passed: !!rec.passed,
        firstPassedAt: rec.firstPassedAt || '',
        lastAttemptAt: rec.lastAttemptAt || '',
        failedChecks: rec.failedChecks || []
      };
    });
    if (data.streak && typeof data.streak === 'object') {
      normalized.streak = {
        lastActiveDate: data.streak.lastActiveDate || '',
        days: data.streak.days || 0
      };
    }
    save(normalized);
    return { ok: true, message: '已导入 ' + ids.length + ' 道题的记录。' };
  }

  /** 首页 4 个指标 */
  function stats() {
    var data = load();
    var passed = 0;
    var attempts = 0;
    var pending = {};
    Object.keys(data.exercises).forEach(function (id) {
      var rec = data.exercises[id];
      if (rec.passed) passed++;
      attempts += rec.attempts || 0;
      (rec.failedChecks || []).forEach(function (t) { pending[t] = true; });
    });
    return {
      passed: passed,
      attempts: attempts,
      pendingChecks: Object.keys(pending).length,
      streakDays: (data.streak && data.streak.days) || 0
    };
  }

  /** 上次动过的练习（首页「继续练习」用） */
  function lastPracticeId() {
    var data = load();
    var best = null;
    var bestAt = '';
    Object.keys(data.exercises).forEach(function (id) {
      var rec = data.exercises[id];
      var at = rec.lastAttemptAt || '';
      if (at > bestAt) { bestAt = at; best = id; }
    });
    return best;
  }

  function reset() {
    save(emptyData());
  }

  Mobai.Progress = {
    KEY: KEY,
    load: load,
    save: save,
    get: get,
    recordAttempt: recordAttempt,
    stats: stats,
    lastPracticeId: lastPracticeId,
    exportText: exportText,
    importText: importText,
    reset: reset
  };
})(typeof window !== 'undefined' ? window : globalThis);