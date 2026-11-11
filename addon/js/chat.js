/*
 * 上下文对话管理：多会话、消息持久化、上下文变量。
 * 存储优先级：localStorage -> wps.PluginStorage -> 内存（file:// 下 localStorage 不可用时的兜底）
 */
(function (global) {
  'use strict';

  var KEY_DATA = 'wpsai.conversations.v1';
  var KEY_SETTINGS = 'wpsai.settings.v1';
  var memory = {};
  var state = null;
  var settings = null;

  function pluginStorage() {
    try {
      if (global.wps && global.wps.PluginStorage) { return global.wps.PluginStorage; }
      if (global.parent && global.parent.wps && global.parent.wps.PluginStorage) { return global.parent.wps.PluginStorage; }
    } catch (e) { /* 跨域或环境不支持 */ }
    return null;
  }

  var storage = {
    get: function (key) {
      try {
        var local = global.localStorage && global.localStorage.getItem(key);
        if (local !== null && local !== undefined) { return local; }
      } catch (e) { /* 忽略 */ }
      var ps = pluginStorage();
      if (ps) {
        try {
          var fromWps = ps.getItem(key);
          if (fromWps) { return String(fromWps); }
        } catch (e2) { /* 忽略 */ }
      }
      return Object.prototype.hasOwnProperty.call(memory, key) ? memory[key] : null;
    },
    set: function (key, value) {
      memory[key] = value;
      try { global.localStorage && global.localStorage.setItem(key, value); } catch (e) { /* 忽略 */ }
      var ps = pluginStorage();
      if (ps) { try { ps.setItem(key, value); } catch (e2) { /* 忽略 */ } }
    },
    remove: function (key) {
      delete memory[key];
      try { global.localStorage && global.localStorage.removeItem(key); } catch (e) { /* 忽略 */ }
    }
  };

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
  }

  function readJSON(key, fallback) {
    var raw = storage.get(key);
    if (!raw) { return fallback; }
    try {
      var value = JSON.parse(raw);
      return value || fallback;
    } catch (e) { return fallback; }
  }

  function defaultSettings() {
    return {
      llmEnabled: false,
      baseUrl: 'https://api.openai.com/v1',
      apiKey: '',
      model: 'gpt-4o-mini',
      temperature: 0.2,
      autoRunReadonly: true,
      contextTurns: 6,
      colorMode: 'auto',
      interactionMode: 'advise'
    };
  }

  function makeSession(title) {
    var now = Date.now();
    return {
      id: uid(),
      title: title || '新会话',
      createdAt: now,
      updatedAt: now,
      context: { lastApp: '', lastRange: '', lastColor: '', lastTopic: '' },
      messages: []
    };
  }

  function normalize(data) {
    if (!data.sessions || !data.sessions.length) {
      var fresh = makeSession('新会话');
      data = { version: 1, activeId: fresh.id, sessions: [fresh] };
    }
    data.version = 1;
    for (var i = 0; i < data.sessions.length; i++) {
      var s = data.sessions[i];
      s.id = s.id || uid();
      s.title = s.title || '新会话';
      s.createdAt = s.createdAt || Date.now();
      s.updatedAt = s.updatedAt || s.createdAt;
      s.context = s.context || { lastApp: '', lastRange: '', lastColor: '', lastTopic: '' };
      s.messages = s.messages || [];
    }
    var exists = false;
    for (var j = 0; j < data.sessions.length; j++) {
      if (data.sessions[j].id === data.activeId) { exists = true; break; }
    }
    if (!exists) { data.activeId = data.sessions[0].id; }
    return data;
  }

  function getState() {
    if (!state) { state = normalize(readJSON(KEY_DATA, null) || {}); }
    return state;
  }

  function persist() {
    storage.set(KEY_DATA, JSON.stringify(getState()));
  }

  function findSession(id) {
    var data = getState();
    for (var i = 0; i < data.sessions.length; i++) {
      if (data.sessions[i].id === id) { return data.sessions[i]; }
    }
    return null;
  }

  function getActiveSession() {
    return findSession(getState().activeId) || getState().sessions[0];
  }

  function createSession(title) {
    var s = makeSession(title);
    getState().sessions.unshift(s);
    getState().activeId = s.id;
    persist();
    return s;
  }

  function setActive(id) {
    if (findSession(id)) { getState().activeId = id; persist(); }
    return getActiveSession();
  }

  function deleteSession(id) {
    var data = getState();
    for (var i = 0; i < data.sessions.length; i++) {
      if (data.sessions[i].id === id) { data.sessions.splice(i, 1); break; }
    }
    if (!data.sessions.length) { data.sessions.push(makeSession('新会话')); }
    if (data.activeId === id) { data.activeId = data.sessions[0].id; }
    persist();
  }

  function renameSession(id, title) {
    var s = findSession(id);
    if (s && title) { s.title = title; s.updatedAt = Date.now(); persist(); }
  }

  function addMessage(sessionId, message) {
    var s = findSession(sessionId) || getActiveSession();
    var msg = {
      id: message.id || uid(),
      role: message.role || 'assistant',
      text: message.text || '',
      time: message.time || Date.now(),
      plan: message.plan || null,
      result: message.result || null,
      advice: message.advice || null,
      status: message.status || 'done',
      source: message.source || ''
    };
    s.messages.push(msg);
    s.updatedAt = Date.now();
    if (msg.role === 'user' && (s.title === '新会话' || !s.title)) {
      var clean = msg.text.replace(/\s+/g, ' ').substring(0, 18);
      s.title = clean || '新会话';
    }
    if (s.messages.length > 200) { s.messages = s.messages.slice(-200); }
    persist();
    return msg;
  }

  function updateMessage(sessionId, messageId, patch) {
    var s = findSession(sessionId) || getActiveSession();
    for (var i = 0; i < s.messages.length; i++) {
      if (s.messages[i].id === messageId) {
        for (var key in patch) {
          if (Object.prototype.hasOwnProperty.call(patch, key)) { s.messages[i][key] = patch[key]; }
        }
        s.updatedAt = Date.now();
        persist();
        return s.messages[i];
      }
    }
    return null;
  }

  function rememberPlan(sessionId, plan) {
    var s = findSession(sessionId) || getActiveSession();
    if (!plan) { return; }
    if (plan.app) { s.context.lastApp = plan.app; }
    if (plan.topic) { s.context.lastTopic = plan.topic; }
    for (var i = 0; i < (plan.actions || []).length; i++) {
      var a = plan.actions[i];
      if (a.range) { s.context.lastRange = a.range; }
      if (a.color) { s.context.lastColor = a.color; }
    }
    persist();
  }

  /* 供规则引擎、大模型与公式顾问共同使用的上下文 */
  function getContext(sessionId, info) {
    var s = findSession(sessionId) || getActiveSession();
    var turns = getSettings().contextTurns || 6;
    var start = Math.max(0, s.messages.length - turns * 2);
    var history = [];
    for (var i = start; i < s.messages.length; i++) {
      var m = s.messages[i];
      var content = m.text || '';
      if (m.plan && m.plan.summary) { content += ' 【操作计划】' + m.plan.summary; }
      history.push({ role: m.role === 'user' ? 'user' : 'assistant', content: content });
    }
    return {
      history: history,
      host: info && info.host,
      mode: info && info.mode,
      selection: info && info.selection,
      selectionForAI: info && info.selectionForAI,
      docName: info && info.docName,
      sheetName: info && info.sheetName,
      usedRange: info && info.usedRange,
      lastRow: info && info.lastRow,
      lastCol: info && info.lastCol,
      lastApp: s.context.lastApp,
      lastRange: s.context.lastRange,
      lastColor: s.context.lastColor,
      lastTopic: s.context.lastTopic
    };
  }

  function getSettings() {
    if (!settings) {
      var saved = readJSON(KEY_SETTINGS, {}) || {};
      var defaults = defaultSettings();
      for (var k in defaults) {
        if (!Object.prototype.hasOwnProperty.call(saved, k)) { saved[k] = defaults[k]; }
      }
      settings = saved;
    }
    return settings;
  }

  function saveSettings(patch) {
    var current = getSettings();
    for (var k in patch) {
      if (Object.prototype.hasOwnProperty.call(patch, k)) { current[k] = patch[k]; }
    }
    settings = current;
    storage.set(KEY_SETTINGS, JSON.stringify(settings));
    return settings;
  }

  function clearAll() {
    state = { version: 1, activeId: '', sessions: [] };
    state = normalize(state);
    persist();
    return state;
  }

  /* 导出会话为 Markdown，便于留痕或复盘 */
  function exportSession(sessionId) {
    var s = findSession(sessionId) || getActiveSession();
    var lines = ['# ' + s.title, '', '> 导出时间：' + new Date().toLocaleString(), ''];
    for (var i = 0; i < s.messages.length; i++) {
      var m = s.messages[i];
      if (m.source === 'welcome') { continue; }
      lines.push('## ' + (m.role === 'user' ? '我' : 'AI 助手') + ' · ' + new Date(m.time).toLocaleString(), '');
      lines.push(m.text || '(空)');
      if (m.advice) {
        lines.push('', '**我理解你要的是**：' + (m.advice.intent || ''));
        if (m.advice.formulas && m.advice.formulas.length) {
          lines.push('', '**公式建议**');
          for (var f = 0; f < m.advice.formulas.length; f++) {
            var fm = m.advice.formulas[f];
            lines.push('- ' + fm.title + '：`' + fm.formula + '`' + (fm.why ? ' — ' + fm.why : ''));
          }
        }
        if (m.advice.gaps && m.advice.gaps.length) {
          var missing = [];
          for (var g = 0; g < m.advice.gaps.length; g++) { missing.push(m.advice.gaps[g].text); }
          lines.push('', '**待补充**：' + missing.join('；'));
        }
      }
      if (m.plan) {
        lines.push('', '**操作计划**：' + (m.plan.summary || ''));
        for (var a = 0; a < (m.plan.actions || []).length; a++) {
          lines.push('- ' + (m.plan.actions[a].type || '') + '：' + JSON.stringify(m.plan.actions[a]));
        }
      }
      if (m.result) {
        lines.push('', '**执行结果**：' + (m.result.ok ? '成功 ' : '部分失败 ') +
          m.result.succeeded + ' 项成功 / ' + m.result.failed + ' 项失败');
        for (var d = 0; d < (m.result.details || []).length; d++) {
          var detail = m.result.details[d];
          lines.push('- ' + (detail.ok ? '成功' : '失败') + ' · ' + (detail.label || '') + '：' + (detail.text || ''));
        }
      }
      lines.push('');
    }
    return lines.join('\n');
  }

  global.WpsAIChat = {
    uid: uid,
    exportSession: exportSession,
    getState: getState,
    persist: persist,
    getActiveSession: getActiveSession,
    findSession: findSession,
    createSession: createSession,
    setActive: setActive,
    deleteSession: deleteSession,
    renameSession: renameSession,
    addMessage: addMessage,
    updateMessage: updateMessage,
    rememberPlan: rememberPlan,
    getContext: getContext,
    getSettings: getSettings,
    saveSettings: saveSettings,
    clearAll: clearAll,
    storage: storage
  };
})(window);
