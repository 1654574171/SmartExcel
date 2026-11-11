/*
 * 侧边栏 UI：公式建议 + 执行操作双模式。
 * 公式建议只给公式文本；执行操作先出计划卡片，用户点「执行」后才写入真实表格。
 * 没有连接真实表格时不提供虚拟工作表，也不伪装写入。
 */
(function (global) {
  'use strict';

  var Chat = global.WpsAIChat;
  var Bridge = global.WpsAIBridge;
  var Advisor = global.WpsAIAdvisor;
  var Actions = global.WpsAIActions;
  var Planner = global.WpsAIPlanner;
  var History = global.WpsAIHistory;
  var Llm = global.WpsAILlm;

  var state = {
    sending: false, renaming: null, clearArmed: false,
    ranges: [], rangeSeq: 1, promptCaret: null,
    mode: 'advise'
  };

  function el(id) { return document.getElementById(id); }

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function toast(message) {
    var node = el('toast');
    node.textContent = message;
    node.className = 'toast';
    global.clearTimeout(toast.timer);
    toast.timer = global.setTimeout(function () { node.className = 'toast hidden'; }, 2600);
  }

  function timeText(ts) {
    var d = new Date(ts || Date.now());
    return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
  }

  function connected() {
    var info = Bridge.info();
    return info.mode === 'live' && info.host === 'et';
  }

  function hintForMode() {
    if (state.sending) { return '正在理解指令…'; }
    return state.mode === 'exec'
      ? 'Enter 发送 · Shift+Enter 换行 · 选区可跨表多次引用 · 先出计划，确认后写入'
      : 'Enter 发送 · Shift+Enter 换行 · 选区可跨表多次引用 · 只给公式建议，不改动表格';
  }

  function updateModeBadges() {
    var badge = el('modeBadge');
    if (state.mode === 'exec') {
      badge.textContent = '执行操作 · 确认后写入';
      badge.className = 'badge mode exec';
      badge.title = '当前会写入表格，点击切回公式建议';
      el('engineBadge').textContent = (Llm && Llm.enabled(Chat.getSettings())) ? '大模型计划' : '规则引擎';
    } else {
      badge.textContent = '仅建议 · 不写入';
      badge.className = 'badge mode';
      badge.title = '当前只给公式，点击切到执行操作';
      el('engineBadge').textContent = '公式顾问';
    }
    el('settingExecMode').checked = state.mode === 'exec';
    el('hintText').textContent = hintForMode();
  }

  function refreshEnv() {
    var info = Bridge.detect(true);
    var live = info.mode === 'live' && info.host === 'et';
    var modeText = live ? '已连接' : '未连接真实表格';
    var name = info.docName || (live ? '当前工作簿' : '—');
    el('envLine').innerHTML = '<span class="dot ' + (live ? 'dot-live' : 'dot-offline') + '"></span>' +
      esc('当前表格 · ' + modeText + (live ? ' · ' + name : ''));
    updateModeBadges();
    el('envDetail').innerHTML =
      '<div>状态：<b>' + esc(modeText) + '</b></div>' +
      '<div>工作簿：<b>' + esc(name) + '</b></div>' +
      '<div>工作表：<b>' + esc(info.sheetName || '未获取') + '</b></div>' +
      '<div>当前选区：<b>' + esc(info.selection || '未选择') + '</b></div>' +
      '<div>数据区：<b>' + esc(info.usedRange || '未获取') + '</b></div>';
    return info;
  }

  function setMode(mode) {
    state.mode = mode === 'exec' ? 'exec' : 'advise';
    Chat.saveSettings({ interactionMode: state.mode });
    updateModeBadges();
    renderMessages();
    toast(state.mode === 'exec'
      ? '已切到执行操作模式：计划需确认后才会写入'
      : '已切回公式建议模式：不会改动表格');
  }

  function renderHeader() {
    var s = Chat.getActiveSession();
    el('sessionTitle').textContent = s.title || '新会话';
  }

  /* ---- 多区域引用：每个引用都记录并展示所属工作表 ---- */

  function rangeToken(id) { return '\u3014' + id + '\u3015'; }

  function findRange(id) {
    for (var i = 0; i < state.ranges.length; i++) {
      if (state.ranges[i].id === id) { return state.ranges[i]; }
    }
    return null;
  }

  function renderRangeBar() {
    var bar = el('rangeBar');
    if (!bar) { return; }
    if (!state.ranges.length) {
      bar.innerHTML = '';
      bar.className = 'range-bar hidden';
      return;
    }
    bar.className = 'range-bar';
    var names = [];
    try { names = Bridge.et.sheetNames ? Bridge.et.sheetNames() : []; } catch (e) { names = []; }
    var html = '<div class="range-bar-head"><span>已引用区域（点序号回表格定位，表名 / 地址都可直接改）</span>' +
      '<button class="range-clear" data-act-range="clear">清空</button></div>' +
      '<datalist id="sheetNameList">';
    for (var n = 0; n < names.length; n++) {
      html += '<option value="' + esc(names[n]) + '"></option>';
    }
    html += '</datalist><div class="range-chips">';
    for (var i = 0; i < state.ranges.length; i++) {
      var r = state.ranges[i];
      html += '<span class="range-chip">' +
        '<button class="range-locate" data-act-range="locate" data-range-id="' + r.id +
        '" title="切到「' + esc(r.sheet || '当前工作表') + '」并选中 ' + esc(r.addr) + '">' + rangeToken(r.id) + '</button>' +
        '<input class="range-sheet" data-range-id="' + r.id + '" list="sheetNameList" value="' + esc(r.sheet || '') +
        '" title="该区域所属工作表，留空表示当前工作表">' +
        '<span class="range-bang">!</span>' +
        '<input class="range-edit" data-range-id="' + r.id + '" value="' + esc(r.addr) +
        '" title="区域地址，也可写成 表名!地址；改完按 Enter 或点别处">' +
        '<button class="range-del" data-act-range="del" data-range-id="' + r.id + '" title="移除该引用">×</button>' +
        '</span>';
    }
    bar.innerHTML = html + '</div>';
  }

  function insertAtCursor(text) {
    var box = el('prompt');
    var start = box.value.length, end = start;
    if (state.promptCaret !== null && state.promptCaret !== undefined) {
      start = end = state.promptCaret;
    } else {
      try { start = box.selectionStart; end = box.selectionEnd; } catch (e) { /* 兜底为末尾 */ }
    }
    var value = box.value;
    box.value = value.substring(0, start) + text + value.substring(end);
    var pos = start + text.length;
    box.focus();
    try { box.setSelectionRange(pos, pos); } catch (e2) { /* 忽略 */ }
    state.promptCaret = null;
  }

  function addRangeReference() {
    var ref = { addr: '', sheet: '', book: '' };
    try {
      ref = Bridge.et.selectionRef ? Bridge.et.selectionRef() :
        { addr: Bridge.et.selectionAddress(), sheet: '', book: '' };
    } catch (e) { ref = { addr: '', sheet: '', book: '' }; }
    if (!ref.addr) {
      toast(connected() ? '请先在左侧表格里选中一个区域，再点「使用选区」'
        : '未连接真实表格，无法读取选区；可以直接手写「表名!区域」');
      return;
    }
    var id = state.rangeSeq++;
    state.ranges.push({ id: id, addr: ref.addr, sheet: ref.sheet || '', book: ref.book || '' });
    insertAtCursor(rangeToken(id));
    renderRangeBar();
    toast('已引用 ' + rangeToken(id) + ' = ' + (ref.sheet ? ref.sheet + '!' : '') + ref.addr +
      '；可继续选别的区域再点一次');
  }

  /* ---- 会话与消息 ---- */

  function renderSessions() {
    var data = Chat.getState();
    var html = '';
    for (var i = 0; i < data.sessions.length; i++) {
      var s = data.sessions[i];
      var active = s.id === data.activeId;
      html += '<div class="session-item' + (active ? ' active' : '') + '" data-sid="' + esc(s.id) + '">' +
        '<div class="si-main">' +
        (state.renaming === s.id
          ? '<input class="rename-input" data-rename-input="' + esc(s.id) + '" value="' + esc(s.title) + '">'
          : '<div class="si-title">' + esc(s.title) + '</div>') +
        '<div class="si-meta">' + s.messages.length + ' 条消息 · ' + timeText(s.updatedAt) + '</div>' +
        '</div>' +
        '<button class="si-rename" data-rename="' + esc(s.id) + '" title="重命名">✎</button>' +
        '<button class="si-del" data-del="' + esc(s.id) + '" title="删除">×</button>' +
        '</div>';
    }
    el('sessionList').innerHTML = html || '<div style="padding:12px;color:var(--text-3);font-size:11px">暂无会话</div>';
  }

  function copyText(text) {
    var area = document.createElement('textarea');
    area.value = text;
    area.style.position = 'fixed';
    area.style.left = '-9999px';
    document.body.appendChild(area);
    area.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(area);
    return ok;
  }

  function adviceHTML(msg) {
    var a = msg.advice;
    if (!a) { return ''; }
    var tag = a.confidence === 'high' ? '<span class="conf-tag high">信息齐全</span>'
      : (a.confidence === 'unclear' ? '<span class="conf-tag low">需要补充</span>' : '<span class="conf-tag low">还差一点信息</span>');
    var html = '<div class="advice-card"><div class="advice-head"><span class="advice-kicker">我理解你要的是</span>' + tag + '</div>' +
      '<div class="advice-intent">' + esc(a.intent) + '</div>';
    if (a.points && a.points.length) {
      html += '<div class="advice-points">';
      for (var i = 0; i < a.points.length; i++) {
        html += '<div class="advice-point"><span>' + esc(a.points[i].label) + '</span><b>' + esc(a.points[i].value) + '</b></div>';
      }
      html += '</div>';
    }
    if (a.gaps && a.gaps.length) {
      html += '<div class="advice-gaps"><div class="advice-sub">还差这些信息，确认一下：</div><ul>';
      for (var g = 0; g < a.gaps.length; g++) { html += '<li>' + esc(a.gaps[g].text) + '</li>'; }
      html += '</ul></div>';
    }
    if (a.questions && a.questions.length) {
      html += '<div class="advice-ask"><div class="advice-sub">点一下补充，我会重新整理：</div><div class="clarify-chips">';
      for (var q = 0; q < a.questions.length; q++) {
        var item = a.questions[q];
        var chipAct = item.fill ? 'fill-input' : 'ask';
        html += '<button class="clarify-chip" data-act="' + chipAct + '" data-prompt="' + esc(item.prompt) +
          '" title="' + esc(item.prompt) + '">' + esc(item.label) + '</button>';
      }
      html += '<button class="clarify-chip" data-act="focus-input">我自己补充</button></div></div>';
    }
    if (a.formulas && a.formulas.length) {
      html += '<div class="advice-formulas">';
      for (var f = 0; f < a.formulas.length; f++) {
        var fm = a.formulas[f];
        html += '<div class="formula-item"><div class="formula-title">' + esc(fm.title) + '</div>' +
          '<div class="formula-row"><code class="formula-code">' + esc(fm.formula) + '</code>' +
          '<button class="ghost-btn copy-btn" data-act="copy-formula" data-formula="' + esc(fm.formula) + '">复制</button></div>' +
          '<div class="formula-why">' + esc(fm.why) + '</div>' +
          (fm.tip ? '<div class="formula-tip">' + esc(fm.tip) + '</div>' : '') + '</div>';
      }
      html += '</div>';
    }
    if (a.notes && a.notes.length) {
      html += '<div class="advice-notes">';
      for (var n = 0; n < a.notes.length; n++) { html += '<div>· ' + esc(a.notes[n]) + '</div>'; }
      html += '</div>';
    }
    return html + '</div>';
  }

  function planHTML(m) {
    var p = m.plan;
    if (!p) { return ''; }
    var html = '<div class="plan-card"><div class="plan-head"><span class="plan-icon">计</span>' +
      '<div><strong>' + esc(p.summary || '操作计划') + '</strong>' +
      '<small>共 ' + p.actions.length + ' 步 · ' + (p.source === 'llm' ? '大模型生成' : '规则引擎生成') +
      ' · 需你确认后执行</small></div></div>';
    if (p.engineNote) {
      html += '<div class="plan-note">' + esc(p.engineNote) + '</div>';
    }

    if (p.intent) {
      var low = p.intent.confidence === 'low';
      html += '<div class="intent-box' + (low ? ' low' : '') + '"><div class="intent-main">' +
        '<span class="intent-key">我理解为</span><span>当前表格</span>' +
        '<span class="intent-scope">影响范围：<b>' + esc(p.intent.scope || '当前选区') + '</b></span>' +
        '<span class="conf-tag ' + (low ? 'low' : 'high') + '">' + (low ? '需要确认' : '信息齐全') + '</span></div>';
      if (low && p.intent.reasons && p.intent.reasons.length) {
        html += '<div class="intent-why">' + esc(p.intent.reasons.join('；')) + '</div>';
      }
      html += '</div>';
    }

    if (p.actions.length) {
      html += '<ul class="plan-steps">';
      for (var i = 0; i < p.actions.length; i++) {
        var a = p.actions[i];
        var risk = Actions.risk(a);
        var tagCls = risk === 'read' ? 'tag-read' : (risk === 'danger' ? 'tag-danger' : 'tag-write');
        var tagText = risk === 'read' ? '只读' : (risk === 'danger' ? '高风险' : '写入');
        html += '<li class="plan-step"><span class="step-index">' + (i + 1) + '</span><div class="step-body">' +
          '<div class="step-title">' + esc(Actions.label(a)) + '<span class="tag ' + tagCls + '">' + tagText + '</span></div>' +
          '<div class="step-desc">' + esc(Actions.describe(a)) + '</div></div></li>';
      }
      html += '</ul>';
    }

    if (p.warnings && p.warnings.length) {
      html += '<div class="plan-warn">' + esc(p.warnings.join('；')) + '</div>';
    }
    if ((p.help || !p.actions.length) && p.clarify && p.clarify.length) {
      html += '<div class="clarify-card"><div class="clarify-head">没确定你要做什么，点一条接近的指令重新生成计划：</div><div class="clarify-chips">';
      for (var c = 0; c < p.clarify.length; c++) {
        html += '<button class="clarify-chip" data-act="ask" data-prompt="' + esc(p.clarify[c].prompt) +
          '" title="' + esc(p.clarify[c].prompt) + '">' + esc(p.clarify[c].label) + '</button>';
      }
      html += '</div></div>';
    }

    html += '<div class="plan-actions">' +
      (p.actions.length
        ? '<button class="primary-btn" data-act="run-plan" data-mid="' + esc(m.id) + '">执行这 ' + p.actions.length + ' 步</button>'
        : '') +
      '<span class="note-mini">' + (connected() ? '执行前会采集快照，可撤回' : '未连接真实表格，暂时无法执行') + '</span></div>';
    return html + '</div>';
  }

  function resultHTML(m) {
    var r = m.result;
    if (!r) { return ''; }
    var html = '<div class="result ' + (r.ok ? 'ok' : 'fail') + '"><div class="result-head">' +
      (r.ok ? '✓ 执行完成' : '! 部分步骤失败') + ' · ' + r.succeeded + ' 成功 / ' + r.failed + ' 失败</div><ul>';
    for (var i = 0; i < r.details.length; i++) {
      var d = r.details[i];
      html += '<li class="' + (d.ok ? 'ok' : 'bad') + '"><span class="mark">' + (d.ok ? '✓' : '×') + '</span>' +
        '<span class="det">' + esc((d.label || '') + '：' + (d.text || '')) + '</span></li>';
    }
    html += '</ul>';
    if (m.txId) {
      var tx = History.find(m.txId);
      html += '<div class="result-actions">' +
        '<button class="ghost-btn undo-btn" data-act="undo-tx" data-mid="' + esc(m.id) + '">撤回本次操作</button>' +
        '<button class="ghost-btn undo-btn" data-act="redo-tx" data-mid="' + esc(m.id) + '">重做</button>' +
        '<span class="note-mini">' + esc(tx && !tx.undoable ? tx.reason : '快照仅在当前面板有效') + '</span></div>';
    }
    return html + '</div>';
  }

  function welcomeHTML() {
    if (state.mode === 'exec') {
      return '<strong>你好，我是智能表格助手，现在是「执行操作」模式。</strong>说出你想做的事，我会先给出操作计划，你确认后才写入当前工作簿：' +
        '<ul class="intro-list">' +
        '<li>例如「对当前选区求和，结果放在选区下方」「把大于 90 的单元格标黄」「表头加粗居中」「统计当前选区」</li>' +
        '<li>每步都会写明操作对象与区域，含删除 / 清空等高风险动作会单独标注</li>' +
        '<li>执行前自动采集区域快照，可在结果卡片或设置里的「操作历史」撤回 / 重做</li>' +
        '<li>在左侧表格选中区域点「使用选区」，可跨工作表一直加多个引用，点引用块切表定位</li>' +
        '<li>没有连接真实表格时仍能查看计划，但不会写入任何内容</li>' +
        '</ul>';
    }
    return '<strong>你好，我是智能表格助手，现在是「公式建议」模式。</strong>告诉我你想算什么，我会先重述你的意图，再给出可直接复制的公式：' +
      '<ul class="intro-list">' +
      '<li>例如「把 C2:C20 求和放到 C21」「统计大于 90 的有几个」「算 F 列占比」「删除 A2:A7 重复行」</li>' +
      '<li>在左侧表格选中区域，点「使用选区」可跨工作表一直加多个引用；每个引用都带工作表名，点引用块能切表定位</li>' +
      '<li>「在此处 Sheet2!C2:C6 中计算 C2:C6 大于 3500 且 D2:D6 …」这类跨表多条件判断也能给出 IF 公式</li>' +
      '<li>描述不清楚时，我会列出缺的信息，点一下就能补充</li>' +
      '<li>这个模式只给公式建议，<b>不会改动任何单元格</b>；想让我直接写入，点顶部模式徽章切到「执行操作」</li>' +
      '</ul>';
  }

  function messageHTML(m) {
    if (m.role === 'user') {
      return '<div class="msg msg-user"><div class="bubble">' + esc(m.text) + '</div></div>';
    }
    var body = '';
    if (m.status === 'thinking') {
      body = '<div class="typing"><i></i><i></i><i></i></div>' +
        '<div style="font-size:11px;color:var(--text-3);margin-top:4px">正在理解指令…</div>';
    } else if (m.source === 'welcome') {
      body = welcomeHTML();
    } else {
      body = '<div class="msg-text">' + esc(m.text) + '</div>';
    }
    if (m.advice) { body += adviceHTML(m); }
    if (m.plan) { body += planHTML(m); }
    if (m.result) { body += resultHTML(m); }
    return '<div class="msg msg-ai"><div class="avatar">AI</div><div class="bubble">' + body +
      '<div class="msg-time">' + timeText(m.time) + '</div></div></div>';
  }

  function renderMessages() {
    var session = Chat.getActiveSession();
    var html = '';
    for (var i = 0; i < session.messages.length; i++) { html += messageHTML(session.messages[i]); }
    var box = el('messages');
    var atBottom = box.scrollHeight - box.scrollTop - box.clientHeight < 40;
    var keepTop = box.scrollTop;
    box.innerHTML = html;
    box.scrollTop = atBottom ? box.scrollHeight : keepTop;
  }

  function renderHistory() {
    var box = el('historyList');
    if (!box) { return; }
    var list = History.list();
    if (!list.length) { box.innerHTML = '<div class="history-empty">暂无执行记录</div>'; return; }
    var html = '';
    for (var i = 0; i < list.length; i++) {
      var tx = list[i];
      var undone = tx.status === 'undone';
      html += '<div class="history-item"><div class="hi-main">' +
        '<div class="hi-title"><span class="hi-status ' + (undone ? 'undone' : 'done') + '">' +
        (undone ? '已撤回' : '已执行') + '</span>' + esc(tx.summary) + '</div>' +
        '<div class="hi-meta">' + timeText(tx.time) + ' · ' + tx.actionCount + ' 步 · ' +
        esc(tx.docName || '当前工作簿') + '</div></div>' +
        (tx.undoable
          ? '<button class="ghost-btn" data-act="hist-toggle" data-tx="' + esc(tx.id) + '">' + (undone ? '重做' : '撤回') + '</button>'
          : '<span class="hi-lock">需表格撤销</span>') +
        '</div>';
    }
    box.innerHTML = html;
  }

  function renderAll() {
    renderHeader();
    renderSessions();
    renderMessages();
  }

  function ensureWelcome() {
    var s = Chat.getActiveSession();
    if (!s.messages.length) {
      Chat.addMessage(s.id, { role: 'assistant', text: '', source: 'welcome', status: 'done' });
    }
  }

  function setSending(flag) {
    state.sending = flag;
    el('btnSend').disabled = flag;
    el('hintText').textContent = hintForMode();
  }

  function findMessage(id) {
    var s = Chat.getActiveSession();
    for (var i = 0; i < s.messages.length; i++) {
      if (s.messages[i].id === id) { return s.messages[i]; }
    }
    return null;
  }

  function send(text) {
    if (state.sending) { return; }
    text = String(text || '').trim();
    if (text && Advisor && Advisor.expandRangeTokens) {
      var baseSheet = '';
      try {
        baseSheet = Bridge.et.selectionRef ? Bridge.et.selectionRef().sheet : (Bridge.info().sheetName || '');
      } catch (e) { baseSheet = ''; }
      /* 展开后每个引用都带「工作表名!」，绝不省略 sheet 属性 */
      text = Advisor.expandRangeTokens(text, state.ranges, baseSheet);
    }
    if (!text) { return; }
    refreshEnv();
    var session = Chat.getActiveSession();
    Chat.addMessage(session.id, { role: 'user', text: text });
    var thinking = Chat.addMessage(session.id, { role: 'assistant', text: '', status: 'thinking' });
    renderAll();
    setSending(true);

    var ctx = Chat.getContext(session.id, Bridge.info());
    /* 把真实工作表清单带给计划生成器与大模型，避免它臆造表名 */
    try { ctx.sheetNames = Bridge.et.sheetNames(); } catch (e) { ctx.sheetNames = []; }

    if (state.mode === 'exec') {
      if (!Planner || !Actions) {
        setSending(false);
        toast('执行模块未加载，请重新安装插件');
        return;
      }
      Planner.parse(text, ctx, function (raw) {
        var plan = Actions.validate(raw || {});
        plan.source = (raw && raw.source) || 'rule';
        plan.engineNote = (raw && raw.engineNote) || '';
        plan.warnings = ((raw && raw.warnings) || []).concat(plan.warnings || []);
        el('engineBadge').textContent = plan.source === 'llm' ? '大模型计划' : '规则引擎';
        var lead = (plan.help || !plan.actions.length)
          ? '我没确定你要做什么，先确认一下再生成计划。'
          : '我把你的指令整理成下面这份操作计划，确认后才会写入表格。';
        Chat.updateMessage(session.id, thinking.id, {
          text: lead, status: 'done', source: 'plan', plan: plan
        });
        Chat.rememberPlan(session.id, plan);
        setSending(false);
        renderAll();
      });
      return;
    }

    if (!Advisor || !Advisor.advise) {
      setSending(false);
      toast('公式顾问模块未加载，请重新安装插件');
      return;
    }
    global.setTimeout(function () {
      var advice = Advisor.advise(text, ctx);
      var lead = advice.kind === 'unclear'
        ? '我没确定你要算什么，先问清楚再给公式。'
        : '我把你的需求整理成下面这样，并给出可直接复制的公式；没有改动任何单元格。';
      Chat.updateMessage(session.id, thinking.id, {
        text: lead, status: 'done', source: 'advice', advice: advice
      });
      setSending(false);
      renderAll();
    }, 60);
  }

  function runPlan(mid) {
    var m = findMessage(mid);
    if (!m || !m.plan || m.result) { return; }
    var problem = Actions.hostCheck();
    if (problem) { toast(problem); return; }
    var session = Chat.getActiveSession();
    var before = History.capture(m.plan);
    var result = Actions.execute(m.plan);
    var tx = History.commit(m.plan, result, before);
    Chat.updateMessage(session.id, m.id, { result: result, txId: tx ? tx.id : '' });
    renderAll();
    renderHistory();
    toast(result.ok ? '执行完成：' + result.succeeded + ' 步成功'
      : '执行结束：' + result.failed + ' 步失败，请查看结果');
  }

  function toggleTx(mid, redo) {
    var m = findMessage(mid);
    if (!m || !m.txId) { toast('该操作没有可撤回的快照'); return; }
    var res = redo ? History.redo(m.txId) : History.undo(m.txId);
    toast(res.message);
    renderHistory();
  }

  function copySession() {
    var md = Chat.exportSession(Chat.getActiveSession().id);
    toast(copyText(md) ? '会话已复制为 Markdown' : '复制受限，请手动选中内容');
  }

  function fillSettings() {
    var s = Chat.getSettings();
    el('settingExecMode').checked = state.mode === 'exec';
    el('settingLlmEnabled').checked = !!s.llmEnabled;
    el('settingBaseUrl').value = s.baseUrl || '';
    el('settingApiKey').value = s.apiKey || '';
    el('settingModel').value = s.model || '';
  }

  function findAttr(node, attr) {
    while (node && node !== document.body) {
      if (node.getAttribute && node.getAttribute(attr) !== null) { return node; }
      node = node.parentNode;
    }
    return null;
  }

  function bind() {
    el('btnSend').addEventListener('click', function () {
      var value = el('prompt').value;
      el('prompt').value = '';
      send(value);
    });
    el('prompt').addEventListener('keydown', function (e) {
      if (e.keyCode === 13 && !e.shiftKey) {
        e.preventDefault();
        var value = el('prompt').value;
        el('prompt').value = '';
        send(value);
      }
    });
    el('prompt').addEventListener('input', function () {
      this.style.height = 'auto';
      this.style.height = Math.min(130, this.scrollHeight) + 'px';
    });

    el('btnUseSelection').addEventListener('mousedown', function () {
      var box = el('prompt');
      try { state.promptCaret = box.selectionStart; } catch (e) { state.promptCaret = box.value.length; }
    });
    el('btnUseSelection').addEventListener('click', addRangeReference);

    el('rangeBar').addEventListener('click', function (e) {
      var node = findAttr(e.target, 'data-act-range');
      if (!node) { return; }
      var act = node.getAttribute('data-act-range');
      if (act === 'clear') {
        state.ranges = [];
        renderRangeBar();
        toast('已清空区域引用');
        return;
      }
      var id = parseInt(node.getAttribute('data-range-id'), 10);
      var r = findRange(id);
      if (!r) { return; }
      if (act === 'del') {
        for (var i = 0; i < state.ranges.length; i++) {
          if (state.ranges[i].id === id) { state.ranges.splice(i, 1); break; }
        }
        renderRangeBar();
        return;
      }
      if (act === 'locate') {
        /* 只切表并选中，不重建任何会话 DOM，因此不会闪屏 */
        var res = Bridge.et.selectRange(r.addr, r.sheet);
        toast(res.message);
      }
    });
    el('rangeBar').addEventListener('input', function (e) {
      if (!e.target.classList) { return; }
      var isSheet = e.target.classList.contains('range-sheet');
      var isAddr = e.target.classList.contains('range-edit');
      if (!isSheet && !isAddr) { return; }
      var r = findRange(parseInt(e.target.getAttribute('data-range-id'), 10));
      if (!r) { return; }
      var value = String(e.target.value || '').trim();
      if (isSheet) { r.sheet = value.replace(/^'|'$/g, ''); return; }
      var bang = value.indexOf('!');
      if (bang > 0) {
        r.sheet = value.slice(0, bang).replace(/^'|'$/g, '');
        r.addr = value.slice(bang + 1).toUpperCase().replace(/\$/g, '');
      } else {
        r.addr = value.toUpperCase().replace(/\$/g, '');
      }
    });
    el('rangeBar').addEventListener('keydown', function (e) {
      if (e.target.classList && (e.target.classList.contains('range-edit') ||
          e.target.classList.contains('range-sheet')) && e.keyCode === 13) {
        e.preventDefault();
        e.target.blur();
      }
    });
    el('rangeBar').addEventListener('change', function (e) {
      if (!e.target.classList) { return; }
      var isSheet = e.target.classList.contains('range-sheet');
      var isAddr = e.target.classList.contains('range-edit');
      if (!isSheet && !isAddr) { return; }
      var r = findRange(parseInt(e.target.getAttribute('data-range-id'), 10));
      if (!r) { return; }
      var chip = e.target.parentNode;
      if (isAddr) { e.target.value = r.addr; }
      if (isSheet) { e.target.value = r.sheet || ''; }
      var sheetInput = chip && chip.querySelector('.range-sheet');
      if (sheetInput) { sheetInput.value = r.sheet || ''; }
      var locate = chip && chip.querySelector('.range-locate');
      if (locate) { locate.title = '切到「' + (r.sheet || '当前工作表') + '」并选中 ' + r.addr; }
      toast('引用已更新');
    });

    el('btnNew').addEventListener('click', function () {
      Chat.createSession('新会话');
      ensureWelcome();
      renderAll();
      el('prompt').focus();
    });
    el('btnSessions').addEventListener('click', function () {
      el('sessionDrawer').classList.toggle('hidden');
      renderSessions();
    });
    el('btnCloseDrawer').addEventListener('click', function () { el('sessionDrawer').classList.add('hidden'); });
    el('btnExportSession').addEventListener('click', copySession);
    el('sessionList').addEventListener('click', function (e) {
      var del = findAttr(e.target, 'data-del');
      if (del) {
        e.stopPropagation();
        Chat.deleteSession(del.getAttribute('data-del'));
        renderAll();
        return;
      }
      var rename = findAttr(e.target, 'data-rename');
      if (rename) {
        e.stopPropagation();
        state.renaming = rename.getAttribute('data-rename');
        renderSessions();
        var input = el('sessionList').querySelector('[data-rename-input]');
        if (input) { input.focus(); input.select(); }
        return;
      }
      var item = findAttr(e.target, 'data-sid');
      if (item) {
        Chat.setActive(item.getAttribute('data-sid'));
        state.renaming = null;
        el('sessionDrawer').classList.add('hidden');
        renderAll();
      }
    });
    el('sessionList').addEventListener('keydown', function (e) {
      if (!e.target.getAttribute || !e.target.getAttribute('data-rename-input')) { return; }
      if (e.keyCode === 13) {
        Chat.renameSession(state.renaming, e.target.value.trim() || '新会话');
        state.renaming = null;
        renderAll();
      } else if (e.keyCode === 27) {
        state.renaming = null;
        renderSessions();
      }
    });
    el('btnClearSessions').addEventListener('click', function () {
      if (!state.clearArmed) {
        state.clearArmed = true;
        this.textContent = '再次点击，确认清空';
        var self = this;
        global.setTimeout(function () { state.clearArmed = false; self.textContent = '清空所有会话'; }, 4000);
        return;
      }
      state.clearArmed = false;
      this.textContent = '清空所有会话';
      Chat.clearAll();
      ensureWelcome();
      renderAll();
      toast('会话已清空');
    });

    el('messages').addEventListener('click', function (e) {
      var node = findAttr(e.target, 'data-act');
      if (!node) { return; }
      var act = node.getAttribute('data-act');
      if (act === 'copy-formula') {
        toast(copyText(node.getAttribute('data-formula')) ? '公式已复制，粘贴到目标单元格即可' : '复制受限，请手动选中公式');
      } else if (act === 'ask') {
        var more = node.getAttribute('data-prompt');
        if (more) { send(more); } else { el('prompt').focus(); }
      } else if (act === 'focus-input') {
        el('prompt').focus();
        toast('在输入框里补充说明后按 Enter');
      } else if (act === 'fill-input') {
        el('prompt').value = node.getAttribute('data-prompt') || '';
        el('prompt').focus();
        toast('已把指令填进输入框，补上缺少的信息后按 Enter');
      } else if (act === 'run-plan') {
        runPlan(node.getAttribute('data-mid'));
      } else if (act === 'undo-tx') {
        toggleTx(node.getAttribute('data-mid'), false);
      } else if (act === 'redo-tx') {
        toggleTx(node.getAttribute('data-mid'), true);
      }
    });

    el('modeBadge').addEventListener('click', function () {
      setMode(state.mode === 'exec' ? 'advise' : 'exec');
    });
    el('btnHistory').addEventListener('click', function () {
      renderHistory();
      el('settingsModal').classList.remove('hidden');
    });
    el('btnSettings').addEventListener('click', function () {
      refreshEnv();
      fillSettings();
      renderHistory();
      el('settingsModal').classList.remove('hidden');
    });
    el('btnCloseSettings').addEventListener('click', function () { el('settingsModal').classList.add('hidden'); });
    el('settingsModal').addEventListener('click', function (e) {
      if (e.target === this) { this.classList.add('hidden'); }
    });
    el('btnRefreshEnv').addEventListener('click', function () { refreshEnv(); toast('已重新检测表格环境'); });

    el('settingExecMode').addEventListener('change', function () {
      setMode(this.checked ? 'exec' : 'advise');
    });
    el('settingLlmEnabled').addEventListener('change', function () {
      Chat.saveSettings({ llmEnabled: !!this.checked });
      updateModeBadges();
    });
    el('settingBaseUrl').addEventListener('change', function () { Chat.saveSettings({ baseUrl: this.value.trim() }); });
    el('settingApiKey').addEventListener('change', function () { Chat.saveSettings({ apiKey: this.value.trim() }); });
    el('settingModel').addEventListener('change', function () { Chat.saveSettings({ model: this.value.trim() }); });
    el('btnUndoLast').addEventListener('click', function () {
      toast(History.undoLast().message);
      renderHistory();
    });
    el('btnRedoLast').addEventListener('click', function () {
      toast(History.redoLast().message);
      renderHistory();
    });
    el('historyList').addEventListener('click', function (e) {
      var node = findAttr(e.target, 'data-act');
      if (!node || node.getAttribute('data-act') !== 'hist-toggle') { return; }
      var id = node.getAttribute('data-tx');
      var tx = History.find(id);
      var res = tx && tx.status === 'undone' ? History.redo(id) : History.undo(id);
      toast(res.message);
      renderHistory();
    });
  }

  function init() {
    state.mode = Chat.getSettings().interactionMode === 'exec' ? 'exec' : 'advise';
    refreshEnv();
    fillSettings();
    ensureWelcome();
    renderAll();
    renderRangeBar();
    renderHistory();
    bind();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})(window);
