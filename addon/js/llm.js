/*
 * 可选大模型接入（OpenAI 兼容 /chat/completions）。
 * 把自然语言转换成结构化表格操作计划，再经 WpsAIActions.validate 白名单校验。
 * 提示词携带真实表格结构；模型输出做归一化；解析失败或空动作重试一次并给出具体原因。
 */
(function (global) {
  'use strict';

  /* 本地部署的兼容接口通常不需要 API Key，因此只要求启用开关 + 接口地址 */
  function enabled(settings) {
    return !!(settings && settings.llmEnabled && settings.baseUrl);
  }

  function baseUrl(url) { return String(url || '').trim().replace(/\/+$/, ''); }

  function safeSelection(ctx) {
    var s = String((ctx && (ctx.selectionForAI || ctx.selection)) || '').trim();
    return s || '未选择（可用 @selection 指代）';
  }

  /* 真实表格结构：模型只能引用这里列出的工作表名，不许自己编 Sheet1 / Sheet2 */
  function contextLines(ctx) {
    ctx = ctx || {};
    var names = (ctx.sheetNames || []).slice(0, 12);
    var lines = [
      '工作簿：' + (ctx.docName || '未获取'),
      '当前工作表：' + (ctx.sheetName || '未获取'),
      '工作表列表：' + (names.length ? names.join('、') : '未获取'),
      '当前选区：' + safeSelection(ctx),
      '已用区域：' + (ctx.usedRange || '未获取')
    ];
    if (ctx.lastRow) { lines.push('数据末行：第 ' + ctx.lastRow + ' 行'); }
    if (ctx.lastRange) { lines.push('上次操作区域：' + ctx.lastRange); }
    return lines.join('\n');
  }

  function systemPrompt(ctx, strict) {
    var lines = [
      '你是「智能表格助手」，把用户对当前表格的一句自然语言指令转换成可执行的操作计划。只输出一个 JSON 对象，不要解释文字，不要 Markdown 代码块。',
      '输出格式：{"app":"et","summary":"一句话说明将做什么","actions":[{"type":"et.xxx", ...参数}]}',
      '',
      '【当前表格的真实结构】',
      contextLines(ctx),
      '',
      '【硬性规则】',
      '1. actions 必须至少有 1 个动作；只有完全无法理解指令时才允许空数组，并同时置 needClarify=true。',
      '2. 工作表名只能取自上面的工作表列表，不要臆造 Sheet1 / Sheet2；跨表区域写成 工作表名!A1:B10，表名含空格时用单引号。',
      '3. 区域地址、阈值、颜色、结果文字、目标单元格都必须来自用户指令；指令里没有写区域时才用 @selection。',
      '4. 下面的示例只说明输出格式，其中的数字、地址、颜色不得照抄到实际计划里。',
      '5. 只生成下面【可用动作】列出的 type，不要生成文字文档类动作。',
      '6. 占位符：@selection 当前选区、@selectionBottom 选区下方一行、@lastRow 数据末行、@lastRowPlus1 末行下一行。',
      '7. 破坏性动作（et.deleteRows、et.deleteDuplicateRows、et.deleteRowsIf、et.clear）在用户明确说删除 / 去重 / 清空时直接输出对应动作，不要因为风险高就返回空 actions 或 needClarify；删除重复行用 et.deleteDuplicateRows，不要用 et.deleteRows 删掉整个区域。',
      '8. 有条件的操作要用条件动作：内容等于「是」的行标黄用 et.conditionalFill（threshold:"是"，fillScope:"row" 表示标整行），把内容为「是」的行删除用 et.deleteRowsIf；不要用 et.fillColor 给整个区域上色，也不要以「没有该动作」为由返回空 actions。',
      '',
      '【可用动作】',
      global.WpsAIActions.schema(),
      '',
      '【格式示例】',
      '用户：把当前选区中大于 3000 的单元格标黄',
      '输出：{"app":"et","summary":"把当前选区中大于 3000 的单元格标黄","actions":[{"type":"et.conditionalFill","range":"@selection","op":">","threshold":3000,"color":"#FFFF00"}]}',
      '用户：对 A1:A10 求和，结果放到 A11',
      '输出：{"app":"et","summary":"对 A1:A10 求和并写入 A11","actions":[{"type":"et.setFormula","range":"A11","formula":"=SUM(A1:A10)"}]}'
    ];
    if (strict) {
      lines.push('', '【重试要求】上一次回复没有可用动作。这次必须只输出 1 个最贴近用户指令的动作，禁止空 actions，禁止任何解释。');
    }
    lines.push('', '确实无法判断时输出：{"app":"et","summary":"说明缺什么信息","actions":[],"needClarify":true,"clarify":[{"label":"候选说明","prompt":"一条完整可执行指令"}]}');
    return lines.join('\n');
  }

  /* 从代码块 / 解释文字 / 多个对象里找出计划对象 */
  function jsonObjects(s) {
    var out = [], depth = 0, start = -1, inStr = false, esc = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (inStr) {
        if (esc) { esc = false; }
        else if (ch === '\\') { esc = true; }
        else if (ch === '"') { inStr = false; }
        continue;
      }
      if (ch === '"') { inStr = true; continue; }
      if (ch === '{') { if (depth === 0) { start = i; } depth++; }
      else if (ch === '}') {
        depth--;
        if (depth === 0 && start >= 0) { out.push(s.slice(start, i + 1)); start = -1; }
      }
    }
    return out;
  }

  function extractPlanObject(text) {
    var objs = jsonObjects(String(text || ''));
    var parsed = [];
    for (var i = 0; i < objs.length; i++) {
      try { parsed.push(JSON.parse(objs[i])); } catch (e) { /* 跳过无效片段 */ }
    }
    for (var j = 0; j < parsed.length; j++) {
      var o = parsed[j];
      if (!o || typeof o !== 'object') { continue; }
      if (o.plan && typeof o.plan === 'object') { return o.plan; }
      if (o.actions || o.steps || o.operations || o.needClarify || o.need_clarify || o.summary) { return o; }
    }
    var bare = parsed.filter(function (o) { return o && o.type; });
    if (bare.length) { return { app: 'et', summary: '', actions: bare }; }
    return null;
  }

  /* 模型常把类型写成 setFormula / et.set_formula / 带空格，统一映射到白名单 */
  var TYPE_ALIAS = {
    removeduplicates: 'et.deleteDuplicateRows',
    deletedupes: 'et.deleteDuplicateRows',
    deduplicaterows: 'et.deleteDuplicateRows',
    dedupe: 'et.deleteDuplicateRows',
    deleterowsmatching: 'et.deleteRowsIf',
    deleterowsbycondition: 'et.deleteRowsIf',
    conditionaldelete: 'et.deleteRowsIf'
  };
  function canonicalType(type) {
    var meta = global.WpsAIActions.meta;
    var t = String(type || '').trim();
    if (meta[t]) { return t; }
    if (t.slice(0, 3) !== 'et.' && meta['et.' + t]) { return 'et.' + t; }
    var key = t.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (TYPE_ALIAS[key]) { return TYPE_ALIAS[key]; }
    for (var k in meta) {
      if (Object.prototype.hasOwnProperty.call(meta, k) &&
          k.toLowerCase().replace(/[^a-z0-9]/g, '') === key) { return k; }
    }
    return t;
  }

  function normalizePlan(raw) {
    var src = (raw.plan && typeof raw.plan === 'object') ? raw.plan : raw;
    var list = src.actions;
    if (!Array.isArray(list)) {
      if (Array.isArray(src.steps)) { list = src.steps; }
      else if (Array.isArray(src.operations)) { list = src.operations; }
      else if (list && typeof list === 'object') { list = [list]; }
      else { list = []; }
    }
    var actions = [];
    for (var i = 0; i < list.length; i++) {
      var a = list[i] || {};
      var clean = {};
      for (var k in a) { if (Object.prototype.hasOwnProperty.call(a, k)) { clean[k] = a[k]; } }
      clean.type = canonicalType(a.type || a.action || a.operation);
      if (!clean.range && (a.cell || a.address || a.region)) { clean.range = a.cell || a.address || a.region; }
      if (!clean.formula && a.expression) { clean.formula = a.expression; }
      if (!clean.color && a.fill) { clean.color = a.fill; }
      if (!clean.value && a.text && clean.type === 'et.setValue') { clean.value = a.text; }
      actions.push(clean);
    }
    var clarify = src.clarify || raw.clarify;
    return {
      app: 'et',
      summary: String(src.summary || raw.summary || ''),
      actions: actions,
      needClarify: !!(src.needClarify || src.need_clarify || raw.needClarify || raw.need_clarify ||
        (clarify && clarify.length && !actions.length)),
      clarify: clarify,
      intent: src.intent,
      unparsed: src.unparsed
    };
  }

  function request(settings, messages, onResult) {
    var payload = {
      model: settings.model || 'gpt-4o-mini',
      messages: messages,
      temperature: Number(settings.temperature) || 0,
      stream: false
    };
    var xhr = new XMLHttpRequest();
    var finished = false;
    function finish(content, error) {
      if (finished) { return; }
      finished = true;
      onResult(content || '', error || '');
    }
    try {
      xhr.open('POST', baseUrl(settings.baseUrl) + '/chat/completions', true);
      xhr.setRequestHeader('Content-Type', 'application/json');
      if (settings.apiKey) { xhr.setRequestHeader('Authorization', 'Bearer ' + settings.apiKey); }
      xhr.timeout = 45000;
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) { return; }
        if (xhr.status < 200 || xhr.status >= 300) {
          finish('', '模型服务返回 HTTP ' + xhr.status);
          return;
        }
        try {
          var data = JSON.parse(xhr.responseText);
          var choice = data.choices && data.choices[0];
          var content = choice && ((choice.message && choice.message.content) || choice.text) || '';
          finish(String(content), '');
        } catch (e) { finish('', '解析模型响应失败：' + e.message); }
      };
      xhr.onerror = function () { finish('', '无法连接模型服务'); };
      xhr.ontimeout = function () { finish('', '模型服务响应超时'); };
      xhr.send(JSON.stringify(payload));
    } catch (e2) { finish('', '调用模型服务失败：' + e2.message); }
  }

  function parse(text, ctx, settings, done) {
    if (!enabled(settings)) { done(null, '未启用大模型'); return; }
    ctx = ctx || {};
    /* 只带最近的用户指令，且明确标注仅供理解指代：避免历史里的公式 / 表名污染本次计划 */
    var userHistory = (ctx.history || [])
      .filter(function (h) { return h.role === 'user'; })
      .slice(-3)
      .map(function (h) { return String(h.content || '').substring(0, 200); });

    function attempt(strict) {
      var messages = [{ role: 'system', content: systemPrompt(ctx, strict) }];
      for (var i = 0; i < userHistory.length; i++) {
        messages.push({ role: 'user', content: '（历史指令，仅供理解指代，不要照抄）' + userHistory[i] });
      }
      messages.push({ role: 'user', content: text });
      request(settings, messages, function (content, error) {
        if (error) { done(null, error); return; }
        var raw = extractPlanObject(content);
        if (!raw) {
          if (!strict) { attempt(true); return; }
          done(null, '模型未返回有效 JSON');
          return;
        }
        var normalized = normalizePlan(raw);
        var plan = global.WpsAIActions.validate(normalized);
        if (!plan.actions.length) {
          if (!strict && !normalized.needClarify) { attempt(true); return; }
          if (normalized.needClarify) {
            plan.help = true;
            plan.needClarify = true;
            plan.source = 'llm';
            plan.intent = {
              confidence: 'low',
              reasons: ['模型没能确定要做什么：' + String(normalized.summary || '').substring(0, 80)],
              scope: global.WpsAIPlanner ? global.WpsAIPlanner.scopeOf('et', []) : '',
              host: '当前表格', basis: []
            };
            if ((!plan.clarify || !plan.clarify.length) && global.WpsAIPlanner) {
              plan.clarify = global.WpsAIPlanner.clarifyOptions(text, ctx);
            }
            done(plan, '');
            return;
          }
          done(null, '模型未给出可执行动作' +
            (normalized.summary ? '：' + String(normalized.summary).substring(0, 80) : ''));
          return;
        }
        plan.source = 'llm';
        done(plan, '');
      });
    }
    attempt(false);
  }

  global.WpsAILlm = {
    enabled: enabled,
    parse: parse,
    systemPrompt: systemPrompt,
    normalizePlan: normalizePlan,
    extractPlanObject: extractPlanObject
  };
})(window);
