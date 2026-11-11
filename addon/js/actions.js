/*
 * 操作计划的定义、校验与执行。
 * 计划 = { app, summary, actions: [{ type, ...params }] }
 * 所有写操作在 UI 中先展示计划，用户确认后才执行。
 */
(function (global) {
  'use strict';

  var Bridge = global.WpsAIBridge;

  var META = {
    'et.summary': { app: 'et', risk: 'read', label: '数据摘要' },
    'et.readRange': { app: 'et', risk: 'read', label: '读取区域' },
    'et.setFormula': { app: 'et', risk: 'write', label: '写入公式' },
    'et.setValue': { app: 'et', risk: 'write', label: '写入内容' },
    'et.fillColor': { app: 'et', risk: 'write', label: '设置底色' },
    'et.fontColor': { app: 'et', risk: 'write', label: '设置字体颜色' },
    'et.bold': { app: 'et', risk: 'write', label: '加粗 / 取消加粗' },
    'et.italic': { app: 'et', risk: 'write', label: '斜体' },
    'et.fontSize': { app: 'et', risk: 'write', label: '字号' },
    'et.fontName': { app: 'et', risk: 'write', label: '字体' },
    'et.numberFormat': { app: 'et', risk: 'write', label: '数字格式' },
    'et.align': { app: 'et', risk: 'write', label: '对齐方式' },
    'et.border': { app: 'et', risk: 'write', label: '边框' },
    'et.columnWidth': { app: 'et', risk: 'write', label: '列宽' },
    'et.autoFit': { app: 'et', risk: 'write', label: '自适应列宽' },
    'et.conditionalFill': { app: 'et', risk: 'write', label: '条件高亮' },
    'et.sort': { app: 'et', risk: 'write', label: '排序' },
    'et.autoFilter': { app: 'et', risk: 'write', label: '筛选' },
    'et.freeze': { app: 'et', risk: 'write', label: '冻结窗格' },
    'et.insertChart': { app: 'et', risk: 'write', label: '插入图表' },
    'et.addSheet': { app: 'et', risk: 'write', label: '新建工作表' },
    'et.renameSheet': { app: 'et', risk: 'write', label: '重命名工作表' },
    'et.copyTo': { app: 'et', risk: 'write', label: '复制区域' },
    'et.fillSeries': { app: 'et', risk: 'write', label: '序列填充' },
    'et.clear': { app: 'et', risk: 'danger', label: '清空区域' },
    'et.deleteRows': { app: 'et', risk: 'danger', label: '删除行' }
  };

  var TYPE_SCHEMA = [
    'et.summary {range?} 读取数据摘要',
    'et.readRange {range?} 读取单元格内容',
    'et.setFormula {range, formula} 写入公式，如 =SUM(B2:B10)',
    'et.setValue {range, value} 写入值',
    'et.fillColor {range, color:"#RRGGBB"} 设置底色',
    'et.fontColor {range, color} 字体颜色',
    'et.bold {range, value:true|false}',
    'et.italic {range, value:true|false}',
    'et.fontSize {range, size:number}',
    'et.fontName {range, name}',
    'et.numberFormat {range, format} 如 "0.00%"、"0.00"、"¥#,##0.00"',
    'et.align {range, align:"left"|"center"|"right"}',
    'et.border {range} 添加全部边框',
    'et.columnWidth {range, width:number}',
    'et.autoFit {range} 自适应列宽',
    'et.conditionalFill {range, op:">"|">="|"<"|"<="|"="|"!=", threshold:number, color, fontColor?} 按条件高亮单元格',
    'et.sort {range, order:"asc"|"desc"}',
    'et.autoFilter {range}',
    'et.freeze {range} 冻结窗格',
    'et.insertChart {range, chartType:"column"|"line"}',
    'et.addSheet {name}',
    'et.renameSheet {name}',
    'et.copyTo {range, target}',
    'et.fillSeries {range, start:number, step:number}',
    'et.clear {range, what:"contents"|"formats"|"all"}',
    'et.deleteRows {range}'
  ];

  function safeRange(range) {
    if (!range) { return '当前选区'; }
    try { return Bridge.et.resolve(range); } catch (e) { return String(range); }
  }

  function describe(action) {
    var a = action || {};
    var r = safeRange(a.range);
    switch (a.type) {
      case 'et.summary': case 'et.readRange': return '读取 ' + r + ' 的数据摘要';
      case 'et.setFormula': return '在 ' + r + ' 写入公式 ' + a.formula;
      case 'et.setValue': return '在 ' + r + ' 写入「' + a.value + '」';
      case 'et.fillColor': return '把 ' + r + ' 的底色设为 ' + a.color;
      case 'et.fontColor': return '把 ' + r + ' 的字体颜色设为 ' + a.color;
      case 'et.bold': return (a.value === false ? '取消 ' : '加粗 ') + r + ' 的字体';
      case 'et.italic': return '设置 ' + r + ' 斜体';
      case 'et.fontSize': return '把 ' + r + ' 字号设为 ' + a.size;
      case 'et.fontName': return '把 ' + r + ' 字体设为 ' + a.name;
      case 'et.numberFormat': return '把 ' + r + ' 的数字格式设为 ' + a.format;
      case 'et.align': return '把 ' + r + ' 设为' + (a.align === 'center' ? '居中' : a.align === 'right' ? '右对齐' : '左对齐');
      case 'et.border': return '给 ' + r + ' 添加边框';
      case 'et.columnWidth': return '把 ' + r + ' 列宽设为 ' + a.width;
      case 'et.autoFit': return '让 ' + r + ' 列宽自适应';
      case 'et.conditionalFill': return '把 ' + r + ' 中 ' + a.op + ' ' + a.threshold + ' 的单元格标为 ' + a.color;
      case 'et.sort': return '对 ' + r + (a.order === 'desc' ? ' 降序' : ' 升序') + '排序';
      case 'et.autoFilter': return '为 ' + r + ' 添加筛选';
      case 'et.freeze': return '按 ' + r + ' 冻结窗格';
      case 'et.insertChart': return '基于 ' + r + ' 插入' + (a.chartType === 'line' ? '折线图' : '柱状图');
      case 'et.addSheet': return '新建工作表「' + a.name + '」';
      case 'et.renameSheet': return '重命名当前工作表为「' + a.name + '」';
      case 'et.copyTo': return '把 ' + r + ' 复制到 ' + safeRange(a.target);
      case 'et.fillSeries': return '在 ' + r + ' 填充序列（起始 ' + a.start + '，步长 ' + a.step + '）';
      case 'et.clear': return '清空 ' + r + (a.what === 'formats' ? ' 的格式' : a.what === 'all' ? ' 的内容和格式' : ' 的内容');
      case 'et.deleteRows': return '删除 ' + r + ' 所在的行';
      default: return '未知操作 ' + a.type;
    }
  }

  function riskOf(action) {
    var m = META[action.type];
    return m ? m.risk : 'write';
  }

  function labelOf(action) {
    var m = META[action.type];
    return m ? m.label : '操作';
  }

  /* 只允许白名单动作；限制条数与危险操作，防止 LLM 输出越界指令 */
  function strList(value, limit) {
    if (!value || Object.prototype.toString.call(value) !== '[object Array]') { return []; }
    var out = [];
    for (var i = 0; i < value.length && out.length < limit; i++) { out.push(String(value[i])); }
    return out;
  }

  function validate(plan) {
    var out = { app: 'et', summary: String(plan.summary || '执行操作'), actions: [], warnings: [] };
    if (!plan.actions || Object.prototype.toString.call(plan.actions) !== '[object Array]') {
      out.actions = [];
    } else {
      var list = plan.actions.slice ? plan.actions.slice(0, 12) : [];
      for (var i = 0; i < list.length; i++) {
        var a = list[i] || {};
        if (!META[a.type]) { out.warnings.push('已忽略不支持的操作：' + a.type); continue; }
        var clean = { type: a.type };
        var keys = ['range', 'target', 'color', 'fontColor', 'formula', 'value', 'text', 'name', 'size',
          'format', 'align', 'width', 'op', 'threshold', 'find', 'replace', 'rows', 'cols', 'data',
          'what', 'order', 'chartType', 'start', 'step', 'position', 'level', 'topic', 'draft', 'items', 'bold', 'italic'];
        for (var k = 0; k < keys.length; k++) {
          if (a[keys[k]] !== undefined && a[keys[k]] !== null) { clean[keys[k]] = a[keys[k]]; }
        }
        if (clean.color) { clean.color = Bridge.util.normalizeColor(clean.color); }
        if (clean.fontColor) { clean.fontColor = Bridge.util.normalizeColor(clean.fontColor); }
        if (clean.range) { clean.range = String(clean.range); }
        out.actions.push(clean);
      }
    }
    if (!out.actions.length) { out.warnings.push('没有可执行的动作'); }
    /* 透传「我理解为」与澄清候选，供界面在不确定时明确询问用户 */
    if (plan.intent && typeof plan.intent === 'object') {
      out.intent = {
        confidence: plan.intent.confidence === 'low' ? 'low' : 'high',
        reasons: strList(plan.intent.reasons, 6),
        scope: String(plan.intent.scope || ''),
        host: '当前表格',
        basis: strList(plan.intent.basis, 12)
      };
    }
    out.unparsed = strList(plan.unparsed, 6);
    if (plan.clarify && Object.prototype.toString.call(plan.clarify) === '[object Array]') {
      var opts = [];
      for (var c = 0; c < plan.clarify.length && opts.length < 4; c++) {
        var item = plan.clarify[c] || {};
        if (item.label && item.prompt) {
          opts.push({
            label: String(item.label).substring(0, 24),
            prompt: String(item.prompt).substring(0, 120)
          });
        }
      }
      if (opts.length) { out.clarify = opts; }
    }
    if (plan.needClarify) { out.needClarify = true; }
    if (plan.help) { out.help = true; }
    return out;
  }

  function hostCheck() {
    var info = Bridge.info();
    if (info.mode !== 'live' || info.host !== 'et') {
      return '未连接真实表格：请在表格客户端中打开侧边栏后再执行操作。';
    }
    return null;
  }

  function runAction(action) {
    var ET = Bridge.et;
    var a = action;
    switch (a.type) {
      case 'et.summary': case 'et.readRange': {
        var s = ET.summary(a.range || '@selection');
        return '区域 ' + s.range + '：' + s.rows + ' 行 × ' + s.cols + ' 列，数值 ' + s.numeric +
          ' 个，合计 ' + s.sum + '，平均 ' + s.average + '，最大 ' + s.max + '，最小 ' + s.min;
      }
      case 'et.setFormula': ET.setFormula(a.range, a.formula); return safeRange(a.range) + ' ← ' + Bridge.et.expand(a.formula);
      case 'et.setValue': ET.setValue(a.range, a.value); return safeRange(a.range) + ' 已写入';
      case 'et.fillColor': ET.setFill(a.range, a.color); return safeRange(a.range) + ' 底色已设为 ' + a.color;
      case 'et.fontColor': ET.setFont(a.range, { color: a.color }); return safeRange(a.range) + ' 字体颜色已更新';
      case 'et.bold': ET.setFont(a.range, { bold: a.value !== false }); return safeRange(a.range) + ' 已' + (a.value === false ? '取消' : '加') + '粗';
      case 'et.italic': ET.setFont(a.range, { italic: a.value !== false }); return safeRange(a.range) + ' 斜体已设置';
      case 'et.fontSize': ET.setFont(a.range, { size: Number(a.size) || 11 }); return safeRange(a.range) + ' 字号已更新';
      case 'et.fontName': ET.setFont(a.range, { name: a.name }); return safeRange(a.range) + ' 字体已更新';
      case 'et.numberFormat': ET.setFormat(a.range, { numberFormat: a.format }); return safeRange(a.range) + ' 数字格式 ' + a.format;
      case 'et.align': ET.setFormat(a.range, { align: a.align }); return safeRange(a.range) + ' 对齐已更新';
      case 'et.border': ET.setFormat(a.range, { border: true }); return safeRange(a.range) + ' 边框已添加';
      case 'et.columnWidth': ET.setFormat(a.range, { width: Number(a.width) || 12 }); return safeRange(a.range) + ' 列宽已更新';
      case 'et.autoFit': ET.setFormat(a.range, { autoFit: true }); return safeRange(a.range) + ' 列宽已自适应';
      case 'et.conditionalFill': {
        var n = ET.conditionalFill(a.range, a.op, Number(a.threshold), a.color, a.fontColor);
        return '符合条件 ' + a.op + ' ' + a.threshold + ' 的单元格共 ' + n + ' 个，已标记';
      }
      case 'et.sort': { var rows = ET.sort(a.range, a.order); return '排序完成' + (rows ? '（' + rows + ' 行）' : ''); }
      case 'et.autoFilter': ET.autoFilter(a.range); return '筛选已添加';
      case 'et.freeze': ET.freeze(a.range || 'A2'); return '窗格已冻结';
      case 'et.insertChart': ET.insertChart(a.range, a.chartType); return '图表已插入';
      case 'et.addSheet': ET.addSheet(a.name || '新工作表'); return '工作表已创建';
      case 'et.renameSheet': ET.renameSheet(a.name || '工作表1'); return '工作表已重命名';
      case 'et.copyTo': ET.copyTo(a.range, a.target); return '复制完成';
      case 'et.fillSeries': ET.fillSeries(a.range, a.start, a.step); return '序列已填充';
      case 'et.clear': ET.clear(a.range, a.what || 'contents'); return safeRange(a.range) + ' 已清空';
      case 'et.deleteRows': ET.deleteRows(a.range); return '行已删除';
      default: throw new Error('未实现的操作：' + a.type);
    }
  }

  function execute(plan, onProgress) {
    var result = { ok: true, succeeded: 0, failed: 0, details: [], startedAt: Date.now() };
    var problem = hostCheck();
    if (problem) { return { ok: false, succeeded: 0, failed: plan.actions.length, details: [{ label: '环境检查', ok: false, text: problem }] }; }
    for (var i = 0; i < plan.actions.length; i++) {
      var a = plan.actions[i];
      var item = { label: labelOf(a), description: describe(a), ok: true, text: '' };
      try {
        item.text = String(runAction(a));
        result.succeeded++;
      } catch (e) {
        item.ok = false;
        item.text = e && e.message ? e.message : String(e);
        result.failed++;
      }
      result.details.push(item);
      if (onProgress) { onProgress(item, i + 1, plan.actions.length); }
    }
    result.ok = result.failed === 0;
    result.finishedAt = Date.now();
    return result;
  }

  function schemaForLLM() { return TYPE_SCHEMA.join('\n'); }
  function listTypes() {
    var arr = [];
    for (var k in META) { if (Object.prototype.hasOwnProperty.call(META, k)) { arr.push(k); } }
    return arr;
  }

  global.WpsAIActions = {
    meta: META,
    describe: describe,
    label: labelOf,
    risk: riskOf,
    validate: validate,
    execute: execute,
    hostCheck: hostCheck,
    schema: schemaForLLM,
    listTypes: listTypes
  };
})(window);
