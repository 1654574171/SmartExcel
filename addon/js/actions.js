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
    'et.deleteRows': { app: 'et', risk: 'danger', label: '删除行' },
    'et.deleteDuplicateRows': { app: 'et', risk: 'danger', label: '按内容去重删除行' },
    'et.deleteRowsIf': { app: 'et', risk: 'danger', label: '按条件删除行' },
    'et.findReplace': { app: 'et', risk: 'write', label: '查找替换' },
    'et.fillBlanks': { app: 'et', risk: 'write', label: '填充空白单元格' },
    'et.fillDown': { app: 'et', risk: 'write', label: '公式向下填充' },
    'et.mergeCells': { app: 'et', risk: 'danger', label: '合并单元格' },
    'et.insertRows': { app: 'et', risk: 'danger', label: '插入行' },
    'et.insertColumns': { app: 'et', risk: 'danger', label: '插入列' },
    'et.deleteColumns': { app: 'et', risk: 'danger', label: '删除列' },
    'et.hideColumns': { app: 'et', risk: 'danger', label: '隐藏 / 显示列' },
    'et.dataValidation': { app: 'et', risk: 'write', label: '设置下拉选项' },
    'et.pageSetup': { app: 'et', risk: 'write', label: '打印区域设置' },
    'et.conditionalFormatRule': { app: 'et', risk: 'write', label: '条件格式规则' },
    'et.pivotTable': { app: 'et', risk: 'danger', label: '数据透视表' },
    'et.textToColumns': { app: 'et', risk: 'write', label: '按分隔符分列' }
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
    'et.conditionalFill {range, op:">"|">="|"<"|"<="|"="|"!="|"contains", threshold:number|string, color, fontColor?, fillScope:"cell"|"row"} 只给满足条件的单元格或整行上色，文本条件如 threshold:"是"',
    'et.sort {range, order:"asc"|"desc", columns?:["A","C"]} 按指定列 / 多关键字排序',
    'et.autoFilter {range}',
    'et.freeze {range} 冻结窗格',
    'et.insertChart {range, chartType:"column"|"line"}',
    'et.addSheet {name}',
    'et.renameSheet {name}',
    'et.copyTo {range, target}',
    'et.fillSeries {range, start:number, step:number}',
    'et.clear {range, what:"contents"|"formats"|"all"}',
    'et.deleteRows {range} 删除指定区域所在的整行',
    'et.deleteDuplicateRows {range, keyColumn?, keep:"first"|"last"} 按内容删除重复行；用户明确要求去重 / 删除重复行时直接使用，不要因为破坏性就返回空 actions',
    'et.deleteRowsIf {range, op:"="|"!="|">"|">="|"<"|"<="|"contains", value:number|string, keyColumn?} 按条件删除整行，如把 D2:D6 中内容为「是」的行删除',
    'et.findReplace {range, find, replace} 把区域内的文本替换成新文本，公式单元格不动',
    'et.fillBlanks {range, value} 只给区域内的空白单元格填值',
    'et.fillDown {range} 把区域首行的公式向下填充到其余行',
    'et.mergeCells {range, center?:true|false} 合并单元格，可同时居中',
    'et.insertRows {range, count?:number} 在指定位置插入整行',
    'et.insertColumns {range, count?:number} 插入整列',
    'et.deleteColumns {range} 删除整列',
    'et.hideColumns {range, hidden:true|false} 隐藏或取消隐藏列',
    'et.dataValidation {range, list:["是","否"]} 设置下拉选项',
    'et.pageSetup {range, titleRows?} 设置打印区域 / 打印标题行',
    'et.conditionalFormatRule {range, formula, color} 新建条件格式规则，如隔行底纹 =MOD(ROW(A1),2)=0',
    'et.pivotTable {range, target} 基于区域创建数据透视表',
    'et.textToColumns {range, delimiter} 按分隔符把一列拆成多列，如按逗号分列'
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
      case 'et.conditionalFill': return '只把 ' + r + ' 中「' + a.op + ' ' + a.threshold + '」' +
        (a.fillScope === 'row' ? '的单元格所在整行' : '的单元格') + '标为 ' + a.color + '，其余单元格不动';
      case 'et.sort': return '对 ' + r + (a.order === 'desc' ? ' 降序' : ' 升序') + '排序' +
        (a.columns && a.columns.length ? '（按 ' + a.columns.join('、') + ' 列）' : '') + '；整行数据会一起重排';
      case 'et.autoFilter': return '为 ' + r + ' 添加筛选';
      case 'et.freeze': return '按 ' + r + ' 冻结窗格';
      case 'et.insertChart': return '基于 ' + r + ' 插入' + (a.chartType === 'line' ? '折线图' : '柱状图');
      case 'et.addSheet': return '新建工作表「' + a.name + '」';
      case 'et.renameSheet': return '重命名当前工作表为「' + a.name + '」';
      case 'et.copyTo': return '把 ' + r + ' 复制到 ' + safeRange(a.target);
      case 'et.fillSeries': return '在 ' + r + ' 填充序列（起始 ' + a.start + '，步长 ' + a.step + '）';
      case 'et.clear': return '清空 ' + r + (a.what === 'formats' ? ' 的格式' : a.what === 'all' ? ' 的内容和格式' : ' 的内容');
      case 'et.deleteRows': return '删除 ' + r + ' 所在的行';
      case 'et.deleteDuplicateRows': return '在 ' + r + ' 中按' +
        (a.keyColumn ? '「' + a.keyColumn + '」列' : '整行内容') + '比对，删除重复行（保留' +
        (a.keep === 'last' ? '最后' : '首次') + '出现）；会删除整行，影响范围外的列也会随之上移';
      case 'et.deleteRowsIf': return '删除 ' + r + ' 中「' + a.op + ' ' + a.value + '」' +
        (a.keyColumn ? '（按 ' + a.keyColumn + ' 列判断）' : '') +
        '的单元格所在整行；会删除整行，影响范围外的列也会随之上移';
      case 'et.findReplace': return '把 ' + r + ' 中的「' + a.find + '」替换为「' + a.replace + '」（公式单元格不动）';
      case 'et.fillBlanks': return '把 ' + r + ' 中的空白单元格填为「' + a.value + '」，有内容的单元格不动';
      case 'et.fillDown': return '把 ' + r + ' 首行的公式向下填充到其余行';
      case 'et.mergeCells': return '合并 ' + r + (a.center ? ' 并居中' : '') + '；只保留左上角单元格的内容';
      case 'et.insertRows': return '在 ' + r + ' 处插入 ' + (a.count || 1) + ' 行，下方数据会下移';
      case 'et.insertColumns': return '在 ' + r + ' 处插入 ' + (a.count || 1) + ' 列，右侧数据会右移';
      case 'et.deleteColumns': return '删除 ' + r + ' 所在的整列，右侧数据会左移';
      case 'et.hideColumns': return (a.hidden === false ? '取消隐藏 ' : '隐藏 ') + r + ' 所在的列';
      case 'et.dataValidation': return '给 ' + r + ' 设置下拉选项：' + ((a.list || []).join('、') || '（未给出选项）');
      case 'et.pageSetup': return '把 ' + r + ' 设为打印区域' + (a.titleRows ? '，并打印标题行 ' + a.titleRows : '');
      case 'et.conditionalFormatRule': return '给 ' + r + ' 新建条件格式规则 ' + a.formula + '，命中显示 ' + a.color;
      case 'et.pivotTable': return '基于 ' + r + ' 创建数据透视表，放在 ' + safeRange(a.target);
      default: {
        var extraMeta = META[a.type];
        return extraMeta ? extraMeta.label + '：' + r : '未知操作 ' + a.type;
      }
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
          'format', 'align', 'width', 'op', 'threshold', 'find', 'replace', 'delimiter', 'rows', 'cols', 'data',
          'what', 'order', 'chartType', 'start', 'step', 'position', 'level', 'topic', 'draft', 'items', 'bold', 'italic',
          'keyColumn', 'keep', 'fillScope', 'count', 'center', 'hidden', 'columns', 'list', 'titleRows'];
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
        /* 阈值可能是「是」这类文本，不能强制转成数字 */
        var n = ET.conditionalFill(a.range, a.op, a.threshold, a.color, a.fontColor,
          { fillScope: a.fillScope });
        return '符合条件 ' + a.op + ' ' + a.threshold + ' 的共 ' + n +
          (a.fillScope === 'row' ? ' 处，已标记（整行）' : ' 个单元格，已标记（仅命中单元格）');
      }
      case 'et.deleteRowsIf': {
        var removed = ET.deleteRowsIf(a.range, a.op, a.value, a.keyColumn);
        return '在 ' + safeRange(a.range) + ' 中按「' + a.op + ' ' + a.value + '」删除了 ' +
          removed.removed + ' 行；请用表格自带的撤销恢复';
      }
      case 'et.textToColumns': {
        var split = ET.textToColumns(a.range, a.delimiter);
        return '在 ' + safeRange(a.range) + ' 中按分隔符拆开了 ' + split + ' 个单元格';
      }
      case 'et.findReplace': {
        var replaced = ET.findReplace(a.range, a.find, a.replace);
        return '在 ' + safeRange(a.range) + ' 中替换了 ' + replaced + ' 处（公式单元格未改动）';
      }
      case 'et.fillBlanks': {
        var blanks = ET.fillBlanks(a.range, a.value);
        return '填充了 ' + blanks + ' 个空白单元格';
      }
      case 'et.fillDown': {
        var down = ET.fillDown(a.range);
        return safeRange(a.range) + ' 首行公式已向下填充' + (down ? ' ' + down + ' 行' : '');
      }
      case 'et.mergeCells': ET.mergeCells(a.range, a.center); return safeRange(a.range) + ' 已合并；请用表格自带的撤销恢复';
      case 'et.insertRows': ET.insertRows(a.range, a.count); return '已插入 ' + (a.count || 1) + ' 行；请用表格自带的撤销恢复';
      case 'et.insertColumns': ET.insertColumns(a.range, a.count); return '已插入列；请用表格自带的撤销恢复';
      case 'et.deleteColumns': ET.deleteColumns(a.range); return '已删除列；请用表格自带的撤销恢复';
      case 'et.hideColumns': ET.hideColumns(a.range, a.hidden); return '列可见性已更新；请用表格自带的撤销恢复';
      case 'et.dataValidation': ET.dataValidation(a.range, a.list); return safeRange(a.range) + ' 的下拉选项已设置';
      case 'et.pageSetup': ET.pageSetup(a.range, { titleRows: a.titleRows }); return '打印区域已设置';
      case 'et.conditionalFormatRule': ET.conditionalFormatRule(a.range, a.formula, a.color);
        return '条件格式规则已创建';
      case 'et.pivotTable': ET.pivotTable(a.range, a.target); return '数据透视表已创建；请用表格自带的撤销恢复';
      case 'et.sort': { var rows = ET.sort(a.range, a.order, a.columns); return '排序完成' + (rows ? '（' + rows + ' 行）' : ''); }
      case 'et.autoFilter': ET.autoFilter(a.range); return '筛选已添加';
      case 'et.freeze': ET.freeze(a.range || 'A2'); return '窗格已冻结';
      case 'et.insertChart': ET.insertChart(a.range, a.chartType); return '图表已插入';
      case 'et.addSheet': ET.addSheet(a.name || '新工作表'); return '工作表已创建';
      case 'et.renameSheet': ET.renameSheet(a.name || '工作表1'); return '工作表已重命名';
      case 'et.copyTo': ET.copyTo(a.range, a.target); return '复制完成';
      case 'et.fillSeries': ET.fillSeries(a.range, a.start, a.step); return '序列已填充';
      case 'et.clear': ET.clear(a.range, a.what || 'contents'); return safeRange(a.range) + ' 已清空';
      case 'et.deleteRows': ET.deleteRows(a.range); return '行已删除';
      case 'et.deleteDuplicateRows': {
        var dedup = ET.deleteDuplicateRows(a.range, { keyColumn: a.keyColumn, keep: a.keep || 'first' });
        return '在 ' + safeRange(a.range) + ' 中删除了 ' + dedup.removed + ' 行重复数据，保留 ' +
          dedup.kept + ' 行；整行已删除，请用表格自带的撤销恢复';
      }
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

  /* 撤回只需恢复动作真正改动的属性：改底色绝不能顺带重写整片区域的字体色 */
  var PROPS_BY_TYPE = {
    'et.setFormula': ['value'],
    'et.setValue': ['value'],
    'et.fillSeries': ['value'],
    'et.copyTo': ['value', 'fill', 'fontColor', 'bold', 'italic', 'size', 'fontName', 'numberFormat', 'align'],
    'et.sort': ['value'],
    'et.fillColor': ['fill'],
    'et.conditionalFill': ['fill'],
    'et.fontColor': ['fontColor'],
    'et.bold': ['bold'],
    'et.italic': ['italic'],
    'et.fontSize': ['size'],
    'et.fontName': ['fontName'],
    'et.numberFormat': ['numberFormat'],
    'et.align': ['align'],
    'et.border': [],
    'et.columnWidth': ['width'],
    'et.autoFit': ['width'],
    'et.autoFilter': [], 'et.freeze': [], 'et.insertChart': [],
    'et.addSheet': [], 'et.renameSheet': [],
    'et.deleteRows': [], 'et.deleteDuplicateRows': [], 'et.deleteRowsIf': [],
    'et.findReplace': ['value'], 'et.fillBlanks': ['value'], 'et.fillDown': ['value'],
    'et.textToColumns': ['value'],
    'et.mergeCells': [], 'et.insertRows': [], 'et.insertColumns': [],
    'et.deleteColumns': [], 'et.hideColumns': [], 'et.dataValidation': [],
    'et.pageSetup': [], 'et.conditionalFormatRule': [], 'et.pivotTable': []
  };

  function affectedProps(plan) {
    var actions = (plan && plan.actions) || [];
    var seen = {}, out = [], unknown = false;
    for (var i = 0; i < actions.length; i++) {
      var a = actions[i];
      if (!a || riskOf(a) === 'read') { continue; }
      var list;
      if (a.type === 'et.clear') {
        list = a.what === 'formats'
          ? ['fill', 'fontColor', 'bold', 'italic', 'size', 'fontName', 'numberFormat', 'align']
          : (a.what === 'all'
            ? ['value', 'fill', 'fontColor', 'bold', 'italic', 'size', 'fontName', 'numberFormat', 'align']
            : ['value']);
      } else {
        list = PROPS_BY_TYPE[a.type];
      }
      if (!list) { unknown = true; continue; }
      /* 条件高亮带 fontColor 参数时才会碰字体色 */
      if (a.type === 'et.conditionalFill' && a.fontColor) { list = list.concat(['fontColor']); }
      for (var j = 0; j < list.length; j++) {
        if (!seen[list[j]]) { seen[list[j]] = true; out.push(list[j]); }
      }
    }
    /* 未知动作或没有可识别属性时退回全量恢复，宁可多读也不漏恢复 */
    return (unknown || !out.length) ? null : out;
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
    affectedProps: affectedProps,
    hostCheck: hostCheck,
    schema: schemaForLLM,
    listTypes: listTypes
  };
})(window);
