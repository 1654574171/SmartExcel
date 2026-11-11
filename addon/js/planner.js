/*
 * 自然语言 -> 操作计划。
 * 本地规则引擎（默认、无需联网）+ 可选大模型（由 llm.js 提供）。
 * 支持一句话多个意图、上下文指代（它 / 同上 / 刚才的区域），仅面向表格。
 */
(function (global) {
  'use strict';

  var Bridge = global.WpsAIBridge;
  var Util = Bridge.util;
  var Actions = global.WpsAIActions;

  var COLOR_WORDS = [
    [/深红|暗红/, '#C00000'],
    [/标红|红色|红底|浅红|警示|异常|不合格|标红/, '#FFC7CE'],
    [/浅黄/, '#FFF2CC'],
    [/标黄|黄色|黄底|高亮|突出/, '#FFFF00'],
    [/标绿|绿色|绿底|浅绿|达标|合格/, '#C6EFCE'],
    [/蓝色|蓝底|浅蓝|标蓝/, '#DDEBF7'],
    [/橙色|橙底|橘色/, '#FCE4D6'],
    [/紫色|紫底/, '#E4DFEC'],
    [/灰色|灰底/, '#E7E6E6'],
    [/黑色|黑底/, '#000000'],
    [/白色|白底/, '#FFFFFF']
  ];
  var RED_FONT = '#9C0006';
  var GREEN_FONT = '#006100';


  function normalize(text) {
    return String(text || '')
      .replace(/[，]/g, ',').replace(/[；]/g, ';').replace(/[：]/g, ':')
      .replace(/[“”]/g, '"').replace(/[‘’]/g, "'")
      .replace(/[（]/g, '(').replace(/[）]/g, ')')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function splitClauses(text) {
    /* 逗号经常只是「如果…,标为…,结果放在…」内部的停顿，只有后面另起一条动作才切分；
       英文 ! 属于「工作表名!区域」，绝不能当句末标点切断跨表引用。 */
    var raw = String(text);
    var boundary = /[,;。?、]|！|并且|然后|接着|同时|另外|以及|并|再(?=[把将给设置标加做])/g;
    var actionStart = /^(把|将|给|再|然后|接着|同时|另外|以及|并且|设置|删除|清空|排序|筛选|冻结|复制|新建|插入|对当前|对选中)/;
    var out = [], start = 0, m;
    while ((m = boundary.exec(raw)) !== null) {
      var rest = raw.slice(m.index + m[0].length).replace(/^\s+/, '');
      if (!actionStart.test(rest)) { continue; }
      var part = raw.slice(start, m.index).trim();
      if (part) { out.push(part); }
      start = m.index + m[0].length;
    }
    var tail = raw.slice(start).trim();
    if (tail) { out.push(tail); }
    return out.length ? out : [raw.trim()];
  }

  /* 地址前面紧邻的「工作表名!」限定必须一起带走，否则跨表引用会退化成当前表 */
  function sheetPrefixBefore(text, index) {
    var m = /(?:\[[^\]]+\])?(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*|[\u4e00-\u9fa5]+)!$/.exec(String(text).slice(0, index));
    return m ? m[0] : '';
  }

  function extractRanges(text) {
    var out = [];
    var re = /\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      out.push(sheetPrefixBefore(text, m.index) + m[0].replace(/\$/g, '').toUpperCase());
    }
    var colRe = /(^|[^A-Z0-9])([A-Z]{1,3})\s*列/g;
    while ((m = colRe.exec(text)) !== null) {
      out.push(sheetPrefixBefore(text, m.index + (m[1] ? m[1].length : 0)) + m[2] + '1:' + m[2] + '@lastRow');
    }
    /* 「自动调整 A:F 的列宽」要识别成整段列范围，而不是退回当前选区 */
    var spanRe = /(^|[^A-Z0-9])([A-Z]{1,3})\s*:\s*([A-Z]{1,3})\s*列?/g;
    while ((m = spanRe.exec(text)) !== null) {
      out.push(sheetPrefixBefore(text, m.index + (m[1] ? m[1].length : 0)) +
        m[2] + '1:' + m[3] + '@lastRow');
    }
    var rowRe = /第\s*(\d+)\s*行/g;
    while ((m = rowRe.exec(text)) !== null) { out.push(m[1] + ':' + m[1]); }
    return out;
  }

  function extractColor(text) {
    var hex = /#([0-9A-F]{6}|[0-9A-F]{3})/i.exec(text);
    if (hex) { return Util.normalizeColor(hex[0]); }
    for (var i = 0; i < COLOR_WORDS.length; i++) {
      if (COLOR_WORDS[i][0].test(text)) { return COLOR_WORDS[i][1]; }
    }
    return null;
  }

  function inferTarget(source) {
    if (!source || source.indexOf('@selection') >= 0) { return '@selectionBottom'; }
    var bang = source.lastIndexOf('!');
    var prefix = bang >= 0 ? source.slice(0, bang + 1) : '';
    var plain = bang >= 0 ? source.slice(bang + 1) : source;
    var m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(plain);
    if (m) {
      var col = m[3] || m[1];
      return prefix + col + ((parseInt(m[4] || m[2], 10) || 1) + 1);
    }
    var colOnly = /^([A-Z]{1,3})1:[A-Z]{1,3}@lastRow$/.exec(plain);
    if (colOnly) { return prefix + colOnly[1] + '@lastRowPlus1'; }
    return '@selectionBottom';
  }

  function splitTarget(clause) {
    var m = /(?:放到|放在|写入|写到|填到|输出到|结果放在|结果在|存到|填在)/.exec(clause);
    if (m) {
      return { sourcePart: clause.substring(0, m.index), targetPart: clause.substring(m.index + m[0].length) };
    }
    return { sourcePart: clause, targetPart: '' };
  }

  /* 只面向表格，保留函数签名以兼容调用方 */
  function detectApp() { return 'et'; }

  /* 公式类指令统一委托公式顾问：区域、阈值、返回值都从当前指令解析，
     不再用关键词套写死阈值 / 文案 / 列号的模板，也不限定某张表、某几列。 */
  function formulaActions(clause, ctx, fallbackTarget) {
    var Advisor = global.WpsAIAdvisor;
    if (!Advisor || !Advisor.advise) { return []; }
    var advice;
    try { advice = Advisor.advise(clause, ctx || {}); } catch (e) { return []; }
    if (!advice || advice.kind !== 'formula' || advice.intentKind === 'conditionalFormat' ||
        !advice.formulas || !advice.formulas.length) { return []; }
    var formula = String(advice.formulas[0].formula || '');
    if (formula.charAt(0) !== '=') { return []; }
    /* 公式里还留着占位说明时不要写进单元格，宁可让用户补全，也不能写入错误公式 */
    if (/"(?:条件1?|条件2|满足条件|不满足条件)"|查找表!|@lookupRange|@区域|替换为/.test(formula)) {
      return [];
    }
    /* SUMIFS / COUNTIFS 要用到「求和区 + 多个条件区」，只凭选区推断出的列不可靠。
       公式建议模式仍给出草案，执行模式等用户把区域说清楚后再落地。 */
    if ((advice.intentKind === 'sumifs' || advice.intentKind === 'countifs') &&
        extractRanges(clause).length < 2) {
      return [];
    }
    var target = '', area = '';
    for (var i = 0; i < (advice.points || []).length; i++) {
      var p = advice.points[i];
      if (p.label === '结果区域') { area = String(p.value); }
      if (p.label === '结果位置' && p.value !== '未指定') { target = String(p.value); }
    }
    var range = advice.intentKind === 'multiIf' ? (area || target) : (target || fallbackTarget);
    if (!range) { range = '@selectionBottom'; }
    return [{ type: 'et.setFormula', range: range, formula: formula }];
  }

  function quotedValue(clause) {
    var m = /["']([^"']{1,60})["']/.exec(clause);
    if (m) { return m[1]; }
    m = /(?:写入|填入|输入|填写|设为|改成|改为|内容为|值是)\s*[:：]?\s*([^\s,;]{1,40})/.exec(clause);
    return m ? m[1] : null;
  }

  function parseEtClause(clause, state) {
    var actions = [];
    var c = clause.toUpperCase();
    var parts = splitTarget(clause);
    var sourceRanges = extractRanges(parts.sourcePart);
    var targetRanges = extractRanges(parts.targetPart);
    var color = extractColor(clause);
    /* 用户明确说「当前选区」时必须用选区，不能沿用上一步的区域 */
    var explicitSelection = /当前选区|选中区域|选区|当前区域|所选区域/.test(clause);
    var source;
    if (sourceRanges.length) { source = sourceRanges[0]; }
    else if (explicitSelection) { source = '@selection'; }
    else { source = state.currentRef || '@selection'; }
    var target = targetRanges.length ? targetRanges[0] : inferTarget(source);
    var range = sourceRanges.length ? sourceRanges[0] : source;

    if (color) { state.color = color; }

    /* 条件高亮：大于 / 小于 X 标黄 */
    var cond = null;
    if (/不小于|>=|≥|大于|超过|高于/.test(clause)) { cond = '>='; }
    else if (/不大于|<=|≤|小于|低于|不足/.test(clause)) { cond = '<='; }
    else if (/不等于|!=/.test(clause)) { cond = '!='; }
    else if (/等于|恰好是/.test(clause)) { cond = '='; }
    if (/大于|超过|高于/.test(clause)) { cond = '>'; }
    if (/小于|低于|不足/.test(clause)) { cond = '<'; }
    var condNum = /(?:大于|超过|高于|小于|低于|不足|等于|不等于|>=|<=|≥|≤|!=)\s*(\d+(?:\.\d+)?)/.exec(clause);
    if (cond && condNum && color) {
      actions.push({
        type: 'et.conditionalFill', range: range, op: cond,
        threshold: Number(condNum[1]), color: color,
        fillScope: /整行|的行/.test(clause) ? 'row' : 'cell',
        fontColor: color === '#FFC7CE' ? RED_FONT : (color === '#C6EFCE' ? GREEN_FONT : '')
      });
      state.currentRef = range;
      return actions;
    }

    /* 「为『是』的行标黄 / 删除」这类文本条件，必须抢在整区域填色和普通删除行之前 */
    var wantsDelete = /(删除|删掉|去掉|去除|清除|清理|剔除)/.test(clause);
    var textCond = /(?:为|是|等于|内容是|值为|标记为)\s*[:：]?\s*["'「『]?([^"'」』\s,，。;；]{1,12})["'」』]?\s*(?:的|，|,)/.exec(clause) ||
      /["'「『]([^"'」』]{1,12})["'」』]\s*的/.exec(clause);
    if (textCond) {
      if (wantsDelete) {
        actions.push({ type: 'et.deleteRowsIf', range: range, op: '=', value: textCond[1] });
        state.currentRef = range;
        return actions;
      }
      if (color) {
        actions.push({
          type: 'et.conditionalFill', range: range, op: '=', threshold: textCond[1], color: color,
          fillScope: /整行|的行/.test(clause) ? 'row' : 'cell'
        });
        state.currentRef = range;
        return actions;
      }
    }
    if (cond && condNum && wantsDelete && !/重复|去重/.test(clause)) {
      actions.push({ type: 'et.deleteRowsIf', range: range, op: cond, value: Number(condNum[1]) });
      state.currentRef = range;
      return actions;
    }

    /* 一、互斥的清洗 / 结构 / 计算动作 */
    var wantsRound = /四舍五入|ROUND/i.test(c);
    var replacePair =
      /["'「『]([^"'」』]{1,20})["'」』]\s*(?:替换成|替换为|替换|改成|改为|换成)\s*["'「『]?([^"'」』\s,，。;；]{1,20})/.exec(clause) ||
      /(?:里的|中的|中|里)\s*["'「『]?([^"'」』\s,，。;；]{1,20}?)["'」』]?\s*(?:替换成|替换为|改成|改为|换成|替换)\s*["'「『]?([^"'」』\s,，。;；]{1,20})["'」』]?/.exec(clause);
    if (replacePair && /替换|改成|改为|换成/.test(clause)) {
      actions.push({ type: 'et.findReplace', range: range, find: replacePair[1], replace: replacePair[2] });
    } else if (/分列|拆分成列|按.*拆分/.test(clause)) {
      var delimWord = /按\s*["'「『]?([^"'」』\s,，。;；]{1,6})["'」』]?\s*(?:分列|拆分)/.exec(clause);
      var word = delimWord ? delimWord[1] : ',';
      var delimMap = { '逗号': ',', '分号': ';', '顿号': '、', '空格': ' ', '制表符': '\t', '竖线': '|' };
      actions.push({ type: 'et.textToColumns', range: range, delimiter: delimMap[word] || word });
    } else if (/空白|空单元格|空值/.test(clause) && /填|补|设为/.test(clause)) {
      var blankValue = /(?:填成|填为|填入|填|补|设为)\s*[:：]?\s*["'「『]?([^"'」』\s,，。;；]{1,20})/.exec(clause);
      actions.push({ type: 'et.fillBlanks', range: range, value: blankValue ? blankValue[1] : '0' });
    } else if (/向下填充|下拉填充|填充公式|公式.*下拉|往下拖/.test(clause)) {
      /* 「把 A2 的公式向下填充到 A20」要把两个地址合成一个区域 */
      var downRange = range;
      if (sourceRanges.length >= 2) {
        var one = Util.parseA1(sourceRanges[0]);
        var two = Util.parseA1(sourceRanges[1]);
        if (one && two && one.c1 === two.c1) {
          var downPrefix = Util.splitSheetAddr(sourceRanges[0]).prefix;
          downRange = downPrefix + Util.numToCol(one.c1) + Math.min(one.r1, two.r1) + ':' +
            Util.numToCol(one.c1) + Math.max(one.r1, two.r1);
        }
      }
      actions.push({ type: 'et.fillDown', range: downRange });
    } else if (/插入.*行|在第.*行.*插入|新增.*行/.test(clause)) {
      /* 只从「插入 N 行」取数量，不能把「第 2 行」的行号当成插入行数 */
      var insCount = /插入\s*(\d+)\s*行/.exec(clause);
      var insCn = /插入\s*([一两二三四五六七八九十])\s*行/.exec(clause);
      var cnNum = { '一': 1, '两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
      var insTotal = insCount ? Number(insCount[1]) : (insCn ? (cnNum[insCn[1]] || 1) : 1);
      actions.push({ type: 'et.insertRows', range: range, count: insTotal });
    } else if (/插入.*列|新增.*列/.test(clause)) {
      actions.push({ type: 'et.insertColumns', range: range, count: 1 });
    } else if (/删除.*列|删掉.*列/.test(clause)) {
      actions.push({ type: 'et.deleteColumns', range: range });
    } else if (/隐藏|取消隐藏|显示.*列/.test(clause) && /列/.test(clause)) {
      actions.push({ type: 'et.hideColumns', range: range, hidden: !/取消隐藏|取消|显示/.test(clause) });
    } else if (/下拉选项|下拉菜单|数据验证|有效性|可选值|选项列表/.test(clause)) {
      var opts = /(?:选项|下拉|列表)\s*[:：]\s*["'「『]?([^"'」』。;；]{1,80})/.exec(clause);
      actions.push({ type: 'et.dataValidation', range: range,
        list: opts ? opts[1].split(/[,，、\/|]/) : [] });
    } else if (/打印区域|打印设置|打印标题|页面设置/.test(clause)) {
      actions.push({ type: 'et.pageSetup', range: range,
        titleRows: /标题行|首行/.test(clause) ? '1:1' : '' });
    } else if (/条件格式规则|新建规则|隔行底纹|隔行填色|斑马纹/.test(clause)) {
      var plainRange = Util.splitSheetAddr(range).plain;
      var firstCell = String(plainRange).split(':')[0].replace(/\$/g, '').toUpperCase();
      var rc = /(大于|小于|不小于|不大于|等于)\s*(\d+(?:\.\d+)?)/.exec(clause);
      var opMap = { '大于': '>', '小于': '<', '不小于': '>=', '不大于': '<=', '等于': '=' };
      var ruleFormula = rc ? '=' + firstCell + (opMap[rc[1]] || '>') + rc[2]
        : '=MOD(ROW(' + firstCell + '),2)=0';
      actions.push({ type: 'et.conditionalFormatRule', range: range,
        formula: ruleFormula, color: color || '#FFFF00' });
    } else if (/透视表|数据透视/.test(clause)) {
      actions.push({ type: 'et.pivotTable', range: range, target: target });
    } else if (/排序|升序|降序/.test(clause)) {
      var sortCols = [], sc, scm = /按\s*\$?([A-Z]{1,3})\s*列/g;
      while ((sc = scm.exec(clause)) !== null) { sortCols.push(sc[1].toUpperCase()); }
      var sortAction = { type: 'et.sort', range: range,
        order: /降序|从大到小|倒序/.test(clause) ? 'desc' : 'asc' };
      if (sortCols.length || /先按|再按|然后按/.test(clause)) { sortAction.columns = sortCols; }
      actions.push(sortAction);
    } else if (/筛选|过滤/.test(clause)) {
      actions.push({ type: 'et.autoFilter', range: range });
    } else if (/冻结/.test(clause)) {
      actions.push({ type: 'et.freeze', range: /冻结首行|冻结第一行/.test(clause) ? 'A2' : range });
    } else if (/图表|柱状图|折线图/.test(clause)) {
      actions.push({ type: 'et.insertChart', range: range, chartType: /折线/.test(clause) ? 'line' : 'column' });
    } else if (/新建工作表|添加工作表|新增工作表/.test(clause)) {
      var sname = /(?:新建|添加|新增)工作表\s*[:：]?\s*[「『"']?([\u4e00-\u9fa5A-Za-z0-9]{1,16})/.exec(clause) ||
        /(?:名为|叫)\s*[:：]?\s*[「『"']?([\u4e00-\u9fa5A-Za-z0-9]{1,16})/.exec(clause);
      actions.push({ type: 'et.addSheet', name: sname ? sname[1] : '新工作表' });
    } else if (/重命名|改名为|工作表名称/.test(clause)) {
      var rname = /(?:为|成|[:：])\s*([\u4e00-\u9fa5A-Za-z0-9]{1,16})\s*$/.exec(clause);
      actions.push({ type: 'et.renameSheet', name: rname ? rname[1] : '数据表' });
    } else if (/复制/.test(clause) && /到/.test(clause)) {
      var all = extractRanges(clause);
      if (all.length >= 2) { actions.push({ type: 'et.copyTo', range: all[0], target: all[1] }); }
    } else if (/序列|自动填充|编号|序号|流水号/.test(clause)) {
      var start = /从\s*(\d+)\s*开始/.exec(clause);
      var step = /步长\s*(\d+)/.exec(clause);
      actions.push({ type: 'et.fillSeries', range: range, start: start ? Number(start[1]) : 1, step: step ? Number(step[1]) : 1 });
    } else if (/(删除|删掉|去掉|去除|清除|清理|去重|剔除)/.test(clause) &&
               /(重复|去重|相同|不重复)/.test(clause)) {
      /* 必须放在「删除…行」之前，否则会被误当成删除整个区域 */
      var keyByCol = /(?:按|根据|以)\s*\$?([A-Z]{1,3})\s*列/.exec(clause) ||
        /(?:按|根据|以)\s*第\s*(\d+)\s*列/.exec(clause);
      var dedupAction = {
        type: 'et.deleteDuplicateRows', range: range,
        keep: /保留最后|最后一次|保留后/.test(clause) ? 'last' : 'first'
      };
      if (keyByCol) { dedupAction.keyColumn = keyByCol[1]; }
      actions.push(dedupAction);
    } else if (/删除.*行|删掉.*行/.test(clause)) {
      actions.push({ type: 'et.deleteRows', range: range });
    } else if (/清空|清除内容|清除格式|清除全部/.test(clause)) {
      actions.push({ type: 'et.clear', range: range, what: /格式/.test(clause) ? 'formats' : (/全部|所有/.test(clause) ? 'all' : 'contents') });
    } else if (/统计|摘要|分析一下|看一下|有多少|多少个|汇总情况/.test(clause)) {
      /* 要往单元格写统计公式时交给公式顾问，只有明确要看数据摘要才做只读统计 */
      var statsAdvice = formulaActions(clause, state.ctx, target);
      if (statsAdvice.length) {
        for (var si = 0; si < statsAdvice.length; si++) { actions.push(statsAdvice[si]); }
      } else {
        actions.push({ type: 'et.summary', range: range });
      }
    }

    /* 二、格式类动作可在同一句里叠加，不再被 if/else 只取第一个 */
    if (!actions.length) {
      var fmt = [];
      if (/加粗|粗体|取消加粗|BOLD/.test(c)) {
        fmt.push({ type: 'et.bold', range: range, value: !/取消|不要|去掉|不加/.test(clause) });
      }
      if (/斜体|ITALIC/.test(c)) {
        fmt.push({ type: 'et.italic', range: range, value: !/取消|不要|去掉/.test(clause) });
      }
      if (/字号|号字|字体大小/.test(clause)) {
        var size = /(?:字号|号字|字体大小)\s*[:：]?\s*(\d{1,2})/.exec(clause) || /(\d{1,2})\s*(?:号字|磅)/.exec(clause);
        fmt.push({ type: 'et.fontSize', range: range, size: size ? Number(size[1]) : 14 });
      }
      if (/字体|字型/.test(clause) && /[:：为成设置]/.test(clause)) {
        var fname = /字体\s*[:：为成设置]?\s*([\u4e00-\u9fa5A-Za-z]{2,12})/.exec(clause);
        if (fname) { fmt.push({ type: 'et.fontName', range: range, name: fname[1] }); }
      }
      if (/居中|左对齐|右对齐|对齐/.test(clause)) {
        var align = /右对齐/.test(clause) ? 'right' : (/左对齐/.test(clause) ? 'left' : 'center');
        fmt.push({ type: 'et.align', range: range, align: align });
      }
      if (/合并单元格|合并居中|合并\s*[^，。;；]{0,24}?\$?[A-Z]{1,3}\$?\d/.test(clause) &&
          !/合并文本|文本合并|合并到|拼接到|连接到/.test(clause)) {
        fmt.push({ type: 'et.mergeCells', range: range, center: /居中/.test(clause) });
      }
      if (/边框|框线/.test(clause)) { fmt.push({ type: 'et.border', range: range }); }
      if (/自适应|自动调整|自动列宽/.test(clause) && /列宽|列|宽/.test(clause)) {
        fmt.push({ type: 'et.autoFit', range: range });
      } else if (/列宽|行高/.test(clause) && !/自适应|自动调整/.test(clause)) {
        var width = /(?:设为|改成|调整到|宽度)\s*[:：]?\s*(\d{1,3})/.exec(clause) ||
          /(\d{1,3})\s*(?:字符|像素|磅)/.exec(clause);
        fmt.push({ type: 'et.columnWidth', range: range, width: width ? Number(width[1]) : 18 });
      }
      /* 同一句里只采用一种数字格式，避免两个动作互相覆盖 */
      var numFmt = '';
      if (!wantsRound && /百分比|百分号/.test(clause)) { numFmt = '0.00%'; }
      else if (!wantsRound && /保留\s*\d+\s*位小数|\d+\s*位小数/.test(clause)) {
        var digits = /(\d+)\s*位小数/.exec(clause);
        var n = digits ? Math.min(6, Number(digits[1])) : 2;
        numFmt = '0.' + new Array(n + 1).join('0');
      } else if (/货币格式|金额格式|人民币格式|¥\s*格式|设为货币|按货币/.test(clause)) { numFmt = '¥#,##0.00'; }
      else if (/千分位/.test(clause)) { numFmt = '#,##0'; }
      else if (/日期格式/.test(clause)) { numFmt = 'yyyy-mm-dd'; }
      if (numFmt) { fmt.push({ type: 'et.numberFormat', range: range, format: numFmt }); }
      if (/字体颜色|文字颜色|字色/.test(clause) && color) {
        fmt.push({ type: 'et.fontColor', range: range, color: color });
      }
      if (/底色|背景|填充|标黄|标红|标绿|高亮|涂色|颜色/.test(clause) && color &&
          !/空白填充|向下填充|填充公式|序列填充/.test(clause)) {
        fmt.push({ type: 'et.fillColor', range: range, color: color });
      }
      if (fmt.length) {
        for (var fi0 = 0; fi0 < fmt.length; fi0++) { actions.push(fmt[fi0]); }
      } else {
        /* 三、公式 / 值兜底：四舍五入这类指令走公式顾问，不再被数字格式抢走 */
        var value = quotedValue(clause);
        if (value !== null && /写入|填入|输入|填写|设为|改成|改为/.test(clause)) {
          actions.push({ type: 'et.setValue', range: range, value: value });
        } else {
          var fromAdvice = formulaActions(clause, state.ctx, target);
          for (var fi = 0; fi < fromAdvice.length; fi++) { actions.push(fromAdvice[fi]); }
        }
      }
    }

    if (actions.length) { state.currentRef = actions[actions.length - 1].range || range; }
    return actions;
  }

  /* 把 @selection 等占位符翻译成用户能核对的对象描述 */
  function readableRange(addr) {
    var raw = String(addr || '');
    if (!raw) { return ''; }
    var shown = raw;
    try { shown = String(Bridge.et.resolve(raw) || raw); } catch (e) { /* 保留原值 */ }
    shown = shown.replace(/@lastRowPlus1/gi, '末行下一行').replace(/@lastRow/gi, '末行');
    if (/@selectionBottom/i.test(raw)) { return '选区下方 ' + shown; }
    if (/@selection/i.test(raw)) { return '当前选区 ' + shown; }
    return shown;
  }

  function scopeOf(app, actions) {
    var seen = {}, list = [];
    for (var i = 0; i < actions.length; i++) {
      var a = actions[i];
      var label = readableRange(a.range || a.target || '');
      if (label && !seen[label]) { seen[label] = true; list.push(label); }
    }
    if (!list.length) { return '当前工作表（未写明区域，将使用当前选区）'; }
    return list.length > 2 ? list.slice(0, 2).join('、') + ' 等 ' + list.length + ' 处' : list.join('、');
  }

  /* 候选指令只描述当前选区 / 当前工作簿，不写死某张表、某列或某个阈值 */
  function clarifyOptions(text, ctx) {
    var sheet = (ctx && ctx.sheetName) || '';
    var out = [
      { label: '对当前选区求和', prompt: '对当前选中区域求和，把结果放在选区下方一行' },
      { label: '给当前选区加底色', prompt: '把当前选中区域加黄色底色' },
      { label: '统计当前选区', prompt: '统计一下当前选区的数据' },
      { label: '表头加粗居中', prompt: '把当前选中区域加粗并居中' }
    ];
    if (/占比|比例|百分比/.test(text)) {
      out.unshift({ label: '算当前选区占比', prompt: '计算当前选中区域每个值占该列合计的比例' });
    } else if (sheet && /工作表|跨表|另一张/.test(text)) {
      out.unshift({ label: '引用当前工作表', prompt: '在 ' + sheet + ' 的当前选中区域求和，结果放在选区下方' });
    }
    return out.slice(0, 4);
  }

  function rulePlan(text, ctx) {
    var raw = normalize(text);
    var clauses = splitClauses(raw);
    var app = detectApp(raw, ctx);
    var state = {
      currentRef: ctx.lastRange || '@selection',
      color: ctx.lastColor || '',
      topic: ctx.lastTopic || '',
      lastApp: ctx.lastApp || '',
      ctx: ctx
    };
    var actions = [];
    var basis = [];
    var unparsed = [];
    for (var i = 0; i < clauses.length; i++) {
      var list = parseEtClause(clauses[i], state);
      if (!list.length) { unparsed.push(clauses[i]); }
      for (var j = 0; j < list.length; j++) {
        actions.push(list[j]);
        basis.push(clauses[i]);
      }
    }
    var plan = {
      app: app,
      summary: '',
      actions: actions,
      source: 'rule',
      warnings: [],
      topic: state.topic,
      needConfirm: true
    };
    var reasons = [];
    if (unparsed.length) {
      var quoted = [];
      for (var u = 0; u < unparsed.length; u++) { quoted.push('「' + unparsed[u] + '」'); }
      reasons.push('有 ' + unparsed.length + ' 句没理解：' + quoted.join('、'));
    }
    var guessed = false;
    for (var g = 0; g < actions.length; g++) {
      var rr = String(actions[g].range || actions[g].target || '');
      if (/^@|@selection/.test(rr)) { guessed = true; break; }
    }
    if (guessed && !ctx.lastRange) { reasons.push('没有写明区域，按当前选区推断'); }
    if (guessed) {
      for (var q = 0; q < actions.length; q++) {
        if (actions[q].type === 'et.setValue') {
          reasons.push('写入内容是按字面推断的，请核对文本与位置');
          break;
        }
      }
    }
    if (actions.length > 4) { reasons.push('一次包含 ' + actions.length + ' 个动作'); }
    if (ctx.host && ctx.host !== 'browser' && ctx.host !== 'et') {
      reasons.push('当前环境不是表格');
      plan.warnings.push('当前环境不是表格，请打开一个工作簿再执行。');
    }
    if (ctx.mode === 'disconnected') {
      plan.warnings.push('当前未连接真实表格；执行操作请回到表格客户端，公式建议仍可正常查看。');
    }
    plan.intent = {
      confidence: reasons.length ? 'low' : 'high',
      reasons: reasons,
      scope: scopeOf(app, actions),
      host: '当前表格',
      basis: basis
    };
    plan.unparsed = unparsed;
    plan.clarify = clarifyOptions(raw, ctx);
    if (actions.length) {
      var labels = [];
      for (var k = 0; k < actions.length; k++) { labels.push(Actions.label(actions[k])); }
      plan.summary = '识别到 ' + actions.length + ' 个操作：' + labels.join('、');
      if (unparsed.length) { plan.summary += '（另有 ' + unparsed.length + ' 句未识别）'; }
    } else {
      plan.summary = '我没确定这句话要做什么：没有找到可识别的操作词或操作对象。可以从下面选一个接近的意图，或把「对象 + 动作 + 区域」写得更具体。';
      plan.help = true;
    }
    return plan;
  }

  function parse(text, ctx, done) {
    var settings = global.WpsAIChat.getSettings();
    if (global.WpsAILlm.enabled(settings)) {
      global.WpsAILlm.parse(text, ctx, settings, function (plan, error) {
        if (plan) {
          /* 模型说「需要确认」而规则引擎能确定时，优先给出可执行计划，别把用户卡在 0 步卡片上 */
          if (!plan.actions || !plan.actions.length) {
            var byRule = rulePlan(text, ctx);
            if (byRule.actions.length) {
              byRule.engineNote = '模型没能给出可执行动作，已用本地规则引擎生成计划';
              done(byRule);
              return;
            }
          }
          done(plan);
          return;
        }
        var fallback = rulePlan(text, ctx);
        /* 模型失败只是引擎降级，不算计划本身的风险，单独放在说明里而不是刺眼的警告条 */
        fallback.engineNote = '模型未返回结果（' + error + '），已自动用本地规则引擎生成计划';
        done(fallback);
      });
      return;
    }
    global.setTimeout(function () { done(rulePlan(text, ctx)); }, 80);
  }

  global.WpsAIPlanner = {
    parse: parse, rulePlan: rulePlan, detectApp: detectApp, normalize: normalize,
    clarifyOptions: clarifyOptions, scopeOf: scopeOf
  };
})(window);
