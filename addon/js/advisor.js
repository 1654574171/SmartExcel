/*
 * 公式顾问：理解指令 -> 整理重述意图 -> 给出公式建议。
 * 只读取环境信息，绝不调用任何写入接口；信息不足时列出反向追问。
 * 纯本地规则解析，不产生任何网络请求，也绝不写入表格。
 */
(function (global) {
  'use strict';

  var Bridge = global.WpsAIBridge;
  var Util = Bridge.util;

  var LABELS = {
    sum: '求和', average: '求平均值', count: '统计数值个数', counta: '统计非空单元格个数',
    max: '找最大值', min: '找最小值', ratio: '计算占整体的比例', growth: '计算增长率',
    rank: '计算排名', sumif: '按条件求和', countif: '按条件计数', averageif: '按条件求平均',
    if: '按条件判断并返回不同结果', round: '四舍五入到指定小数位',
    concat: '把多个单元格的文本拼接起来', vlookup: '按关键字从另一张表带出对应值',
    datedif: '计算两个日期相差的天数', conditionalFormat: '给满足条件的单元格设置条件格式',
    multiIf: '按多个区域的条件逐行判断并给出标记',
    dedup: '标记重复数据 / 提取唯一值',
    direct: '整理你给出的公式'
  };

  var GAP_TEXT = {
    region: '没说清对哪个区域的数据计算',
    target: '没说清结果写在哪里',
    cond: '没给出判断用的阈值或条件',
    denominator: '没说清占比的分母是什么',
    lookup: '没说清按哪一列分类 / 从哪张表查找',
    intent: '没看出明确的公式需求'
  };

  function normalize(text) {
    return String(text || '')
      .replace(/[，]/g, ',').replace(/[；]/g, ';').replace(/[：]/g, ':').replace(/[。]/g, '.')
      .replace(/[“”]/g, '"').replace(/[‘’]/g, "'")
      /* 中文与「表名!地址」之间留边界，避免把介词一起吞进工作表名 */
      .replace(/([\u4e00-\u9fa5])(?=(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d)/g, '$1 ')
      .replace(/\s+/g, ' ').trim();
  }

  function sheetInfo() {
    var i = {};
    try { i = Bridge.info() || {}; } catch (e) { /* 忽略 */ }
    return {
      host: i.host || 'browser', mode: i.mode || 'disconnected',
      sheetName: i.sheetName || '', selection: i.selection || 'A1',
      usedRange: i.usedRange || '', lastRow: Number(i.lastRow) || 12,
      lastCol: Number(i.lastCol) || 6
    };
  }

  function resolveAddr(addr, env) {
    try {
      return String(Bridge.et.resolve(addr))
        .replace(/@lastRowPlus1/gi, String(env.lastRow + 1))
        .replace(/@lastRow/gi, String(env.lastRow));
    } catch (e) { return String(addr); }
  }

  /* 引用可以是 A2:A6，也可以是 员工信息!B2:B6 / '员工 信息'!B2:B6 */
  function splitRef(ref) {
    var s = String(ref || '').replace(/\$/g, '');
    var m = /^('(?:[^']|'')+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!(.+)$/.exec(s);
    if (!m) { return { sheet: '', addr: s }; }
    return { sheet: m[1], addr: m[2] };
  }

  function addrOf(ref) { return splitRef(ref).addr.toUpperCase(); }

  function sheetPrefixOf(ref) {
    var part = splitRef(ref);
    return part.sheet ? part.sheet + '!' : '';
  }

  function quoteSheet(name) {
    var s = String(name || '').trim();
    if (/^'.*'$/.test(s)) { return s; }
    if (!s) { return ''; }
    return /^[A-Za-z_\u4e00-\u9fa5][A-Za-z0-9_.\u4e00-\u9fa5]*$/.test(s)
      ? s : "'" + s.replace(/'/g, "''") + "'";
  }

  function colOf(range) {
    var m = /^([A-Z]{1,3})/i.exec(addrOf(range || 'A'));
    return m ? m[1].toUpperCase() : 'A';
  }

  function boundsOf(range) {
    var m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(addrOf(range || ''));
    if (!m) { return { first: 2, last: 2 }; }
    var first = Number(m[2]) || 2;
    return { first: first, last: Number(m[4]) || first };
  }

  function nextCell(cell) {
    var part = splitRef(cell);
    var m = /^([A-Z]{1,3})(\d+)$/i.exec(part.addr);
    if (!m) { return 'B2'; }
    var out = Util.numToCol(Util.colToNum(m[1].toUpperCase()) + 1) + m[2];
    return part.sheet ? part.sheet + '!' + out : out;
  }

  /* 按原文出现顺序抽取区域；先抓带表名限定的完整引用并遮掉，避免表名里的字母数字被误判成单元格 */
  function extractRanges(text, env) {
    var found = [], m;
    var source = String(text || '');
    var masked = source;
    var qualified = /(?:'(?:[^']|'')+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?/gi;
    while ((m = qualified.exec(masked)) !== null) {
      var part = splitRef(m[0]);
      found.push({ ref: part.sheet + '!' + part.addr.toUpperCase(), at: m.index });
      masked = masked.slice(0, m.index) + m[0].replace(/[^\s]/g, ' ') + masked.slice(m.index + m[0].length);
    }
    var re = /\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?/g;
    while ((m = re.exec(masked)) !== null) {
      found.push({ ref: m[0].replace(/\$/g, '').toUpperCase(), at: m.index });
    }
    var colRe = /(^|[^A-Z0-9])([A-Z]{1,3})\s*列/g;
    while ((m = colRe.exec(masked)) !== null) {
      found.push({ ref: m[2].toUpperCase() + '2:' + m[2].toUpperCase() + env.lastRow, at: m.index });
    }
    found.sort(function (a, b) { return a.at - b.at; });
    var out = [];
    for (var i = 0; i < found.length; i++) {
      if (out.indexOf(found[i].ref) < 0) { out.push(found[i].ref); }
    }
    return out;
  }

  function opOf(word) {
    if (word === '>' || word === '大于' || word === '超过' || word === '高于') { return '>'; }
    if (word === '<' || word === '小于' || word === '低于' || word === '不足') { return '<'; }
    if (word === '>=' || word === '不小于' || word === '大于等于' || word === '≥') { return '>='; }
    if (word === '<=' || word === '不大于' || word === '小于等于' || word === '≤') { return '<='; }
    if (word === '!=' || word === '不等于') { return '<>'; }
    return '=';
  }

  function extractConditions(text) {
    var re = /(>=|<=|不等于|不小于|大于等于|不大于|小于等于|大于|超过|高于|小于|低于|不足|等于|≥|≤|!=|>|<)\s*(\d+(?:\.\d+)?)/g;
    var out = [], m;
    while ((m = re.exec(text)) !== null) {
      out.push({ op: opOf(m[1]), value: Number(m[2]), raw: m[0] });
    }
    return out;
  }

  function extractCondition(text) {
    var list = extractConditions(text);
    return list.length ? list[0] : null;
  }

  function quoted(text) {
    var m = /["']([^"']{1,40})["']/.exec(text);
    return m ? m[1] : '';
  }

  function directFormula(text) {
    /* 中文表名先换成占位符，避免把公式后面的中文说明一起吞进公式 */
    var names = [];
    var masked = String(text).replace(/(?:'[^']+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!/g, function (all) {
      names.push(all);
      return '\u0001' + (names.length - 1) + '\u0001';
    });
    var m = /=\s*[$A-Za-z0-9_(),:+*\/"'&<>=.\s!\u0001-]+/.exec(masked);
    if (!m) { return null; }
    var out = m[0].replace(/\s+/g, '');
    for (var i = 0; i < names.length; i++) {
      out = out.split('\u0001' + i + '\u0001').join(names[i]);
    }
    return out;
  }

  function detectKind(text) {
    if (directFormula(text)) { return 'direct'; }
    if (/占比|占比率|所占比例|比例/.test(text)) { return 'ratio'; }
    if (/增长率|同比|环比|增幅/.test(text)) { return 'growth'; }
    if (/排名|名次|RANK/i.test(text)) { return 'rank'; }
    if (/条件求和|按.+求和|分类汇总|SUMIF/i.test(text)) { return 'sumif'; }
    if (/条件平均|按.+平均|AVERAGEIF/i.test(text)) { return 'averageif'; }
    if (/条件计数|有几个|多少个|多少条|COUNTIF/i.test(text)) { return 'countif'; }
    if (/去重|删除重复|去除重复|重复行|不重复|唯一值|UNIQUE/i.test(text)) { return 'dedup'; }
    if (/条件格式/.test(text)) { return 'conditionalFormat'; }
    if (/标黄|标红|标绿|高亮|底色|填充色|涂色/.test(text)) { return 'conditionalFormat'; }
    if (/查找匹配|带出|匹配过来|VLOOKUP|LOOKUP/i.test(text)) { return 'vlookup'; }
    if (/相差.*天|日期差|间隔天数|DATEDIF/i.test(text)) { return 'datedif'; }
    if (/拼接|合并文本|连起来|CONCAT/i.test(text)) { return 'concat'; }
    if (/如果|条件判断|达标|不及格|IF函数/i.test(text)) { return 'if'; }
    if (/四舍五入|保留\s*\d+\s*位|\d+\s*位小数|ROUND/i.test(text)) { return 'round'; }
    if (/非空|COUNTA/i.test(text)) { return 'counta'; }
    if (/计数|个数|数量|COUNT/i.test(text)) { return 'count'; }
    if (/平均|均值|AVERAGE/i.test(text)) { return 'average'; }
    if (/最大|最大值|MAX/i.test(text)) { return 'max'; }
    if (/最小|最小值|MIN/i.test(text)) { return 'min'; }
    if (/求和|合计|总计|加总|汇总|加起来|SUM/i.test(text)) { return 'sum'; }
    return null;
  }

  function targetOf(text, range) {
    var m = /(?:放到|放在|写入|写到|填到|结果放在|输出到|存到)\s*((?:'[^']+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!)?\$?([A-Z]{1,3})\$?(\d{1,7})/i.exec(text);
    if (m) { return (m[1] || '') + m[2].toUpperCase() + m[3]; }
    if (/下方|下面|下一行|末行|底部/.test(text)) {
      return sheetPrefixOf(range) + colOf(range) + (boundsOf(range).last + 1);
    }
    return '';
  }

  function formulaItem(title, formula, why, tip) {
    return { title: title, formula: formula, why: why, tip: tip || '' };
  }

  function multiIfFormula(o) {
    var parts = [];
    for (var i = 0; i < o.conditions.length; i++) {
      var area = o.ranges[i + 1] || o.ranges[o.ranges.length - 1] || o.range;
      parts.push(sheetPrefixOf(area) + colOf(area) + boundsOf(area).first +
        o.conditions[i].op + o.conditions[i].value);
    }
    if (!parts.length) { return ''; }
    var expr = parts.length > 1
      ? (o.connector === 'or' ? 'OR' : 'AND') + '(' + parts.join(',') + ')'
      : parts[0];
    var yes = o.resultText || '是';
    var no = o.resultText ? '非' + o.resultText : '否';
    return '=IF(' + expr + ',"' + yes + '","' + no + '")';
  }

  function buildFormulas(kind, o) {
    var r = o.range;
    var p = sheetPrefixOf(r);
    var col = colOf(r);
    var b = boundsOf(r);
    var first = p + col + b.first;
    var lastRow = b.last;
    var absolute = p + '$' + col + '$' + b.first + ':$' + col + '$' + lastRow;
    /* 阈值只从用户指令里取，缺了就留空并在待补充里提示，不套任何默认值 */
    var cond = o.cond ? (o.cond.op + o.cond.value) : '';
    var second = p + col + (b.first + 1);
    var neighbour = nextCell(first);
    switch (kind) {
      case 'multiIf': return [formulaItem('逐行判断', multiIfFormula(o),
        '公式写在结果区域第一行再向下填充，条件列用相对引用会随行自动偏移。',
        '若条件之间是「或者」关系，把公式里的 AND 改成 OR。')];
      case 'direct':
        return [formulaItem('你给出的公式', o.direct,
          '按原样整理，可直接复制到目标单元格。', '若含相对引用，向下填充时会自动偏移。')];
      case 'sum': return [formulaItem('求和', '=SUM(' + r + ')',
        'SUM 只累加数值，文本和空单元格会被忽略。')];
      case 'average': return [formulaItem('平均值', '=AVERAGE(' + r + ')',
        'AVERAGE 只对数值求平均，空单元格不计入分母。')];
      case 'count': return [formulaItem('数值个数', '=COUNT(' + r + ')',
        'COUNT 只统计数值单元格；要算非空个数请改用 COUNTA。')];
      case 'counta': return [formulaItem('非空个数', '=COUNTA(' + r + ')',
        'COUNTA 统计所有非空单元格，包括文本。')];
      case 'max': return [formulaItem('最大值', '=MAX(' + r + ')', '返回区域内的最大数值。')];
      case 'min': return [formulaItem('最小值', '=MIN(' + r + ')', '返回区域内的最小数值。')];
      case 'dedup': return [
        formulaItem('标记重复行（筛选后自行删除）',
          '=IF(COUNTIF(' + p + '$' + col + '$' + b.first + ':' + p + col + b.first + ',' + first + ')=1,"保留","重复")',
          'COUNTIF 的区间从首行固定起点向下扩展：第一次出现返回「保留」，之后再出现就标记为「重复」。',
          '把公式粘贴到 ' + p + nextCell(first) + ' 并向下填充，筛选出「重复」的行再删除；不会自动改动原数据。'),
        formulaItem('直接提取唯一值', '=UNIQUE(' + r + ')',
          'UNIQUE 会溢出返回整列不重复值；较老版本可用「数据 → 删除重复项」，或改用上面的标记法。',
          '删除行属于破坏性操作，公式顾问只给公式，不替你删除。')
      ];
      case 'ratio': return [formulaItem('逐行占比', '=' + first + '/SUM(' + absolute + ')',
        '分母用绝对引用锁住整列合计，向下填充时不会漂移。',
        '把结果单元格设为「百分比」格式即可显示为 12.34%。')];
      case 'growth': return [formulaItem('增长率', '=(' + second + '-' + first + ')/' + first,
        '用本期减上期再除以上期；首行没有上期数据，公式从第二行开始填。',
        '结果设为百分比格式。')];
      case 'rank': return [formulaItem('排名', '=RANK(' + first + ',' + absolute + ',0)',
        '最后一个参数 0 表示降序；改成 1 为升序。')];
      case 'countif': return [formulaItem('条件计数', '=COUNTIF(' + r + ',"' + cond + '")',
        '条件要写在英文双引号里，例如 ">90"。')];
      case 'sumif': return [formulaItem('条件求和', '=SUMIF(' + o.condRange + ',"' + (o.condText || '条件') + '",' + o.sumRange + ')',
        '第一个参数是条件列，第三个参数是要求和的数值列。',
        '把 "条件" 换成实际分类名，或改成单元格引用如 A2。')];
      case 'averageif': return [formulaItem('条件平均', '=AVERAGEIF(' + o.condRange + ',"' + (o.condText || '条件') + '",' + o.sumRange + ')',
        '只对满足条件的行计算平均值。')];
      case 'if': {
        /* 返回值来自指令里的「是否为 X / 标记为 X / 标为 X」，不再写死达标文案 */
        var yesText = o.resultText || '满足条件';
        var noText = o.resultText ? '非' + o.resultText : '不满足条件';
        return [formulaItem('条件判断', '=IF(' + first + cond + ',"' + yesText + '","' + noText + '")',
          '条件与返回值都按你的指令生成；缺阈值时先把 ' + first + ' 后面的比较条件补全。')];
      }
      case 'round': return [formulaItem('四舍五入', '=ROUND(' + first + ',' + o.digits + ')',
        '第二个参数是保留的小数位数。')];
      case 'concat': return [formulaItem('文本拼接', '=' + first + '&" "&' + neighbour,
        '用 & 连接文本，中间加了空格分隔。')];
      case 'vlookup': {
        /* 查找表取当前工作簿里真实存在的另一张工作表，未取到才用占位名 */
        var lookupName = '';
        try {
          var names = Bridge.et.sheetNames ? Bridge.et.sheetNames() : [];
          for (var ni = 0; ni < names.length; ni++) {
            if (names[ni] !== o.env.sheetName) { lookupName = names[ni]; break; }
          }
        } catch (e) { /* 保持占位 */ }
        var lookupRef = quoteSheet(lookupName || '查找表') + '!$A:$B';
        return [formulaItem('查找匹配', '=VLOOKUP(' + first + ',' + lookupRef + ',2,0)',
          '在查找表首列找 ' + first + '，返回第 2 列；0 表示精确匹配。',
          '请把 ' + lookupRef + ' 换成实际的查找表区域。')];
      }
      case 'datedif': return [formulaItem('日期差', '=DATEDIF(' + first + ',' + neighbour + ',"D")',
        '"D" 返回天数，也可用 "M"（月）或 "Y"（年）。')];
      case 'conditionalFormat': return [formulaItem('条件格式公式', '=' + first + cond,
        '这不是写在单元格里的公式：在「条件格式 → 新建规则 → 使用公式」中填入，再设置填充色。',
        '公式相对所选区域的左上角单元格书写，应用到整个区域时会自动偏移。')];
      default: return [];
    }
  }

  function condText(cond) {
    if (!cond) { return ''; }
    var names = { '>': '大于', '<': '小于', '>=': '不小于', '<=': '不大于', '=': '等于', '<>': '不等于' };
    return (names[cond.op] || cond.op) + ' ' + cond.value;
  }

  function restate(kind, o) {
    if (kind === 'multiIf') {
      var desc = [];
      for (var mi = 0; mi < o.conditions.length; mi++) {
        var area = o.ranges[mi + 1] || o.ranges[o.ranges.length - 1] || o.range;
        desc.push(area + ' ' + condText(o.conditions[mi]));
      }
      var rp = sheetPrefixOf(o.ranges[0]);
      return '对 ' + o.ranges[0] + ' 逐行判断是否' + (o.resultText || '满足条件') +
        '，条件为 ' + desc.join(o.connector === 'or' ? ' 或 ' : ' 且 ') +
        '；公式写在 ' + rp + colOf(o.ranges[0]) + boundsOf(o.ranges[0]).first +
        ' 并向下填充到 ' + rp + colOf(o.ranges[0]) + boundsOf(o.ranges[0]).last + '。';
    }
    var action = LABELS[kind] || '计算';
    var s = '对 ' + o.range + ' ' + action;
    if (o.cond && kind !== 'sum' && kind !== 'average') { s += '（' + condText(o.cond) + '）'; }
    if (o.target) { s += '，把结果放在 ' + o.target; }
    else { s += '，结果位置未指定'; }
    s += '。';
    if (o.rangeGuessed) { s += ' 其中数据区域是按当前选区推断的，请在下面确认。'; }
    return s;
  }

  function questionsFor(text, kind, o, gaps) {
    var has = function (key) {
      for (var i = 0; i < gaps.length; i++) { if (gaps[i].key === key) { return true; } }
      return false;
    };
    var qs = [];
    if (has('region')) {
      qs.push({ label: '用当前选区 ' + o.range, prompt: text + ',数据区域是当前选区 ' + o.range });
      qs.push({ label: '用整列 ' + colOf(o.range), prompt: text + ',数据区域是 ' + colOf(o.range) + '2:' + colOf(o.range) + o.env.lastRow });
    }
    if (has('cond')) {
      qs.push({ label: '我来补充阈值', prompt: text + ',阈值是 ', fill: true });
      qs.push({ label: '条件用当前选区', prompt: text + ',条件区域是当前选区 ' + o.range });
    }
    if (has('denominator')) {
      qs.push({ label: '分母是整列合计', prompt: text + ',分母是整列合计' });
      qs.push({ label: '分母是固定单元格', prompt: text + ',分母是固定单元格' });
    }
    if (has('lookup')) {
      var c = colOf(o.range);
      var next = Util.numToCol(Util.colToNum(c) + 1);
      qs.push({ label: '条件列用 ' + c + ' 列', prompt: text + ',条件列是 ' + c + ' 列' });
      qs.push({ label: '数值列用 ' + next + ' 列', prompt: text + ',求和的数值列是 ' + next + ' 列' });
    }
    return qs.slice(0, 4);
  }

  function offTopic(text) {
    return {
      kind: 'unclear', confidence: 'unclear',
      intent: '这句话里我没看出明确的公式需求：' + text,
      points: [], formulas: [],
      gaps: [{ key: 'intent', text: GAP_TEXT.intent }],
      questions: [
        { label: '求和', prompt: '把当前选区求和,结果放在下方一行' },
        { label: '条件计数', prompt: '统计当前选区中大于指定阈值的有几个' },
        { label: '算占比', prompt: '计算当前选区每个值占该列合计的比例' },
        { label: 'VLOOKUP', prompt: '从另一张表按编号带出名称用什么公式' }
      ],
      notes: ['可以换个说法，或点上面的常见场景。', '这里只提供公式建议，不会读取或改动表格数据。']
    };
  }

  function advise(text, ctx) {
    var raw = normalize(text);
    var env = sheetInfo();

    var explicitRanges = extractRanges(raw, env);
    var ranges = [];
    for (var ri = 0; ri < explicitRanges.length; ri++) {
      if (ranges.indexOf(explicitRanges[ri]) < 0) { ranges.push(explicitRanges[ri]); }
    }
    var conditions = extractConditions(raw);

    /* 多个引用区域 + 判断条件：整句按「逐行判断」处理，优先于单公式关键词 */
    var multi = ranges.length >= 2 && conditions.length >= 1 &&
      /条件为|是否为|判断|满足|优质|且|并且|同时|或者/.test(raw);
    var kind = multi ? 'multiIf' : detectKind(raw);
    if (!kind) { return offTopic(raw); }

    var rangeGuessed = false;
    var range;
    if (ranges.length) {
      range = ranges[0];
    } else if (/它|刚才|同上|这个区域|同一区域/.test(raw) && ctx && ctx.lastRange) {
      range = String(ctx.lastRange);
    } else {
      range = resolveAddr('@selection', env) || env.selection;
      rangeGuessed = true;
    }
    var b = boundsOf(range);
    var col = colOf(range);
    var cond = extractCondition(raw);
    var target = targetOf(raw, range);
    var direct = directFormula(raw);
    var digits = /(\d+)\s*位小数|保留\s*(\d+)\s*位/.exec(raw);
    var condColMatch = /条件列是\s*([A-Z]{1,3})\s*列|按\s*([A-Z]{1,3})\s*列/.exec(raw);
    var sumColMatch = /(?:求和|数值)列是\s*([A-Z]{1,3})\s*列/.exec(raw);
    var condCol = condColMatch ? (condColMatch[1] || condColMatch[2]).toUpperCase() : col;
    var sumCol = sumColMatch ? sumColMatch[1].toUpperCase() : Util.numToCol(Util.colToNum(col) + 1);

    var labelMatch = /是否为([^，,。;；且并或\s]{1,12})/.exec(raw) ||
      /标记为([^，,。;；且并或\s]{1,12})/.exec(raw) ||
      /标为([^，,。;；且并或\s]{1,12})/.exec(raw);
    var resultText = labelMatch ? labelMatch[1].trim() : '';
    var connector = /或者|或/.test(raw) ? 'or' : 'and';

    var rangePrefix = sheetPrefixOf(range);
    var o = {
      range: range, env: env, cond: cond, target: target, rangeGuessed: rangeGuessed,
      direct: direct || '',
      digits: digits ? Math.min(6, Number(digits[1] || digits[2])) : 2,
      condRange: rangePrefix + condCol + b.first + ':' + condCol + b.last,
      sumRange: rangePrefix + sumCol + b.first + ':' + sumCol + b.last,
      condText: quoted(raw),
      ranges: ranges, conditions: conditions,
      connector: connector, resultText: resultText
    };

    var gaps = [];
    if (kind !== 'direct' && rangeGuessed && !/当前选区|选中区域|选区|这一列|这块/.test(raw)) {
      gaps.push({ key: 'region', text: GAP_TEXT.region });
    }
    if (!cond && (kind === 'countif' || kind === 'conditionalFormat' || kind === 'if')) {
      gaps.push({ key: 'cond', text: GAP_TEXT.cond });
    }
    if (kind === 'multiIf' && conditions.length < ranges.length - 1) {
      gaps.push({ key: 'cond', text: '还有引用的区域没给出判断条件' });
    }
    if (kind === 'ratio' && !/分母|整列合计|合计/.test(raw)) {
      gaps.push({ key: 'denominator', text: GAP_TEXT.denominator });
    }
    if ((kind === 'sumif' || kind === 'averageif' || kind === 'vlookup') &&
        !/条件列|按\s*[A-Z]{1,3}\s*列|查找表|Sheet/i.test(raw)) {
      gaps.push({ key: 'lookup', text: GAP_TEXT.lookup });
    }

    var points = [];
    if (kind === 'multiIf') {
      points.push({ label: '结果区域', value: ranges[0] });
      for (var pi = 0; pi < conditions.length; pi++) {
        var carea = ranges[pi + 1] || ranges[ranges.length - 1];
        points.push({ label: '条件 ' + (pi + 1), value: carea + ' ' + condText(conditions[pi]) });
      }
      points.push({ label: '满足时标记', value: resultText || '是' });
    } else {
      points.push({ label: '要做什么', value: LABELS[kind] });
      points.push({ label: '数据区域', value: range + (rangeGuessed ? '（按当前选区推断）' : '') });
      if (cond) { points.push({ label: '判断条件', value: condText(cond) }); }
      if (target) { points.push({ label: '结果位置', value: target }); }
      else { points.push({ label: '结果位置', value: '未指定' }); }
      if (kind === 'ratio') {
        points.push({ label: '占比分母', value: /固定单元格/.test(raw) ? '指定单元格' : '整列合计（默认假设）' });
      }
    }

    var notes = [
      '以上只是公式建议，没有读取单元格内容，也没有改动任何单元格。',
      '复制公式后粘贴到你希望显示结果的单元格即可' + (target ? '（你指定的是 ' + target + '）' : '') + '。',
      '公式中的区域和阈值请按实际数据调整；主流表格软件都支持这些函数。'
    ];
    if (kind === 'conditionalFormat') {
      notes.push('条件格式公式不是普通单元格公式，需要配合「条件格式」功能使用。');
    }
    if (kind === 'multiIf') {
      notes.push('公式按第一行书写，向下填充即可覆盖 ' + ranges[0] + '；条件之间是「' +
        (connector === 'or' ? '或' : '且') + '」关系。');
      notes.push('结果写在 ' + ranges[0] + ' 会替换该区域原有内容；要保留原数据，请粘贴到旁边的空列。');
    }
    var crossSheet = false;
    for (var ci = 0; ci < ranges.length; ci++) {
      if (sheetPrefixOf(ranges[ci])) { crossSheet = true; break; }
    }
    if (!crossSheet && /[A-Za-z0-9_'\u4e00-\u9fa5]+!\$?[A-Z]{1,3}\$?\d/.test(raw)) { crossSheet = true; }
    if (crossSheet) {
      notes.push('带「工作表名!」前缀的区域来自其它工作表；公式粘贴在当前工作表即可跨表引用它们，表名含空格时要用单引号包起来。');
    }

    return {
      kind: 'formula',
      intentKind: kind,
      confidence: gaps.length ? 'low' : 'high',
      intent: restate(kind, o),
      points: points,
      formulas: buildFormulas(kind, o),
      gaps: gaps,
      questions: questionsFor(raw, kind, o, gaps),
      notes: notes
    };
  }

  /* 每个标记都还原成带工作表限定的完整引用：绝不因为「同表」就省略表名，避免跨表时丢失归属 */
  function expandRangeTokens(text, ranges, baseSheet) {
    var out = String(text || '');
    var list = ranges || [];
    var base = baseSheet;
    if (base === undefined || base === null) {
      try { base = (Bridge.info() || {}).sheetName || ''; } catch (e) { base = ''; }
    }
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      if (!item || !item.addr) { continue; }
      var part = splitRef(String(item.addr).replace(/\s/g, ''));
      var ref;
      if (part.sheet) {
        ref = part.sheet + '!' + part.addr.toUpperCase();
      } else {
        var sheet = String(item.sheet || '').trim() || String(base || '').trim();
        ref = sheet ? quoteSheet(sheet) + '!' + part.addr.toUpperCase() : part.addr.toUpperCase();
      }
      /* 前后留空格，避免中文表名 / 地址与中文说明粘连后被解析成一个整体 */
      out = out.split('\u3014' + item.id + '\u3015').join(' ' + ref + ' ');
    }
    return out.replace(/\s+/g, ' ').trim();
  }

  global.WpsAIAdvisor = {
    advise: advise, detectKind: detectKind,
    expandRangeTokens: expandRangeTokens, quoteSheet: quoteSheet
  };
})(window);
