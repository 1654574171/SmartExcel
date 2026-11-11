/* 逻辑冒烟测试（npm test）：公式建议 + 执行操作双模式。
 * 不再使用浏览器虚拟表格；执行链路通过一个模拟真实表格 JSAPI 的宿主验证。 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
global.window = global;
global.localStorage = {
  _d: {},
  getItem(k) { return this._d[k] === undefined ? null : this._d[k]; },
  setItem(k, v) { this._d[k] = String(v); },
  removeItem(k) { delete this._d[k]; }
};

['chat.js', 'bridge.js', 'advisor.js', 'actions.js', 'planner.js', 'history.js', 'llm.js'].forEach((name) => {
  const file = path.join(ROOT, 'addon/js', name);
  vm.runInThisContext(fs.readFileSync(file, 'utf8'), { filename: name });
});

const Bridge = global.WpsAIBridge;
const Chat = global.WpsAIChat;
const Advisor = global.WpsAIAdvisor;
const Actions = global.WpsAIActions;
const Planner = global.WpsAIPlanner;
const History = global.WpsAIHistory;
const Llm = global.WpsAILlm;

let failed = 0;
function assert(ok, label) {
  if (ok) { console.log('PASS  ' + label); }
  else { failed++; console.error('FAIL  ' + label); }
}

function source(name) { return fs.readFileSync(path.join(ROOT, name), 'utf8'); }

/* ---------------- 1. 未连接真实表格时：只给建议，不伪装执行 ---------------- */
const offline = Bridge.detect(true);
assert(offline.mode === 'disconnected' && offline.host === 'browser', '无客户端时进入「未连接」状态');
assert(offline.docName === '' && offline.sheetName === '' && offline.selection === '',
  '未连接时不伪造工作簿 / 工作表 / 选区数据');
assert(typeof Bridge.getDemo === 'undefined' && typeof Bridge.et.demoPreview === 'undefined',
  '桥接层不再提供虚拟表格与演示预览');
assert(Actions.hostCheck().indexOf('未连接真实表格') >= 0, '执行前检查会阻止在未连接状态下写入');
assert(Bridge.et.selectRange('A1').ok === false, '未连接时定位引用会给出明确失败提示');
const offlineAdvice = Advisor.advise('把 C2:C20 求和放到 C21', { host: 'browser', mode: 'disconnected' });
assert(offlineAdvice.formulas[0].formula === '=SUM(C2:C20)', '未连接时公式建议仍可正常使用');

/* ---------------- 2. 模拟真实表格宿主 ---------------- */
function colToNum(s) {
  let n = 0;
  for (const ch of String(s).toUpperCase()) { n = n * 26 + (ch.charCodeAt(0) - 64); }
  return n;
}
function numToCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s || 'A';
}
function parseA1(addr) {
  const raw = String(addr || '').toUpperCase().replace(/\$/g, '').replace(/\s/g, '');
  const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(raw);
  if (!m) { return null; }
  const c1 = colToNum(m[1]), r1 = parseInt(m[2], 10);
  const c2 = m[3] ? colToNum(m[3]) : c1, r2 = m[4] ? parseInt(m[4], 10) : r1;
  return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
}

let app;
function makeSheet(name, rows) {
  const cells = {};
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      cells[numToCol(c + 1) + (r + 1)] = { value: rows[r][c], formula: '', style: {} };
    }
  }
  function cell(a) {
    if (!cells[a]) { cells[a] = { value: '', formula: '', style: {} }; }
    return cells[a];
  }
  const sheet = {
    Name: name, cells: cells,
    UsedRange: {
      Row: 1, Column: 1,
      Rows: { Count: rows.length },
      Columns: { Count: rows[0].length },
      Address: '$A$1:$' + numToCol(rows[0].length) + '$' + rows.length
    },
    Range(addr) {
      const raw = String(addr).replace(/\$/g, '');
      const p = parseA1(raw);
      const list = [];
      if (p) {
        for (let r = p.r1; r <= p.r2; r++) {
          for (let c = p.c1; c <= p.c2; c++) { list.push(numToCol(c) + r); }
        }
      } else { list.push(raw); }
      const cols = p ? p.c2 - p.c1 + 1 : 1;
      const range = { _addr: raw };
      Object.defineProperty(range, 'Value2', {
        get() {
          if (!p) { return cell(list[0]).value; }
          const out = [];
          for (let r = p.r1; r <= p.r2; r++) {
            const line = [];
            for (let c = p.c1; c <= p.c2; c++) { line.push(cell(numToCol(c) + r).value); }
            out.push(line);
          }
          return out;
        },
        set(v) {
          const matrix = Array.isArray(v) ? (Array.isArray(v[0]) ? v : [v]) : [[v]];
          let i = 0;
          for (const a of list) {
            const ri = Math.min(Math.floor(i / cols), matrix.length - 1);
            const ci = Math.min(i % cols, matrix[ri].length - 1);
            cell(a).value = matrix[ri][ci];
            cell(a).formula = '';
            i++;
          }
        }
      });
      Object.defineProperty(range, 'Formula', {
        get() {
          if (!p) { return cell(list[0]).formula || cell(list[0]).value; }
          const out = [];
          for (let r = p.r1; r <= p.r2; r++) {
            const line = [];
            for (let c = p.c1; c <= p.c2; c++) {
              const cl = cell(numToCol(c) + r);
              line.push(cl.formula || cl.value);
            }
            out.push(line);
          }
          return out;
        },
        set(f) { for (const a of list) { cell(a).formula = f; cell(a).value = f; } }
      });
      function styleProp(key, field) {
        Object.defineProperty(range, key, {
          get() { const v = cell(list[0]).style[field]; return v === undefined ? null : v; },
          set(v) { for (const a of list) { cell(a).style[field] = v; } }
        });
      }
      styleProp('NumberFormat', 'numberFormat');
      styleProp('HorizontalAlignment', 'align');
      styleProp('ColumnWidth', 'width');
      Object.defineProperty(range, 'Interior', {
        get() {
          const o = {};
          Object.defineProperty(o, 'Color', {
            get() { return cell(list[0]).style.color === undefined ? '' : cell(list[0]).style.color; },
            set(v) { for (const a of list) { cell(a).style.color = v; } }
          });
          return o;
        }
      });
      Object.defineProperty(range, 'Font', {
        get() {
          const o = {};
          ['Bold', 'Italic', 'Size', 'Name', 'Color'].forEach((k) => {
            Object.defineProperty(o, k, {
              get() { const v = cell(list[0]).style['font' + k]; return v === undefined ? null : v; },
              set(v) { for (const a of list) { cell(a).style['font' + k] = v; } }
            });
          });
          return o;
        }
      });
      range.Borders = { LineStyle: 0, Weight: 0 };
      range.Columns = { Count: cols, Item() { return range; }, AutoFit() {} };
      range.Rows = { Count: p ? p.r2 - p.r1 + 1 : 1 };
      range.Select = function () { app.Selection = { Address: raw }; };
      range.ClearContents = function () { for (const a of list) { cell(a).value = ''; cell(a).formula = ''; } };
      range.ClearFormats = function () { for (const a of list) { cell(a).style = {}; } };
      range.Clear = function () { range.ClearContents(); range.ClearFormats(); };
      range.Copy = function (dest) { dest.Value2 = range.Value2; };
      range.EntireRow = { Delete() {} };
      range.Sort = function () {};
      range.AutoFilter = function () {};
      return range;
    },
    Activate() { app.ActiveSheet = sheet; },
    ChartObjects() { return { Add() { return { Chart: { SetSourceData() {} } }; } }; }
  };
  return sheet;
}

const sheets = [
  makeSheet('演示销售数据', [
    ['月份', '产品', '销量', '单价', '销售额', '完成率'],
    ['1月', '笔记本', 80, 199, 15920, 0.6], ['2月', '显示器', 120, 299, 35880, 0.7],
    ['3月', '键盘', 90, 399, 35910, 0.8], ['4月', '笔记本', 110, 199, 21890, 0.9],
    ['5月', '显示器', 130, 299, 38870, 0.7], ['6月', '键盘', 100, 399, 39900, 0.8],
    ['7月', '笔记本', 140, 199, 27860, 0.6], ['8月', '显示器', 150, 299, 44850, 0.7],
    ['9月', '键盘', 160, 399, 63840, 0.9], ['10月', '笔记本', 170, 199, 33830, 0.8],
    ['11月', '显示器', 180, 299, 53820, 0.9]
  ]),
  makeSheet('员工信息', [
    ['姓名', '年龄', '月薪', '部门'],
    ['张三', 26, 4250, '技术部'], ['李四', 33, 6800, '市场部'],
    ['王五', 28, 3900, '技术部'], ['赵六', 41, 8600, '市场部'], ['钱七', 24, 2800, '技术部']
  ]),
  makeSheet('参数表', [['条件', '阈值'], ['年龄上限', 30], ['月薪下限', 3000]])
];
const worksheets = {
  Count: sheets.length,
  Item(k) { return sheets[k - 1]; },
  Add() { return sheets[0]; }
};
app = {
  ActiveWorkbook: { Name: '测试工作簿.xlsx' },
  ActiveSheet: sheets[0],
  Worksheets: worksheets,
  Selection: { Address: '$B$2:$B$11' },
  ActiveWindow: { FreezePanes: false, SplitRow: 0, SplitColumn: 0 }
};
global.wps = { Application: app };

const live = Bridge.detect(true);
assert(live.mode === 'live' && live.host === 'et', '检测到真实表格宿主后进入已连接状态');
assert(live.sheetName === '演示销售数据' && live.selection === '$B$2:$B$11', '可读取当前工作表与选区');
assert(Actions.hostCheck() === null, '已连接时执行前检查通过');

const adviseCtx = {
  lastRange: '', selection: 'B2:B11', selectionForAI: 'B2:B11',
  host: 'et', mode: 'live'
};

/* ---------------- 3. 公式建议能力（全部保留） ---------------- */
const sum = Advisor.advise('把 C2:C20 求和放到 C21', adviseCtx);
assert(sum.intentKind === 'sum' && sum.formulas[0].formula === '=SUM(C2:C20)', '求和公式正确');
assert(sum.intent.indexOf('C2:C20') >= 0 && sum.points.some((p) => p.label === '结果位置' && p.value === 'C21'),
  '重述包含数据区域与结果位置');

const ratio = Advisor.advise('计算 F 列的占比，用百分比格式', adviseCtx);
assert(ratio.formulas[0].formula.indexOf('/SUM($F$') >= 0, '占比分母使用绝对引用');

const countif = Advisor.advise('统计当前选区中大于 90 的有几个', adviseCtx);
assert(countif.formulas[0].formula.indexOf('COUNTIF') >= 0 && countif.formulas[0].formula.indexOf('>90') >= 0,
  '条件计数公式带上阈值');

const direct = Advisor.advise('=SUM(B2:B11)', adviseCtx);
assert(direct.formulas[0].formula === '=SUM(B2:B11)', '用户直接给出的公式原样整理');

const unclear = Advisor.advise('随便弄一下', adviseCtx);
assert(unclear.kind === 'unclear' && unclear.formulas.length === 0 && unclear.questions.length > 0,
  '描述不清时反向追问而不是硬给公式');

const dedup = Advisor.advise('删除 A2:A7 重复行', adviseCtx);
assert(dedup.intentKind === 'dedup' && dedup.formulas.length === 2, '删除重复行给出去重公式');
assert(dedup.formulas[0].formula.indexOf('COUNTIF($A$2:A2,A2)=1') >= 0, '给出逐行标记重复的公式');
assert(dedup.formulas[1].formula === '=UNIQUE(A2:A7)', '同时给出提取唯一值的公式');
assert(dedup.formulas[1].tip.indexOf('破坏性操作') >= 0, '明确说明不会替用户删除行');

const marked = Advisor.advise('Sheet1!A2:A6 标为优质嘉宾,条件为 Sheet2!B2:B6 大于 3500', adviseCtx);
assert(marked.formulas[0].formula.indexOf('优质嘉宾') >= 0 && marked.formulas[0].formula.indexOf('Sheet2!B2>3500') >= 0,
  '「标为 X」的结果文字与跨表条件都进入公式');

/* ---------------- 4. 引用展开：每个区域都强制带 sheet ---------------- */
const refs = [
  { id: 1, addr: 'A2:A6', sheet: 'Sheet1' },
  { id: 2, addr: 'B2:B6', sheet: 'Sheet2' },
  { id: 3, addr: 'C2:C6', sheet: 'Sheet1' }
];
const expanded = Advisor.expandRangeTokens(
  '〔1〕此处计算是否为优质男生,条件为 〔2〕小于30,且 〔3〕大于3000', refs, 'Sheet1');
assert(expanded.indexOf('Sheet1!A2:A6') === 0 && expanded.indexOf('Sheet2!B2:B6') >= 0 &&
  expanded.indexOf('Sheet1!C2:C6') >= 0, '每个引用都带工作表名，同表引用也不省略');
assert(expanded.split('Sheet1!A2:A6').join('').indexOf('A2:A6') < 0 &&
  expanded.split('Sheet2!B2:B6').join('').indexOf('B2:B6') < 0, '表达式里不存在裸区域地址');

const cross = Advisor.advise(expanded, adviseCtx);
assert(cross.intentKind === 'multiIf', '跨表多区域识别为逐行判断');
assert(cross.formulas[0].formula.indexOf('Sheet2!B2<30') >= 0 &&
  cross.formulas[0].formula.indexOf('Sheet1!C2>3000') >= 0 &&
  cross.formulas[0].formula.indexOf('优质男生') >= 0, '跨表条件与结果文字都带表名限定');
assert(cross.notes.join('').indexOf('替换该区域原有内容') >= 0, '结果覆盖原数据时给出提示');

const glued = Advisor.advise(
  '在此处Sheet2!C2:C6中计算Sheet1!C2:C6大于3500,且Sheet1!D2:D6为男,标为优质嘉宾', adviseCtx);
assert(glued.intent.indexOf('Sheet2!C2:C6') >= 0 && glued.intent.indexOf('在此处Sheet2') < 0,
  '中文紧邻表名时仍能正确解析工作表');

const quoted = Advisor.expandRangeTokens('〔1〕小于30',
  [{ id: 1, addr: 'B2:B6', sheet: '员工 信息' }], 'Sheet1');
assert(quoted === "'员工 信息'!B2:B6 小于30", '工作表名含空格时自动加单引号');

/* ---------------- 5. 执行操作链路 ---------------- */
assert(Actions.listTypes().every((t) => t.indexOf('et.') === 0), '动作白名单仅包含表格动作');
const docPlan = Planner.rulePlan('写一封请假条并插入文档', adviseCtx);
assert(docPlan.actions.every((a) => String(a.type).indexOf('wps.') !== 0), '文字撰写类指令不再产生文档动作');

const writePlan = Actions.validate(Planner.rulePlan('把 C1 写入 123', adviseCtx));
assert(writePlan.actions.length === 1 && writePlan.actions[0].type === 'et.setValue', '规则引擎生成写入值的计划');
const originalC1 = sheets[0].cells.C1.value;
const before = History.capture(writePlan);
const result = Actions.execute(writePlan);
assert(result.ok && result.succeeded === 1, '执行计划成功写入真实宿主');
assert(String(sheets[0].cells.C1.value) === '123', '单元格内容确实被写入');
const tx = History.commit(writePlan, result, before);
assert(tx && tx.undoable, '执行后生成可撤回的事务快照');
assert(History.undo(tx.id).ok && String(sheets[0].cells.C1.value) === String(originalC1), '撤回可恢复单元格原值');
assert(History.redo(tx.id).ok && String(sheets[0].cells.C1.value) === '123', '重做可恢复执行结果');

const boldPlan = Actions.validate(Planner.rulePlan('把 A1:A3 加粗', adviseCtx));
assert(boldPlan.actions[0].type === 'et.bold', '格式类指令也能生成计划');
assert(Actions.execute(boldPlan).ok, '格式类动作可执行');
assert(sheets[0].cells.A1.style.fontBold === true, '加粗写入宿主字体属性');

const formulaPlan = Actions.validate({
  summary: '跨表写入公式',
  actions: [{ type: 'et.setFormula', range: '员工信息!D1', formula: '=SUM(员工信息!C2:C6)' }]
});
assert(Actions.execute(formulaPlan).ok, '带工作表名限定的目标区域可执行');
assert(sheets[1].cells.D1.formula === '=SUM(员工信息!C2:C6)', '跨表目标地址写入到正确工作表');

/* 回归：跨表条件高亮曾把底色写到 A 列的对应行，而不是目标工作表的命中单元格 */
const highlightPlan = Actions.validate({
  summary: '跨表条件高亮',
  actions: [{ type: 'et.conditionalFill', range: '员工信息!C2:C6', op: '>', threshold: 5000, color: '#FFFF00' }]
});
const highlightResult = Actions.execute(highlightPlan);
assert(highlightResult.ok && highlightResult.succeeded === 1 &&
  highlightResult.details[0].text.indexOf('2 个') >= 0, '跨表条件高亮执行成功并命中 2 个单元格');
assert(sheets[1].cells.C3.style.color === '#FFFF00' && sheets[1].cells.C5.style.color === '#FFFF00',
  '底色落在目标工作表 C 列的命中行，不再是 A 列偏移位置');
assert(!sheets[1].cells.A2.style.color && !sheets[1].cells.A4.style.color,
  '不会把颜色错写到 A 列对应行');
assert(!sheets[0].cells.C3.style.color && !sheets[0].cells.C5.style.color,
  '不会把颜色错写到当前活动工作表');
const parsed = Bridge.util.parseA1('Sheet1!C2:C6');
assert(parsed && parsed.c1 === 3 && parsed.r1 === 2, '坐标解析能剥离工作表前缀');
assert(Bridge.et.resolve('Sheet1!C2:C6') === 'Sheet1!C2:C6', '地址规范化保留工作表名原始大小写');

const sheetRule = Planner.rulePlan('把 Sheet1!C2:C6 中大于 3000 的单元格标黄', adviseCtx);
assert(sheetRule.actions.some((a) => a.range === 'Sheet1!C2:C6'),
  '规则引擎也保留跨表引用的工作表名');
const sheetSum = Planner.rulePlan('对 Sheet1!C2:C6 求和，结果放在下面', adviseCtx);
assert(sheetSum.actions[0].range === 'Sheet1!C7' &&
  sheetSum.actions[0].formula === '=SUM(Sheet1!C2:C6)', '跨表求和的结果位置与公式都带工作表名');

/* ---------------- 6. 跨表定位 ---------------- */
assert(Bridge.et.selectionRef().sheet === '演示销售数据', '选区引用记录了所属工作表');
assert(Bridge.et.selectionAddress() === 'B2:B11', '纯地址接口保持兼容');
const sheetNames = Bridge.et.sheetNames();
assert(sheetNames.indexOf('员工信息') >= 0 && sheetNames.indexOf('参数表') >= 0, '可枚举多张工作表');
const crossLocate = Bridge.et.selectRange('B2:B6', '员工信息');
assert(crossLocate.ok && Bridge.info().sheetName === '员工信息', '定位跨表引用会自动激活对应工作表');
assert(String(app.Selection.Address).toUpperCase() === 'B2:B6', '定位后真实选中目标区域');
assert(Bridge.et.activateSheet('演示销售数据').ok && Bridge.info().sheetName === '演示销售数据',
  '可以切回原来的工作表');

/* ---------------- 7. 会话导出与产品收敛 ---------------- */
const session = Chat.getActiveSession();
Chat.addMessage(session.id, { role: 'user', text: '导出测试指令' });
Chat.addMessage(session.id, { role: 'assistant', text: '公式建议', advice: sum });
Chat.addMessage(session.id, {
  role: 'assistant', text: '计划', plan: { summary: '测试计划', actions: [{ type: 'et.bold', range: 'A1' }] }
});
Chat.addMessage(session.id, {
  role: 'assistant', text: '结果',
  result: { ok: true, succeeded: 1, failed: 0, details: [{ label: '加粗', ok: true, text: '完成' }] }
});
const md = Chat.exportSession(session.id);
assert(md.indexOf('**公式建议**') >= 0 && md.indexOf('=SUM(C2:C20)') >= 0, '会话导出包含公式建议');
assert(md.indexOf('**操作计划**') >= 0 && md.indexOf('**执行结果**') >= 0, '会话导出恢复操作计划与执行结果');
assert(Chat.getContext(session.id, Bridge.info()).history.some((h) => h.content.indexOf('【操作计划】') >= 0),
  '操作计划摘要进入多轮上下文');

const html = source('addon/ui/taskpane.html');
assert(/actions\.js/.test(html) && /planner\.js/.test(html) && /history\.js/.test(html) && /llm\.js/.test(html),
  '侧边栏重新加载执行操作链路');
assert(/modeBadge/.test(html) && /btnHistory/.test(html) && /settingExecMode/.test(html),
  '页面提供模式切换与操作历史入口');
assert(!/quickbar|quickChips|bridge-word\.js|draft\.js/.test(html), '「试试这样问」与文字文档脚本仍被移除');
assert(!/btnDemoReset|demoPreview|查看演示数据|重置演示数据/.test(html), '页面不含浏览器演示 / 虚拟表格入口');
assert(!/\bWPS\b/i.test(html), '页面不出现 WPS 可见字样');

const appSource = source('addon/js/app.js');
assert(/WpsAIActions/.test(appSource) && /WpsAIPlanner/.test(appSource) &&
  /WpsAIHistory/.test(appSource) && /WpsAILlm/.test(appSource), '侧边栏代码引用执行模块');
assert(/run-plan/.test(appSource) && /undo-tx/.test(appSource) && /setMode/.test(appSource),
  '侧边栏保留执行按钮、撤回重做与模式切换');
const locateStart = appSource.indexOf("act === 'locate'");
const locateEnd = appSource.indexOf("el('rangeBar').addEventListener('input'");
const locateBlock = appSource.slice(locateStart, locateEnd);
assert(locateStart > 0 && locateEnd > locateStart && locateBlock.indexOf('renderMessages()') < 0 &&
  locateBlock.indexOf('updatePreview') < 0, '定位引用时不会重建会话 DOM，也不依赖演示预览');
const bridgeSource = source('addon/js/bridge.js');
assert(bridgeSource.indexOf('getDemo') < 0 && bridgeSource.indexOf('initDemo') < 0 &&
  bridgeSource.indexOf('demoPreview') < 0, '桥接层已彻底移除虚拟表格');

/* ---------------- 8. 大模型结构化动作（仅表格） ---------------- */
let sentBodies = [];
let queuedReplies = [];
function modelContent(obj) { return typeof obj === 'string' ? obj : JSON.stringify(obj); }
function setModelReplies(list) { queuedReplies = list.slice(); sentBodies = []; }
global.XMLHttpRequest = function () {
  this.open = function () {};
  this.setRequestHeader = function () {};
  this.send = function (body) {
    sentBodies.push(body);
    const content = queuedReplies.length ? queuedReplies.shift() : '{}';
    this.readyState = 4;
    this.status = 200;
    this.responseText = JSON.stringify({ choices: [{ message: { content: content } }] });
    if (this.onreadystatechange) { this.onreadystatechange(); }
  };
};

const structCtx = {
  history: [{ role: 'user', content: '之前算过优质男生' }, { role: 'assistant', content: '一大段公式' }],
  sheetNames: ['演示销售数据', '员工信息', '参数表'],
  sheetName: '演示销售数据', docName: '测试工作簿.xlsx',
  usedRange: '$A$1:$F$12', selection: 'B2:B11', lastRow: 12
};
const fullPrompt = Llm.systemPrompt(structCtx, false);
assert(fullPrompt.indexOf('员工信息') >= 0 && fullPrompt.indexOf('演示销售数据') >= 0 &&
  fullPrompt.indexOf('$A$1:$F$12') >= 0 && fullPrompt.indexOf('B2:B11') >= 0,
  '提示词注入真实工作表名 / 选区 / 已用区域');
assert(fullPrompt.indexOf('不要臆造 Sheet1') >= 0 && fullPrompt.indexOf('不得照抄') >= 0,
  '提示词禁止臆造表名与照抄示例数值');
assert(fullPrompt.indexOf('wps.') < 0 && fullPrompt.indexOf('"app":"et"') >= 0,
  '模型提示词只含表格动作 schema');

Chat.saveSettings({ llmEnabled: true, apiKey: 'test-key', baseUrl: 'https://example.test/v1' });
assert(Llm.enabled(Chat.getSettings()), '大模型通道可按设置启用');
assert(Llm.enabled({ llmEnabled: true, baseUrl: 'http://127.0.0.1:11434/v1' }),
  '本地模型无需 API Key 也能启用');

setModelReplies([modelContent({
  app: 'et', summary: '模型计划', actions: [{ type: 'et.bold', range: 'A1:A3', value: true }]
})]);
let llmPlan = null;
Planner.parse('把 A1:A3 加粗', structCtx, function (plan) { llmPlan = plan; });
assert(llmPlan && llmPlan.source === 'llm' && llmPlan.actions.length === 1,
  '大模型可生成结构化表格计划');
const sentMessages = sentBodies.length ? JSON.parse(sentBodies[0]).messages : [];
assert(sentMessages.filter((m) => m.content.indexOf('历史指令') >= 0).length === 1,
  '只带最近的用户指令，不带助手长文本污染计划');

/* 输出归一化：解释文字 / 代码块 / 嵌套 plan / 类型别名 / 裸动作数组 */
const nested = Llm.extractPlanObject('模型说明\n```json\n' + JSON.stringify({
  plan: { summary: '嵌套计划', actions: [{ type: 'setFormula', range: 'B1', formula: '=SUM(A1:A10)' }] }
}) + '\n```\n结束');
const normalized = Llm.normalizePlan(nested);
assert(normalized.actions.length === 1 && normalized.actions[0].type === 'et.setFormula',
  '嵌套 plan 与类型别名自动归一化到白名单');
assert(Actions.validate(normalized).actions.length === 1, '归一化后的动作能通过白名单校验');
const bare = Llm.extractPlanObject(JSON.stringify([{ type: 'et.bold', range: 'A1', value: true }]));
assert(Llm.normalizePlan(bare).actions.length === 1, '模型直接输出裸动作数组也能识别');
assert(Llm.normalizePlan(Llm.extractPlanObject(JSON.stringify({
  summary: '需要确认', actions: [], need_clarify: true
}))).needClarify === true, '兼容 need_clarify 字段，不再误判成空结果');

/* 空动作自动重试一次，重试提示更严格 */
setModelReplies([
  modelContent({ app: 'et', summary: '没看懂', actions: [] }),
  modelContent({ app: 'et', summary: '重试成功', actions: [{ type: 'et.bold', range: 'A1', value: true }] })
]);
let retryPlan = null;
Planner.parse('把 A1 加粗', structCtx, function (plan) { retryPlan = plan; });
assert(retryPlan && retryPlan.source === 'llm' && retryPlan.actions.length === 1,
  '模型返回空动作时自动重试并生成可用计划');
assert(sentBodies.length === 2, '重试只额外发出一次请求');
assert(JSON.parse(sentBodies[1]).messages[0].content.indexOf('重试要求') >= 0,
  '重试使用更严格的提示词');

/* 两次都失败时给出具体原因，并降级为规则引擎但不冒充风险警告 */
setModelReplies([
  modelContent({ app: 'et', summary: '缺少明确区域', actions: [] }),
  modelContent({ app: 'et', summary: '仍然缺少区域', actions: [] })
]);
let fallbackPlan = null;
Planner.parse('随便弄一下', structCtx, function (plan) { fallbackPlan = plan; });
assert(fallbackPlan && fallbackPlan.engineNote.indexOf('仍然缺少区域') >= 0,
  '失败原因带上模型自己的说明，便于定位');
assert(fallbackPlan.engineNote.indexOf('本地规则引擎') >= 0 &&
  (fallbackPlan.warnings || []).join('').indexOf('大模型不可用') < 0,
  '引擎降级用中性说明，不再显示刺眼的风险警告');
Chat.saveSettings({ llmEnabled: false });

/* ---------------- 8.1 通用化：不套固定模板 ---------------- */
const generic = Planner.rulePlan('如果 Sheet1!C2 大于 3000，标为优质，结果放在 Sheet1!D2', adviseCtx);
assert(generic.actions.length === 1 && generic.actions[0].type === 'et.setFormula',
  '条件判断类指令生成写入公式动作');
const genericFormula = generic.actions.length ? generic.actions[0].formula : '';
assert(genericFormula.indexOf('>3000') >= 0 && genericFormula.indexOf('优质') >= 0,
  '公式里的阈值和结果文字都来自当前指令');
assert(genericFormula.indexOf('>=60') < 0 && genericFormula.indexOf('达标') < 0,
  '不再出现写死的 >=60 / 达标 模板');
assert(generic.actions[0].range.indexOf('Sheet1!') >= 0, '目标单元格保留工作表名');
const statPlan = Planner.rulePlan('统计 Sheet1!C2:C6 中大于 3000 的有几个', adviseCtx);
assert(statPlan.actions[0].type === 'et.setFormula' &&
  statPlan.actions[0].formula.indexOf('COUNTIF') >= 0 && statPlan.actions[0].formula.indexOf('>3000') >= 0,
  '统计数量的简单指令写出带真实阈值的计数公式，而不是只给只读摘要');
const summaryPlan = Planner.rulePlan('统计一下当前选区的数据', adviseCtx);
assert(summaryPlan.actions[0].type === 'et.summary', '明确要看数据摘要时仍做只读统计');
const plannerSrc = source('addon/js/planner.js');
assert(plannerSrc.indexOf('高亮大于 90') < 0 && plannerSrc.indexOf('计算 F 列占比') < 0 &&
  plannerSrc.indexOf('按 A 列分类') < 0, '澄清候选不再写死 90 / F 列 / A 列');
const advisorSrc = source('addon/js/advisor.js');
assert(advisorSrc.indexOf('"达标","未达标"') < 0 && advisorSrc.indexOf("'>60'") < 0,
  '公式顾问不再写死达标文案与 60 阈值');

/* ---------------- 9. 规则引擎异步分支 ---------------- */
setTimeout(function () {
  Planner.parse('把 A1:A3 加粗', adviseCtx, function (plan) {
    assert(plan && plan.source === 'rule' && plan.actions.length === 1,
      '未启用大模型时自动走本地规则引擎');
    console.log(failed ? '\n冒烟测试失败：' + failed + ' 项' : '\n全部冒烟测试通过');
    process.exit(failed ? 1 : 0);
  });
}, 30);
