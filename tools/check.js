/* 静态自检：运行产物的 JS 语法、页面资源引用、仅表格注册、执行链路完整、无演示模式 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const errors = [];
let checked = 0;

function read(relative) { return fs.readFileSync(path.join(ROOT, relative), 'utf8'); }

function checkSyntax(relative) {
  const file = path.join(ROOT, relative);
  if (!fs.existsSync(file)) { errors.push('缺少文件：' + relative); return; }
  try {
    new vm.Script(fs.readFileSync(file, 'utf8'), { filename: relative });
    checked++;
  } catch (e) {
    errors.push(relative + ' 语法错误：' + e.message);
  }
}

/* 公式建议与执行操作两条链路都属于运行产物 */
[
  'addon/main.js',
  'addon/js/main.js',
  'addon/js/chat.js',
  'addon/js/bridge.js',
  'addon/js/advisor.js',
  'addon/js/actions.js',
  'addon/js/planner.js',
  'addon/js/history.js',
  'addon/js/llm.js',
  'addon/js/app.js',
  'tools/serve.js'
].forEach(checkSyntax);

['addon/ribbon.xml', 'publish.xml', 'addon/ui/taskpane.html', 'addon/css/app.css', 'package.json'].forEach((f) => {
  const file = path.join(ROOT, f);
  if (!fs.existsSync(file)) { errors.push('缺少文件：' + f); return; }
  checked++;
});

const html = read('addon/ui/taskpane.html');
const refs = [];
const re = /(?:src|href)\s*=\s*"([^"]+)"/g;
let m;
while ((m = re.exec(html)) !== null) {
  if (!/^https?:|^data:/.test(m[1])) { refs.push(m[1]); }
}
refs.forEach((ref) => {
  const target = path.resolve(path.join(ROOT, 'addon/ui'), ref);
  if (!fs.existsSync(target)) { errors.push('页面引用了不存在的资源：' + ref); }
  else { checked++; }
});

if (/quickbar|quickChips|bridge-word\.js|draft\.js/.test(html)) {
  errors.push('页面仍包含已移除的快捷指令区或文字文档脚本');
}
['actions.js', 'planner.js', 'history.js', 'llm.js', 'modeBadge', 'btnHistory', 'settingExecMode'].forEach((token) => {
  if (html.indexOf(token) < 0) { errors.push('页面缺少执行操作链路入口：' + token); }
});
if (/btnDemoReset|demoPreview|查看演示数据|重置演示数据/.test(html)) {
  errors.push('页面仍包含浏览器演示 / 虚拟表格入口');
}
if (/\bWPS\b/i.test(html)) { errors.push('页面出现 WPS 可见字样'); }

const appSource = read('addon/js/app.js');
['WpsAIActions', 'WpsAIPlanner', 'WpsAIHistory', 'WpsAILlm'].forEach((token) => {
  if (appSource.indexOf(token) < 0) { errors.push('侧边栏代码缺少执行模块引用：' + token); }
});
if (/getDemo|demoPreview|resetDemo|updatePreviewHighlights|updatePreviewDOM/.test(appSource)) {
  errors.push('侧边栏代码仍包含浏览器演示预览逻辑');
}
const domIds = [];
const idRe = /\bel\('([A-Za-z][\w-]*)'\)/g;
let idMatch;
while ((idMatch = idRe.exec(appSource)) !== null) { domIds.push(idMatch[1]); }
Array.from(new Set(domIds)).forEach((id) => {
  if (html.indexOf('id="' + id + '"') < 0) {
    errors.push('侧边栏代码引用了页面不存在的元素：' + id);
  }
});

const bridgeSource = read('addon/js/bridge.js');
if (/getDemo|initDemo|demoPreview|logDemo|mode === 'demo'/.test(bridgeSource)) {
  errors.push('桥接层仍包含虚拟表格 / 演示模式');
}
if (bridgeSource.indexOf("mode: 'disconnected'") < 0) {
  errors.push('桥接层缺少「未连接真实表格」状态');
}

/* 防止把具体表名 / 列 / 阈值 / 文案重新写死回计划生成器 */
const plannerSource = read('addon/js/planner.js');
const advisorSource = read('addon/js/advisor.js');
const llmSource = read('addon/js/llm.js');
if (/makeFormula\(|formulaKind\(|>=60,"达标","未达标"/.test(plannerSource)) {
  errors.push('规则引擎仍包含写死阈值或文案的公式模板');
}
if (/高亮大于 90|计算 F 列占比|按 A 列分类/.test(plannerSource)) {
  errors.push('澄清候选仍写死具体列或阈值');
}
if (/"达标","未达标"|'>60'/.test(advisorSource)) {
  errors.push('公式顾问仍包含硬编码默认值');
}
if (llmSource.indexOf('sheetNames') < 0 || llmSource.indexOf('attempt(true)') < 0) {
  errors.push('大模型通道缺少真实表格上下文或空动作重试');
}
const actionsSource = read('addon/js/actions.js');
if (actionsSource.indexOf('et.deleteDuplicateRows') < 0 ||
    bridgeSource.indexOf('deleteDuplicateRows:') < 0) {
  errors.push('按内容去重删除行未同时接入动作白名单与桥接层');
}
const dedupIdx = plannerSource.indexOf("type: 'et.deleteDuplicateRows'");
const plainDeleteIdx = plannerSource.indexOf("type: 'et.deleteRows'");
if (dedupIdx < 0 || plainDeleteIdx < dedupIdx) {
  errors.push('去重删除分支必须排在普通删除行之前，避免把重复行误删成整个区域');
}

/* 条件动作、整行标色与撤回链路必须四处一致，缺一处就会出现「撤回不回原样」 */
const historySource = read('addon/js/history.js');
if (actionsSource.indexOf('et.deleteRowsIf') < 0 || bridgeSource.indexOf('deleteRowsIf:') < 0 ||
    historySource.indexOf("'et.deleteRowsIf'") < 0) {
  errors.push('按条件删除行未同时接入动作白名单、桥接层与撤回限制');
}
if (actionsSource.indexOf('fillScope') < 0 || bridgeSource.indexOf("opts.fillScope === 'row'") < 0 ||
    historySource.indexOf('wideRangeOf') < 0) {
  errors.push('标整行时快照范围未同步扩展，撤回会留下残留底色');
}
if (bridgeSource.indexOf('state.cells') < 0 || bridgeSource.indexOf('-4142') < 0) {
  errors.push('桥接层缺少逐单元格快照或「无填充」状态恢复');
}
if (appSource.indexOf('History.preview') < 0 || appSource.indexOf('canUndo') < 0) {
  errors.push('界面未提前提示不可撤回的计划，或仍对不可撤回事务显示撤回按钮');
}
/* 颜色类动作撤回不得把整片区域的字体色重写成兜底色 */
if (/String\(color \|\| ''\)/.test(bridgeSource)) {
  errors.push('颜色规范化仍把数值 0（自动色）当成空值，撤回会把字体色写成兜底浅黄');
}
if (actionsSource.indexOf('affectedProps:') < 0 || bridgeSource.indexOf('state.props') < 0 ||
    historySource.indexOf('Actions.affectedProps') < 0) {
  errors.push('撤回快照缺少属性级恢复，颜色动作会顺带重写整片区域字体色');
}
/* 高频任务能力必须四处齐备：白名单、桥接实现、规则分支、大模型 schema */
const highFreq = ['et.findReplace', 'et.fillBlanks', 'et.fillDown', 'et.mergeCells',
  'et.insertRows', 'et.insertColumns', 'et.deleteColumns', 'et.hideColumns',
  'et.dataValidation', 'et.pageSetup', 'et.conditionalFormatRule',
  'et.pivotTable', 'et.textToColumns'];
highFreq.forEach((t) => {
  const method = t.replace('et.', '') + ':';
  if (actionsSource.indexOf("'" + t + "'") < 0 || bridgeSource.indexOf(method) < 0 ||
      plannerSource.indexOf("'" + t + "'") < 0 || actionsSource.indexOf(t + ' {') < 0) {
    errors.push('高频动作未同时接入白名单 / 桥接层 / 规则引擎 / 大模型 schema：' + t);
  }
});
if (plannerSource.indexOf('fmt.push') < 0) {
  errors.push('一句话里的多个格式意图仍需可叠加输出');
}
if (advisorSource.indexOf("case 'sumifs'") < 0 || advisorSource.indexOf("case 'trim'") < 0 ||
    advisorSource.indexOf("case 'iferror'") < 0) {
  errors.push('高频公式类型（多条件 / 文本清洗 / 错误值）未补齐');
}
if (plannerSource.indexOf('@lookupRange') < 0) {
  errors.push('缺少占位公式写入拦截，错误公式会落到单元格里');
}
/* 新动作的参数必须能通过 validate 白名单，否则计划生成了但执行时参数丢失 */
if (!/'find', 'replace', 'delimiter'/.test(actionsSource)) {
  errors.push('动作校验白名单缺少分列分隔符，分列会退化成默认分隔符');
}
if (plannerSource.indexOf("advice.intentKind === 'sumifs'") < 0) {
  errors.push('缺少多条件公式的区域校验，凭选区推断的 SUMIFS 会写进单元格');
}

const ribbon = read('addon/ribbon.xml');
if (!/onAction="OnAction"/.test(ribbon)) { errors.push('ribbon.xml 缺少 OnAction 回调'); }
if ((ribbon.match(/<button\b/g) || []).length !== 1) {
  errors.push('ribbon.xml 应只保留一个「打开侧边栏」按钮');
}

const publish = read('publish.xml');
if (publish.indexOf('type="et"') < 0) { errors.push('publish.xml 缺少表格组件注册项'); }
if (/type="wps"|type="wpp"/.test(publish)) { errors.push('publish.xml 不应注册文字 / 演示组件'); }

if (errors.length) {
  console.error('自检未通过：');
  errors.forEach((e) => console.error('  - ' + e));
  process.exit(1);
}
console.log('自检通过：' + checked + ' 个运行文件 / 资源检查完成。');
