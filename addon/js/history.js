/*
 * 操作事务与撤回/重做。
 * 每次执行计划前采集受影响表格区域的快照，撤回时恢复，重做时恢复执行后快照。
 * 历史仅保存在当前面板的内存中：刷新或重开面板后清空，避免跨工作簿误恢复。
 */
(function (global) {
  'use strict';

  var Bridge = global.WpsAIBridge;
  var Actions = global.WpsAIActions;
  var MAX = 20;
  var stack = [];

  /* 结构性 / 视图类操作无法用区域快照可靠还原，明确提示用户改用表格自带撤销 */
  var BLOCKED = {
    'et.deleteRows': '删除行',
    'et.addSheet': '新建工作表',
    'et.renameSheet': '重命名工作表',
    'et.insertChart': '插入图表',
    'et.freeze': '冻结窗格',
    'et.autoFilter': '添加筛选'
  };

  function uid() {
    return 'tx' + Date.now().toString(36) + Math.random().toString(36).substring(2, 7);
  }

  function rangesOf(plan) {
    var seen = {};
    var list = [];
    for (var i = 0; i < plan.actions.length; i++) {
      var a = plan.actions[i];
      var keys = [];
      if (a.range) { keys.push(a.range); }
      if (a.target) { keys.push(a.target); }
      for (var j = 0; j < keys.length; j++) {
        var addr;
        try { addr = Bridge.et.resolve(keys[j]); } catch (e) { addr = String(keys[j]); }
        if (addr && !seen[addr]) { seen[addr] = true; list.push(addr); }
      }
    }
    return list;
  }

  function capture(plan) {
    var info = Bridge.info();
    var before = {
      app: 'et', docName: info.docName || '', ranges: [],
      undoable: true, reason: '', noop: false
    };
    var writes = 0;
    var blocked = [];
    for (var i = 0; i < plan.actions.length; i++) {
      var a = plan.actions[i];
      if (Actions.risk(a) === 'read') { continue; }
      writes++;
      if (BLOCKED[a.type]) { blocked.push(BLOCKED[a.type]); }
    }
    if (!writes) { before.noop = true; return before; }
    if (blocked.length) {
      before.undoable = false;
      before.reason = '含无法自动撤回的操作（' + blocked.join('、') + '），请用表格的 Ctrl+Z';
      return before;
    }
    try {
      var addrs = rangesOf(plan);
      for (var k = 0; k < addrs.length; k++) { before.ranges.push(Bridge.et.getState(addrs[k])); }
      if (!before.ranges.length) {
        before.undoable = false;
        before.reason = '计划未指定受影响区域，无法自动撤回，请用表格的 Ctrl+Z';
      }
    } catch (e) {
      before.undoable = false;
      before.reason = '读取执行前状态失败：' + (e && e.message ? e.message : e);
    }
    return before;
  }

  function commit(plan, result, before) {
    if (!before || before.noop || !result || !result.succeeded) { return null; }
    /* 写入新操作后，丢弃已撤回的事务（清空重做栈） */
    var kept = [];
    for (var i = 0; i < stack.length; i++) {
      if (stack[i].status !== 'undone') { kept.push(stack[i]); }
    }
    stack = kept;
    var tx = {
      id: uid(), time: Date.now(), app: 'et',
      summary: String(plan.summary || '执行操作').substring(0, 60),
      actionCount: plan.actions.length, docName: before.docName,
      before: before, after: null, undoable: before.undoable,
      reason: before.reason, status: 'done'
    };
    stack.push(tx);
    if (stack.length > MAX) { stack.shift(); }
    return tx;
  }

  function snapshotFrom(before) {
    var snap = { app: 'et', ranges: [] };
    for (var i = 0; i < before.ranges.length; i++) {
      snap.ranges.push(Bridge.et.getState(before.ranges[i].address));
    }
    return snap;
  }

  function restore(snap) {
    if (!snap) { return { ok: false, message: '缺少快照' }; }
    for (var i = 0; i < snap.ranges.length; i++) { Bridge.et.setState(snap.ranges[i]); }
    return { ok: true, message: '' };
  }

  function find(id) {
    for (var i = 0; i < stack.length; i++) { if (stack[i].id === id) { return stack[i]; } }
    return null;
  }

  function undo(id) {
    var tx = find(id);
    if (!tx) { return { ok: false, message: '该操作记录已失效（面板刷新后历史会清空）' }; }
    if (!tx.undoable) { return { ok: false, message: tx.reason }; }
    if (tx.status === 'undone') { return { ok: false, message: '该操作已撤回' }; }
    try {
      tx.after = snapshotFrom(tx.before);
      var res = restore(tx.before);
      if (res && res.ok === false) { return { ok: false, message: res.message }; }
      tx.status = 'undone';
      return { ok: true, message: '已撤回：' + tx.summary };
    } catch (e) {
      return { ok: false, message: '撤回失败：' + (e && e.message ? e.message : e) };
    }
  }

  function redo(id) {
    var tx = find(id);
    if (!tx) { return { ok: false, message: '该操作记录已失效' }; }
    if (tx.status !== 'undone') { return { ok: false, message: '该操作无需重做' }; }
    if (!tx.after) { return { ok: false, message: '缺少重做快照' }; }
    try {
      var res = restore(tx.after);
      if (res && res.ok === false) { return { ok: false, message: res.message }; }
      tx.status = 'done';
      return { ok: true, message: '已重做：' + tx.summary };
    } catch (e) {
      return { ok: false, message: '重做失败：' + (e && e.message ? e.message : e) };
    }
  }

  function lastWith(status) {
    for (var i = stack.length - 1; i >= 0; i--) {
      if (stack[i].status === status && stack[i].undoable) { return stack[i]; }
    }
    return null;
  }

  function undoLast() {
    var tx = lastWith('done');
    return tx ? undo(tx.id) : { ok: false, message: '没有可撤回的操作' };
  }

  function redoLast() {
    var tx = lastWith('undone');
    return tx ? redo(tx.id) : { ok: false, message: '没有可重做的操作' };
  }

  function list() { return stack.slice().reverse(); }
  function clear() { stack = []; }

  global.WpsAIHistory = {
    capture: capture, commit: commit, undo: undo, redo: redo,
    undoLast: undoLast, redoLast: redoLast,
    list: list, find: find, clear: clear, max: MAX
  };
})(window);
