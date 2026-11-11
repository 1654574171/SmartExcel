/*
 * 表格宿主桥接层：统一表格操作接口，只面向真实表格客户端的 JSAPI。
 * 没有宿主时进入「未连接」状态：公式顾问仍可离线给出建议，
 * 但不会构造虚拟表格，也不会伪装执行任何写操作。
 */
(function (global) {
  'use strict';

  var ALIGN = { left: -4131, center: -4108, right: -4152 };
  var cached = null;
  var colorMode = 'auto';
  var sniffed = null;

  var Util = {
    colToNum: function (letters) {
      var n = 0;
      var s = String(letters).toUpperCase().replace(/[^A-Z]/g, '');
      for (var i = 0; i < s.length; i++) { n = n * 26 + (s.charCodeAt(i) - 64); }
      return n;
    },
    numToCol: function (n) {
      var s = '';
      while (n > 0) {
        var m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s || 'A';
    },
    /* 区域地址可能带「工作表名!」前缀；坐标计算只用地址部分，前缀由调用方拼回 */
    splitSheetAddr: function (addr) {
      var raw = String(addr || '');
      var bang = raw.lastIndexOf('!');
      if (bang < 0) { return { prefix: '', plain: raw }; }
      return { prefix: raw.slice(0, bang + 1), plain: raw.slice(bang + 1) };
    },
    parseA1: function (addr) {
      var s = Util.splitSheetAddr(addr).plain.toUpperCase().replace(/\$/g, '').replace(/\s/g, '');
      var m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(s);
      if (m) {
        var c1 = Util.colToNum(m[1]);
        var c2 = m[3] ? Util.colToNum(m[3]) : c1;
        var r1 = parseInt(m[2], 10);
        var r2 = m[4] ? parseInt(m[4], 10) : r1;
        return {
          c1: Math.min(c1, c2), r1: Math.min(r1, r2),
          c2: Math.max(c1, c2), r2: Math.max(r1, r2)
        };
      }
      m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s);
      if (m) { return { c1: Util.colToNum(m[1]), r1: 1, c2: Util.colToNum(m[2]), r2: 1048576 }; }
      m = /^(\d+):(\d+)$/.exec(s);
      if (m) { return { c1: 1, r1: parseInt(m[1], 10), c2: 16384, r2: parseInt(m[2], 10) }; }
      return null;
    },
    toMatrix: function (value) {
      if (Object.prototype.toString.call(value) === '[object Array]') {
        if (value.length && Object.prototype.toString.call(value[0]) === '[object Array]') { return value; }
        var row = [];
        for (var i = 0; i < value.length; i++) { row.push(value[i]); }
        return [row];
      }
      return [[value]];
    },
    isNumber: function (v) {
      if (typeof v === 'number') { return isFinite(v); }
      if (typeof v === 'string' && v.trim() !== '') { return isFinite(Number(v.replace(/,/g, ''))); }
      return false;
    },
    toNumber: function (v) { return Number(String(v).replace(/,/g, '')); },
    hexToBgr: function (hex) {
      var h = Util.normalizeColor(hex).replace('#', '');
      var r = parseInt(h.substring(0, 2), 16);
      var g = parseInt(h.substring(2, 4), 16);
      var b = parseInt(h.substring(4, 6), 16);
      return r + g * 256 + b * 65536;
    },
    normalizeColor: function (color) {
      var c = String(color || '').trim();
      if (/^#[0-9a-f]{6}$/i.test(c)) { return c.toUpperCase(); }
      if (/^#[0-9a-f]{3}$/i.test(c)) {
        return ('#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3]).toUpperCase();
      }
      var n = parseInt(c, 10);
      if (isFinite(n) && n >= 0 && n <= 16777215) {
        var r = n % 256, g = Math.floor(n / 256) % 256, b = Math.floor(n / 65536);
        var pad = function (x) { return ('0' + x.toString(16)).slice(-2); };
        return ('#' + pad(r) + pad(g) + pad(b)).toUpperCase();
      }
      return '#FFF2CC';
    },
    /* 有些版本的属性是 getter，有些是方法，统一取值 */
    propOrCall: function (obj, name, args) {
      if (!obj) { return null; }
      try {
        var v = obj[name];
        if (typeof v === 'function') { return v.apply(obj, args || []); }
        return v;
      } catch (e) { return null; }
    }
  };

  function getWps() {
    try {
      if (global.wps) { return global.wps; }
      if (global.parent && global.parent.wps) { return global.parent.wps; }
      if (global.top && global.top.wps) { return global.top.wps; }
    } catch (e) { /* 跨域保护 */ }
    return null;
  }

  /* 未连接时返回空白环境，绝不伪造工作簿 / 单元格数据 */
  function disconnectedInfo() {
    return {
      host: 'browser', mode: 'disconnected', app: null,
      docName: '', sheetName: '', selection: '', selectionForAI: '',
      lastRow: 12, lastCol: 6, usedRange: ''
    };
  }

  function detect(force) {
    if (cached && !force) { return cached; }
    var info = disconnectedInfo();
    var w = getWps();
    if (!w) { cached = info; return info; }

    /* 只探测表格组件；文字 / 演示组件不再注册或识别 */
    var app = null;
    try { app = w.Application; } catch (e) { app = null; }
    var candidates = [];
    if (app) { candidates.push(app); }
    try {
      if (typeof w.EtApplication === 'function') {
        var made = w.EtApplication();
        if (made) { candidates.push(made); }
      }
    } catch (e2) { /* 当前进程不支持表格组件 */ }

    for (var i = 0; i < candidates.length; i++) {
      var a = candidates[i];
      var hasWorkbook = false;
      try { hasWorkbook = !!(a.ActiveWorkbook || a.ActiveSheet || a.Workbooks); } catch (e3) { /* 忽略 */ }
      if (!hasWorkbook) { continue; }
      info = {
        host: 'et', mode: 'live', app: a, docName: '', sheetName: '',
        selection: '', selectionForAI: '', lastRow: 12, lastCol: 6, usedRange: ''
      };
      try {
        var wb = a.ActiveWorkbook;
        info.docName = wb ? String(wb.Name) : '当前工作簿';
        var sh = getSheetOf(a);
        if (sh) { info.sheetName = String(sh.Name); }
        var used = sh && sh.UsedRange;
        if (used) {
          var firstRow = Number(Util.propOrCall(used, 'Row')) || 1;
          var firstCol = Number(Util.propOrCall(used, 'Column')) || 1;
          var rowCount = Number(Util.propOrCall(Util.propOrCall(used, 'Rows'), 'Count')) || 1;
          var colCount = Number(Util.propOrCall(Util.propOrCall(used, 'Columns'), 'Count')) || 1;
          info.lastRow = firstRow + rowCount - 1;
          info.lastCol = firstCol + colCount - 1;
          info.usedRange = String(Util.propOrCall(used, 'Address') || '');
        }
        info.selection = String(Util.propOrCall(a.Selection, 'Address') || '') || info.usedRange;
        info.selectionForAI = info.selection;
      } catch (e4) { /* 部分属性不可用，保留默认值 */ }
      break;
    }
    cached = info;
    return info;
  }

  function info() { return cached || detect(); }

  function getSheetOf(app) {
    if (!app) { return null; }
    try { if (app.ActiveSheet) { return app.ActiveSheet; } } catch (e) { /* 忽略 */ }
    try { return app.ActiveWorkbook.ActiveSheet; } catch (e2) { /* 忽略 */ }
    try { return app.Workbooks.Item(1).ActiveSheet; } catch (e3) { /* 忽略 */ }
    return null;
  }

  function getSheet() { return getSheetOf(info().app); }

  function colorValue(hex) {
    var mode = colorMode;
    if (mode === 'auto') {
      if (!sniffed) {
        sniffed = 'hex';
        var sheet = getSheet();
        if (sheet) {
          try {
            var current = sheet.Range('A1').Interior.Color;
            if (typeof current === 'number') { sniffed = 'bgr'; }
          } catch (e) { /* 保持 hex */ }
        }
      }
      mode = sniffed;
    }
    return mode === 'bgr' ? Util.hexToBgr(hex) : Util.normalizeColor(hex);
  }

  function findSheetByName(name) {
    var app = info().app;
    if (!app || !name) { return null; }
    var sheets = null;
    try {
      sheets = Util.propOrCall(app, 'Worksheets') ||
        Util.propOrCall(app.ActiveWorkbook, 'Worksheets');
    } catch (e) { /* 忽略 */ }
    if (!sheets) { return null; }
    var count = Number(Util.propOrCall(sheets, 'Count')) || 0;
    var key = String(name).trim().toLowerCase();
    for (var k = 1; k <= count; k++) {
      var sh = Util.propOrCall(sheets, 'Item', [k]);
      var sname = sh ? String(Util.propOrCall(sh, 'Name') || '') : '';
      if (sname && sname.trim().toLowerCase() === key) { return sh; }
    }
    return null;
  }

  /* 目标区域允许写成 员工信息!A2（跨表限定地址） */
  function getRange(addr) {
    var raw = String(addr || '');
    var bang = raw.lastIndexOf('!');
    var sheet = getSheet();
    var target = raw;
    if (bang >= 0) {
      var name = raw.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
      target = raw.slice(bang + 1);
      var wanted = findSheetByName(name);
      if (!wanted) { throw new Error('未找到工作表「' + name + '」'); }
      sheet = wanted;
    }
    if (!sheet) { throw new Error('未连接真实表格，无法定位区域 ' + target); }
    var range = sheet.Range(target);
    if (!range) { throw new Error('无法定位区域 ' + target); }
    return range;
  }

  function listSheetNames() {
    var app = info().app;
    if (!app) { return []; }
    var sheets = null;
    try {
      sheets = Util.propOrCall(app, 'Worksheets') ||
        Util.propOrCall(app.ActiveWorkbook, 'Worksheets');
    } catch (e) { /* 忽略 */ }
    if (!sheets) { return []; }
    var count = Number(Util.propOrCall(sheets, 'Count')) || 0;
    var names = [];
    for (var k = 1; k <= count; k++) {
      var sh = Util.propOrCall(sheets, 'Item', [k]);
      var name = sh ? String(Util.propOrCall(sh, 'Name') || '') : '';
      if (name) { names.push(name); }
    }
    return names;
  }

  function sameSheet(a, b) {
    return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
  }

  /* 定位跨表引用前先激活对应工作表 */
  function activateSheet(name) {
    var target = String(name || '').trim();
    if (!target) { return { ok: false, message: '工作表名为空' }; }
    var app = info().app;
    if (!app) { return { ok: false, message: '未连接真实表格，无法切换工作表' }; }
    var sheets = null;
    try {
      sheets = Util.propOrCall(app, 'Worksheets') ||
        Util.propOrCall(app.ActiveWorkbook, 'Worksheets');
    } catch (e) { /* 忽略 */ }
    if (!sheets) { return { ok: false, message: '当前表格版本无法枚举工作表，请手动切换' }; }
    var count = Number(Util.propOrCall(sheets, 'Count')) || 0;
    for (var k = 1; k <= count; k++) {
      var sh = Util.propOrCall(sheets, 'Item', [k]);
      var sname = sh ? String(Util.propOrCall(sh, 'Name') || '') : '';
      if (!sameSheet(sname, target)) { continue; }
      try {
        Util.propOrCall(sh, 'Activate');
        cached = null;
        return { ok: true, message: '已切换到工作表 ' + sname, sheet: sname };
      } catch (e2) { return { ok: false, message: '无法激活工作表 ' + sname }; }
    }
    return { ok: false, message: '未找到工作表「' + target + '」，请手动切换' };
  }

  function currentSelectionRef() {
    var i = detect(true);
    var addr = String(i.selection || '').replace(/\$/g, '').replace(/\s/g, '').toUpperCase();
    if (!addr) { addr = String(i.usedRange || '').replace(/\$/g, '').replace(/\s/g, '').toUpperCase(); }
    /* Selection.Address 可能形如 [BOOK1]SHEET1!$B$2:$B$11，只取地址部分 */
    var bang = addr.lastIndexOf('!');
    if (bang >= 0) { addr = addr.slice(bang + 1); }
    return { addr: addr, sheet: i.sheetName || '', book: i.docName || '' };
  }

  function selectionBottom(selection) {
    var p = Util.parseA1(selection);
    if (!p) { return 'A2'; }
    return Util.numToCol(p.c1) + (p.r2 + 1);
  }

  function expand(text) {
    var i = info();
    return String(text == null ? '' : text)
      .replace(/@lastRowPlus1/gi, String(i.lastRow + 1))
      .replace(/@lastRow/gi, String(i.lastRow))
      .replace(/@lastCol/gi, Util.numToCol(i.lastCol))
      .replace(/@selectionBottom/gi, selectionBottom(i.selection))
      .replace(/@selection/gi, i.selection || 'A1');
  }

  function resolveRange(range) {
    var i = info();
    var parts = Util.splitSheetAddr(String(range == null ? '' : range).trim().replace(/\$/g, ''));
    /* 只把地址部分转成大写，工作表名保留原样，避免界面显示成 SHEET1 */
    var s = parts.prefix + parts.plain.toUpperCase();
    if (!s || s === '@SELECTION' || s === '当前选区' || s === '选中区域' || s === '当前区域' || s === '同一区域' || s === '刚才的区域' || s === '它') {
      s = i.selection || i.usedRange || 'A1';
    }
    s = expand(s);
    var colMatch = /^([A-Z]{1,3})\s*列$/.exec(s);
    if (colMatch) { return colMatch[1] + '1:' + colMatch[1] + i.lastRow; }
    var rowMatch = /^第?\s*(\d+)\s*行$/.exec(s);
    if (rowMatch) { return rowMatch[1] + ':' + rowMatch[1]; }
    if (/^[A-Z]{1,3}$/.test(s)) { return s + '1:' + s + i.lastRow; }
    return s;
  }

  var ET = {
    expand: expand,
    resolve: resolveRange,
    info: function () { return info(); },

    getValues: function (addr) {
      return Util.toMatrix(getRange(resolveRange(addr)).Value2);
    },

    setValue: function (addr, value) {
      getRange(resolveRange(addr)).Value2 = value;
    },

    setFormula: function (addr, formula) {
      getRange(resolveRange(addr)).Formula = expand(formula);
    },

    setFill: function (addr, color) {
      addr = resolveRange(addr);
      var hex = Util.normalizeColor(color);
      try { getRange(addr).Interior.Color = colorValue(hex); }
      catch (e) {
        try { getRange(addr).Interior.Color = hex; }
        catch (e2) { throw new Error('设置底色失败：' + e2.message); }
      }
    },

    setFont: function (addr, opts) {
      addr = resolveRange(addr);
      var font = getRange(addr).Font;
      if (!font) { throw new Error('无法获取字体对象'); }
      if (opts.bold !== undefined) { font.Bold = !!opts.bold; }
      if (opts.italic !== undefined) { font.Italic = !!opts.italic; }
      if (opts.size) { font.Size = opts.size; }
      if (opts.name) { font.Name = opts.name; }
      if (opts.color) {
        try { font.Color = colorValue(Util.normalizeColor(opts.color)); }
        catch (e) { font.Color = Util.normalizeColor(opts.color); }
      }
    },

    setFormat: function (addr, opts) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      if (opts.numberFormat) { range.NumberFormat = opts.numberFormat; }
      if (opts.align && ALIGN[opts.align] !== undefined) { range.HorizontalAlignment = ALIGN[opts.align]; }
      if (opts.width) { range.ColumnWidth = opts.width; }
      if (opts.autoFit) { Util.propOrCall(range.Columns, 'AutoFit'); }
      if (opts.border) {
        try { range.Borders.LineStyle = 1; range.Borders.Weight = 2; } catch (e) { /* 忽略 */ }
      }
    },

    clear: function (addr, what) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      if (what === 'formats') { Util.propOrCall(range, 'ClearFormats'); }
      else if (what === 'all') { Util.propOrCall(range, 'Clear'); }
      else { Util.propOrCall(range, 'ClearContents'); }
    },

    conditionalFill: function (addr, op, threshold, color, fontColor) {
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var matrix = ET.getValues(addr);
      var p = Util.parseA1(parts.plain);
      /* 解析失败时明确报错，绝不回退到 A1 起点把颜色写到错误的位置 */
      if (!p) { throw new Error('条件高亮需要明确的区域：' + addr); }
      var matched = 0;
      for (var r = 0; r < matrix.length; r++) {
        for (var c = 0; c < matrix[r].length; c++) {
          var v = matrix[r][c];
          if (!Util.isNumber(v)) { continue; }
          var n = Util.toNumber(v);
          var hit = false;
          if (op === '>') { hit = n > threshold; }
          else if (op === '>=') { hit = n >= threshold; }
          else if (op === '<') { hit = n < threshold; }
          else if (op === '<=') { hit = n <= threshold; }
          else if (op === '=' || op === '==') { hit = n === threshold; }
          else if (op === '!=') { hit = n !== threshold; }
          if (hit) {
            /* 单个单元格也要带上原来的工作表前缀，否则会标到当前活动表 */
            var cellAddr = parts.prefix + Util.numToCol(p.c1 + c) + (p.r1 + r);
            ET.setFill(cellAddr, color);
            if (fontColor) { ET.setFont(cellAddr, { color: fontColor }); }
            matched++;
          }
        }
      }
      return matched;
    },

    sort: function (addr, order) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      var keyRange = Util.propOrCall(range.Columns, 'Item', [1]) || range;
      if (typeof range.Sort === 'function') { range.Sort(keyRange, order === 'desc' ? 2 : 1, 1); }
      else { Util.propOrCall(range, 'Sort', [keyRange, order === 'desc' ? 2 : 1, 1]); }
      return 0;
    },

    autoFilter: function (addr) {
      Util.propOrCall(getRange(resolveRange(addr)), 'AutoFilter');
    },

    freeze: function (addr) {
      addr = resolveRange(addr);
      var p = Util.parseA1(addr);
      var rows = p ? p.r1 - 1 : 1;
      var win = info().app && info().app.ActiveWindow;
      if (!win) { throw new Error('当前没有活动窗口'); }
      win.FreezePanes = false;
      win.SplitRow = rows;
      win.SplitColumn = 0;
      win.FreezePanes = true;
    },

    insertChart: function (addr, type) {
      addr = resolveRange(addr);
      var sheet = getSheet();
      var holder = Util.propOrCall(sheet, 'ChartObjects');
      if (!holder) { throw new Error('当前版本不支持图表对象'); }
      var obj = holder.Add(320, 40, 360, 220);
      var chart = obj.Chart;
      chart.ChartType = type === 'line' ? 4 : 51;
      if (typeof chart.SetSourceData === 'function') { chart.SetSourceData(getRange(addr)); }
    },

    addSheet: function (name) {
      var app = info().app;
      var sheets = Util.propOrCall(app, 'Worksheets') || Util.propOrCall(app.ActiveWorkbook, 'Worksheets');
      var added = Util.propOrCall(sheets, 'Add');
      if (added && name) { added.Name = name; }
    },

    renameSheet: function (name) {
      var sheet = getSheet();
      sheet.Name = name;
    },

    deleteRows: function (addr) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      if (range.EntireRow && typeof range.EntireRow.Delete === 'function') { range.EntireRow.Delete(); }
      else { Util.propOrCall(range.EntireRow, 'Delete'); }
    },

    copyTo: function (addr, target) {
      getRange(resolveRange(addr)).Copy(getRange(resolveRange(target)));
    },

    fillSeries: function (addr, start, step) {
      addr = resolveRange(addr);
      var p = Util.parseA1(addr);
      if (!p) { throw new Error('序列填充需要明确的区域'); }
      var value = Number(start) || 1;
      var delta = Number(step) || 1;
      var matrix = [];
      for (var r = p.r1; r <= Math.min(p.r2, p.r1 + 499); r++) {
        var line = [];
        for (var c = p.c1; c <= Math.min(p.c2, p.c1 + 59); c++) {
          line.push(value + delta * ((r - p.r1) * (p.c2 - p.c1 + 1) + (c - p.c1)));
        }
        matrix.push(line);
      }
      getRange(addr).Value2 = matrix;
    },

    summary: function (addr) {
      addr = resolveRange(addr || '@selection');
      var matrix = ET.getValues(addr);
      var count = 0, sum = 0, max = null, min = null;
      for (var r = 0; r < matrix.length; r++) {
        for (var c = 0; c < matrix[r].length; c++) {
          var v = matrix[r][c];
          if (Util.isNumber(v)) {
            var n = Util.toNumber(v);
            count++; sum += n;
            max = max === null ? n : Math.max(max, n);
            min = min === null ? n : Math.min(min, n);
          }
        }
      }
      return {
        range: addr, rows: matrix.length, cols: matrix.length ? matrix[0].length : 0,
        numeric: count, sum: Math.round(sum * 100) / 100,
        average: count ? Math.round(sum / count * 100) / 100 : 0,
        max: max, min: min, preview: matrix.slice(0, 3)
      };
    },

    /* ---- 撤回快照 ---- */
    getState: function (addr) {
      addr = resolveRange(addr);
      var state = { address: addr, live: true, cells: null, values: null, formulas: null, style: null };
      var range = getRange(addr);
      function read(obj, prop) {
        try {
          if (!obj) { return null; }
          var v = obj[prop];
          return (v === undefined || v === null) ? null : v;
        } catch (e) { return null; }
      }
      state.values = Util.toMatrix(range.Value2);
      try { state.formulas = Util.toMatrix(range.Formula); } catch (e2) { state.formulas = state.values; }
      var interior = null, font = null;
      try { interior = range.Interior; } catch (e3) { /* 忽略 */ }
      try { font = range.Font; } catch (e4) { /* 忽略 */ }
      state.style = {
        color: read(interior, 'Color'),
        bold: read(font, 'Bold'), italic: read(font, 'Italic'), size: read(font, 'Size'),
        fontColor: read(font, 'Color'), fontName: read(font, 'Name'),
        numberFormat: read(range, 'NumberFormat'),
        align: read(range, 'HorizontalAlignment'),
        width: read(range, 'ColumnWidth')
      };
      if (state.style.color !== null) { state.style.color = Util.normalizeColor(state.style.color); }
      if (state.style.fontColor !== null) { state.style.fontColor = Util.normalizeColor(state.style.fontColor); }
      return state;
    },

    setState: function (state) {
      if (!state || !state.address) { return; }
      var addr = state.address;
      var range = getRange(addr);
      if (state.values) { range.Value2 = state.values; }
      var formulas = state.formulas || state.values;
      var stateParts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(stateParts.plain);
      if (p && formulas) {
        var rows = Math.min(p.r2 - p.r1 + 1, 500);
        var cols = Math.min(p.c2 - p.c1 + 1, 60);
        for (var r = 0; r < rows; r++) {
          for (var c = 0; c < cols; c++) {
            var f = formulas[r] && formulas[r][c];
            if (typeof f === 'string' && f.charAt(0) === '=') {
              try { getRange(stateParts.prefix + Util.numToCol(p.c1 + c) + (p.r1 + r)).Formula = f; } catch (e) { /* 忽略 */ }
            }
          }
        }
      }
      var st = state.style || {};
      try {
        var interior = range.Interior;
        if (st.color !== null && st.color !== undefined) { interior.Color = colorValue(st.color); }
      } catch (e2) { /* 忽略 */ }
      try {
        var font = range.Font;
        if (st.bold !== null && st.bold !== undefined) { font.Bold = st.bold; }
        if (st.italic !== null && st.italic !== undefined) { font.Italic = st.italic; }
        if (st.size !== null && st.size !== undefined) { font.Size = st.size; }
        if (st.fontName !== null && st.fontName !== undefined) { font.Name = st.fontName; }
        if (st.fontColor !== null && st.fontColor !== undefined) { font.Color = colorValue(st.fontColor); }
      } catch (e3) { /* 忽略 */ }
      try { if (st.numberFormat !== null && st.numberFormat !== undefined) { range.NumberFormat = st.numberFormat; } } catch (e4) { /* 忽略 */ }
      try { if (st.align !== null && st.align !== undefined) { range.HorizontalAlignment = st.align; } } catch (e5) { /* 忽略 */ }
      try { if (st.width !== null && st.width !== undefined) { range.ColumnWidth = st.width; } } catch (e6) { /* 忽略 */ }
    },

    /* ---- 多区域引用：读取当前选区（含工作表）/ 在表格中定位高亮 ---- */
    sheetNames: listSheetNames,
    activateSheet: activateSheet,
    selectionRef: currentSelectionRef,
    selectionAddress: function () { return currentSelectionRef().addr; },

    selectRange: function (addr, sheet) {
      var target = String(addr || '').replace(/\$/g, '').replace(/\s/g, '').toUpperCase();
      if (!target) { return { ok: false, message: '区域地址为空' }; }
      if (!info().app) { return { ok: false, message: '未连接真实表格，无法在表格中定位' }; }
      var wantSheet = String(sheet || '').trim();
      if (wantSheet && !sameSheet(wantSheet, info().sheetName)) {
        var switched = activateSheet(wantSheet);
        if (!switched.ok) { return switched; }
      }
      try {
        var range = getRange(target);
        if (typeof range.Select === 'function') { range.Select(); }
        else { Util.propOrCall(range, 'Select'); }
        return { ok: true, message: '已在「' + info().sheetName + '」选中 ' + target };
      } catch (e) {
        return { ok: false, message: '当前表格版本无法自动定位（' + (e && e.message ? e.message : e) + '），请手动选中该区域' };
      }
    }
  };

  global.WpsAIBridge = {
    util: Util,
    info: info,
    detect: detect,
    raw: getWps,
    et: ET,
    getSheet: getSheet,
    setColorMode: function (mode) { colorMode = mode; sniffed = null; },
    align: ALIGN
  };
})(window);
