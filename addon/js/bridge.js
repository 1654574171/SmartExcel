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
      /* 数值 0 是 JSAPI 里合法的「自动 / 黑色」色值，不能用 || 把它丢成空串再落到兜底色 */
      var c = (color === null || color === undefined) ? '' : String(color).trim();
      if (/^#[0-9a-f]{6}$/i.test(c)) { return c.toUpperCase(); }
      if (/^#[0-9a-f]{3}$/i.test(c)) {
        return ('#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3]).toUpperCase();
      }
      var n = Number(c);
      if (c !== '' && isFinite(n) && n >= 0 && n <= 16777215 && Math.floor(n) === n) {
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

  function findSheetByName(name, sheetsArg) {
    var app = info().app;
    if (!app || !name) { return null; }
    var sheets = sheetsArg || null;
    if (!sheets) {
      try {
        sheets = Util.propOrCall(app, 'Worksheets') ||
          Util.propOrCall(app.ActiveWorkbook, 'Worksheets');
      } catch (e) { /* 忽略 */ }
    }
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

  /* 目标区域允许写成 员工信息!A2，也允许 [另一工作簿.xlsx]Sheet1!A1 跨工作簿引用 */
  function getRange(addr) {
    var raw = String(addr || '');
    var bang = raw.lastIndexOf('!');
    var sheet = getSheet();
    var target = raw;
    if (bang >= 0) {
      var name = raw.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
      target = raw.slice(bang + 1);
      var sheetsArg = null;
      var ext = /^\[([^\]]+)\](.+)$/.exec(name);
      if (ext) {
        var books = Util.propOrCall(info().app, 'Workbooks');
        var other = books && Util.propOrCall(books, 'Item', [ext[1]]);
        sheetsArg = other && Util.propOrCall(other, 'Worksheets');
        name = ext[2];
        if (!sheetsArg) { throw new Error('未找到已打开的工作簿「' + ext[1] + '」'); }
      }
      var wanted = findSheetByName(name, sheetsArg);
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

  /* 条件匹配既支持数值比较，也支持「是 / 否」这类文本值 */
  function matchCell(value, op, threshold) {
    if (value === null || value === undefined) { return false; }
    var raw = String(value).trim();
    var wanted = (threshold === null || threshold === undefined) ? '' : String(threshold).trim();
    var numeric = Util.isNumber(value) && wanted !== '' && Util.isNumber(threshold);
    var n = numeric ? Util.toNumber(value) : NaN;
    var t = numeric ? Util.toNumber(threshold) : NaN;
    switch (op) {
      case '>': return numeric && n > t;
      case '>=': return numeric && n >= t;
      case '<': return numeric && n < t;
      case '<=': return numeric && n <= t;
      case '=': case '==': case '等于':
        return numeric ? n === t : raw.toLowerCase() === wanted.toLowerCase();
      case '!=': case '<>': case '不等于':
        return numeric ? n !== t : raw.toLowerCase() !== wanted.toLowerCase();
      case 'contains': case '包含':
        return wanted !== '' && raw.toLowerCase().indexOf(wanted.toLowerCase()) >= 0;
      default: return false;
    }
  }

  /* 找出满足条件的行号；条件区域可指定只按某一列判断 */
  function matchedRows(addr, op, threshold, keyColumn) {
    var parts = Util.splitSheetAddr(addr);
    var p = Util.parseA1(parts.plain);
    if (!p) { throw new Error('条件操作需要明确的区域：' + addr); }
    var matrix = ET.getValues(addr);
    var colIndex = -1;
    if (keyColumn) {
      colIndex = /^\d+$/.test(String(keyColumn))
        ? Number(keyColumn) - 1
        : Util.colToNum(keyColumn) - p.c1;
    }
    var rows = [];
    for (var r = 0; r < matrix.length; r++) {
      var cells = matrix[r] || [];
      var hit = false;
      if (colIndex >= 0) { hit = matchCell(cells[colIndex], op, threshold); }
      else {
        for (var c = 0; c < cells.length; c++) {
          if (matchCell(cells[c], op, threshold)) { hit = true; break; }
        }
      }
      if (hit) { rows.push(p.r1 + r); }
    }
    return { parts: parts, p: p, rows: rows };
  }

  /* 从下往上删除整行，避免前面删行导致后面的行号偏移 */
  function deleteRowsByNumbers(parts, rowNumbers) {
    var sorted = rowNumbers.slice().sort(function (a, b) { return b - a; });
    var seen = {}, removed = 0;
    for (var i = 0; i < sorted.length; i++) {
      if (seen[sorted[i]]) { continue; }
      seen[sorted[i]] = 1;
      ET.deleteRows(parts.prefix + sorted[i] + ':' + sorted[i]);
      removed++;
    }
    return removed;
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

    /* 整行标色的末列按目标工作表自己的已用区域算；撤回快照复用它，避免宽表残留底色 */
    rowLastCol: function (addr) {
      var parts = Util.splitSheetAddr(String(addr || ''));
      var name = parts.prefix ? parts.prefix.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : '';
      var sheet = name ? findSheetByName(name) : getSheet();
      var used = sheet && sheet.UsedRange;
      if (used) {
        var firstCol = Number(Util.propOrCall(used, 'Column')) || 1;
        var count = Number(Util.propOrCall(Util.propOrCall(used, 'Columns'), 'Count')) || 1;
        return Math.max(1, firstCol + count - 1);
      }
      return Math.max(1, Number(info().lastCol) || 1);
    },

    conditionalFill: function (addr, op, threshold, color, fontColor, opts) {
      opts = opts || {};
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var matrix = ET.getValues(addr);
      var p = Util.parseA1(parts.plain);
      /* 解析失败时明确报错，绝不回退到 A1 起点把颜色写到错误的位置 */
      if (!p) { throw new Error('条件高亮需要明确的区域：' + addr); }
      var scope = opts.fillScope === 'row' ? 'row' : 'cell';
      var lastCol = Math.max(p.c2, ET.rowLastCol(addr));
      var matched = 0;
      for (var r = 0; r < matrix.length; r++) {
        for (var c = 0; c < matrix[r].length; c++) {
          if (!matchCell(matrix[r][c], op, threshold)) { continue; }
          var row = p.r1 + r;
          /* 单个单元格 / 整行都要带上原来的工作表前缀，否则会标到当前活动表 */
          var target = scope === 'row'
            ? parts.prefix + 'A' + row + ':' + Util.numToCol(lastCol) + row
            : parts.prefix + Util.numToCol(p.c1 + c) + row;
          ET.setFill(target, color);
          if (fontColor) { ET.setFont(target, { color: fontColor }); }
          matched++;
        }
      }
      return matched;
    },

    /* 按条件删除整行：内容等于「是」、大于某数值等都能用 */
    deleteRowsIf: function (addr, op, value, keyColumn) {
      addr = resolveRange(addr);
      var hit = matchedRows(addr, op, value, keyColumn);
      var removed = deleteRowsByNumbers(hit.parts, hit.rows);
      return { removed: removed, range: addr };
    },

    /* columns 可指定排序列：按 C 列、或按多列先后排序，避免永远只按第一列 */
    sort: function (addr, order, columns) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      var cols = Object.prototype.toString.call(columns) === '[object Array]' ? columns : [];
      var desc = order === 'desc' ? 2 : 1;
      if (!cols.length) {
        var keyRange = Util.propOrCall(range.Columns, 'Item', [1]) || range;
        if (typeof range.Sort === 'function') { range.Sort(keyRange, desc, 1); }
        else { Util.propOrCall(range, 'Sort', [keyRange, desc, 1]); }
        return 0;
      }
      var args = [];
      var prefix = Util.splitSheetAddr(addr).prefix;
      for (var i = 0; i < Math.min(cols.length, 3); i++) {
        var letter = String(cols[i]).toUpperCase().replace(/[^A-Z]/g, '');
        if (!letter) { continue; }
        args.push(getRange(prefix + letter + '2'), desc);
      }
      if (!args.length) {
        var fallback = Util.propOrCall(range.Columns, 'Item', [1]) || range;
        if (typeof range.Sort === 'function') { range.Sort(fallback, desc, 1); }
        else { Util.propOrCall(range, 'Sort', [fallback, desc, 1]); }
        return 0;
      }
      while (args.length < 6) { args.push(null); }
      args.push(1);
      if (typeof range.Sort === 'function') { range.Sort.apply(range, args); }
      else { Util.propOrCall(range, 'Sort', args); }
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

    /* 按内容删除重复行：区域可跨表，可按某列或整行比较，从下往上删避免行号偏移 */
    deleteDuplicateRows: function (addr, opts) {
      opts = opts || {};
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      if (!p) { throw new Error('去重删除需要明确的区域：' + addr); }
      var matrix = ET.getValues(addr);
      var keepLast = opts.keep === 'last';
      var colIndex = -1;
      if (opts.keyColumn) {
        colIndex = /^\d+$/.test(String(opts.keyColumn))
          ? Number(opts.keyColumn) - 1
          : Util.colToNum(opts.keyColumn) - p.c1;
      }
      function keyOf(row) {
        var cells = matrix[row] || [];
        if (colIndex >= 0) {
          return String(cells[colIndex] == null ? '' : cells[colIndex]).trim();
        }
        var partsOfRow = [];
        for (var i = 0; i < cells.length; i++) {
          partsOfRow.push(String(cells[i] == null ? '' : cells[i]).trim());
        }
        return partsOfRow.join('\u0001');
      }
      var seen = {}, dupRows = [], order = [];
      for (var r = 0; r < matrix.length; r++) { order.push(r); }
      if (keepLast) { order.reverse(); }
      for (var j = 0; j < order.length; j++) {
        var idx = order[j];
        var key = keyOf(idx);
        if (!key) { continue; }
        if (seen[key]) { dupRows.push(p.r1 + idx); }
        else { seen[key] = 1; }
      }
      var removed = deleteRowsByNumbers(parts, dupRows);
      return { removed: removed, kept: matrix.length - removed, range: addr };
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

    /* ---- 高频数据清洗与结构操作 ---- */
    textToColumns: function (addr, delimiter) {
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      if (!p) { throw new Error('分列需要明确的区域：' + addr); }
      var sep = String(delimiter == null || delimiter === '' ? ',' : delimiter);
      var range = getRange(addr);
      if (typeof range.TextToColumns === 'function') {
        range.TextToColumns(range, 1, false, false, false, false, sep);
        return 0;
      }
      /* 没有 TextToColumns 接口时逐格拆分，写到该行右侧的空列里 */
      var matrix = ET.getValues(addr);
      var changed = 0;
      for (var r = 0; r < matrix.length; r++) {
        var line = matrix[r] || [];
        for (var c = 0; c < line.length; c++) {
          var text = String(line[c] == null ? '' : line[c]);
          if (text.indexOf(sep) < 0) { continue; }
          var bits = text.split(sep);
          for (var k = 0; k < bits.length; k++) {
            getRange(parts.prefix + Util.numToCol(p.c1 + c + k) + (p.r1 + r)).Value2 = bits[k].trim();
          }
          changed++;
        }
      }
      return changed;
    },

    /* 逐格替换文本；公式单元格保持不动，避免把公式结果写死成常量 */
    findReplace: function (addr, find, replacement) {
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      if (!p) { throw new Error('查找替换需要明确的区域：' + addr); }
      var needle = String(find == null ? '' : find);
      if (!needle) { throw new Error('没有给出要查找的内容'); }
      var to = String(replacement == null ? '' : replacement);
      var matrix = ET.getValues(addr);
      var changed = 0;
      for (var r = 0; r < matrix.length; r++) {
        for (var c = 0; c < matrix[r].length; c++) {
          var v = matrix[r][c];
          if (v === null || v === undefined) { continue; }
          var text = String(v);
          if (text.indexOf(needle) < 0) { continue; }
          var one = getRange(parts.prefix + Util.numToCol(p.c1 + c) + (p.r1 + r));
          var f = null;
          try { f = one.Formula; } catch (e) { /* 忽略 */ }
          if (typeof f === 'string' && f.charAt(0) === '=') { continue; }
          one.Value2 = text.split(needle).join(to);
          changed++;
        }
      }
      return changed;
    },

    fillBlanks: function (addr, value) {
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      if (!p) { throw new Error('填充空白需要明确的区域：' + addr); }
      var matrix = ET.getValues(addr);
      var filled = 0;
      for (var r = 0; r < matrix.length; r++) {
        for (var c = 0; c < matrix[r].length; c++) {
          var v = matrix[r][c];
          if (v !== null && v !== undefined && String(v).trim() !== '') { continue; }
          getRange(parts.prefix + Util.numToCol(p.c1 + c) + (p.r1 + r)).Value2 = value;
          filled++;
        }
      }
      return filled;
    },

    fillDown: function (addr) {
      addr = resolveRange(addr);
      var range = getRange(addr);
      if (typeof range.FillDown === 'function') { range.FillDown(); return 0; }
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      if (!p) { throw new Error('向下填充需要明确的区域：' + addr); }
      var src = getRange(parts.prefix + Util.numToCol(p.c1) + p.r1);
      var formula = '';
      try { formula = String(src.Formula || ''); } catch (e) { /* 忽略 */ }
      if (!formula) { return 0; }
      var filled = 0;
      for (var r = p.r1 + 1; r <= p.r2; r++) {
        var delta = r - p.r1;
        var shifted = formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, function (all, col, abs, row) {
          return abs === '$' ? all : col + abs + (Number(row) + delta);
        });
        getRange(parts.prefix + Util.numToCol(p.c1) + r).Formula = shifted;
        filled++;
      }
      return filled;
    },

    mergeCells: function (addr, center) {
      var range = getRange(resolveRange(addr));
      if (typeof range.Merge === 'function') { range.Merge(); }
      else { Util.propOrCall(range, 'Merge'); }
      if (center) { range.HorizontalAlignment = ALIGN.center; }
    },

    insertRows: function (addr, count) {
      var range = getRange(resolveRange(addr));
      var n = Math.max(1, Math.min(100, Number(count) || 1));
      var row = range.EntireRow;
      for (var i = 0; i < n; i++) {
        if (row && typeof row.Insert === 'function') { row.Insert(); }
        else { Util.propOrCall(row, 'Insert'); }
      }
    },

    insertColumns: function (addr, count) {
      var range = getRange(resolveRange(addr));
      var n = Math.max(1, Math.min(100, Number(count) || 1));
      var col = range.EntireColumn;
      for (var i = 0; i < n; i++) {
        if (col && typeof col.Insert === 'function') { col.Insert(); }
        else { Util.propOrCall(col, 'Insert'); }
      }
    },

    deleteColumns: function (addr) {
      var range = getRange(resolveRange(addr));
      var col = range.EntireColumn;
      if (col && typeof col.Delete === 'function') { col.Delete(); }
      else { Util.propOrCall(col, 'Delete'); }
    },

    hideColumns: function (addr, hidden) {
      var range = getRange(resolveRange(addr));
      var col = range.EntireColumn;
      if (!col) { throw new Error('当前客户端不支持隐藏列'); }
      col.Hidden = hidden !== false;
    },

    dataValidation: function (addr, list) {
      var values = Object.prototype.toString.call(list) === '[object Array]'
        ? list : String(list || '').split(/[,，、]/);
      var clean = [];
      for (var i = 0; i < values.length; i++) {
        var v = String(values[i]).trim();
        if (v) { clean.push(v); }
      }
      if (!clean.length) { throw new Error('没有给出下拉选项，请写成「设置下拉选项：是,否」'); }
      var range = getRange(resolveRange(addr));
      var validation = range.Validation;
      if (!validation) { throw new Error('当前客户端不支持数据验证'); }
      try { Util.propOrCall(validation, 'Delete'); } catch (e) { /* 忽略 */ }
      Util.propOrCall(validation, 'Add', [3, 1, 3, '"' + clean.join(',') + '"']);
      return clean.length;
    },

    pageSetup: function (addr, opts) {
      opts = opts || {};
      var parts = Util.splitSheetAddr(resolveRange(addr));
      /* 打印区域必须设在地址指定的工作表上，不能顺手写到当前活动表 */
      var name = String(parts.prefix || '').replace(/!$/, '').replace(/^'|'$/g, '').replace(/''/g, "'");
      var sheet = getSheet();
      var ext = /^\[([^\]]+)\](.+)$/.exec(name);
      if (ext) {
        var books = Util.propOrCall(info().app, 'Workbooks');
        var other = books && Util.propOrCall(books, 'Item', [ext[1]]);
        var wanted = findSheetByName(ext[2], other && Util.propOrCall(other, 'Worksheets'));
        if (!wanted) { throw new Error('未找到已打开的工作簿「' + ext[1] + '」中的工作表'); }
        sheet = wanted;
      } else if (name) {
        var found = findSheetByName(name);
        if (!found) { throw new Error('未找到工作表「' + name + '」'); }
        sheet = found;
      }
      var setup = sheet && sheet.PageSetup;
      if (!setup) { throw new Error('当前客户端不支持页面设置'); }
      setup.PrintArea = parts.plain;
      if (opts.titleRows) { setup.PrintTitleRows = String(opts.titleRows); }
    },

    conditionalFormatRule: function (addr, formula, color) {
      var range = getRange(resolveRange(addr));
      var fc = range.FormatConditions;
      if (!fc) { throw new Error('当前客户端不支持条件格式规则'); }
      var added = Util.propOrCall(fc, 'Add', [2, formula]);
      if (!added) { throw new Error('条件格式规则创建失败'); }
      if (color) { added.Interior.Color = colorValue(Util.normalizeColor(color)); }
    },

    pivotTable: function (addr, target) {
      var app = info().app;
      var wb = app && app.ActiveWorkbook;
      var caches = wb && Util.propOrCall(wb, 'PivotCaches');
      if (!caches) { throw new Error('当前客户端未提供数据透视表接口，请在表格里手动插入'); }
      var obj = typeof caches === 'function' ? caches.call(wb) : caches;
      var cache = Util.propOrCall(obj, 'Create', [1, resolveRange(addr)]);
      var created = cache && Util.propOrCall(cache, 'CreatePivotTable',
        [resolveRange(target || '@selectionBottom'), '智能表格透视表']);
      if (!created) { throw new Error('数据透视表创建失败，请在表格里手动插入'); }
      return true;
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

    /* ---- 撤回快照：逐单元格记录，只读取 / 恢复本次动作实际改动的属性 ----
     * props 为空表示兼容旧的全量模式；颜色类动作不含 fontColor，撤回底色时
     * 就不会把整片区域的字体色重写一遍。 */
    getState: function (addr, props) {
      addr = resolveRange(addr);
      var parts = Util.splitSheetAddr(addr);
      var p = Util.parseA1(parts.plain);
      var state = { address: addr, live: true, cells: [], props: props || null };
      function has(name) { return !props || props.indexOf(name) >= 0; }
      function read(obj, prop) {
        try {
          if (!obj) { return null; }
          var v = obj[prop];
          return (v === undefined || v === null || v === '') ? null : v;
        } catch (e) { return null; }
      }
      if (!p) {
        state.values = Util.toMatrix(getRange(addr).Value2);
        return state;
      }
      var rows = Math.min(p.r2 - p.r1 + 1, 200);
      /* 与整行标色的实际写入列数对齐，撤回时不能只还原前 20 列 */
      var cols = Math.min(p.c2 - p.c1 + 1, 256);
      for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
          var cellAddr = parts.prefix + Util.numToCol(p.c1 + c) + (p.r1 + r);
          var range = getRange(cellAddr);
          var rec = { addr: cellAddr, value: null, formula: null, style: {} };
          try {
            if (has('value')) { rec.value = range.Value2; rec.formula = range.Formula; }
          } catch (e1) { /* 忽略 */ }
          try {
            if (has('fill')) {
              var interior = range.Interior;
              rec.style.color = read(interior, 'Color');
              rec.style.pattern = read(interior, 'Pattern');
              if (rec.style.color !== null) { rec.style.color = Util.normalizeColor(rec.style.color); }
            }
          } catch (e3) { /* 忽略 */ }
          try {
            if (has('fontColor') || has('bold') || has('italic') || has('size') || has('fontName')) {
              var font = range.Font;
              if (has('bold')) { rec.style.bold = read(font, 'Bold'); }
              if (has('italic')) { rec.style.italic = read(font, 'Italic'); }
              if (has('size')) { rec.style.size = read(font, 'Size'); }
              /* 字体色保留宿主原始值：自动色 0 不能被转成一个具体的实色 */
              if (has('fontColor')) { rec.style.fontColor = read(font, 'Color'); }
              if (has('fontName')) { rec.style.fontName = read(font, 'Name'); }
            }
          } catch (e4) { /* 忽略 */ }
          try { if (has('numberFormat')) { rec.style.numberFormat = read(range, 'NumberFormat'); } } catch (e5) { /* 忽略 */ }
          try { if (has('align')) { rec.style.align = read(range, 'HorizontalAlignment'); } } catch (e6) { /* 忽略 */ }
          try { if (has('width')) { rec.style.width = read(range, 'ColumnWidth'); } } catch (e7) { /* 忽略 */ }
          state.cells.push(rec);
        }
      }
      return state;
    },

    setState: function (state) {
      if (!state || !state.address) { return; }
      if (!state.cells || !state.cells.length) {
        try { if (state.values) { getRange(state.address).Value2 = state.values; } } catch (e) { /* 忽略 */ }
        return;
      }
      var props = state.props || null;
      function want(name) { return !props || props.indexOf(name) >= 0; }
      for (var i = 0; i < state.cells.length; i++) {
        var rec = state.cells[i];
        var st = rec.style || {};
        var range = getRange(rec.addr);
        if (want('fill')) {
          /* 先还原填充状态：原本「无填充」必须显式清掉后来加上的底色 */
          var noFill = st.pattern === null || st.pattern === undefined ||
            st.pattern === -4142 || st.pattern === 0;
          try {
            var interior = range.Interior;
            if (noFill) {
              try { interior.Pattern = -4142; }
              catch (eNone) { interior.Color = colorValue('#FFFFFF'); }
            } else {
              interior.Pattern = st.pattern;
              if (st.color !== null && st.color !== undefined) { interior.Color = colorValue(st.color); }
            }
          } catch (e2) { /* 忽略 */ }
        }
        if (want('fontColor') || want('bold') || want('italic') || want('size') || want('fontName')) {
          try {
            var font = range.Font;
            if (want('bold') && st.bold !== null && st.bold !== undefined) { font.Bold = st.bold; }
            if (want('italic') && st.italic !== null && st.italic !== undefined) { font.Italic = st.italic; }
            if (want('size') && st.size !== null && st.size !== undefined) { font.Size = st.size; }
            if (want('fontName') && st.fontName !== null && st.fontName !== undefined) { font.Name = st.fontName; }
            /* 原样写回，保留「自动色」语义，不经过 hex / BGR 来回转换 */
            if (want('fontColor') && st.fontColor !== null && st.fontColor !== undefined) { font.Color = st.fontColor; }
          } catch (e3) { /* 忽略 */ }
        }
        try { if (want('numberFormat') && st.numberFormat !== null && st.numberFormat !== undefined) { range.NumberFormat = st.numberFormat; } } catch (e4) { /* 忽略 */ }
        try { if (want('align') && st.align !== null && st.align !== undefined) { range.HorizontalAlignment = st.align; } } catch (e5) { /* 忽略 */ }
        try { if (want('width') && st.width !== null && st.width !== undefined) { range.ColumnWidth = st.width; } } catch (e6) { /* 忽略 */ }
        if (want('value')) {
          try {
            if (typeof rec.formula === 'string' && rec.formula.charAt(0) === '=') {
              range.Formula = rec.formula;
            } else {
              range.Value2 = rec.value == null ? '' : rec.value;
            }
          } catch (e7) { /* 忽略 */ }
        }
      }
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
