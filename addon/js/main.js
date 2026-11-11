/*
 * Add-in interface functions (callbacks for ribbon.xml onAction / onLoad).
 * ASCII-only on purpose: this file may be loaded by a host-generated index.html
 * whose charset is out of our control. Localized strings live in the task pane.
 * APIs: api.CreateTaskPane(url, title), api.PluginStorage, api.GetTaskPane(id)
 */
if (!window.__SMART_SHEET_MAIN_LOADED__) {
  window.__SMART_SHEET_MAIN_LOADED__ = true;

  (function () {
    var PANE_KEY = 'smartsheet_taskpane_id';
    /* Unicode escapes keep this file ASCII while showing a Chinese pane title */
    var PANE_TITLE = '\u667a\u80fd\u8868\u683c\u52a9\u624b';

    function getApi() {
      if (window.wps) { return window.wps; }
      try { return parent.wps; } catch (e) { /* cross-frame guard */ }
      try { return top.wps; } catch (e2) { /* cross-frame guard */ }
      return null;
    }

    function storageGet(key) {
      var api = getApi();
      if (!api || !api.PluginStorage) { return null; }
      try { return api.PluginStorage.getItem(key); } catch (e) { return null; }
    }

    function storageSet(key, value) {
      var api = getApi();
      if (!api || !api.PluginStorage) { return; }
      try { api.PluginStorage.setItem(key, value); } catch (e) { /* ignore */ }
    }

    /* Add-in root = folder of the auto-generated index.html */
    function getRoot() {
      var url = String(document.location);
      var q = url.indexOf('?');
      if (q >= 0) { url = url.substring(0, q); }
      var h = url.indexOf('#');
      if (h >= 0) { url = url.substring(0, h); }
      var slash = url.lastIndexOf('/');
      return slash >= 0 ? url.substring(0, slash + 1) : url + '/';
    }

    function getTaskPane(api, id) {
      if (!id) { return null; }
      try {
        if (typeof api.GetTaskPane === 'function') { return api.GetTaskPane(id); }
        if (typeof api.GetTaskpane === 'function') { return api.GetTaskpane(id); }
      } catch (e) { /* stale id */ }
      return null;
    }

    function createTaskPane(api, url) {
      var pane = null;
      try {
        if (typeof api.CreateTaskPane === 'function') { pane = api.CreateTaskPane(url, PANE_TITLE); }
        else if (typeof api.CreateTaskpane === 'function') { pane = api.CreateTaskpane(url, PANE_TITLE); }
      } catch (e) { /* try a simpler signature */ }
      if (!pane) {
        try {
          if (typeof api.CreateTaskPane === 'function') { pane = api.CreateTaskPane(url); }
          else if (typeof api.CreateTaskpane === 'function') { pane = api.CreateTaskpane(url); }
        } catch (e2) { /* ignore */ }
      }
      return pane;
    }

    function showTaskPane() {
      var api = getApi();
      if (!api) {
        window.alert('Spreadsheet host not detected. Please open a workbook first.');
        return false;
      }

      var pane = getTaskPane(api, storageGet(PANE_KEY));
      if (pane) {
        try { pane.Visible = true; return true; } catch (e) { /* recreate */ }
      }

      var url = getRoot() + 'ui/taskpane.html';
      pane = createTaskPane(api, url);

      /* Fallback: absolute local path (some builds reject relative / http URLs) */
      if (!pane && api.Env && typeof api.Env.GetAppDataPath === 'function') {
        try {
          var appData = String(api.Env.GetAppDataPath()).replace(/\\/g, '/');
          var dirMatch = /\/jsaddons\/([^/]+)\//i.exec(getRoot());
          if (dirMatch) {
            var local = appData + '/kingsoft/wps/jsaddons/' + dirMatch[1] + '/ui/taskpane.html';
            pane = createTaskPane(api, local);
          }
        } catch (e3) { /* ignore */ }
      }

      if (!pane) {
        window.alert('Failed to create task pane: ' + url + '\nPlease reinstall the add-in and restart the spreadsheet app.');
        return false;
      }

      try { if (pane.ID) { storageSet(PANE_KEY, pane.ID); } } catch (e4) { /* ignore */ }
      try { pane.Width = 390; } catch (e5) { /* ignore */ }
      try { pane.DockPosition = 2; } catch (e6) { /* dock right */ }
      try { pane.Visible = true; } catch (e7) { /* ignore */ }
      return true;
    }

    /* ---- ribbon.xml callbacks ---- */

    window.OnAddinLoad = function (ribbonUI) {
      try { getApi().ribbonUI = ribbonUI; } catch (e) { /* ignore */ }
    };

    window.GetVisible = function () { return true; };
    window.GetEnabled = function () { return true; };
    window.GetImage = function () { return ''; };

    /* Only one button remains: clicking it opens the sidebar directly. */
    window.OnAction = function () { return showTaskPane(); };

    /* available to the task pane page as well */
    window.SmartSheet_ShowTaskPane = showTaskPane;
    window.SmartSheet_GetApi = getApi;
  })();
}
