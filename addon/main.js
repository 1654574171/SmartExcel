/*
 * Add-in entry shim.
 * Different host builds generate index.html that loads either ./main.js or ./js/main.js.
 * Both paths end up in js/main.js, which defines the ribbon callbacks.
 * ASCII-only because the generated index.html charset is not under our control.
 */
(function () {
  if (window.__SMART_SHEET_ENTRY_LOADED__) { return; }
  window.__SMART_SHEET_ENTRY_LOADED__ = true;
  document.write('<scr' + 'ipt type="text/javascript" charset="utf-8" src="js/main.js"><\/scr' + 'ipt>');
})();
