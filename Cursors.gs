/**
 * Per-response-tab "how far the daily append has read" watermark (a row
 * number). Small footprint, stored in Script Properties — you never edit
 * this directly. The weekly rebuild fast-forwards it after every full
 * rewrite so the next daily append doesn't re-append rows weekly already
 * wrote fresh.
 */

function getResponseCursors_() {
  const raw = PropertiesService.getScriptProperties().getProperty('ELP_RESPONSE_CURSORS');
  const stored = raw ? JSON.parse(raw) : {};
  const cursors = {};
  ELP_RESPONSE_TABS.forEach((tabName) => { cursors[tabName] = Number(stored[tabName] || 1); });
  return cursors;
}

function commitResponseCursors_(targetRowByTab) {
  PropertiesService.getScriptProperties().setProperty('ELP_RESPONSE_CURSORS', JSON.stringify(targetRowByTab));
}
