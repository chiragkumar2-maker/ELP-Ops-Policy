/** Persistent operational logging and one-email-per-run alerting. */

function createRunContext_(kind) {
  return {
    runId: uuid_(),
    kind: kind,
    startedAt: now_(),
    entries: [],
    errors: [],
    metrics: {}
  };
}

function recordLog_(context, severity, action, entity, outcome, details) {
  context.entries.push([now_(), context.runId, severity, action, entity, outcome, normalizeId_(details)]);
  if (severity === 'ERROR') context.errors.push(`${action} — ${entity}: ${details}`);
}

function recordError_(context, action, entity, error) {
  const message = error && error.message ? error.message : String(error);
  recordLog_(context, 'ERROR', action, entity, 'Failed', message);
  console.error(`${action} — ${entity}: ${message}`);
}

function flushLogs_(context) {
  if (!context.entries.length) return;
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.LOGS, ELP_HEADERS.LOGS);
  const startRow = sheet.getLastRow() + 1;
  ensureGridCapacity_(sheet, startRow + context.entries.length - 1, ELP_HEADERS.LOGS.length);
  sheet.getRange(startRow, 1, context.entries.length, ELP_HEADERS.LOGS.length).setValues(context.entries);
  context.entries = [];
}

function sendAlertIfNeeded_(context, config, extraMessage) {
  const to = config && config.ALERT_EMAIL;
  if (!to) return;
  const durationSec = Math.round((now_() - context.startedAt) / 1000);
  const status = context.errors.length ? 'FAILED' : 'OK';
  const subject = `ELP Connects for Ops — ${context.kind} — ${status}`;
  const lines = [`Run ID: ${context.runId}`, `Status: ${status}`, `Duration: ${durationSec}s`];
  Object.keys(context.metrics || {}).forEach((key) => lines.push(`${key}: ${context.metrics[key]}`));
  if (extraMessage) { lines.push(''); lines.push(extraMessage); }
  if (context.errors.length) {
    lines.push('', 'Errors:');
    context.errors.slice(0, 25).forEach((err) => lines.push(`- ${err}`));
    if (context.errors.length > 25) lines.push(`...and ${context.errors.length - 25} more. See the Logs tab.`);
  }
  MailApp.sendEmail(to, subject, lines.join('\n'));
}

function updateControlPanel_(context, extraMetrics) {
  Object.assign(context.metrics, extraMetrics || {});
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTROL_PANEL, ELP_HEADERS.CONTROL_PANEL);
  const table = readTable_(sheet, ELP_HEADERS.CONTROL_PANEL);
  const byMetric = new Map(table.rows.map((row) => [row['Metric'], row]));
  const updatedAt = now_();
  Object.keys(context.metrics).forEach((metric) => {
    const value = context.metrics[metric];
    const existing = byMetric.get(metric);
    if (existing) {
      sheet.getRange(existing.__rowNumber, 2, 1, 2).setValues([[value, updatedAt]]);
    } else {
      appendObjectRow_(sheet, ELP_HEADERS.CONTROL_PANEL, { Metric: metric, Value: value, 'Updated At': updatedAt });
    }
  });
}
