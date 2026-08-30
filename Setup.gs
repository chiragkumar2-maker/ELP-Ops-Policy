/** One-time, idempotent initialization of the System Config spreadsheet. */

function setupSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Open/bind this script to your ELP System Config spreadsheet, then run setupSystem() from there.');
  }
  PropertiesService.getScriptProperties().setProperty('SYSTEM_CONFIG_SPREADSHEET_ID', spreadsheet.getId());

  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONFIG, ELP_HEADERS.CONFIG);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EMPLOYEES, ELP_HEADERS.EMPLOYEES);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EMPLOYEES_LAST_RUN, ELP_HEADERS.EMPLOYEES);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EXITED_EMPLOYEES, ELP_HEADERS.EXITED_EMPLOYEES);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTACTS, ELP_HEADERS.CONTACTS);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EXITED_HRBPS, ELP_HEADERS.EXITED_HRBPS);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.HRBP_REGISTRY, ELP_HEADERS.HRBP_REGISTRY);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.LOGS, ELP_HEADERS.LOGS);
  ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTROL_PANEL, ELP_HEADERS.CONTROL_PANEL);

  initializeConfigRows_(spreadsheet.getSheetByName(ELP_SHEETS.CONFIG));

  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1) spreadsheet.deleteSheet(defaultSheet);

  return { ready: true, spreadsheetId: spreadsheet.getId() };
}

function initializeConfigRows_(sheet) {
  const table = readTable_(sheet, ELP_HEADERS.CONFIG);
  const existingKeys = new Set(table.rows.map((row) => normalizeId_(row['Key'])));
  Object.keys(ELP_DEFAULT_CONFIG).forEach((key) => {
    if (existingKeys.has(key)) return;
    appendObjectRow_(sheet, ELP_HEADERS.CONFIG, {
      Key: key, Value: ELP_DEFAULT_CONFIG[key], Description: ELP_CONFIG_DESCRIPTIONS[key] || ''
    });
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ELP Connects for Ops')
    .addItem('Run daily append now', 'runDailyAppend')
    .addItem('Run weekly rebuild now', 'runWeeklyRebuild')
    .addSeparator()
    .addItem('Install daily trigger (2 AM)', 'installDailyTrigger')
    .addItem('Install weekly trigger (Sun 3 AM)', 'installWeeklyTrigger')
    .addItem('Remove daily trigger', 'removeDailyTrigger')
    .addItem('Remove weekly trigger', 'removeWeeklyTrigger')
    .addSeparator()
    .addItem('Clear stuck job (emergency)', 'forceReleaseJobSlot')
    .addToUi();
}
