/**
 * Static names, schema, and the Config tab (Key/Value) loader.
 *
 * Simplified architecture: no staging/publish/versioning, no exceptions
 * workflow, no workbook "generations", no Drive permission-graph
 * verification. See the Logs tab / the run-summary email for what
 * happened on each run of runDailyAppend() or runWeeklyRebuild().
 */

const ELP_TIMEZONE = 'Asia/Kolkata';

const ELP_SHEETS = Object.freeze({
  EMPLOYEES: 'Employee Data',
  EMPLOYEES_LAST_RUN: 'Employee Data (Last Run)',
  EXITED_EMPLOYEES: 'Exited Employees',
  CONTACTS: 'HRBP Contacts',
  EXITED_HRBPS: 'Exited HRBPs',
  HRBP_REGISTRY: 'HRBP Registry',
  CONFIG: 'Config',
  LOGS: 'Logs',
  CONTROL_PANEL: 'Control Panel'
});

// The 9 Master Response Spreadsheet tabs (read-only source of truth for
// responses). Agent/Supervisor tabs route by looking up "Employee Code" in
// Employee Data. HRBP tabs route directly: "Employee Code" IS the HRBP's own
// ECode (see isHrbpResponseTab_ in Domain.gs).
const ELP_RESPONSE_TABS = Object.freeze([
  'Agent Day 30', 'Agent Day 60', 'Agent Day 90',
  'Supervisor Day 30', 'Supervisor Day 60', 'Supervisor Day 90',
  'HRBP Day 30', 'HRBP Day 60', 'HRBP Day 90'
]);

const ELP_SOURCE_HEADERS = Object.freeze({
  TIMESTAMP: 'Timestamp',
  EMPLOYEE_CODE: 'Employee Code',
  EMPLOYEE_NAME: 'Employee Name'
});

const ELP_HEADERS = Object.freeze({
  EMPLOYEES: Object.freeze([
    'ECode', 'EmpName', 'UniqueID', 'DOJ', 'Level', 'HRSpocEcode', 'HRBPSPOCNAME', 'HRBP Email'
  ]),
  CONTACTS: Object.freeze([
    'HRBP ECode', 'Manager Email', "Manager's Manager Email", 'Additional Email', 'HRBP Self Email'
  ]),
  EXITED_HRBPS: Object.freeze([
    'HRBP ECode', 'Manager Email', "Manager's Manager Email", 'Additional Email', 'HRBP Self Email', 'Exited Date'
  ]),
  EXITED_EMPLOYEES: Object.freeze([
    'ECode', 'EmpName', 'UniqueID', 'DOJ', 'Level',
    'Last Known HRBP ECode', 'Last Known HRBP Name', 'Last Known HRBP Email', 'Exited Date'
  ]),
  HRBP_REGISTRY: Object.freeze([
    'HRBP ECode', 'HRBP Name', 'Workbook ID', 'Workbook URL', 'Created Date', 'Last Synced'
  ]),
  CONFIG: Object.freeze(['Key', 'Value', 'Description']),
  LOGS: Object.freeze(['Timestamp', 'Run ID', 'Severity', 'Action', 'Entity', 'Outcome', 'Details']),
  CONTROL_PANEL: Object.freeze(['Metric', 'Value', 'Updated At'])
});

const ELP_DEFAULT_CONFIG = Object.freeze({
  MASTER_RESPONSE_SPREADSHEET_ID: '',
  HRBP_WORKBOOKS_FOLDER_ID: '',
  ALERT_EMAIL: 'chiragkumar2@policybazaar.com',
  APPROVED_VIEWER_DOMAINS: 'policybazaar.com'
});

const ELP_CONFIG_DESCRIPTIONS = Object.freeze({
  MASTER_RESPONSE_SPREADSHEET_ID: 'The Master Response Spreadsheet (Google Forms responses). Read-only.',
  HRBP_WORKBOOKS_FOLDER_ID: 'Drive folder where per-HRBP workbooks are created and kept.',
  ALERT_EMAIL: 'Where the one run-summary email goes after each run.',
  APPROVED_VIEWER_DOMAINS: 'Comma-separated email domains allowed to receive viewer access, e.g. policybazaar.com'
});

function getSystemSpreadsheet_() {
  const propertyId = PropertiesService.getScriptProperties().getProperty('SYSTEM_CONFIG_SPREADSHEET_ID');
  if (propertyId) return SpreadsheetApp.openById(propertyId);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('SYSTEM_CONFIG_SPREADSHEET_ID is not set. Run setupSystem() from the bound Config spreadsheet first.');
  }
  return active;
}

function loadConfig_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(ELP_SHEETS.CONFIG);
  if (!sheet) throw new Error('Missing Config tab. Run setupSystem().');

  const values = sheet.getDataRange().getDisplayValues();
  const config = Object.assign({}, ELP_DEFAULT_CONFIG);
  const seen = new Set();
  for (let i = 1; i < values.length; i += 1) {
    const key = normalizeId_(values[i][0]);
    if (!key) continue;
    if (seen.has(key)) throw new Error(`Duplicate Config key: ${key}`);
    seen.add(key);
    config[key] = normalizeId_(values[i][1]);
  }
  config.SYSTEM_CONFIG_SPREADSHEET_ID = spreadsheet.getId();
  return config;
}

function requireRuntimeConfig_(config) {
  const required = [
    'MASTER_RESPONSE_SPREADSHEET_ID', 'HRBP_WORKBOOKS_FOLDER_ID', 'ALERT_EMAIL', 'APPROVED_VIEWER_DOMAINS'
  ];
  const missing = required.filter((key) => !normalizeId_(config[key]));
  if (missing.length) throw new Error(`Missing required Config values: ${missing.join(', ')}`);
  if (!isValidEmail_(config.ALERT_EMAIL)) throw new Error('ALERT_EMAIL must be a valid email address.');
  const approvedDomains = parseApprovedDomains_(config.APPROVED_VIEWER_DOMAINS);
  if (!approvedDomains.size) throw new Error('APPROVED_VIEWER_DOMAINS must contain at least one domain.');
}
