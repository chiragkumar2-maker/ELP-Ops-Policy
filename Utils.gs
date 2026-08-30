/** General-purpose pure helpers and small Google Sheets adapters. */

function normalizeId_(value) {
  return String(value == null ? '' : value).trim();
}

function normalizeEmail_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeId_(value));
}

function uuid_() {
  return Utilities.getUuid();
}

function now_() {
  return new Date();
}

function todayIst_(date) {
  return Utilities.formatDate(date || new Date(), ELP_TIMEZONE, 'yyyy-MM-dd');
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const key = normalizeId_(header);
    if (key) map[key] = index;
  });
  return map;
}

function assertRequiredHeaders_(headers, required) {
  const map = buildHeaderMap_(headers);
  const missing = required.filter((header) => !(header in map));
  if (missing.length) throw new Error(`Missing required column(s): ${missing.join(', ')}`);
  return map;
}

function ensureGridCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

/**
 * Wrap text, left-align, top-align — applied to the sheet's FULL grid (every
 * row/column up to current max), not just the used range. That way rows
 * added later by ensureGridCapacity_ are already covered without needing to
 * reformat again. Cheap to call repeatedly — re-applying the same formatting
 * to already-formatted cells is a no-op in effect.
 */
function applyStandardFormatting_(sheet) {
  const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  range.setWrap(true);
  range.setHorizontalAlignment('left');
  range.setVerticalAlignment('top');
}

function ensureSheetWithHeaders_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getDisplayValues()[0];
  const matches = headers.every((header, index) => normalizeId_(currentHeaders[index]) === header);
  if (!matches) {
    ensureGridCapacity_(sheet, 1, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Reads a sheet into { headers, rows, display }.
 * Each row object exposes header values by bracket key (e.g. row['HRBP ECode'])
 * plus a raw display-string twin under '__display_<Header>', and '__rowNumber'
 * (1-indexed sheet row) for in-place writes/deletes.
 */
function readTable_(sheet, requiredHeaders) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const display = dataRange.getDisplayValues();
  const headers = (display[0] || []).map((value) => normalizeId_(value));
  assertRequiredHeaders_(headers, requiredHeaders || []);

  const rows = [];
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const displayRow = display[i];
    const hasContent = row.some((value) => normalizeId_(value)) || displayRow.some((value) => normalizeId_(value));
    if (!hasContent) continue;
    const object = { __rowNumber: i + 1 };
    headers.forEach((header, column) => {
      if (!header) return;
      object[header] = row[column];
      object[`__display_${header}`] = displayRow[column];
    });
    rows.push(object);
  }
  return { headers: headers, rows: rows };
}

/** Wipes all data rows (row 2 down) and writes fresh rows in header order. */
function writeRows_(sheet, headers, rows) {
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  }
  if (!rows.length) return;
  ensureGridCapacity_(sheet, rows.length + 1, headers.length);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function appendObjectRow_(sheet, headers, object) {
  const row = headers.map((header) => (object[header] === undefined || object[header] === null ? '' : object[header]));
  const targetRow = sheet.getLastRow() + 1;
  ensureGridCapacity_(sheet, targetRow, headers.length);
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
}

function writeObjectRowAt_(sheet, headers, object, rowNumber) {
  const row = headers.map((header) => (object[header] === undefined || object[header] === null ? '' : object[header]));
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function parseApprovedDomains_(value) {
  return new Set(normalizeId_(value).split(',').map((part) => part.trim().toLowerCase()).filter(Boolean));
}

function isApprovedDomain_(email, approvedDomains) {
  const domain = (normalizeEmail_(email).split('@')[1] || '');
  return approvedDomains.has(domain);
}

/** Moves a Drive file into exactly one target folder (adds if missing, removes other parents). */
function moveFileToFolder_(file, folder) {
  const targetId = folder.getId();
  const parents = file.getParents();
  let alreadyThere = false;
  const toRemove = [];
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() === targetId) alreadyThere = true;
    else toRemove.push(parent);
  }
  if (!alreadyThere) folder.addFile(file);
  toRemove.forEach((parent) => parent.removeFile(file));
}
