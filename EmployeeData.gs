/**
 * Employee Data and HRBP Contacts reading. You paste directly into the
 * "Employee Data" tab — this reads it as-is every run. No staging, no
 * publish step, no versioning.
 */

function readEmployeeData_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(ELP_SHEETS.EMPLOYEES);
  if (!sheet) throw new Error(`Missing "${ELP_SHEETS.EMPLOYEES}" tab. Run setupSystem().`);
  return readEmployeeTable_(sheet);
}

/** Shared by Employee Data and Employee Data (Last Run) — same headers. */
function readEmployeeTable_(sheet) {
  const table = readTable_(sheet, ELP_HEADERS.EMPLOYEES);
  const byEcode = new Map();
  table.rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_ECode']);
    if (!ecode) return;
    byEcode.set(ecode, {
      ecode: ecode,
      empName: normalizeId_(row['EmpName']),
      uniqueId: normalizeId_(row['__display_UniqueID']),
      doj: row['__display_DOJ'],
      level: normalizeId_(row['Level']),
      hrbpEcode: normalizeId_(row['__display_HRSpocEcode']),
      hrbpName: normalizeId_(row['HRBPSPOCNAME']),
      hrbpEmail: normalizeEmail_(row['HRBP Email'])
    });
  });
  return byEcode;
}

/** HRBP ECode -> array of up to 4 viewer emails (Manager / Manager's Manager / Additional / HRBP Self). */
function readHrbpContacts_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTACTS, ELP_HEADERS.CONTACTS);
  const table = readTable_(sheet, ELP_HEADERS.CONTACTS);
  const byEcode = new Map();
  table.rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_HRBP ECode']);
    if (!ecode) return;
    const emails = [row['Manager Email'], row["Manager's Manager Email"], row['Additional Email'], row['HRBP Self Email']]
      .map(normalizeEmail_)
      .filter(Boolean);
    byEcode.set(ecode, emails);
  });
  return byEcode;
}
