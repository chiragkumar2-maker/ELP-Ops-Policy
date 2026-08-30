/**
 * Keeps "Exited Employees" in sync by diffing today's Employee Data against
 * "Employee Data (Last Run)" — a snapshot the script itself maintains, taken
 * at the end of the last successful run. You never edit either of these two
 * system tabs.
 *
 * - Someone who drops out of Employee Data is frozen into Exited Employees
 *   at their last known HRBP, so the nightly full rebuild keeps placing
 *   their past responses instead of losing them.
 * - Someone who reappears in Employee Data is removed from Exited Employees
 *   and goes back to normal (live) routing.
 * - Reassignment (HRBP change while still active) needs no special handling
 *   here — the full rebuild in ResponseTransfer.gs always routes an active
 *   employee to their *current* HRBP, so the "move" happens automatically.
 *   This function just logs that it happened.
 */
function reconcileEmployeeLifecycle_(context, currentByEcode) {
  const spreadsheet = getSystemSpreadsheet_();
  const lastRunSheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EMPLOYEES_LAST_RUN, ELP_HEADERS.EMPLOYEES);
  const lastRunByEcode = readEmployeeTable_(lastRunSheet);
  const exited = readExitedEmployees_();

  let newlyExited = 0;
  let rejoined = 0;
  let reassigned = 0;

  // Newly exited: present last run, absent now, not already frozen.
  lastRunByEcode.forEach((employee, ecode) => {
    if (currentByEcode.has(ecode) || exited.byEcode.has(ecode)) return;
    appendObjectRow_(exited.sheet, ELP_HEADERS.EXITED_EMPLOYEES, {
      ECode: employee.ecode,
      EmpName: employee.empName,
      UniqueID: employee.uniqueId,
      DOJ: employee.doj,
      Level: employee.level,
      'Last Known HRBP ECode': employee.hrbpEcode,
      'Last Known HRBP Name': employee.hrbpName,
      'Last Known HRBP Email': employee.hrbpEmail,
      'Exited Date': todayIst_()
    });
    exited.byEcode.set(ecode, employee);
    newlyExited += 1;
    recordLog_(context, 'INFO', 'Employee exit detected', ecode, 'Frozen',
      `No longer in Employee Data. Frozen at last known HRBP ${employee.hrbpEcode || '(none)'}.`);
  });

  // Rejoined: present now, still sitting in Exited Employees — un-freeze.
  const rowsToDelete = [];
  exited.byEcode.forEach((exitedEmployee, ecode) => {
    if (!currentByEcode.has(ecode)) return;
    rowsToDelete.push(exitedEmployee.__rowNumber);
    exited.byEcode.delete(ecode);
    rejoined += 1;
    recordLog_(context, 'INFO', 'Employee rejoined', ecode, 'Unfrozen', 'Back in Employee Data; routing resumes normally.');
  });
  rowsToDelete.sort((a, b) => b - a).forEach((rowNumber) => exited.sheet.deleteRow(rowNumber));

  // Reassignment: HRBP changed while the employee stayed active. Log-only —
  // the full rebuild handles the actual move.
  currentByEcode.forEach((employee, ecode) => {
    const previous = lastRunByEcode.get(ecode);
    if (previous && previous.hrbpEcode && employee.hrbpEcode && previous.hrbpEcode !== employee.hrbpEcode) {
      reassigned += 1;
      recordLog_(context, 'INFO', 'HRBP reassignment', ecode, 'Moved',
        `${previous.hrbpEcode} → ${employee.hrbpEcode}. Full history moves to the new HRBP's workbook this run.`);
    }
  });

  return { exitedMap: exited.byEcode, newlyExited: newlyExited, rejoined: rejoined, reassigned: reassigned };
}

function readExitedEmployees_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EXITED_EMPLOYEES, ELP_HEADERS.EXITED_EMPLOYEES);
  const table = readTable_(sheet, ELP_HEADERS.EXITED_EMPLOYEES);
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
      hrbpEcode: normalizeId_(row['__display_Last Known HRBP ECode']),
      hrbpName: normalizeId_(row['Last Known HRBP Name']),
      hrbpEmail: normalizeEmail_(row['Last Known HRBP Email']),
      exitedDate: row['Exited Date'],
      __rowNumber: row.__rowNumber
    });
  });
  return { sheet: sheet, byEcode: byEcode };
}

/** Only called after a fully successful run — keeps the diff baseline honest. */
function commitEmployeeLastRunSnapshot_(currentByEcode) {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EMPLOYEES_LAST_RUN, ELP_HEADERS.EMPLOYEES);
  const rows = [];
  currentByEcode.forEach((employee) => {
    rows.push(ELP_HEADERS.EMPLOYEES.map((header) => {
      switch (header) {
        case 'ECode': return employee.ecode;
        case 'EmpName': return employee.empName;
        case 'UniqueID': return employee.uniqueId;
        case 'DOJ': return employee.doj;
        case 'Level': return employee.level;
        case 'HRSpocEcode': return employee.hrbpEcode;
        case 'HRBPSPOCNAME': return employee.hrbpName;
        case 'HRBP Email': return employee.hrbpEmail;
        default: return '';
      }
    }));
  });
  writeRows_(sheet, ELP_HEADERS.EMPLOYEES, rows);
}
