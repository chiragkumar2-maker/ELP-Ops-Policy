/**
 * Reads Master Response rows and decides where each one belongs.
 *
 * One shared scanner, two calling conventions:
 *  - Daily append passes its committed cursor as "from" and a frozen
 *    row-number snapshot as "to", so it only sees genuinely new rows.
 *  - Weekly rebuild passes {} as "from" and null as "to", so it always
 *    sees every row in the Master sheet.
 *
 * Every destination ECode this returns is guaranteed to already be a
 * member of hrbpUniverse: Agent/Supervisor rows route via hrbpEcode values
 * that came from Employee Data / Exited Employees (exactly how
 * hrbpUniverse was built), and HRBP-tab rows are only accepted when the
 * code is already a hrbpUniverse member.
 */
function resolveDestination_(isHrbpTab, employeeCode, activeByEcode, exitedByEcode, hrbpUniverse) {
  if (!employeeCode) return { destinationEcode: null, failReason: 'Blank Employee Code.' };

  if (isHrbpTab) {
    if (hrbpUniverse.has(employeeCode)) return { destinationEcode: employeeCode, failReason: '' };
    return {
      destinationEcode: null,
      failReason: `"${employeeCode}" is not a recognized HRBP ECode (add them to Employee Data or HRBP Contacts).`
    };
  }

  const active = activeByEcode.get(employeeCode);
  if (active && active.hrbpEcode) return { destinationEcode: active.hrbpEcode, failReason: '' };
  if (active && !active.hrbpEcode) {
    return { destinationEcode: null, failReason: `${employeeCode} has no HRSpocEcode set in Employee Data.` };
  }
  const exited = exitedByEcode.get(employeeCode);
  if (exited && exited.hrbpEcode) return { destinationEcode: exited.hrbpEcode, failReason: '' };
  return { destinationEcode: null, failReason: `Employee Code "${employeeCode}" not found in Employee Data or Exited Employees.` };
}

/**
 * Scans all 9 Master tabs from fromCursorByTab (exclusive) to toTargetByTab
 * (inclusive; pass null to mean "scan to whatever the current last row is").
 * Returns placements grouped by destination HRBP ECode, ready to append or
 * to fully rewrite a workbook's tabs from.
 */
function readNewResponsesSinceCursor_(config, fromCursorByTab, toTargetByTab, activeByEcode, exitedByEcode, hrbpUniverse) {
  const master = SpreadsheetApp.openById(config.MASTER_RESPONSE_SPREADSHEET_ID);
  const placementsByEcode = new Map();
  const failures = [];
  const targetRowByTab = {};
  const schemas = new Map();
  let totalRead = 0;
  let totalPlaced = 0;
  let totalFailed = 0;

  ELP_RESPONSE_TABS.forEach((tabName) => {
    const source = master.getSheetByName(tabName);
    if (!source) throw new Error(`Master response tab missing: ${tabName}`);
    const lastColumn = Math.max(source.getLastColumn(), 1);
    const headers = source.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map((value) => normalizeId_(value));
    assertRequiredHeaders_(headers, [
      ELP_SOURCE_HEADERS.TIMESTAMP, ELP_SOURCE_HEADERS.EMPLOYEE_CODE, ELP_SOURCE_HEADERS.EMPLOYEE_NAME
    ]);
    schemas.set(tabName, headers);
    const headerMap = buildHeaderMap_(headers);

    const currentLastRow = source.getLastRow();
    const fromRow = Math.max(2, (fromCursorByTab[tabName] || 1) + 1);
    const toRow = toTargetByTab ? Math.min(toTargetByTab[tabName] || 1, currentLastRow) : currentLastRow;
    targetRowByTab[tabName] = toTargetByTab ? (toTargetByTab[tabName] || 1) : currentLastRow;

    if (toRow < fromRow) return; // nothing new for this tab

    const values = source.getRange(fromRow, 1, toRow - fromRow + 1, headers.length).getDisplayValues();
    const isHrbpTab = isHrbpResponseTab_(tabName);

    values.forEach((row, index) => {
      if (!row.some((value) => normalizeId_(value))) return;
      totalRead += 1;
      const employeeCode = normalizeId_(row[headerMap[ELP_SOURCE_HEADERS.EMPLOYEE_CODE]]);
      const employeeName = normalizeId_(row[headerMap[ELP_SOURCE_HEADERS.EMPLOYEE_NAME]]);
      const timestamp = row[headerMap[ELP_SOURCE_HEADERS.TIMESTAMP]];

      const resolved = resolveDestination_(isHrbpTab, employeeCode, activeByEcode, exitedByEcode, hrbpUniverse);
      if (resolved.destinationEcode) {
        if (!placementsByEcode.has(resolved.destinationEcode)) {
          placementsByEcode.set(resolved.destinationEcode, new Map(ELP_RESPONSE_TABS.map((t) => [t, []])));
        }
        placementsByEcode.get(resolved.destinationEcode).get(tabName).push(row);
        totalPlaced += 1;
      } else {
        totalFailed += 1;
        failures.push({
          entity: `${tabName} row ${fromRow + index}`,
          details: `${resolved.failReason} (Employee Code: "${employeeCode}", Name: "${employeeName}", Timestamp: ${timestamp})`
        });
      }
    });
  });

  return { placementsByEcode, failures, totalRead, totalPlaced, totalFailed, schemas, targetRowByTab };
}

/** Daily: appends new rows to the bottom of each relevant tab. Never wipes. */
function appendPlacementsForHrbp_(workbook, perTabRows, schemas) {
  ELP_RESPONSE_TABS.forEach((tabName) => {
    const rows = perTabRows.get(tabName);
    if (!rows || !rows.length) return;
    const headers = schemas.get(tabName);
    const sheet = ensureDestinationTabHeaders_(workbook, tabName, headers);
    const startRow = sheet.getLastRow() + 1;
    ensureGridCapacity_(sheet, startRow + rows.length - 1, headers.length);
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    applyStandardFormatting_(sheet);
  });
}

/** Weekly: wipes and rewrites every tab from scratch (including to empty, if an HRBP now has no matches). */
function writeAllTabsForHrbp_(workbook, perTabRows, schemas) {
  ELP_RESPONSE_TABS.forEach((tabName) => {
    const headers = schemas.get(tabName);
    const sheet = ensureDestinationTabHeaders_(workbook, tabName, headers);
    writeRows_(sheet, headers, perTabRows.get(tabName) || []);
    applyStandardFormatting_(sheet);
  });
}

function ensureDestinationTabHeaders_(workbook, tabName, headers) {
  const sheet = workbook.getSheetByName(tabName) || workbook.insertSheet(tabName);
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getDisplayValues()[0];
  const matches = headers.every((header, index) => normalizeId_(currentHeaders[index]) === header);
  if (!matches) {
    ensureGridCapacity_(sheet, 1, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
