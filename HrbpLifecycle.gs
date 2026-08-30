/**
 * Keeps HRBP Contacts in sync with who is actually acting as an HRBP right
 * now, purely from Employee Data (the single source of truth):
 *   - An ECode that shows up as some active employee's HRSpocEcode is a
 *     current HRBP -> gets a Contacts row if it doesn't have one.
 *   - An ECode that used to show up but no longer does (no active employee
 *     points to them anymore) is treated as exited -> archived out of
 *     Contacts rather than deleted, so a manually-filled Manager Email
 *     chain isn't lost, and restored automatically if they reappear.
 *
 * Weekly-only, same as sharing sync. Removing someone from live Contacts
 * here is what makes syncHrbpSharing_ revoke their workbook access this
 * same run — their workbook itself is untouched and keeps being rebuilt
 * as long as any exited employee still points to them.
 */

function computeActiveHrbpRoster_(activeByEcode) {
  const roster = new Map(); // HRBP ECode -> best-known name
  activeByEcode.forEach((employee) => {
    if (!employee.hrbpEcode) return;
    if (!roster.has(employee.hrbpEcode) || (!roster.get(employee.hrbpEcode) && employee.hrbpName)) {
      roster.set(employee.hrbpEcode, employee.hrbpName || roster.get(employee.hrbpEcode) || '');
    }
  });
  return roster;
}

function reconcileHrbpRoster_(context, activeRoster) {
  const spreadsheet = getSystemSpreadsheet_();
  const contactsSheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTACTS, ELP_HEADERS.CONTACTS);
  const exitedSheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.EXITED_HRBPS, ELP_HEADERS.EXITED_HRBPS);

  const contactsByEcode = new Map();
  readTable_(contactsSheet, ELP_HEADERS.CONTACTS).rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_HRBP ECode']);
    if (ecode) contactsByEcode.set(ecode, row);
  });

  const exitedByEcode = new Map();
  readTable_(exitedSheet, ELP_HEADERS.EXITED_HRBPS).rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_HRBP ECode']);
    if (ecode) exitedByEcode.set(ecode, row);
  });

  let added = 0;
  let rejoined = 0;
  let exited = 0;

  // New HRBPs: in the active roster, not already in Contacts.
  activeRoster.forEach((name, ecode) => {
    if (contactsByEcode.has(ecode)) return;
    const archived = exitedByEcode.get(ecode);

    if (archived) {
      appendObjectRow_(contactsSheet, ELP_HEADERS.CONTACTS, {
        'HRBP ECode': ecode,
        'Manager Email': archived['Manager Email'],
        "Manager's Manager Email": archived["Manager's Manager Email"],
        'Additional Email': archived['Additional Email'],
        'HRBP Self Email': archived['HRBP Self Email']
      });
      exitedSheet.deleteRow(archived.__rowNumber);
      exitedByEcode.delete(ecode);
      rejoined += 1;
      recordLog_(context, 'INFO', 'HRBP rejoined', ecode, 'Restored',
        'Reappeared as an active HRSpocEcode; contact chain restored from Exited HRBPs.');
    } else {
      appendObjectRow_(contactsSheet, ELP_HEADERS.CONTACTS, { 'HRBP ECode': ecode });
      added += 1;
      recordLog_(context, 'INFO', 'HRBP added', ecode, 'New row in HRBP Contacts',
        `First appearance as an active HRSpocEcode${name ? ` (${name})` : ''}. `
        + `Fill Manager Email / Manager's Manager Email in HRBP Contacts — HRBP Self Email fills in automatically if learnable.`);
    }
  });

  // Exited HRBPs: in Contacts, no active employee points to them anymore.
  const rowsToRemove = [];
  contactsByEcode.forEach((row, ecode) => {
    if (activeRoster.has(ecode)) return;
    appendObjectRow_(exitedSheet, ELP_HEADERS.EXITED_HRBPS, {
      'HRBP ECode': ecode,
      'Manager Email': row['Manager Email'],
      "Manager's Manager Email": row["Manager's Manager Email"],
      'Additional Email': row['Additional Email'],
      'HRBP Self Email': row['HRBP Self Email'],
      'Exited Date': todayIst_()
    });
    rowsToRemove.push(row.__rowNumber);
    exited += 1;
    recordLog_(context, 'INFO', 'HRBP exited', ecode, 'Archived',
      'No active employee lists this ECode as HRSpocEcode anymore. Moved to Exited HRBPs; '
      + 'viewer access on their workbook is revoked this run. Historical responses are unaffected.');
  });
  rowsToRemove.sort((a, b) => b - a).forEach((rowNumber) => contactsSheet.deleteRow(rowNumber));

  if (added || rejoined || exited) {
    recordLog_(context, 'INFO', 'HRBP roster reconciled', 'HRBP Contacts', 'Success',
      `${added} added, ${rejoined} rejoined, ${exited} exited.`);
  }
  return { added: added, rejoined: rejoined, exited: exited };
}
