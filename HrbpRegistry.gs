/**
 * HRBP Registry: one row per HRBP, with the Workbook URL as a plain
 * clickable link — that link IS "open anyone's book."
 *
 * Functions here work one ECode at a time (not the whole registry in one
 * pass), so the daily/weekly checkpoint loops in Code.gs can call them
 * per item and safely stop between items when the time budget runs out.
 */

/** Every HRBP ECode that needs a workbook this run, with the best name we know. */
function computeHrbpUniverse_(activeByEcode, exitedByEcode, contactsByEcode) {
  const universe = new Map();
  const consider = (ecode, name) => {
    if (!ecode) return;
    if (!universe.has(ecode) || (name && !universe.get(ecode))) {
      universe.set(ecode, name || universe.get(ecode) || '');
    }
  };
  activeByEcode.forEach((employee) => consider(employee.hrbpEcode, employee.hrbpName));
  exitedByEcode.forEach((employee) => consider(employee.hrbpEcode, employee.hrbpName));
  contactsByEcode.forEach((emails, ecode) => consider(ecode, ''));
  return universe;
}

/**
 * The HRBP's own email — sourced from the "HRBP Email" column in Employee
 * Data / Exited Employees (whoever they manage carries it), not from HRBP
 * Contacts (that sheet is the manager chain, not the HRBP themselves).
 */
function computeHrbpSelfEmails_(activeByEcode, exitedByEcode) {
  const emails = new Map();
  const consider = (ecode, email) => {
    if (!ecode || !email || emails.has(ecode)) return;
    emails.set(ecode, email);
  };
  activeByEcode.forEach((employee) => consider(employee.hrbpEcode, employee.hrbpEmail));
  exitedByEcode.forEach((employee) => consider(employee.hrbpEcode, employee.hrbpEmail));
  return emails;
}

/**
 * Writes each HRBP's own email into the "HRBP Self Email" column of HRBP
 * Contacts, so sharing reads from one place instead of cross-referencing
 * Employee Data at share-time. Only overwrites when a value was actually
 * learned this run — if Employee Data has nothing for an HRBP, whatever's
 * already in Contacts (including a manually-typed fallback) is left alone.
 */
function syncHrbpSelfEmailsIntoContacts_(context, hrbpSelfEmails, hrbpUniverse) {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.CONTACTS, ELP_HEADERS.CONTACTS);
  const table = readTable_(sheet, ELP_HEADERS.CONTACTS);
  const byEcode = new Map();
  table.rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_HRBP ECode']);
    if (ecode) byEcode.set(ecode, row);
  });
  const selfEmailColumn = ELP_HEADERS.CONTACTS.indexOf('HRBP Self Email') + 1;

  let written = 0;
  hrbpUniverse.forEach((name, ecode) => {
    const learnedEmail = hrbpSelfEmails.get(ecode) || '';
    const existingRow = byEcode.get(ecode);
    const existingSelfEmail = existingRow ? normalizeEmail_(existingRow['HRBP Self Email']) : '';

    if (learnedEmail && learnedEmail !== existingSelfEmail) {
      if (existingRow) {
        sheet.getRange(existingRow.__rowNumber, selfEmailColumn).setValue(learnedEmail);
      } else {
        appendObjectRow_(sheet, ELP_HEADERS.CONTACTS, { 'HRBP ECode': ecode, 'HRBP Self Email': learnedEmail });
      }
      written += 1;
    } else if (!learnedEmail && !existingSelfEmail) {
      recordLog_(context, 'WARNING', 'HRBP self-email not found', ecode, 'Missing',
        'No Employee Data row lists this HRBP as HRSpocEcode with a filled HRBP Email, and HRBP Contacts '
        + 'has no fallback either. Add their email directly to the "HRBP Self Email" column in HRBP Contacts.');
    }
  });

  if (written) {
    recordLog_(context, 'INFO', 'HRBP self-email synced into Contacts', 'HRBP Contacts', 'Success', `${written} row(s) updated.`);
  }
  return written;
}

/** Reads the HRBP Registry sheet once per execution — cheap, re-read fresh on every continuation. */
function loadHrbpRegistryState_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, ELP_SHEETS.HRBP_REGISTRY, ELP_HEADERS.HRBP_REGISTRY);
  const table = readTable_(sheet, ELP_HEADERS.HRBP_REGISTRY);
  const byEcode = new Map();
  table.rows.forEach((row) => {
    const ecode = normalizeId_(row['__display_HRBP ECode']);
    if (ecode) byEcode.set(ecode, row);
  });
  return { sheet: sheet, byEcode: byEcode };
}

/** Weekly: full ensure — reopens/recreates, reparents, renames, and upserts the registry row. */
function ensureHrbpWorkbookForEcode_(context, config, ecode, name, registryState) {
  const folder = DriveApp.getFolderById(config.HRBP_WORKBOOKS_FOLDER_ID);
  const existingRow = registryState.byEcode.get(ecode);
  const displayName = name || (existingRow && existingRow['HRBP Name']) || ecode;
  let workbook = null;

  if (existingRow && normalizeId_(existingRow['__display_Workbook ID'])) {
    try {
      workbook = SpreadsheetApp.openById(normalizeId_(existingRow['__display_Workbook ID']));
      moveFileToFolder_(DriveApp.getFileById(workbook.getId()), folder);
      ensureHrbpWorkbookTabs_(workbook);
      const expectedName = hrbpWorkbookName_(ecode, displayName);
      if (workbook.getName() !== expectedName) workbook.rename(expectedName);
    } catch (error) {
      recordLog_(context, 'WARNING', 'Workbook missing', ecode, 'Recreating',
        `Could not open registered workbook (${error.message}); creating a new one.`);
      workbook = null;
    }
  }
  if (!workbook) {
    workbook = createHrbpWorkbook_(folder, ecode, displayName);
    recordLog_(context, 'INFO', 'Workbook created', ecode, 'Success', workbook.getUrl());
  }

  upsertRegistryRow_(registryState, ecode, displayName, workbook, existingRow, now_());
  return workbook;
}

/** Daily: lightweight — just open the known workbook, or create one if this HRBP is brand new. No rename/reparent checks. */
function lookupOrCreateHrbpWorkbookForDailyAppend_(context, config, ecode, name, registryState) {
  const existingRow = registryState.byEcode.get(ecode);
  if (existingRow && normalizeId_(existingRow['__display_Workbook ID'])) {
    try {
      return SpreadsheetApp.openById(normalizeId_(existingRow['__display_Workbook ID']));
    } catch (error) {
      recordLog_(context, 'WARNING', 'Workbook missing', ecode, 'Recreating',
        `Could not open registered workbook (${error.message}); creating a new one. Sharing will be set up on the next weekly rebuild.`);
    }
  }
  const folder = DriveApp.getFolderById(config.HRBP_WORKBOOKS_FOLDER_ID);
  const workbook = createHrbpWorkbook_(folder, ecode, name);
  recordLog_(context, 'INFO', 'Workbook created (daily)', ecode, 'Success',
    `${workbook.getUrl()} — sharing will be set up on the next weekly rebuild.`);
  upsertRegistryRow_(registryState, ecode, name, workbook, existingRow, '');
  return workbook;
}

function upsertRegistryRow_(registryState, ecode, displayName, workbook, existingRow, lastSynced) {
  const registryFields = {
    'HRBP ECode': ecode,
    'HRBP Name': displayName,
    'Workbook ID': workbook.getId(),
    'Workbook URL': workbook.getUrl(),
    'Created Date': (existingRow && existingRow['Created Date']) || todayIst_(),
    'Last Synced': lastSynced
  };
  if (existingRow) {
    writeObjectRowAt_(registryState.sheet, ELP_HEADERS.HRBP_REGISTRY, registryFields, existingRow.__rowNumber);
  } else {
    appendObjectRow_(registryState.sheet, ELP_HEADERS.HRBP_REGISTRY, registryFields);
    // Mirror plain keys into __display_-prefixed keys too, so this cached row
    // reads back consistently with readTable_()'s output if it's ever re-checked.
    const cached = Object.assign({ __rowNumber: registryState.sheet.getLastRow() }, registryFields);
    Object.keys(registryFields).forEach((header) => { cached[`__display_${header}`] = registryFields[header]; });
    registryState.byEcode.set(ecode, cached);
  }
}
