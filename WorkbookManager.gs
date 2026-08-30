/** Creates and maintains the per-HRBP Google Sheets workbook itself. */

function hrbpWorkbookName_(ecode, name) {
  return `ELP Ops | ${ecode} | ${name}`.trim().replace(/\s*\|\s*$/, '');
}

function createHrbpWorkbook_(folder, ecode, name) {
  const workbook = SpreadsheetApp.create(hrbpWorkbookName_(ecode, name));
  moveFileToFolder_(DriveApp.getFileById(workbook.getId()), folder);
  ensureHrbpWorkbookTabs_(workbook);
  return workbook;
}

function ensureHrbpWorkbookTabs_(workbook) {
  ELP_RESPONSE_TABS.forEach((tabName) => {
    let sheet = workbook.getSheetByName(tabName);
    if (!sheet) {
      sheet = workbook.insertSheet(tabName);
      applyStandardFormatting_(sheet);
    }
  });
  const defaultSheet = workbook.getSheetByName('Sheet1');
  if (defaultSheet && workbook.getSheets().length > ELP_RESPONSE_TABS.length) {
    workbook.deleteSheet(defaultSheet);
  }
}
