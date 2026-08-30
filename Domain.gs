/**
 * Pure domain decisions. Kept tiny and separate on purpose so it's easy to
 * eyeball and (if you ever want to) unit test outside Apps Script.
 */

/**
 * Agent/Supervisor tabs: "Employee Code" is the employee being surveyed —
 * route via their HRSpocEcode in Employee Data / Exited Employees.
 * HRBP tabs: "Employee Code" IS the HRBP's own ECode — route to themselves.
 */
function isHrbpResponseTab_(sheetName) {
  return /^HRBP Day /.test(sheetName);
}
