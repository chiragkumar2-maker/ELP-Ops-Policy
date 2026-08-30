/**
 * Viewer-only sharing, synced exactly against HRBP Contacts each run
 * (which now includes the HRBP's own email — see syncHrbpSelfEmailsIntoContacts_):
 * missing viewers are added, anyone no longer in that list is auto-revoked.
 * Uses plain DriveApp — no Advanced Drive service needed.
 */
function syncHrbpSharing_(context, workbook, ecode, contactsByEcode, approvedDomains) {
  const file = DriveApp.getFileById(workbook.getId());
  const candidateEmails = contactsByEcode.get(ecode) || [];

  const desired = [];
  candidateEmails.forEach((email) => {
    if (!isValidEmail_(email)) return;
    if (!isApprovedDomain_(email, approvedDomains)) {
      recordLog_(context, 'WARNING', 'Sharing skipped', ecode, 'Domain not approved',
        `${email} is not in APPROVED_VIEWER_DOMAINS.`);
      return;
    }
    desired.push(email);
  });
  const desiredSet = new Set(desired);

  const currentSet = new Set(file.getViewers().map((user) => normalizeEmail_(user.getEmail())));

  let added = 0;
  let revoked = 0;
  desiredSet.forEach((email) => {
    if (currentSet.has(email)) return;
    try {
      file.addViewer(email);
      added += 1;
    } catch (error) {
      recordLog_(context, 'WARNING', 'Sharing failed', ecode, 'Add failed', `${email}: ${error.message}`);
    }
  });
  currentSet.forEach((email) => {
    if (desiredSet.has(email)) return;
    try {
      file.removeViewer(email);
      revoked += 1;
    } catch (error) {
      recordLog_(context, 'WARNING', 'Sharing failed', ecode, 'Revoke failed', `${email}: ${error.message}`);
    }
  });

  if (added || revoked) {
    recordLog_(context, 'INFO', 'Sharing synced', ecode, 'Success',
      `+${added}/-${revoked}. Current viewers: ${desired.join(', ') || '(none)'}`);
  }
  return { added: added, revoked: revoked };
}
