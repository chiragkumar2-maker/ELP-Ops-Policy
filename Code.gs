/**
 * Public entry points.
 *
 *  - runDailyAppend()   — trigger + manual menu item for the 2 AM job.
 *  - continueDailyAppend() — only ever called by a self-scheduled trigger
 *    when runDailyAppend() ran out of time budget. Same underlying work.
 *  - runWeeklyRebuild() / continueWeeklyRebuild() — same pattern, weekly.
 */

function runDailyAppend() { return runDailyAppendInternal_(); }
function continueDailyAppend() { return runDailyAppendInternal_(); }

function runWeeklyRebuild() { return runWeeklyRebuildInternal_(); }
function continueWeeklyRebuild() { return runWeeklyRebuildInternal_(); }

function runDailyAppendInternal_() {
  const jobKey = 'daily';
  if (!tryAcquireJobSlot_(jobKey)) {
    console.log('Daily append skipped: another ELP job is already in progress.');
    return { skipped: true };
  }
  const context = createRunContext_('Daily append');
  const startedAtMs = Date.now();
  let checkpoint = loadCheckpoint_(jobKey);
  const isResume = Boolean(checkpoint);

  try {
    const config = loadConfig_();
    requireRuntimeConfig_(config);
    const activeByEcode = readEmployeeData_();
    const lifecycle = reconcileEmployeeLifecycle_(context, activeByEcode);
    const contactsByEcode = readHrbpContacts_();
    const hrbpUniverse = computeHrbpUniverse_(activeByEcode, lifecycle.exitedMap, contactsByEcode);
    const committedCursors = getResponseCursors_();

    const scan = readNewResponsesSinceCursor_(
      config, committedCursors, isResume ? checkpoint.targetRowByTab : null,
      activeByEcode, lifecycle.exitedMap, hrbpUniverse
    );

    if (!isResume) {
      checkpoint = {
        targetRowByTab: scan.targetRowByTab,
        remainingEcodes: Array.from(scan.placementsByEcode.keys()),
        metrics: { totalRead: scan.totalRead, totalPlaced: scan.totalPlaced, totalFailed: scan.totalFailed, workbooksTouched: 0 }
      };
      scan.failures.forEach((failure) => recordLog_(context, 'WARNING', 'Response not placed', failure.entity, 'Failed', failure.details));
      recordLog_(context, 'INFO', 'Daily append scan', 'Master responses', 'Success',
        `${scan.totalRead} new row(s), ${scan.totalPlaced} placed, ${scan.totalFailed} failed, `
        + `${checkpoint.remainingEcodes.length} HRBP workbook(s) to touch.`);
      saveCheckpoint_(jobKey, checkpoint);
    } else {
      recordLog_(context, 'INFO', 'Daily append resumed', 'ELP', 'Continuing',
        `${checkpoint.remainingEcodes.length} HRBP workbook(s) left from an earlier run that hit the time budget.`);
    }

    const registryState = loadHrbpRegistryState_();
    while (checkpoint.remainingEcodes.length) {
      const ecode = checkpoint.remainingEcodes[0];
      const perTab = scan.placementsByEcode.get(ecode);
      if (perTab) {
        const workbook = lookupOrCreateHrbpWorkbookForDailyAppend_(context, config, ecode, hrbpUniverse.get(ecode) || ecode, registryState);
        appendPlacementsForHrbp_(workbook, perTab, scan.schemas);
        checkpoint.metrics.workbooksTouched += 1;
      }
      checkpoint.remainingEcodes.shift();

      if (timeBudgetExceeded_(startedAtMs) && checkpoint.remainingEcodes.length) {
        saveCheckpoint_(jobKey, checkpoint);
        scheduleContinuation_('continueDailyAppend');
        recordLog_(context, 'INFO', 'Daily append paused', 'ELP', 'Time budget',
          `${checkpoint.remainingEcodes.length} HRBP workbook(s) left; resuming in ~${ELP_CONTINUATION_DELAY_SECONDS}s.`);
        flushLogs_(context);
        return { paused: true, remaining: checkpoint.remainingEcodes.length };
      }
    }

    // Fully done: commit the frozen cursor boundary so nothing is re-appended tomorrow.
    commitResponseCursors_(checkpoint.targetRowByTab);
    commitEmployeeLastRunSnapshot_(activeByEcode);
    clearCheckpoint_(jobKey);
    clearContinuationTriggers_('continueDailyAppend');

    recordLog_(context, 'INFO', 'Daily append complete', 'ELP', 'Success',
      `${checkpoint.metrics.totalRead} read, ${checkpoint.metrics.totalPlaced} placed, ${checkpoint.metrics.totalFailed} failed, `
      + `${checkpoint.metrics.workbooksTouched} workbook(s) touched. `
      + `${lifecycle.newlyExited} newly exited, ${lifecycle.rejoined} rejoined, ${lifecycle.reassigned} reassigned.`);
    updateControlPanel_(context, {
      'Last Daily Append': now_(), 'Daily — Placed': checkpoint.metrics.totalPlaced, 'Daily — Failed': checkpoint.metrics.totalFailed
    });
    flushLogs_(context);
    sendAlertIfNeeded_(context, config);
    return checkpoint.metrics;
  } catch (error) {
    recordError_(context, 'Daily append', 'ELP', error);
    try { flushLogs_(context); sendAlertIfNeeded_(context, loadConfig_()); } catch (secondary) { console.error(secondary); }
    throw error;
  } finally {
    if (!loadCheckpoint_(jobKey)) releaseJobSlot_(jobKey);
  }
}

function runWeeklyRebuildInternal_() {
  const jobKey = 'weekly';
  if (!tryAcquireJobSlot_(jobKey)) {
    console.log('Weekly rebuild skipped: another ELP job is already in progress.');
    return { skipped: true };
  }
  const context = createRunContext_('Weekly rebuild');
  const startedAtMs = Date.now();
  let checkpoint = loadCheckpoint_(jobKey);
  const isResume = Boolean(checkpoint);

  try {
    const config = loadConfig_();
    requireRuntimeConfig_(config);
    const activeByEcode = readEmployeeData_();
    const lifecycle = reconcileEmployeeLifecycle_(context, activeByEcode);
    const activeHrbpRoster = computeActiveHrbpRoster_(activeByEcode);
    const rosterChanges = reconcileHrbpRoster_(context, activeHrbpRoster);
    const contactsByEcode = readHrbpContacts_();
    const hrbpUniverse = computeHrbpUniverse_(activeByEcode, lifecycle.exitedMap, contactsByEcode);
    const hrbpSelfEmails = computeHrbpSelfEmails_(activeByEcode, lifecycle.exitedMap);
    syncHrbpSelfEmailsIntoContacts_(context, hrbpSelfEmails, hrbpUniverse);
    const contactsByEcodeSynced = readHrbpContacts_(); // re-read to pick up what we just wrote
    const approvedDomains = parseApprovedDomains_(config.APPROVED_VIEWER_DOMAINS);

    // Always a full scan — the Master snapshot is the truth, every run.
    const scan = readNewResponsesSinceCursor_(config, {}, null, activeByEcode, lifecycle.exitedMap, hrbpUniverse);

    if (!isResume) {
      checkpoint = {
        remainingEcodes: Array.from(hrbpUniverse.keys()),
        metrics: {
          totalRead: scan.totalRead, totalPlaced: scan.totalPlaced, totalFailed: scan.totalFailed,
          workbooksWritten: 0, sharingAdded: 0, sharingRevoked: 0
        }
      };
      scan.failures.forEach((failure) => recordLog_(context, 'WARNING', 'Response not placed', failure.entity, 'Failed', failure.details));
      recordLog_(context, 'INFO', 'Weekly rebuild scan', 'Master responses', 'Success',
        `${scan.totalRead} read, ${scan.totalPlaced} placed, ${scan.totalFailed} failed. `
        + `${checkpoint.remainingEcodes.length} HRBP workbook(s) to rebuild.`);
      saveCheckpoint_(jobKey, checkpoint);
    } else {
      recordLog_(context, 'INFO', 'Weekly rebuild resumed', 'ELP', 'Continuing', `${checkpoint.remainingEcodes.length} workbook(s) left.`);
    }

    const registryState = loadHrbpRegistryState_();
    const emptyPerTab = () => new Map(ELP_RESPONSE_TABS.map((t) => [t, []]));

    while (checkpoint.remainingEcodes.length) {
      const ecode = checkpoint.remainingEcodes[0];
      const name = hrbpUniverse.get(ecode) || ecode;
      const workbook = ensureHrbpWorkbookForEcode_(context, config, ecode, name, registryState);
      const sharingResult = syncHrbpSharing_(context, workbook, ecode, contactsByEcodeSynced, approvedDomains);
      checkpoint.metrics.sharingAdded += sharingResult.added;
      checkpoint.metrics.sharingRevoked += sharingResult.revoked;
      writeAllTabsForHrbp_(workbook, scan.placementsByEcode.get(ecode) || emptyPerTab(), scan.schemas);
      checkpoint.metrics.workbooksWritten += 1;
      checkpoint.remainingEcodes.shift();

      if (timeBudgetExceeded_(startedAtMs) && checkpoint.remainingEcodes.length) {
        saveCheckpoint_(jobKey, checkpoint);
        scheduleContinuation_('continueWeeklyRebuild');
        recordLog_(context, 'INFO', 'Weekly rebuild paused', 'ELP', 'Time budget',
          `${checkpoint.remainingEcodes.length} HRBP workbook(s) left; resuming in ~${ELP_CONTINUATION_DELAY_SECONDS}s.`);
        flushLogs_(context);
        return { paused: true, remaining: checkpoint.remainingEcodes.length };
      }
    }

    // Fast-forward the daily cursor so tomorrow's append doesn't re-append what we just wrote fresh.
    commitResponseCursors_(scan.targetRowByTab);
    commitEmployeeLastRunSnapshot_(activeByEcode);
    clearCheckpoint_(jobKey);
    clearContinuationTriggers_('continueWeeklyRebuild');

    recordLog_(context, 'INFO', 'Weekly rebuild complete', 'ELP', 'Success',
      `${checkpoint.metrics.totalRead} read, ${checkpoint.metrics.totalPlaced} placed, ${checkpoint.metrics.totalFailed} failed. `
      + `${checkpoint.metrics.workbooksWritten} workbook(s) rewritten. Sharing +${checkpoint.metrics.sharingAdded}/-${checkpoint.metrics.sharingRevoked}. `
      + `${lifecycle.newlyExited} newly exited, ${lifecycle.rejoined} rejoined, ${lifecycle.reassigned} reassigned. `
      + `HRBPs: ${rosterChanges.added} added, ${rosterChanges.rejoined} rejoined, ${rosterChanges.exited} exited.`);
    updateControlPanel_(context, {
      'Last Weekly Rebuild': now_(), 'Weekly — Placed': checkpoint.metrics.totalPlaced,
      'Weekly — Failed': checkpoint.metrics.totalFailed, 'HRBP Workbooks': hrbpUniverse.size,
      'HRBPs Added (last run)': rosterChanges.added, 'HRBPs Exited (last run)': rosterChanges.exited
    });
    flushLogs_(context);
    sendAlertIfNeeded_(context, config);
    return checkpoint.metrics;
  } catch (error) {
    recordError_(context, 'Weekly rebuild', 'ELP', error);
    try { flushLogs_(context); sendAlertIfNeeded_(context, loadConfig_()); } catch (secondary) { console.error(secondary); }
    throw error;
  } finally {
    if (!loadCheckpoint_(jobKey)) releaseJobSlot_(jobKey);
  }
}
