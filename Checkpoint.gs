/**
 * Lets a job that touches many HRBP workbooks safely span more than one
 * Apps Script execution. A job saves its remaining work (just a list of
 * HRBP ECodes, not response data) and running totals to Script Properties,
 * schedules a one-off trigger ~1 minute later to pick up where it left
 * off, and ends the current execution cleanly — instead of getting killed
 * mid-write by the 6-minute execution limit.
 *
 * A second, separate guard (the "job slot") stops the daily append and
 * weekly rebuild from running at the same time and stepping on each
 * other's writes to the same workbooks.
 */

const ELP_RUN_TIME_BUDGET_SECONDS = 300; // ~5 min soft budget; leaves buffer under the 6-min hard limit.
const ELP_CONTINUATION_DELAY_SECONDS = 60;
const ELP_JOB_STALE_MINUTES = 40; // safety valve if a job ever errors out without releasing its slot.

function timeBudgetExceeded_(startedAtMs) {
  return (Date.now() - startedAtMs) > ELP_RUN_TIME_BUDGET_SECONDS * 1000;
}

// --- Checkpoint state (per job: 'daily' or 'weekly') ---

function checkpointPropertyKey_(jobKey) {
  return `ELP_CHECKPOINT_${jobKey}`;
}

function loadCheckpoint_(jobKey) {
  const raw = PropertiesService.getScriptProperties().getProperty(checkpointPropertyKey_(jobKey));
  return raw ? JSON.parse(raw) : null;
}

function saveCheckpoint_(jobKey, state) {
  PropertiesService.getScriptProperties().setProperty(checkpointPropertyKey_(jobKey), JSON.stringify(state));
}

function clearCheckpoint_(jobKey) {
  PropertiesService.getScriptProperties().deleteProperty(checkpointPropertyKey_(jobKey));
}

// --- Continuation triggers (one-off, distinct handler names from the
// recurring installed triggers so clearing them never touches your 2 AM /
// weekly schedule) ---

function scheduleContinuation_(handlerFunctionName) {
  clearContinuationTriggers_(handlerFunctionName);
  ScriptApp.newTrigger(handlerFunctionName)
    .timeBased()
    .after(ELP_CONTINUATION_DELAY_SECONDS * 1000)
    .create();
}

function clearContinuationTriggers_(handlerFunctionName) {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === handlerFunctionName)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

// --- Job slot: only one of {daily, weekly} runs at a time ---

function tryAcquireJobSlot_(jobKey) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const current = props.getProperty('ELP_JOB_IN_PROGRESS');
    const since = Number(props.getProperty('ELP_JOB_IN_PROGRESS_SINCE') || 0);
    const stale = current && since && (Date.now() - since) > ELP_JOB_STALE_MINUTES * 60000;
    if (current && current !== jobKey && !stale) return false;
    props.setProperty('ELP_JOB_IN_PROGRESS', jobKey);
    props.setProperty('ELP_JOB_IN_PROGRESS_SINCE', String(Date.now()));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function releaseJobSlot_(jobKey) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('ELP_JOB_IN_PROGRESS') === jobKey) {
      props.deleteProperty('ELP_JOB_IN_PROGRESS');
      props.deleteProperty('ELP_JOB_IN_PROGRESS_SINCE');
    }
  } finally {
    lock.releaseLock();
  }
}

/** Emergency manual escape hatch — menu item. Only needed if a job ever gets stuck. */
function forceReleaseJobSlot() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('ELP_JOB_IN_PROGRESS');
  props.deleteProperty('ELP_JOB_IN_PROGRESS_SINCE');
  clearCheckpoint_('daily');
  clearCheckpoint_('weekly');
  clearContinuationTriggers_('continueDailyAppend');
  clearContinuationTriggers_('continueWeeklyRebuild');
  return { cleared: true };
}
