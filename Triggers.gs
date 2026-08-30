/**
 * Installable recurring triggers. Separate handler names from the
 * continuation triggers in Checkpoint.gs, so removing/reinstalling one
 * never touches the other.
 */

const ELP_DAILY_HANDLER = 'runDailyAppend';
const ELP_WEEKLY_HANDLER = 'runWeeklyRebuild';
const ELP_DAILY_HOUR_IST = 2;
const ELP_WEEKLY_HOUR_IST = 3; // Sundays, so it doesn't collide with the daily 2 AM run.

function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger(ELP_DAILY_HANDLER)
    .timeBased()
    .atHour(ELP_DAILY_HOUR_IST)
    .everyDays(1)
    .inTimezone(ELP_TIMEZONE)
    .create();
  return { installed: true, hour: ELP_DAILY_HOUR_IST };
}

function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers().filter((trigger) => trigger.getHandlerFunction() === ELP_DAILY_HANDLER);
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  return { removed: triggers.length };
}

function installWeeklyTrigger() {
  removeWeeklyTrigger();
  ScriptApp.newTrigger(ELP_WEEKLY_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(ELP_WEEKLY_HOUR_IST)
    .inTimezone(ELP_TIMEZONE)
    .create();
  return { installed: true, day: 'Sunday', hour: ELP_WEEKLY_HOUR_IST };
}

function removeWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers().filter((trigger) => trigger.getHandlerFunction() === ELP_WEEKLY_HANDLER);
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  return { removed: triggers.length };
}
