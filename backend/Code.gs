/**
 * Habit Tracker <-> Google Sheets bridge.
 *
 * Deploy this bound to the "Habit Tracker Data" spreadsheet as a Web App:
 *   Deploy -> New deployment -> type: Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the resulting /exec URL and paste it into SHEET_API_URL in index.html.
 *
 * Storage model: three tabs, each fully rewritten on every save (fine for a
 * single-user personal tracker). "Habits" holds the habit list, "Logs" holds
 * one row per completed day, "Queue" holds the wheel's current pick queue.
 *
 * Only the columns listed below are persisted: any other field on a habit is
 * silently dropped. To store a new habit field, add a column to HABIT_HEADER
 * and handle it in both readState() and writeHabits().
 * "everyDays" (optional, positive integer) = "do this once every N days".
 */

var HABIT_HEADER = ['id', 'text', 'wheel', 'createdAt', 'everyDays'];

function doGet(e) {
  return jsonOutput(readState());
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  writeState(body);
  return jsonOutput(readState());
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
  // Clean up the default "Sheet1" once we've created our own tabs.
  var stray = ss.getSheetByName('Sheet1');
  if (stray && ss.getSheets().length > 1) {
    var isEmpty = stray.getLastRow() === 0 && stray.getLastColumn() === 0;
    if (isEmpty) ss.deleteSheet(stray);
  }
  return sheet;
}

function readState() {
  var habitsSheet = getOrCreateSheet('Habits', HABIT_HEADER);
  var logsSheet = getOrCreateSheet('Logs', ['habitId', 'date']);
  var queueSheet = getOrCreateSheet('Queue', ['habitId']);

  var habits = [];
  var hVals = habitsSheet.getDataRange().getValues();
  for (var i = 1; i < hVals.length; i++) {
    var row = hVals[i];
    if (!row[0]) continue;
    var habit = {
      id: String(row[0]),
      text: String(row[1]),
      wheel: row[2] !== false && row[2] !== 'FALSE' && row[2] !== 'false',
      createdAt: row[3] ? String(row[3]) : ''
    };
    var every = Number(row[4]);
    if (every >= 1 && Math.floor(every) === every) habit.everyDays = every;
    habits.push(habit);
  }

  var logs = {};
  var lVals = logsSheet.getDataRange().getValues();
  for (var i = 1; i < lVals.length; i++) {
    var row = lVals[i];
    if (!row[0] || !row[1]) continue;
    var hid = String(row[0]);
    var date = row[1] instanceof Date
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[1]);
    if (!logs[hid]) logs[hid] = [];
    logs[hid].push(date);
  }

  var queue = [];
  var qVals = queueSheet.getDataRange().getValues();
  for (var i = 1; i < qVals.length; i++) {
    if (qVals[i][0]) queue.push(String(qVals[i][0]));
  }

  return { habits: habits, logs: logs, queue: queue };
}

function writeState(state) {
  writeHabits(state.habits || []);
  writeLogs(state.logs || {});
  writeQueue(state.queue || []);
}

function writeHabits(habits) {
  var sheet = getOrCreateSheet('Habits', HABIT_HEADER);
  sheet.clearContents();
  var rows = [HABIT_HEADER];
  habits.forEach(function (h) {
    rows.push([h.id, h.text, h.wheel !== false, h.createdAt || '', h.everyDays || '']);
  });
  sheet.getRange(1, 1, rows.length, HABIT_HEADER.length).setValues(rows);
}

function writeLogs(logs) {
  var sheet = getOrCreateSheet('Logs', ['habitId', 'date']);
  sheet.clearContents();
  var rows = [['habitId', 'date']];
  Object.keys(logs).forEach(function (hid) {
    (logs[hid] || []).forEach(function (d) {
      rows.push([hid, d]);
    });
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function writeQueue(queue) {
  var sheet = getOrCreateSheet('Queue', ['habitId']);
  sheet.clearContents();
  var rows = [['habitId']];
  queue.forEach(function (hid) {
    rows.push([hid]);
  });
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}
