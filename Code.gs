/**
 * LIFE & WORK PLANNER — Apps Script backend (API-only)
 * Data store: this spreadsheet (two sheets: "Tasks" and "Projects")
 *
 * This project is the BACKEND ONLY — same pattern as Toka. The frontend
 * (Index.html) is a plain static page hosted on GitHub Pages; it talks to
 * this script over HTTP instead of via google.script.run.
 *
 * ONE-TIME SETUP:
 *   1. Open Extensions > Apps Script from a Google Sheet. Paste this file
 *      in as Code.gs. (No HTML file needed here anymore.)
 *   2. Run the `setup` function once (Run menu > select "setup" > Run).
 *      This creates the sheets/headers, seeds default projects, and
 *      installs the daily trigger that generates recurring tasks.
 *   3. Deploy > New deployment > type "Web app" > Execute as "Me",
 *      Who has access "Anyone" (required — GitHub Pages calling in is a
 *      different origin, and "Only myself" will block it).
 *   4. Copy the deployed Web app URL (ends in /exec). Paste it into the
 *      API_URL constant near the top of Index.html's <script>, then push
 *      Index.html to your GitHub Pages repo.
 *   5. Open your GitHub Pages URL — that's your planner. Use "Import from
 *      Apple" to pull events in from a public Apple Calendar link (one-way;
 *      there's no export back out to Apple Calendar).
 *
 *   Already set up and just changing project colors? Run `updateProjectColors`
 *   to re-apply DEFAULT_PROJECTS colors to your existing Projects sheet rows.
 */

const TASKS_SHEET = 'Tasks';
const PROJECTS_SHEET = 'Projects';

// Column order matters — it's how createTask/createRecurringTemplate/
// generateRecurringTasks build each row. Keep in sync if you add columns.
const TASK_HEADERS = [
  'ID', 'Name', 'Project', 'ShowDate', 'ShowTime', 'DurationMinutes', 'Deadline',
  'Status', 'IsTemplate', 'RecurrenceType', 'RecurrenceInterval', 'NextRunDate',
  'ParentTemplateId', 'SyncToCalendar', 'CreatedAt', 'ExternalUID', 'Notes'
];

const PROJECT_HEADERS = ['Name', 'Color'];

// Script Property key under which the user's public Apple Calendar share
// link (from Calendar app > Share Calendar > Public Calendar) is stored.
const APPLE_ICS_PROP_KEY = 'APPLE_ICS_URL';

// Palette: Collection 029 (cranberry / midnight / steel blue / camel coat /
// ceramic mug / knit sweater), plus a neutral fallback for "Other".
const DEFAULT_PROJECTS = [
  ['Toka', '#37514D'],   // midnight
  ['House', '#BE845E'],  // camel coat
  ['Yoga', '#90AEB2'],   // steel blue
  ['Arabic', '#B6594C'], // cranberry
  ['Life', '#DD8E75'],   // ceramic mug
  ['Matcha', '#EEE6DE'], // knit sweater (light — client renders dark text on it)
  ['Other', '#8A8578']   // neutral fallback, not part of the palette photo
];

// ---------- Web app entry point ----------
// GET is only used for a health-check message. Every real action goes
// through POST as a small JSON-RPC dispatcher: {action: 'createTask',
// args: [...]}. POST with a text/plain content-type (see Index.html's
// apiCall helper) avoids triggering a CORS preflight, so this works from
// GitHub Pages with no extra CORS handling needed.

function doGet(e) {
  return jsonResponse_({ ok: true, message: 'Life & Work Planner API is running.' });
}

// Maps action names the client can call to the actual functions.
const API_ACTIONS = {
  getAppData: getAppData,
  createTask: createTask,
  updateTask: updateTask,
  deleteTask: deleteTask,
  toggleDone: toggleDone,
  addProject: addProject,
  createRecurringTemplate: createRecurringTemplate,
  updateRecurringTemplate: updateRecurringTemplate,
  deleteTemplate: deleteTemplate,
  importFromAppleCalendar: importFromAppleCalendar,
  getAppleCalendarUrl: getAppleCalendarUrl
};

function doPost(e) {
  let response;
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = API_ACTIONS[body.action];
    if (!fn) throw new Error('Unknown action: ' + body.action);
    const data = fn.apply(null, body.args || []);
    response = { success: true, data: data };
  } catch (err) {
    response = { success: false, error: err.message };
  }
  return jsonResponse_(response);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Setup ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let tasksSheet = ss.getSheetByName(TASKS_SHEET);
  if (!tasksSheet) tasksSheet = ss.insertSheet(TASKS_SHEET);
  if (tasksSheet.getLastRow() === 0) {
    tasksSheet.appendRow(TASK_HEADERS);
    tasksSheet.setFrozenRows(1);
  }

  let projectsSheet = ss.getSheetByName(PROJECTS_SHEET);
  if (!projectsSheet) projectsSheet = ss.insertSheet(PROJECTS_SHEET);
  if (projectsSheet.getLastRow() === 0) {
    projectsSheet.appendRow(PROJECT_HEADERS);
    projectsSheet.setFrozenRows(1);
    DEFAULT_PROJECTS.forEach(row => projectsSheet.appendRow(row));
  }

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'generateRecurringTasks') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('generateRecurringTasks')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  generateRecurringTasks();

  Logger.log('Setup complete. Tasks and Projects sheets are ready, daily trigger installed.');
}

/** Run once any time you want existing Projects rows to pick up the palette above. */
function updateProjectColors() {
  const sheet = getSheet_(PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const colorMap = {};
  DEFAULT_PROJECTS.forEach(([name, color]) => { colorMap[name] = color; });
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    if (colorMap[name]) sheet.getRange(i + 1, 2).setValue(colorMap[name]);
  }
  Logger.log('Project colors updated.');
}

/**
 * ONE-TIME FIX: your Tasks sheet's row 1 (headers) is stale — it still says
 * the old 13-column layout (…Project, ShowDate, Deadline, Status…) from
 * before ShowTime/DurationMinutes/ExternalUID were added, even though rows
 * created since then are physically written in the new 16-column layout.
 * That mismatch is why ShowTime always read as blank.
 *
 * This function:
 *  1. Finds any row still physically laid out the old way (recognized by
 *     having nothing in the new CreatedAt position, column O) and shifts
 *     its Deadline..CreatedAt values two columns right, leaving ShowTime/
 *     DurationMinutes blank for those (they never had a time anyway).
 *  2. Rewrites row 1 to the current TASK_HEADERS so every column label
 *     matches what's actually stored beneath it.
 *
 * Safe to run more than once — already-migrated rows are left untouched.
 */
function fixTaskSheetLayout() {
  const sheet = getSheet_(TASKS_SHEET);
  const range = sheet.getDataRange();
  const values = range.getValues();
  let migrated = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // blank row, skip

    // New-scheme rows have a CreatedAt timestamp in physical column O (idx 14).
    // Old-scheme rows only ever wrote 13 columns, so O is empty for them.
    if (row[14]) continue; // already new-scheme, nothing to do

    const old = row.slice(0, 13); // ID..CreatedAt under the old 13-column layout
    const fixed = [
      old[0], old[1], old[2], old[3], // ID, Name, Project, ShowDate
      '', '',                          // ShowTime, DurationMinutes (legacy rows never had these)
      old[4], old[5], old[6], old[7], old[8], old[9], old[10], old[11], old[12], // Deadline..CreatedAt, shifted right by 2
      ''                                // ExternalUID
    ];
    sheet.getRange(i + 1, 1, 1, fixed.length).setValues([fixed]);
    migrated++;
  }

  sheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]);

  Logger.log('Header row fixed. Migrated ' + migrated + ' legacy row(s) to the new column layout.');
  return migrated;
}

/**
 * ONE-TIME CLEANUP: if you clicked "Import from Apple" more than once before
 * fixTaskSheetLayout() was run, the duplicate-check silently failed (same
 * root cause as above — it also depends on reading the ExternalUID column
 * correctly), so the same events may have been inserted multiple times.
 * Run this AFTER fixTaskSheetLayout() to remove the extras, keeping the
 * first copy of each imported event.
 */
function removeDuplicateImports() {
  const sheet = getSheet_(TASKS_SHEET);
  const values = sheet.getDataRange().getValues();
  const seen = new Set();
  const rowsToDelete = [];

  for (let i = 1; i < values.length; i++) {
    const uid = values[i][15]; // ExternalUID, physical column P
    if (!uid) continue;
    if (seen.has(uid)) {
      rowsToDelete.push(i + 1); // 1-indexed sheet row
    } else {
      seen.add(uid);
    }
  }

  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  Logger.log('Removed ' + rowsToDelete.length + ' duplicate imported row(s).');
  return rowsToDelete.length;
}

/**
 * ONE-TIME FIX: Apple Calendar imports used to default to the "Life"
 * project. Now they import with no project at all, so they show up on the
 * calendar without cluttering the Projects board. This clears the project
 * on any row that was actually imported from Apple (identified by having
 * an ExternalUID) rather than a real "Life" task you created yourself.
 */
function clearProjectFromImportedEvents() {
  const sheet = getSheet_(TASKS_SHEET);
  const values = sheet.getDataRange().getValues();
  let cleared = 0;

  for (let i = 1; i < values.length; i++) {
    const uid = values[i][15]; // ExternalUID, physical column P
    if (!uid) continue; // not an Apple import, leave it alone
    if (!values[i][2]) continue; // already has no project
    sheet.getRange(i + 1, 3).setValue(''); // column C = Project
    cleared++;
  }

  Logger.log('Cleared project from ' + cleared + ' imported event(s).');
  return cleared;
}

// ---------- Sheet helpers ----------

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setup() first.');
  return sheet;
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .map((row, i) => {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = row[idx]; });
      obj._row = i + 2;
      return obj;
    })
    .filter(obj => obj.ID !== '' && obj.ID !== undefined && obj.ID !== null);
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function newId_() {
  return Utilities.getUuid();
}

function formatDate_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTime_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}

// ---------- Public API (called from client via google.script.run) ----------

function getAppData() {
  const tasksSheet = getSheet_(TASKS_SHEET);
  const projectsSheet = getSheet_(PROJECTS_SHEET);

  const allTasks = sheetToObjects_(tasksSheet).map(t => ({
    id: t.ID,
    name: t.Name,
    project: t.Project,
    showDate: formatDate_(t.ShowDate),
    showTime: formatTime_(t.ShowTime),
    durationMinutes: t.DurationMinutes || '',
    deadline: formatDate_(t.Deadline),
    status: t.Status || 'Not started',
    isTemplate: t.IsTemplate === true || t.IsTemplate === 'TRUE',
    recurrenceType: t.RecurrenceType || '',
    recurrenceInterval: t.RecurrenceInterval || '',
    nextRunDate: formatDate_(t.NextRunDate),
    parentTemplateId: t.ParentTemplateId || '',
    syncToCalendar: t.SyncToCalendar === true || t.SyncToCalendar === 'TRUE',
    notes: t.Notes || ''
  }));

  const tasks = allTasks.filter(t => !t.isTemplate);
  const templates = allTasks.filter(t => t.isTemplate);

  // Projects sheet has no ID column, so sheetToObjects_ (which filters on
  // ID) can't be reused here — read the raw rows directly instead.
  const projects = projectsSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ name: r[0], color: r[1] }));

  return { tasks, templates, projects };
}

/**
 * One entry point for both plain and recurring tasks — the client sends
 * `repeats: true` plus recurrenceType/recurrenceInterval when the "Repeats"
 * checkbox is on, otherwise a normal task is created.
 */
function createTask(task) {
  if (task.repeats) {
    return createRecurringTemplate(task);
  }
  const sheet = getSheet_(TASKS_SHEET);
  const id = newId_();
  sheet.appendRow([
    id,
    task.name || 'Untitled',
    task.project || '',
    task.showDate || '',
    task.showTime || '',
    task.showTime ? (task.durationMinutes || 60) : '',
    task.deadline || '',
    task.status || 'Not started',
    false,
    '', '', '', '',
    !!task.syncToCalendar,
    new Date(),
    '',
    task.notes || ''
  ]);
  return id;
}

function updateTask(id, updates) {
  const sheet = getSheet_(TASKS_SHEET);
  const row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found: ' + id);

  const fieldToCol = {
    name: 2, project: 3, showDate: 4, showTime: 5, durationMinutes: 6,
    deadline: 7, status: 8, syncToCalendar: 14, notes: 17
  };
  Object.keys(updates).forEach(key => {
    if (fieldToCol[key]) {
      sheet.getRange(row, fieldToCol[key]).setValue(updates[key]);
    }
  });
  return true;
}

function deleteTask(id) {
  const sheet = getSheet_(TASKS_SHEET);
  const row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found: ' + id);
  sheet.deleteRow(row);
  return true;
}

function toggleDone(id) {
  const sheet = getSheet_(TASKS_SHEET);
  const row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found: ' + id);
  const statusCell = sheet.getRange(row, 8);
  const current = statusCell.getValue();
  statusCell.setValue(current === 'Done' ? 'Not started' : 'Done');
  return true;
}

function addProject(name, color) {
  const sheet = getSheet_(PROJECTS_SHEET);
  sheet.appendRow([name, color]);
  return true;
}

// ---------- Recurring tasks ----------

/**
 * Creates a recurring template. Called directly from createTask() when the
 * client sets `repeats: true`.
 * recurrenceType: 'DAILY' | 'WEEKLY' | 'EVERY_N_DAYS' | 'MONTHLY'
 * recurrenceInterval: number of days for EVERY_N_DAYS, ignored otherwise.
 * startDate: 'yyyy-MM-dd' — first date an instance should be created for.
 * showTime/durationMinutes optionally carried onto every generated instance.
 */
function createRecurringTemplate(template) {
  const sheet = getSheet_(TASKS_SHEET);
  const id = newId_();
  const startDate = template.startDate || template.showDate || formatDate_(new Date());
  const showTime = template.showTime || '';
  const duration = showTime ? (template.durationMinutes || 60) : '';
  const sync = !!template.syncToCalendar;

  sheet.appendRow([
    id,
    template.name || 'Untitled',
    template.project || '',
    '',
    showTime,
    duration,
    '',
    'Not started',
    true,
    template.recurrenceType,
    template.recurrenceInterval || 1,
    // NextRunDate = the occurrence AFTER the first one, since the first
    // occurrence is created directly below, on the exact date chosen —
    // same as Apple Calendar showing a new recurring event immediately
    // on its start date rather than waiting for the next occurrence.
    nextRunAfter_(startDate, template.recurrenceType, template.recurrenceInterval),
    '',
    sync,
    new Date(),
    '',
    template.notes || ''
  ]);

  sheet.appendRow([
    newId_(), template.name || 'Untitled', template.project || '', startDate,
    showTime, duration, '', 'Not started', false, '', '', '', id, sync, new Date(),
    '', template.notes || ''
  ]);

  // Catch up on any further occurrences if the chosen start date is more
  // than one cycle in the past (e.g. backfilling a recurring task that
  // "should have" started weeks ago).
  generateRecurringTasks();
  return id;
}

/** Edits an existing recurring template's settings (name, project, time,
 *  duration, frequency, interval, sync). Doesn't touch NextRunDate, so the
 *  next scheduled instance still lands on its existing date — only future
 *  ones after that follow the newly edited frequency/interval. */
function updateRecurringTemplate(id, updates) {
  const sheet = getSheet_(TASKS_SHEET);
  const row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Recurring task not found: ' + id);

  const fieldToCol = {
    name: 2, project: 3, showTime: 5, durationMinutes: 6,
    recurrenceType: 10, recurrenceInterval: 11, syncToCalendar: 14, notes: 17
  };
  Object.keys(updates).forEach(key => {
    if (fieldToCol[key]) {
      sheet.getRange(row, fieldToCol[key]).setValue(updates[key]);
    }
  });
  return true;
}

function deleteTemplate(id) {
  return deleteTask(id);
}

function addDays_(dateStr, days) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + days);
  return formatDate_(d);
}

function nextRunAfter_(nextRunDate, recurrenceType, interval) {
  switch (recurrenceType) {
    case 'DAILY': return addDays_(nextRunDate, 1);
    case 'WEEKLY': return addDays_(nextRunDate, 7);
    case 'EVERY_N_DAYS': return addDays_(nextRunDate, Number(interval) || 1);
    case 'MONTHLY': {
      const d = new Date(nextRunDate + 'T00:00:00');
      d.setMonth(d.getMonth() + 1);
      return formatDate_(d);
    }
    default: return addDays_(nextRunDate, 7);
  }
}

/**
 * Runs daily (installed by setup()). For every recurring template whose
 * NextRunDate has arrived, creates a real task instance (inheriting the
 * template's time/duration) and advances the template's NextRunDate.
 */
function generateRecurringTasks() {
  const sheet = getSheet_(TASKS_SHEET);
  const data = sheet.getDataRange().getValues();
  const today = formatDate_(new Date());

  // Column indices (0-based) per TASK_HEADERS order.
  const COL = { project: 2, showTime: 4, duration: 5, isTemplate: 8,
                recType: 9, recInterval: 10, nextRun: 11, sync: 13, notes: 16 };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const name = row[1];
    const project = row[COL.project];
    const showTime = row[COL.showTime];
    const duration = row[COL.duration];
    const isTemplate = row[COL.isTemplate];
    const recurrenceType = row[COL.recType];
    const interval = row[COL.recInterval];
    const nextRunDate = row[COL.nextRun];
    const sync = row[COL.sync];
    const notes = row[COL.notes];

    if (!isTemplate) continue;
    if (!nextRunDate) continue;

    let runDate = formatDate_(nextRunDate);
    let guard = 0;
    while (runDate <= today && guard < 20) {
      sheet.appendRow([
        newId_(), name, project, runDate, showTime, duration, '',
        'Not started', false, '', '', '', id, sync, new Date(), '', notes
      ]);
      runDate = nextRunAfter_(runDate, recurrenceType, interval);
      guard++;
    }
    sheet.getRange(i + 1, COL.nextRun + 1).setValue(runDate);
  }
}

// ---------- Apple Calendar import ----------
// One-way pull: Apple Calendar > right-click a calendar > Share Calendar >
// check "Public Calendar" > copy the link. That link is itself an ICS feed —
// this fetches it, parses each VEVENT, and inserts any not already imported
// (matched by UID, stored in the ExternalUID column, so re-running never
// duplicates).
//
// Known limitation: DTSTART/DTEND are read as literal wall-clock values.
// If your Mac's calendar and this script's timezone differ, timed events
// may land an hour or two off — ping me if that turns out to matter and
// I'll add proper TZID conversion.

function getAppleCalendarUrl() {
  return PropertiesService.getScriptProperties().getProperty(APPLE_ICS_PROP_KEY) || '';
}

function importFromAppleCalendar(url) {
  const props = PropertiesService.getScriptProperties();
  if (url) props.setProperty(APPLE_ICS_PROP_KEY, url);
  const savedUrl = props.getProperty(APPLE_ICS_PROP_KEY);
  if (!savedUrl) throw new Error('No Apple Calendar link provided yet.');

  const fetchUrl = savedUrl.replace(/^webcal:\/\//i, 'https://');
  const text = UrlFetchApp.fetch(fetchUrl).getContentText();
  const events = parseIcs_(text);

  // Public calendar feeds include the entire history — only pull events
  // from two weeks back through a year out, so this doesn't flood the
  // sheet with years of past events every run.
  const today = new Date();
  const earliest = formatDate_(new Date(today.getTime() - 14 * 86400000));
  const latest = formatDate_(new Date(today.getTime() + 365 * 86400000));

  const sheet = getSheet_(TASKS_SHEET);
  const existingUids = new Set(
    sheetToObjects_(sheet).map(t => t.ExternalUID).filter(Boolean)
  );

  let imported = 0;
  events.forEach(ev => {
    if (!ev.uid || existingUids.has(ev.uid)) return;
    if (!ev.date || ev.date < earliest || ev.date > latest) return;
    sheet.appendRow([
      newId_(),
      ev.summary || 'Untitled',
      '', // no project — Apple Calendar events are calendar-only, not part of the Projects board
      ev.date || '',
      ev.time || '',
      ev.time ? (ev.durationMinutes || 60) : '',
      '',
      'Not started',
      false, '', '', '', '',
      false,
      new Date(),
      ev.uid,
      [ev.location ? ('📍 ' + ev.location) : '', ev.notes || ''].filter(Boolean).join('\n\n')
    ]);
    imported++;
  });

  Logger.log('Imported ' + imported + ' event(s) from Apple Calendar.');
  return imported;
}

/** Minimal ICS parser: unfolds continuation lines, extracts VEVENT blocks. */
/**
 * Parses a yyyyMMddTHHmmss(Z)? ICS datetime into a JS Date. UTC ("Z" suffix)
 * is converted properly; a bare local/floating value (no Z, no TZID
 * handling) is treated as a wall-clock time — see the known-limitation note
 * on importFromAppleCalendar above.
 */
function icsDateTimeToDate_(value) {
  const m = value.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z)?$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  const h = Number(m[4]), mi = Number(m[5]), s = Number(m[6]);
  return m[7]
    ? new Date(Date.UTC(y, mo, d, h, mi, s))
    : new Date(y, mo, d, h, mi, s);
}

function parseIcs_(text) {
  const rawLines = text.split(/\r\n|\n|\r/);
  const lines = [];
  rawLines.forEach(line => {
    if (/^[ \t]/.test(line) && lines.length) {
      lines[lines.length - 1] += line.slice(1);
    } else {
      lines.push(line);
    }
  });

  const events = [];
  let cur = null;

  function finalize(ev) {
    if (ev.dtstartDate) {
      const h = String(ev.dtstartDate.getHours()).padStart(2, '0');
      const mi = String(ev.dtstartDate.getMinutes()).padStart(2, '0');
      ev.time = h + ':' + mi;
      if (ev.dtendDate && ev.dtendDate > ev.dtstartDate) {
        ev.durationMinutes = Math.round((ev.dtendDate - ev.dtstartDate) / 60000);
      }
    }
    delete ev.dtstartDate;
    delete ev.dtendDate;
    return ev;
  }

  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { cur = {}; return; }
    if (line === 'END:VEVENT') { if (cur) events.push(finalize(cur)); cur = null; return; }
    if (!cur) return;

    const idx = line.indexOf(':');
    if (idx === -1) return;
    const keyPart = line.slice(0, idx);
    const value = line.slice(idx + 1);
    const key = keyPart.split(';')[0];

    if (key === 'UID') cur.uid = value;
    if (key === 'SUMMARY') cur.summary = value.replace(/\\,/g, ',').replace(/\\n/gi, ' ');
    if (key === 'DESCRIPTION') cur.notes = value.replace(/\\,/g, ',').replace(/\\n/gi, '\n');
    if (key === 'LOCATION') cur.location = value.replace(/\\,/g, ',');

    if (key === 'DTSTART') {
      const isDateOnly = /^\d{8}$/.test(value);
      if (isDateOnly) {
        cur.date = value.slice(0, 4) + '-' + value.slice(4, 6) + '-' + value.slice(6, 8);
      } else {
        const dt = icsDateTimeToDate_(value);
        if (dt) {
          cur.date = formatDate_(dt);
          cur.dtstartDate = dt;
        }
      }
    }

    if (key === 'DTEND') {
      const dt = icsDateTimeToDate_(value);
      if (dt) cur.dtendDate = dt;
    }
  });

  return events;
}
