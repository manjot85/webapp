// ==========================================
// 1. SHEET & CONSTANTS
// ==========================================

const TASK_SHEET = "Task Master";
const TEAM_SHEET = "Team";
const LOG_SHEET = "Log";
const PERSONAL_STATUS_SHEET = "Personal Task Status";

// ==========================================
// DAY-NAME NORMALIZATION
// ==========================================
// "Working Days" (Team sheet) and "Recurring Days" (Task Master sheet) are both
// free-text comma lists. If they're not typed in the exact "Mon,Tue,Wed..."
// format the frontend compares against, a day-of-week check can silently fail —
// e.g. a Team row using "Tuesday" instead of "Tue" would never match today's
// CST day abbreviation, quietly blocking that person's entire My Tasks list
// while every other view (which doesn't do this exact-match check) still works.
// This normalizes common variants down to the canonical 3-letter abbreviation
// so entry-format inconsistencies between rows can't cause that.
const DAY_ALIASES_ = {
  'sun': 'Sun', 'sunday': 'Sun',
  'mon': 'Mon', 'monday': 'Mon',
  'tue': 'Tue', 'tues': 'Tue', 'tuesday': 'Tue',
  'wed': 'Wed', 'weds': 'Wed', 'wednesday': 'Wed',
  'thu': 'Thu', 'thur': 'Thu', 'thurs': 'Thu', 'thursday': 'Thu',
  'fri': 'Fri', 'friday': 'Fri',
  'sat': 'Sat', 'saturday': 'Sat'
};

// Coerces a "records completed" value to a non-negative integer. The frontend
// requires a positive count on every completion, but this guards the backend so
// a malformed/legacy call can't poison the Log/Personal Status sheets.
function parseCount_(val) {
  const n = parseInt(val, 10);
  return isNaN(n) || n < 0 ? 0 : n;
}

// Returns the Log sheet, creating it (or back-filling the Records Count header)
// on first use so both new and legacy spreadsheets get the extra column.
function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.appendRow(["Timestamp", "User", "Sheet Name", "Task", "Status", "Records Count"]);
  } else if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Timestamp", "User", "Sheet Name", "Task", "Status", "Records Count"]);
  } else if (logSheet.getRange(1, 6).getValue() !== "Records Count") {
    logSheet.getRange(1, 6).setValue("Records Count");
  }
  return logSheet;
}

function normalizeDayAbbrev_(raw) {
  const key = (raw || '').toString().trim().toLowerCase();
  if (!key) return '';
  return DAY_ALIASES_[key] || raw.toString().trim();
}

// Parses a free-text days cell into a normalized ["Mon","Tue",...] array.
// Also expands a few common shorthand phrases people tend to type by hand.
function parseDaysList_(raw, defaultDays) {
  const str = (raw || '').toString().trim();
  if (!str) return defaultDays ? defaultDays.slice() : [];

  const lower = str.toLowerCase();
  if (lower === 'm-f' || lower === 'mon-fri' || lower === 'weekdays' || lower === 'all weekdays' || lower === 'weekday') {
    return ["Mon", "Tue", "Wed", "Thu", "Fri"];
  }
  if (lower === 'daily' || lower === 'everyday' || lower === 'every day' || lower === '7 days' || lower === 'all days' || lower === 'every day of the week') {
    return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  }

  return str.split(',').map(s => normalizeDayAbbrev_(s)).filter(Boolean);
}

// ==========================================
// 2. TEAM & WORKING DAYS MANAGEMENT
// ==========================================

function getTeamData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let teamSheet = ss.getSheetByName(TEAM_SHEET);
  
  if (!teamSheet) {
    teamSheet = ss.insertSheet(TEAM_SHEET);
    teamSheet.appendRow(["Name", "Title", "Status", "Working Days"]);
  }

  const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
  const data = teamSheet.getDataRange().getValues();
  const members = [];

  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : "";
    const title = data[i][1] ? data[i][1].toString().trim() : "Coordinator";
    const status = data[i][2] ? data[i][2].toString().trim() : "Active";
    const daysRaw = data[i][3] ? data[i][3].toString().trim() : "Mon,Tue,Wed,Thu,Fri";

    if (name && status.toLowerCase() === "active") {
      members.push({
        row: i + 1,
        name: name,
        title: title,
        workingDays: parseDaysList_(daysRaw, ["Mon", "Tue", "Wed", "Thu", "Fri"])
      });
    }
  }

  function getRolePriority(title) {
    const lower = title.toLowerCase();
    if (lower.includes('manager') || lower.includes('am')) return 1;
    if (lower.includes('supervisor')) return 2;
    if (lower.includes('sr.') || lower.includes('senior')) return 3;
    return 4;
  }

  members.sort((a, b) => getRolePriority(a.title) - getRolePriority(b.title));

  return {
    currentUser: { email: currentUserEmail },
    members: members
  };
}

function saveTeamMember(memberData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let teamSheet = ss.getSheetByName(TEAM_SHEET);
  const daysStr = Array.isArray(memberData.workingDays) ? memberData.workingDays.join(',') : memberData.workingDays;

  if (memberData.row) {
    teamSheet.getRange(memberData.row, 1, 1, 4).setValues([[
      memberData.name,
      memberData.title,
      "Active",
      daysStr
    ]]);
  } else {
    teamSheet.appendRow([memberData.name, memberData.title, "Active", daysStr]);
  }
  return { success: true };
}

function deleteTeamMember(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamSheet = ss.getSheetByName(TEAM_SHEET);
  teamSheet.getRange(row, 3).setValue("Inactive");
  return { success: true };
}

// ==========================================
// 3. TASK MASTER ENGINE
// ==========================================

function getTaskMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_SHEET);
    sheet.appendRow([
      "ID", 
      "Task Name", 
      "Assigned To", 
      "Is Recurring", 
      "Recurring Days", 
      "Status", 
      "Completed", 
      "Last Updated", 
      "Updated By", 
      "Task Type", 
      "Dashboard Name", 
      "Dashboard Link",
      "Due Date"
    ]);
  } else if (sheet.getRange(1, 13).getValue() !== "Due Date") {
    // Backward-compatible migration: older sheets won't have the Due Date column yet.
    sheet.getRange(1, 13).setValue("Due Date");
  }
  return sheet;
}

function getAllTasksMaster() {
  const sheet = getTaskMasterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const todayDay = getTodayCstDayName_();

  return data.map((row, idx) => {
    const id = row[0] || (idx + 2);
    const taskName = row[1] ? row[1].toString().trim() : "";
    const assignees = row[2] ? row[2].toString().split(',').map(s => s.trim()).filter(Boolean) : [];
    const isRecurring = Boolean(row[3]);
    const recurringDays = parseDaysList_(row[4], []);
    const status = row[5] ? row[5].toString().trim() : "Active";
    const isActive = status.toLowerCase() === "active";
    const completed = Boolean(row[6]);
    const lastUpdatedRaw = row[7] ? new Date(row[7]) : null;
    const lastUpdated = lastUpdatedRaw && !isNaN(lastUpdatedRaw) ? Utilities.formatDate(lastUpdatedRaw, "America/Chicago", "hh:mm a") : '';
    const updatedBy = row[8] || '';
    const taskType = row[9] || 'Report';
    const dashboardName = row[10] ? row[10].toString().trim() : '';
    const dashboardLink = row[11] ? row[11].toString().trim() : '';

    // The global "Completed" flag carries no date, but its Last Updated timestamp
    // does. Reports need to know WHICH day a task was completed for all so a
    // recurring task finished on Monday shows as pending again on Tuesday.
    const completedDate = completed && lastUpdatedRaw && !isNaN(lastUpdatedRaw)
      ? Utilities.formatDate(lastUpdatedRaw, "America/Chicago", "yyyy-MM-dd")
      : '';

    const dueDateRaw = row[12] ? new Date(row[12]) : null;
    const hasDueDate = dueDateRaw && !isNaN(dueDateRaw);
    const dueDate = hasDueDate ? Utilities.formatDate(dueDateRaw, "America/Chicago", "MMM d, yyyy") : '';
    const dueDateISO = hasDueDate ? Utilities.formatDate(dueDateRaw, "America/Chicago", "yyyy-MM-dd") : '';

    const isScheduledToday = !isRecurring || recurringDays.includes(todayDay);

    return {
      row: idx + 2,
      id: id,
      taskName: taskName,
      assignedTo: assignees,
      isRecurring: isRecurring,
      recurringDays: recurringDays,
      status: status,
      isActive: isActive,
      completed: completed,
      lastUpdated: lastUpdated,
      updatedBy: updatedBy,
      taskType: taskType,
      dashboardName: dashboardName,
      dashboardLink: dashboardLink,
      dueDate: dueDate,
      dueDateISO: dueDateISO,
      isScheduledToday: isScheduledToday,
      completedDate: completedDate
    };
  }).filter(t => t.taskName !== "");
}

function saveTaskMaster(taskData) {
  const sheet = getTaskMasterSheet();
  const assigneesStr = Array.isArray(taskData.assignedTo) ? taskData.assignedTo.join(',') : (taskData.assignedTo || '');
  const daysStr = Array.isArray(taskData.recurringDays) ? taskData.recurringDays.join(',') : (taskData.recurringDays || '');
  const todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy HH:mm:ss");
  const updatedBy = taskData.currentUserName || "Web App";
  const statusStr = (taskData.isActive === false || taskData.isActive === 'false') ? "Inactive" : "Active";
  // Due Date only applies to non-recurring tasks; ignore it otherwise so stale dates don't linger.
  const isRecurringBool = (taskData.isRecurring === true || taskData.isRecurring === 'true');
  const dueDateStr = (!isRecurringBool && taskData.dueDate) ? taskData.dueDate : "";

  if (taskData.row) {
    const rowNum = parseInt(taskData.row, 10);
    sheet.getRange(rowNum, 2, 1, 5).setValues([[
      taskData.taskName,
      assigneesStr,
      taskData.isRecurring,
      daysStr,
      statusStr
    ]]);
    sheet.getRange(rowNum, 8, 1, 2).setValues([[
      todayStr,
      updatedBy
    ]]);
    sheet.getRange(rowNum, 10, 1, 3).setValues([[
      taskData.taskType || "Report",
      taskData.dashboardName || "",
      taskData.dashboardLink || ""
    ]]);
    sheet.getRange(rowNum, 13, 1, 1).setValues([[dueDateStr]]);
  } else {
    const id = "TASK-" + new Date().getTime();
    sheet.appendRow([
      id, 
      taskData.taskName, 
      assigneesStr, 
      taskData.isRecurring, 
      daysStr, 
      statusStr, 
      false, 
      todayStr, 
      updatedBy, 
      taskData.taskType || "Report", 
      taskData.dashboardName || "", 
      taskData.dashboardLink || "",
      dueDateStr
    ]);
  }
  return { success: true };
}

function setTaskActiveStatus(row, isActive, userFullName) {
  const sheet = getTaskMasterSheet();
  const statusStr = isActive ? "Active" : "Inactive";
  const todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy HH:mm:ss");
  
  sheet.getRange(row, 6).setValue(statusStr);
  sheet.getRange(row, 8).setValue(todayStr);
  sheet.getRange(row, 9).setValue(userFullName || "System");
  return { success: true };
}

function deleteTaskMaster(row) {
  return setTaskActiveStatus(row, false, "System");
}

// ==========================================
// 3b. DUPLICATE TASK DETECTION & CLEANUP
// ==========================================

// Normalizes a task name for duplicate comparison (trim + lowercase + collapse spaces)
function normalizeTaskName_(name) {
  return (name || "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}

// Scans the Task Master sheet and returns groups of ACTIVE tasks that share
// the same (normalized) task name. Each group has 2+ entries.
// This is used by the "Find & Clean Duplicates" tool in Task Setup.
function findDuplicateTaskGroups() {
  const sheet = getTaskMasterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const groups = {};

  data.forEach((row, idx) => {
    const taskName = row[1] ? row[1].toString().trim() : "";
    if (!taskName) return;

    const status = row[5] ? row[5].toString().trim() : "Active";
    if (status.toLowerCase() !== "active") return;

    const key = normalizeTaskName_(taskName);
    if (!groups[key]) groups[key] = [];

    groups[key].push({
      row: idx + 2,
      taskName: taskName,
      assignedTo: row[2] ? row[2].toString().split(',').map(s => s.trim()).filter(Boolean) : [],
      dashboardName: row[10] ? row[10].toString().trim() : '',
      lastUpdated: row[7] ? row[7].toString() : '',
      updatedBy: row[8] || ''
    });
  });

  return Object.keys(groups)
    .map(key => groups[key])
    .filter(g => g.length > 1);
}

// Merges one or more groups of duplicate task rows into a single "keep" row per group,
// combining assignees + recurring days, then removes the duplicate rows entirely.
// groups: [{ keepRow: <rowNum>, duplicateRows: [<rowNum>, ...] }, ...]
function mergeDuplicateTaskGroups(groups) {
  if (!groups || groups.length === 0) return { success: true, merged: 0 };

  const sheet = getTaskMasterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, merged: 0 };

  const allData = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const getRowData = (rowNum) => allData[rowNum - 2];

  const todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy HH:mm:ss");
  const rowsToDelete = [];

  groups.forEach(g => {
    const keepRowNum = parseInt(g.keepRow, 10);
    const dupRowNums = (g.duplicateRows || []).map(r => parseInt(r, 10)).filter(r => r && r !== keepRowNum);
    const keepData = getRowData(keepRowNum);
    if (!keepData) return;

    const allRowNums = [keepRowNum].concat(dupRowNums);

    const mergedAssignees = new Set();
    const mergedDays = new Set();
    let mergedIsRecurring = false;
    let mergedTaskType = '';
    let mergedDashboardName = '';
    let mergedDashboardLink = '';
    let mergedDueDate = '';

    allRowNums.forEach(rn => {
      const rowData = getRowData(rn);
      if (!rowData) return;

      const assignees = rowData[2] ? rowData[2].toString().split(',').map(s => s.trim()).filter(Boolean) : [];
      assignees.forEach(a => mergedAssignees.add(a));

      const isRec = Boolean(rowData[3]);
      const days = rowData[4] ? rowData[4].toString().split(',').map(s => s.trim()).filter(Boolean) : [];
      if (isRec) {
        mergedIsRecurring = true;
        days.forEach(d => mergedDays.add(d));
      }

      if (!mergedTaskType && rowData[9]) mergedTaskType = rowData[9].toString().trim();
      if (!mergedDashboardName && rowData[10]) mergedDashboardName = rowData[10].toString().trim();
      if (!mergedDashboardLink && rowData[11]) mergedDashboardLink = rowData[11].toString().trim();
      if (!mergedDueDate && rowData[12]) mergedDueDate = rowData[12];
    });

    // Recurring tasks don't carry a Due Date.
    if (mergedIsRecurring) mergedDueDate = '';

    // A task with no assignee is now invisible in everyone's My Tasks list, so if
    // ANY duplicate copy had real assignees, those take precedence — the merged
    // task should still be visible to whoever was actually assigned to work on it.
    // Only stays unassigned if every single copy was unassigned.
    const finalAssignees = Array.from(mergedAssignees);

    sheet.getRange(keepRowNum, 2, 1, 5).setValues([[
      keepData[1], // preserve the kept row's original task name text/casing
      finalAssignees.join(','),
      mergedIsRecurring,
      Array.from(mergedDays).join(','),
      "Active"
    ]]);
    sheet.getRange(keepRowNum, 8, 1, 2).setValues([[todayStr, "Duplicate Cleanup"]]);
    sheet.getRange(keepRowNum, 13, 1, 1).setValues([[mergedDueDate]]);
    sheet.getRange(keepRowNum, 10, 1, 3).setValues([[
      mergedTaskType || "Report",
      mergedDashboardName,
      mergedDashboardLink
    ]]);

    dupRowNums.forEach(r => rowsToDelete.push(r));
  });

  // Delete rows highest-to-lowest so earlier deletions never shift the row
  // numbers of rows still queued for deletion.
  const uniqueRowsToDelete = Array.from(new Set(rowsToDelete)).sort((a, b) => b - a);
  uniqueRowsToDelete.forEach(r => sheet.deleteRow(r));

  return { success: true, merged: uniqueRowsToDelete.length };
}

function toggleTaskStatus(row, isChecked, userFullName, userTitle, recordsCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTaskMasterSheet();
  const logSheet = getLogSheet_();
  
  const formattedActor = userTitle ? `${userFullName} (${userTitle})` : userFullName;
  const timestamp = new Date();
  const taskName = sheet.getRange(row, 2).getValue();

  sheet.getRange(row, 7).setValue(isChecked);
  sheet.getRange(row, 8).setValue(timestamp);
  sheet.getRange(row, 9).setValue(formattedActor);

  const count = isChecked ? parseCount_(recordsCount) : '';
  logSheet.appendRow([timestamp, formattedActor, "Task Master", taskName, isChecked ? "Completed" : "Removed", count]);

  // Also record the actor's per-day personal status so per-user record counts
  // resolve in My Tasks rows and the drilldown even for single-assignee tasks
  // and team-wide completions. A Pending row on uncheck keeps "latest wins" sane.
  const personalSheet = getPersonalStatusSheet_();
  personalSheet.appendRow([
    getTodayCstDateStr_(), row, taskName, userFullName,
    isChecked ? "Completed" : "Pending", "", timestamp, count
  ]);

  return { success: true };
}

// ==========================================
// 3c. PERSONAL (PER-USER, PER-DAY) TASK STATUS
// ==========================================
// Shared/multi-assignee "team report" tasks (e.g. one report covered by any of
// several Coordinators) need completion tracked per person per day, separate
// from the task's own shared "Completed" flag — which is reserved for the
// explicit "Mark Complete for All" action that clears the report from
// everyone's queue at once. This is an append-only log; the most recent entry
// for a given (date, row, user) wins.

// Returns "America/Chicago" (CST/CDT) formatted as yyyy-MM-dd, matching what the
// client computes for "today" so both sides agree on the shift date.
function getTodayCstDateStr_() {
  return Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd");
}

// Returns today's CST weekday abbreviation ("Sun".."Sat"), computed via pure date
// math (not locale-dependent formatting) so it always matches the WEEK_DAYS values
// used throughout the app regardless of the script/account's locale settings.
function getTodayCstDayName_() {
  const cstDateStr = getTodayCstDateStr_();
  const parts = cstDateStr.split("-");
  const utcDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
  return ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][utcDate.getUTCDay()];
}

// Normalizes a sheet cell that may have been auto-converted to a Date by Sheets
// even though it was written as a "yyyy-MM-dd" string.
function normalizeDateCell_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "America/Chicago", "yyyy-MM-dd");
  }
  return (val || "").toString();
}

function getPersonalStatusSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PERSONAL_STATUS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PERSONAL_STATUS_SHEET);
    sheet.appendRow(["Date", "Task Row", "Task Name", "User Name", "Status", "Note", "Timestamp", "Records Count"]);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Task Row", "Task Name", "User Name", "Status", "Note", "Timestamp", "Records Count"]);
  } else if (sheet.getRange(1, 8).getValue() !== "Records Count") {
    sheet.getRange(1, 8).setValue("Records Count");
  }
  return sheet;
}

// Records a personal status for one user on one shared task, for today only.
// status: "Completed" (I finished my copy) | "Pending" (undo/revert) | "Deferred" (hand-off for today)
// note: optional free text (e.g. who is covering it, for "Deferred")
// recordsCount: required for "Completed" status — number of records processed.
function setPersonalTaskStatus(row, userName, userTitle, status, note, recordsCount) {
  const sheet = getPersonalStatusSheet_();
  const taskSheet = getTaskMasterSheet();
  const taskName = taskSheet.getRange(row, 2).getValue();
  const todayStr = getTodayCstDateStr_();
  const timestamp = new Date();
  const count = status === "Completed" ? parseCount_(recordsCount) : '';

  sheet.appendRow([todayStr, row, taskName, userName, status, note || "", timestamp, count]);

  // Only "Completed" earns Log-sheet credit (feeds the Progress report's per-user
  // completed count), matching how individual task completions are credited
  // elsewhere. "Pending" (undo) and "Deferred" (hand-off) are functional-only and
  // don't need to show up as report activity.
  if (status === "Completed") {
    const logSheet = getLogSheet_();
    const formattedActor = userTitle ? `${userName} (${userTitle})` : userName;
    logSheet.appendRow([timestamp, formattedActor, "Task Master", taskName, "Completed", count]);
  }

  return { success: true };
}

// Returns the latest personal status per (task row, user) for a given CST date
// as a flat array: [{ row, userName, status, note, recordsCount }, ...]
function getPersonalStatusesForDate_(dateStr) {
  const sheet = getPersonalStatusSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  const latest = {};
  data.forEach(r => {
    const dateValue = normalizeDateCell_(r[0]);
    if (dateValue !== dateStr) return;

    const row = r[1];
    const userName = r[3] ? r[3].toString() : "";
    if (!userName) return;
    const status = r[4] ? r[4].toString() : "";
    const note = r[5] ? r[5].toString() : "";
    const recordsCount = parseCount_(r[7]);
    const ts = r[6] instanceof Date ? r[6].getTime() : new Date(r[6]).getTime();

    const key = row + "|" + userName;
    if (!latest[key] || ts >= latest[key].ts) {
      latest[key] = { row: row, userName: userName, status: status, note: note, recordsCount: recordsCount, ts: ts };
    }
  });

  return Object.keys(latest).map(k => {
    const entry = latest[k];
    return { row: entry.row, userName: entry.userName, status: entry.status, note: entry.note, recordsCount: entry.recordsCount };
  });
}

// Today's per-user statuses (used by the My Tasks view).
function getTodayPersonalTaskStatuses() {
  return getPersonalStatusesForDate_(getTodayCstDateStr_());
}

// Per-user statuses for an arbitrary report date (yyyy-MM-dd, CST). The Daily
// Reports tab asks for a specific date so pending/completed can be shown for any
// shift, not just today's.
function getDailyReportStatuses(dateStr) {
  const targetDate = dateStr || getTodayCstDateStr_();
  return getPersonalStatusesForDate_(targetDate);
}

// Resolves a Log-sheet "User" cell to a canonical team member name. The app
// writes "Name (Title)" rows; legacy rows hold emails. Emails are matched by
// looking for a member's surname in the email's local part, but only a unique
// match is accepted — unmappable or ambiguous rows are ignored rather than
// guessed, so attribution stays conservative.
function resolveLogUser_(rawUser, memberNames) {
  const value = (rawUser || "").toString().trim();
  if (!value) return "";

  let name = value;
  const parenIdx = name.indexOf(" (");
  if (parenIdx > 0) name = name.substring(0, parenIdx).trim();

  if (memberNames[name]) return name;

  if (name.indexOf("@") > 0) {
    const local = name.split("@")[0].toLowerCase();
    const matches = Object.keys(memberNames).filter(n => {
      const tokens = n.toLowerCase().split(/\s+/);
      const surname = tokens[tokens.length - 1];
      return surname.length >= 4 && local.indexOf(surname) !== -1;
    });
    return matches.length === 1 ? matches[0] : "";
  }
  return "";
}

// Per-day completion matrix for a date range (yyyy-MM-dd, CST). For each active
// team member and each date in the range, counts the tasks they logged as
// "Completed" in the Log sheet (the app's completion log: Timestamp, User, Sheet
// Name, Task, Status, Records Count) plus the sum of their record counts. The
// Personal Task Status sheet is merged in as a fallback for per-user entries.
// Latest entry wins per (date, user, task) so undo/"Removed" entries are
// respected. Changing the date range changes what this returns, because the
// report is derived purely from when each completion was logged by whom.
function getDailyCompletionReport(dateFrom, dateTo) {
  const fromStr = dateFrom || getTodayCstDateStr_();
  const toStr = dateTo || getTodayCstDateStr_();
  const members = getTeamData().members.map(m => ({ name: m.name, title: m.title }));
  const memberNames = {};
  members.forEach(m => { memberNames[m.name] = true; });

  // Collapse to the latest status per (date, user, task) across both sheets,
  // then only count entries whose final status for that day was "Completed".
  const latest = {};
  const addEntry = (dateStr, userName, taskName, status, recordsCount, tsMs) => {
    if (dateStr < fromStr || dateStr > toStr) return;
    if (!userName || !taskName) return;
    const key = dateStr + "|" + userName + "|" + taskName;
    if (!latest[key] || tsMs >= latest[key].ts) {
      latest[key] = { date: dateStr, userName: userName, taskName: taskName, status: status, recordsCount: recordsCount, ts: tsMs };
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Log sheet — primary source. Legacy rows carry an email in the User column,
  // current app rows carry "Name (Title)".
  const logSheet = ss.getSheetByName(LOG_SHEET);
  if (logSheet) {
    const logRows = logSheet.getLastRow();
    if (logRows >= 2) {
      const logCols = Math.max(1, Math.min(logSheet.getLastColumn(), 6));
      logSheet.getRange(2, 1, logRows - 1, logCols).getValues().forEach(r => {
        const tsRaw = r[0];
        const tsDate = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
        if (!tsDate || isNaN(tsDate.getTime())) return;
        const dateStr = Utilities.formatDate(tsDate, "America/Chicago", "yyyy-MM-dd");
        if (dateStr < fromStr || dateStr > toStr) return;
        addEntry(
          dateStr,
          resolveLogUser_(r[1], memberNames),
          r[3] ? r[3].toString().trim() : "",
          r[4] ? r[4].toString().trim() : "",
          logCols >= 6 ? parseCount_(r[5]) : 0,
          tsDate.getTime()
        );
      });
    }
  }

  // 2) Personal Task Status sheet — fallback for per-user entries and record
  // counts that only exist there.
  const pSheet = ss.getSheetByName(PERSONAL_STATUS_SHEET);
  if (pSheet) {
    const pRows = pSheet.getLastRow();
    if (pRows >= 2) {
      pSheet.getRange(2, 1, pRows - 1, 8).getValues().forEach(r => {
        const dateStr = normalizeDateCell_(r[0]);
        if (dateStr < fromStr || dateStr > toStr) return;
        const userName = r[3] ? r[3].toString().trim() : "";
        if (!memberNames[userName]) return;
        const tsRaw = r[6] instanceof Date ? r[6].getTime() : new Date(r[6]).getTime();
        if (isNaN(tsRaw)) return;
        addEntry(
          dateStr,
          userName,
          r[2] ? r[2].toString().trim() : "",
          r[4] ? r[4].toString().trim() : "",
          parseCount_(r[7]),
          tsRaw
        );
      });
    }
  }

  // 3) Aggregate completed entries per (date, member), keeping the task list so
  // the frontend can show per-cell details on click.
  const dataByDate = {};
  Object.keys(latest).forEach(k => {
    const e = latest[k];
    if (e.status.toLowerCase() !== "completed") return;
    if (!dataByDate[e.date]) dataByDate[e.date] = {};
    const u = dataByDate[e.date][e.userName] ||
      (dataByDate[e.date][e.userName] = { completed: 0, records: 0, tasks: [] });
    u.completed += 1;
    u.records += e.recordsCount;
    u.tasks.push({ taskName: e.taskName, records: e.recordsCount });
  });

  // 4) Ordered day list built with pure UTC math so the yyyy-MM-dd strings stay
  // in the same frame as the sheet's dates (no DST/locale drift).
  const days = [];
  const startParts = fromStr.split("-");
  const endParts = toStr.split("-");
  const start = Date.UTC(parseInt(startParts[0], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[2], 10));
  const end = Date.UTC(parseInt(endParts[0], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[2], 10));
  for (let t = start; t <= end; t += 86400000) {
    const d = new Date(t);
    days.push(d.getUTCFullYear() + "-" + ("0" + (d.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + d.getUTCDate()).slice(-2));
  }

  return { members: members, days: days, data: dataByDate };
}

// ==========================================
// 4. APP INITIALIZATION & REPORTS
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dept of Responsibility')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
