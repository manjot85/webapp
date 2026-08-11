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
      isScheduledToday: isScheduledToday
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

function toggleTaskStatus(row, isChecked, userFullName, userTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTaskMasterSheet();
  const logSheet = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  
  const formattedActor = userTitle ? `${userFullName} (${userTitle})` : userFullName;
  const timestamp = new Date();
  const taskName = sheet.getRange(row, 2).getValue();

  sheet.getRange(row, 7).setValue(isChecked);
  sheet.getRange(row, 8).setValue(timestamp);
  sheet.getRange(row, 9).setValue(formattedActor);

  logSheet.appendRow([timestamp, formattedActor, "Task Master", taskName, isChecked ? "Completed" : "Removed"]);
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
    sheet.appendRow(["Date", "Task Row", "Task Name", "User Name", "Status", "Note", "Timestamp"]);
  }
  return sheet;
}

// Records a personal status for one user on one shared task, for today only.
// status: "Completed" (I finished my copy) | "Pending" (undo/revert) | "Deferred" (hand-off for today)
// note: optional free text (e.g. who is covering it, for "Deferred")
function setPersonalTaskStatus(row, userName, userTitle, status, note) {
  const sheet = getPersonalStatusSheet_();
  const taskSheet = getTaskMasterSheet();
  const taskName = taskSheet.getRange(row, 2).getValue();
  const todayStr = getTodayCstDateStr_();
  const timestamp = new Date();

  sheet.appendRow([todayStr, row, taskName, userName, status, note || "", timestamp]);

  // Only "Completed" earns Log-sheet credit (feeds the Progress report's per-user
  // completed count), matching how individual task completions are credited
  // elsewhere. "Pending" (undo) and "Deferred" (hand-off) are functional-only and
  // don't need to show up as report activity.
  if (status === "Completed") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
    const formattedActor = userTitle ? `${userName} (${userTitle})` : userName;
    logSheet.appendRow([timestamp, formattedActor, "Task Master", taskName, "Completed"]);
  }

  return { success: true };
}

// Returns today's latest personal status per (task row, user) as a flat array:
// [{ row, userName, status, note, timestamp }, ...]
function getTodayPersonalTaskStatuses() {
  const sheet = getPersonalStatusSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const todayStr = getTodayCstDateStr_();
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  const latest = {};
  data.forEach(r => {
    const dateStr = normalizeDateCell_(r[0]);
    if (dateStr !== todayStr) return;

    const row = r[1];
    const userName = r[3] ? r[3].toString() : "";
    if (!userName) return;
    const status = r[4] ? r[4].toString() : "";
    const note = r[5] ? r[5].toString() : "";
    const ts = r[6] instanceof Date ? r[6].getTime() : new Date(r[6]).getTime();

    const key = row + "|" + userName;
    if (!latest[key] || ts >= latest[key].ts) {
      latest[key] = { row: row, userName: userName, status: status, note: note, ts: ts };
    }
  });

  return Object.keys(latest).map(k => {
    const entry = latest[k];
    return { row: entry.row, userName: entry.userName, status: entry.status, note: entry.note };
  });
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

function getReportSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) return { totalCompleted: 0, totalRemoved: 0, userStats: {} };

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return { totalCompleted: 0, totalRemoved: 0, userStats: {} };

  const logs = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  // Compare dates using explicit CST (America/Chicago), not the server runtime's
  // ambient default timezone — otherwise "today" here can silently drift from the
  // CST shift date the rest of the app (and the person using it) is working off of.
  const todayStr = getTodayCstDateStr_();
  
  let totalCompleted = 0;
  let totalRemoved = 0;
  const userStats = {};

  logs.forEach(row => {
    const logDateRaw = row[0] ? new Date(row[0]) : null;
    if (!logDateRaw || isNaN(logDateRaw)) return;
    const logDate = Utilities.formatDate(logDateRaw, "America/Chicago", "yyyy-MM-dd");
    const rawUser = row[1] ? row[1].toString() : "Unknown";
    const cleanUser = rawUser.split(' (')[0].trim();
    const status = row[4];

    if (logDate === todayStr) {
      if (status === "Completed") totalCompleted++;
      if (status === "Removed") totalRemoved++;

      if (!userStats[cleanUser]) userStats[cleanUser] = { completed: 0, removed: 0 };
      if (status === "Completed") userStats[cleanUser].completed++;
      if (status === "Removed") userStats[cleanUser].removed++;
    }
  });

  return { todayStr, totalCompleted, totalRemoved, userStats };
}
