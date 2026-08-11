// ==========================================
// GLOBALS & CONSTANTS
// ==========================================
const SHEET_QUESTIONS  = "Questions Tracker";
const SHEET_ANSWERED   = "Answered";
const SHEET_TEAM       = "Team Setup";

// Titles that count as "Support" (supervisor-tier) even if Category isn't set to "Support"
const SUPERVISOR_TITLES = ["Manager", "Assistant Manager", "Supervisor", "Escalation Supervisor"];

function doGet() {
  ensureSheetsExist();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Support Escalation Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_QUESTIONS)) {
    const s = ss.insertSheet(SHEET_QUESTIONS);
    s.appendRow(["Question", "Event/Wedding Name", "Asked By", "Created", "Due Date & Time", "Auto Priority", "Hours Left", "Case / Event Link", "Answer", "Status", "Assigned To"]);
  }

  if (!ss.getSheetByName(SHEET_ANSWERED)) {
    const s = ss.insertSheet(SHEET_ANSWERED);
    // Column 12 used to be a single "IsRead" boolean shared by everyone.
    // It is now "Read By" - a comma-separated list of the emails of everyone
    // who has personally marked the ticket read (see the read-tracking helpers below).
    s.appendRow(["Question", "Event/Wedding Name", "Asked By", "Created", "Answered Date", "Auto Priority", "Turnaround Hours", "Case / Event Link", "Answer", "Status", "Answered By", "Read By"]);
  }

  if (!ss.getSheetByName(SHEET_TEAM)) {
    const s = ss.insertSheet(SHEET_TEAM);
    // PasswordHash / PasswordSalt are only ever read/written server-side and are
    // NEVER returned by getTeamMembers() - the client never sees a password or hash.
    s.appendRow(["Name", "Title", "Status", "Category", "Email", "PasswordHash", "PasswordSalt"]);
    s.appendRow(["Amreen Sultana", "Coordinator", "active", "Coordinator", "asultana@orionphotogroup.com", "", ""]);
    s.appendRow(["Manjot Bhasin", "Assistant Manager", "active", "Support", "mbhasin@orionphotogroup.com", "", ""]);
    s.appendRow(["Harshita Jain", "Coordinator", "active", "Coordinator", "hjain@orionphotogroup.com", "", ""]);
  }
}

// ==========================================
// ROLE / AUTH HELPERS
// ==========================================
// IMPORTANT: This app's "login" is just a dropdown picker on the client, so
// currentUser on the browser side can never be trusted for permissions - anyone
// with the URL could open dev tools and call google.script.run functions directly
// pretending to be anyone. Every privileged server function below therefore
// re-derives the caller's role from the Team Setup sheet itself before acting.
// For real access control, deploy this as "Execute as: User accessing" restricted
// to your Workspace domain and cross-check Session.getActiveUser().getEmail()
// against the email passed in.

function isSupportMember(member) {
  if (!member) return false;
  if (String(member.status).toLowerCase() !== 'active') return false;
  return member.category === 'Support' || member.category === 'Admin' || SUPERVISOR_TITLES.indexOf(member.title) !== -1;
}

// Admin is a distinct, higher-trust tier from Support/Supervisor: only Admins
// may edit or permanently delete question/answer records. Set a team member's
// Category to "Admin" in the Team Setup tab to grant this.
function isAdminMember(member) {
  if (!member) return false;
  if (String(member.status).toLowerCase() !== 'active') return false;
  return member.category === 'Admin';
}

function findTeamMemberByEmail(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return null;
  const team = getTeamMembers();
  return team.find(m => m.email === email) || null;
}

// Throws unless the email belongs to a recognized, active team member (any role).
function requireTeamMember(email) {
  const member = findTeamMemberByEmail(email);
  if (!member) throw new Error("Access denied: unrecognized or inactive user.");
  return member;
}

// Throws unless the email belongs to an active Support / Supervisor-tier member.
function requireSupportRole(email) {
  const member = requireTeamMember(email);
  if (!isSupportMember(member)) {
    throw new Error("Access denied: this action requires a Supervisor / Support role.");
  }
  return member;
}

// Throws unless the email belongs to an active Admin.
function requireAdminRole(email) {
  const member = requireTeamMember(email);
  if (!isAdminMember(member)) {
    throw new Error("Access denied: this action requires the Admin role.");
  }
  return member;
}

function getTeamMembers() {
  ensureSheetsExist();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let team = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).trim() !== "") {
      team.push({
        rowIndex: i + 1,
        name: String(data[i][0]).trim(),
        title: String(data[i][1]).trim(),
        status: String(data[i][2]).trim(),
        category: String(data[i][3]).trim(),
        email: String(data[i][4]).trim().toLowerCase()
        // NOTE: PasswordHash / PasswordSalt (columns 6-7) are deliberately
        // omitted here. This function backs the login dropdown and is
        // callable before anyone is authenticated, so it must never leak
        // password material to the browser.
      });
    }
  }
  return team;
}

// ---- Password hashing (Admin accounts only) ----
// Server-side only. The client only ever sends a plaintext password attempt
// over the call itself (same as any login form); it never receives a hash,
// salt, or stored password back.
function hashPassword(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt));
  return Utilities.base64Encode(digest);
}

// Counts Admins that can actually log in (active + a password has been set).
// This is deliberately stricter than "any row tagged Admin" - an Admin row
// with no password yet must not close the bootstrap door, or nobody could
// ever fix it again.
function countWorkingAdmins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const category = String(data[i][3] || '').trim();
    const status = String(data[i][2] || '').trim().toLowerCase();
    const passwordHash = String(data[i][5] || '').trim();
    if (category === 'Admin' && status === 'active' && passwordHash) count++;
  }
  return count;
}

// Server-only: includes the password hash/salt, unlike getTeamMembers(). Never expose this return value to the client.
function getTeamMemberRawByEmail(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toLowerCase() === email) {
      return {
        rowIndex: i + 1,
        name: String(data[i][0]).trim(),
        title: String(data[i][1]).trim(),
        status: String(data[i][2]).trim(),
        category: String(data[i][3]).trim(),
        email: email,
        passwordHash: String(data[i][5] || '').trim(),
        passwordSalt: String(data[i][6] || '').trim()
      };
    }
  }
  return null;
}

// Called at login time whenever the selected profile is an Admin account.
// Throws on any failure so the client never has to infer "wrong password" vs "no account".
function verifyAdminPassword(email, passwordAttempt) {
  const member = getTeamMemberRawByEmail(email);
  if (!member || String(member.status).toLowerCase() !== 'active') {
    throw new Error("Access denied: unrecognized or inactive user.");
  }
  if (member.category !== 'Admin') {
    throw new Error("This account is not an Admin account.");
  }
  if (!member.passwordHash) {
    throw new Error("This Admin account has no password set yet. Ask another Admin to set one in Team Setup.");
  }
  const attemptHash = hashPassword(passwordAttempt, member.passwordSalt);
  if (attemptHash !== member.passwordHash) {
    throw new Error("Incorrect password.");
  }
  return { success: true };
}

function safeIsoDate(val) {
  if (!val) return new Date().toISOString();
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val.toISOString();
  }
  try {
    const parsed = new Date(val);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
  } catch (e) {}
  return new Date().toISOString();
}

// forcedPriority comes from the "Auto Priority" column. If the ticket was
// submitted with "Mark as Urgent Escalation" checked, that column is written
// as "P1" at submit time and stays P1 forever, regardless of elapsed hours.
function calculatePriorityAndHours(createdVal, forcedPriority) {
  if (!createdVal) return { priority: "P4", hoursElapsed: "0.0" };
  let created = (createdVal instanceof Date) ? createdVal : new Date(createdVal);
  if (isNaN(created.getTime())) return { priority: "P4", hoursElapsed: "0.0" };

  const now = new Date();
  const diffHours = (now - created) / (1000 * 60 * 60);
  const hoursElapsed = Math.max(0, diffHours).toFixed(1);

  const forced = String(forcedPriority || '').trim().toUpperCase();
  if (forced === 'P1' || forced === 'URGENT') {
    return { priority: "P1", hoursElapsed: hoursElapsed };
  }

  let priority = "P4";
  if (diffHours >= 6) priority = "P1";
  else if (diffHours >= 3) priority = "P2";
  else if (diffHours >= 1) priority = "P3";

  return { priority: priority, hoursElapsed: hoursElapsed };
}

// ---- Per-user "Read By" helpers ----
// Stored as a comma-separated list of lowercase emails in the "Read By" column.
// '*' is a legacy marker (from the old shared boolean) meaning "read by everyone".
function parseReadByList(raw) {
  if (raw === true) return ['*'];
  if (!raw) return [];
  return String(raw).split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}
function isReadByUser(raw, email) {
  const list = parseReadByList(raw);
  if (list.indexOf('*') !== -1) return true;
  return list.indexOf(String(email || '').trim().toLowerCase()) !== -1;
}

function getQuestionsData(requestingEmail) {
  ensureSheetsExist();
  requireTeamMember(requestingEmail); // must be a recognized logged-in user

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);

  let records = [];

  try {
    if (qSheet && qSheet.getLastRow() > 1) {
      const qData = qSheet.getDataRange().getValues();
      for (let i = 1; i < qData.length; i++) {
        const row = qData[i];
        const qText = String(row[0] || "").trim();
        if (!qText) continue;

        const createdIso = safeIsoDate(row[3]);
        const prioInfo = calculatePriorityAndHours(row[3], row[5]);

        records.push({
          id: "Q_" + (i + 1),
          rowIndex: i + 1,
          sheet: SHEET_QUESTIONS,
          question: qText,
          eventName: String(row[1] || "").trim(),
          askedBy: String(row[2] || "").trim(),
          created: createdIso,
          caseLink: String(row[7] || "").trim(),
          answer: String(row[8] || "").trim(),
          status: String(row[9] || "Open").trim(),
          assignedTo: String(row[10] || "").trim().toLowerCase(),
          priority: prioInfo.priority,
          hoursElapsed: prioInfo.hoursElapsed,
          isRead: true
        });
      }
    }

    if (aSheet && aSheet.getLastRow() > 1) {
      const aData = aSheet.getDataRange().getValues();
      for (let i = 1; i < aData.length; i++) {
        const row = aData[i];
        const qText = String(row[0] || "").trim();
        if (!qText) continue;

        const createdIso = safeIsoDate(row[3]);
        const answeredIso = safeIsoDate(row[4]);

        let turnaround = row[6];
        if (turnaround === "" || turnaround === undefined || isNaN(turnaround) || Number(turnaround) < 0) {
          const cDate = new Date(createdIso);
          const aDate = new Date(answeredIso);
          const diff = (aDate - cDate) / (1000 * 60 * 60);
          turnaround = diff > 0 ? diff.toFixed(1) : "0.5";
        } else {
          turnaround = Math.abs(Number(turnaround)).toFixed(1);
        }

        records.push({
          id: "A_" + (i + 1),
          rowIndex: i + 1,
          sheet: SHEET_ANSWERED,
          question: qText,
          eventName: String(row[1] || "").trim(),
          askedBy: String(row[2] || "").trim(),
          created: createdIso,
          answeredDate: answeredIso,
          turnaroundHours: turnaround,
          caseLink: String(row[7] || "").trim(),
          answer: String(row[8] || "").trim(),
          status: String(row[9] || "Answered").trim(),
          answeredBy: String(row[10] || "Supervisor").trim(),
          priority: String(row[5] || "Resolved").trim(),
          hoursElapsed: 0,
          // Per-user: reading it as an admin/supervisor never marks it read for the asker,
          // and vice versa - each person's read state is tracked independently.
          isRead: isReadByUser(row[11], requestingEmail)
        });
      }
    }
  } catch (err) {
    Logger.log("Error in getQuestionsData: " + err.toString());
  }

  return records;
}

function submitQuestion(payload) {
  ensureSheetsExist();
  requireTeamMember(payload.askedByEmail); // must be submitted by a recognized user

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_QUESTIONS);
  const now = new Date();

  sheet.appendRow([
    payload.question,
    payload.eventName,
    payload.askedBy,          // Full Name passed from client, shown in the UI
    now,
    "",
    payload.isUrgent ? "P1" : "",  // Forces top priority permanently when checked
    "",
    payload.caseLink,
    "",
    "Open",
    ""                         // Assigned To - starts unassigned
  ]);
  return { success: true };
}

function answerQuestion(rowIndex, answerText, supervisorEmail, supervisorName) {
  requireSupportRole(supervisorEmail); // only Support/Supervisor roles may answer

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);

  const rowData = qSheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
  const now = new Date();
  const createdDate = rowData[3] ? new Date(rowData[3]) : now;
  const turnaroundHours = Math.max(0.1, (now - createdDate) / (1000 * 60 * 60)).toFixed(1);

  aSheet.appendRow([
    rowData[0], // Question
    rowData[1], // Event Name
    rowData[2], // Asked By Name
    rowData[3], // Created
    now,        // Answered Date
    "Resolved", // Priority
    turnaroundHours, // Turnaround Hours
    rowData[7], // Case Link
    answerText, // Answer
    "Answered", // Status
    supervisorName, // Answered By Name
    ""          // Read By - empty, i.e. unread for EVERYONE including the asker,
                // until each person individually opens/marks it
  ]);

  qSheet.deleteRow(rowIndex);
  return { success: true };
}

// Assign an open question to another supervisor as a task. Any active
// Support/Supervisor-tier member can assign to any other Support/Supervisor-tier
// member; the ticket then shows up in that person's own Supervisor Desk queue.
function assignQuestion(rowIndex, assigneeEmail, requestingEmail) {
  requireSupportRole(requestingEmail);
  const assignee = requireSupportRole(assigneeEmail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (rowIndex < 2 || rowIndex > qSheet.getLastRow()) return { success: false };

  qSheet.getRange(rowIndex, 11).setValue(assignee.email);
  return { success: true };
}

// Assign several open questions to a supervisor at once (Supervisor Desk multi-select).
function assignQuestionsBulk(rowIndexes, assigneeEmail, requestingEmail) {
  requireSupportRole(requestingEmail);
  const assignee = requireSupportRole(assigneeEmail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const lastRow = qSheet.getLastRow();
  let updated = 0;

  (rowIndexes || []).forEach(raw => {
    const rowIndex = Number(raw);
    if (rowIndex >= 2 && rowIndex <= lastRow) {
      qSheet.getRange(rowIndex, 11).setValue(assignee.email);
      updated++;
    }
  });

  return { success: true, updated: updated };
}

function unassignQuestion(rowIndex, requestingEmail) {
  requireSupportRole(requestingEmail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (rowIndex < 2 || rowIndex > qSheet.getLastRow()) return { success: false };

  qSheet.getRange(rowIndex, 11).setValue("");
  return { success: true };
}

function toggleReadStatus(rowIndex, isRead, requestingEmail) {
  requireTeamMember(requestingEmail);
  const email = String(requestingEmail).trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (!aSheet || rowIndex < 2 || rowIndex > aSheet.getLastRow()) return { success: false };

  const cell = aSheet.getRange(rowIndex, 12);
  let list = parseReadByList(cell.getValue()).filter(e => e !== '*');
  if (isRead) {
    if (list.indexOf(email) === -1) list.push(email);
  } else {
    list = list.filter(e => e !== email);
  }
  cell.setValue(list.join(', '));
  return { success: true };
}

function saveTeamMember(member, requestingEmail) {
  ensureSheetsExist();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);

  const newCategory = String(member.category || '').trim();
  const existingCategory = member.rowIndex ? String(sheet.getRange(member.rowIndex, 4).getValue()).trim() : '';
  const touchesAdmin = newCategory === 'Admin' || existingCategory === 'Admin';

  if (touchesAdmin) {
    // Bootstrap exception: if there is no Admin at all yet, any Support/Supervisor
    // may create the very first one. Once an Admin exists, only an Admin can
    // create further Admins, edit an Admin's row, or rotate an Admin's password.
    if (countWorkingAdmins() > 0) {
      requireAdminRole(requestingEmail);
    } else {
      requireSupportRole(requestingEmail);
    }
  } else {
    requireSupportRole(requestingEmail); // only Support/Supervisor/Admin roles manage the roster
  }

  let passwordHash = '';
  let passwordSalt = '';
  if (member.rowIndex) {
    const existing = sheet.getRange(member.rowIndex, 1, 1, 7).getValues()[0];
    passwordHash = existing[5] || '';
    passwordSalt = existing[6] || '';
  }

  if (newCategory === 'Admin') {
    if (member.password) {
      // A new password was supplied - rotate it. Leaving the field blank on an edit keeps the existing password.
      passwordSalt = Utilities.getUuid();
      passwordHash = hashPassword(member.password, passwordSalt);
    }
    if (!member.rowIndex && !passwordHash) {
      throw new Error("A password is required when creating a new Admin account.");
    }
  } else {
    // Not an Admin (anymore) - clear any stored credential.
    passwordHash = '';
    passwordSalt = '';
  }

  const row = [member.name, member.title, member.status, member.category, member.email, passwordHash, passwordSalt];

  if (member.rowIndex) {
    sheet.getRange(member.rowIndex, 1, 1, 7).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

function deleteTeamMember(rowIndex, requestingEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  const targetCategory = (rowIndex >= 2 && rowIndex <= sheet.getLastRow()) ? String(sheet.getRange(rowIndex, 4).getValue()).trim() : '';

  if (targetCategory === 'Admin') {
    if (countWorkingAdmins() > 0) {
      requireAdminRole(requestingEmail); // only an Admin can remove another Admin
    } else {
      requireSupportRole(requestingEmail); // bootstrap: no working Admin exists, so Support can clean up a broken Admin row
    }
  } else {
    requireSupportRole(requestingEmail); // only Support/Supervisor/Admin roles manage the roster
  }

  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ==========================================
// ADMIN-ONLY: EDIT / DELETE QUESTIONS & ANSWERS
// These act directly on the connected Google Sheet - a delete here is
// permanent and removes the row from the underlying spreadsheet/database,
// not just from the app's view.
// ==========================================

function updateQuestion(rowIndex, updates, requestingEmail) {
  requireAdminRole(requestingEmail);
  updates = updates || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (rowIndex < 2 || rowIndex > qSheet.getLastRow()) return { success: false };

  const current = qSheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
  const merged = [
    updates.question !== undefined ? updates.question : current[0],
    updates.eventName !== undefined ? updates.eventName : current[1],
    updates.askedBy !== undefined ? updates.askedBy : current[2],
    current[3], // Created - immutable, preserves accurate priority/turnaround math
    current[4],
    updates.isUrgent !== undefined ? (updates.isUrgent ? "P1" : "") : current[5],
    current[6],
    updates.caseLink !== undefined ? updates.caseLink : current[7],
    current[8],
    current[9],
    current[10]
  ];
  qSheet.getRange(rowIndex, 1, 1, 11).setValues([merged]);
  return { success: true };
}

function deleteQuestion(rowIndex, requestingEmail) {
  requireAdminRole(requestingEmail);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (rowIndex < 2 || rowIndex > qSheet.getLastRow()) return { success: false };
  qSheet.deleteRow(rowIndex); // permanent - removes the row from the Questions Tracker sheet
  return { success: true };
}

function updateAnswer(rowIndex, updates, requestingEmail) {
  requireAdminRole(requestingEmail);
  updates = updates || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (rowIndex < 2 || rowIndex > aSheet.getLastRow()) return { success: false };

  const current = aSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
  const merged = [
    updates.question !== undefined ? updates.question : current[0],
    updates.eventName !== undefined ? updates.eventName : current[1],
    updates.askedBy !== undefined ? updates.askedBy : current[2],
    current[3],
    current[4],
    current[5],
    current[6],
    updates.caseLink !== undefined ? updates.caseLink : current[7],
    updates.answer !== undefined ? updates.answer : current[8],
    current[9],
    updates.answeredBy !== undefined ? updates.answeredBy : current[10],
    current[11] // Read By - untouched by an edit
  ];
  aSheet.getRange(rowIndex, 1, 1, 12).setValues([merged]);
  return { success: true };
}

function deleteAnswer(rowIndex, requestingEmail) {
  requireAdminRole(requestingEmail);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (rowIndex < 2 || rowIndex > aSheet.getLastRow()) return { success: false };
  aSheet.deleteRow(rowIndex); // permanent - removes the row from the Answered sheet
  return { success: true };
}
