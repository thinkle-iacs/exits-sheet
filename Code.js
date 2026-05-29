const CONFIG = {
  progressSheetName: "Account Suspension Progress",
  sourceHeaderSearchRows: 3,
  reminderLeadDays: 7,
  dailyTriggerHour: 6,
  adminNoticeEmails: ["thinkle@innovationcharter.org"],
  leavingAccountGuideUrl:
    "https://docs.google.com/document/d/1URfLLiZJ91qkv6P27pG7SYSp9hpNCrzJsHAZ_kP0ox0/edit?tab=t.0",
};

const SOURCE_FIELDS = {
  submitterEmail: "submitterEmail",
  departingEmail: "departingEmail",
  staffName: "staffName",
  expectedLastDay: "expectedLastDay",
  daysBeyondExit: "daysBeyondExit",
};

const SOURCE_HEADER_MATCHERS = {
  [SOURCE_FIELDS.submitterEmail]: [/^email\s*address$/i],
  [SOURCE_FIELDS.departingEmail]: [/^departing\s*email\s*address$/i],
  [SOURCE_FIELDS.staffName]: [/staff.*name/i],
  [SOURCE_FIELDS.expectedLastDay]: [/expected.*last.*day/i, /last.*day/i],
  [SOURCE_FIELDS.daysBeyondExit]: [/days.*beyond.*exit.*suspend/i],
};

const PROGRESS_HEADERS = [
  "Account",
  "Submitter Email",
  "Staff Name",
  "Expected Last Day",
  "Days Beyond Exit",
  "Initial Email Date",
  "Initial Email Sent",
  "Reminder Date",
  "Reminder Sent",
  "Suspension Date",
  "Groups Removed",
  "Account Suspended",
  "Current Status",
  "Last Checked",
  "Last Alert Sent",
  "Last Error",
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Exit Automation")
    .addItem("Set up triggers and progress sheet", "setupExitAutomation")
    .addItem("Sync existing form rows", "syncExistingExitRows")
    .addItem("Run daily check now", "dailyExitAccountCheck")
    .addToUi();
}

function setupExitAutomation() {
  ensureProgressSheet_();
  installTrigger_("handleExitFormSubmit", ScriptApp.EventType.ON_FORM_SUBMIT);
  installDailyTrigger_("dailyExitAccountCheck", CONFIG.dailyTriggerHour);
}

function handleExitFormSubmit(e) {
  if (!e || !e.range) return;
  ingestExitSourceRow_(e.range.getSheet(), e.range.getRow());
}

function syncExistingExitRows() {
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getSheets().forEach((sheet) => {
    if (sheet.getName() === CONFIG.progressSheetName) return;

    const headerInfo = getSourceHeaderInfo_(sheet);
    if (!headerInfo.headerRow) return;

    for (let row = headerInfo.headerRow + 1; row <= sheet.getLastRow(); row++) {
      ingestExitSourceRow_(sheet, row);
    }
  });
}

function dailyExitAccountCheck() {
  const progressSheet = ensureProgressSheet_();
  const headerMap = getProgressHeaderMap_(progressSheet);
  const today = startOfDay_(new Date());
  const lastRow = progressSheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    processProgressRow_(progressSheet, headerMap, row, today);
  }
}

function ingestExitSourceRow_(sourceSheet, row) {
  const headerInfo = getSourceHeaderInfo_(sourceSheet);
  if (!headerInfo.headerRow || row <= headerInfo.headerRow) return;

  const sourceValues = getSourceRowValues_(sourceSheet, headerInfo.headerMap, row);
  const account = normalizeEmail_(sourceValues[SOURCE_FIELDS.departingEmail]);
  if (!account) return;

  const progressSheet = ensureProgressSheet_();
  const progressHeaderMap = getProgressHeaderMap_(progressSheet);
  if (findProgressRowByAccount_(progressSheet, progressHeaderMap, account)) {
    return;
  }

  const expectedLastDay = asDate_(sourceValues[SOURCE_FIELDS.expectedLastDay]);
  const daysBeyondExit = getNonNegativeInteger_(
    sourceValues[SOURCE_FIELDS.daysBeyondExit]
  );
  const suspensionDate = expectedLastDay
    ? addDays_(expectedLastDay, daysBeyondExit)
    : "";
  const reminderDate = suspensionDate
    ? addDays_(suspensionDate, -CONFIG.reminderLeadDays)
    : "";

  appendProgressRow_(progressSheet, {
    "Account": account,
    "Submitter Email": normalizeEmail_(sourceValues[SOURCE_FIELDS.submitterEmail]),
    "Staff Name": sourceValues[SOURCE_FIELDS.staffName],
    "Expected Last Day": expectedLastDay || "",
    "Days Beyond Exit": daysBeyondExit,
    "Initial Email Date": "",
    "Initial Email Sent": false,
    "Reminder Date": reminderDate || "",
    "Reminder Sent": false,
    "Suspension Date": suspensionDate || "",
    "Groups Removed": false,
    "Account Suspended": false,
    "Current Status": "Pending initial email",
    "Last Checked": "",
    "Last Alert Sent": "",
    "Last Error": "",
  });
}

function processProgressRow_(sheet, headerMap, row, today) {
  const rowValues = getProgressRowValues_(sheet, headerMap, row);
  const account = normalizeEmail_(rowValues["Account"]);
  if (!account) return;

  try {
    const user = getUser_(account);
    const currentStatus = user.suspended ? "Suspended" : "Active";
    setProgressValue_(sheet, headerMap, row, "Current Status", currentStatus);
    setProgressValue_(sheet, headerMap, row, "Last Checked", new Date());
    setProgressValue_(sheet, headerMap, row, "Last Error", "");

    if (wasReenabled_(rowValues, user)) {
      notifyAccountReenabled_(sheet, headerMap, row, account);
      return;
    }

    const suspensionDate = asDate_(rowValues["Suspension Date"]);
    const reminderDate = asDate_(rowValues["Reminder Date"]);

    if (!asBoolean_(rowValues["Initial Email Sent"])) {
      sendInitialEmail_(account, rowValues);
      setProgressValue_(sheet, headerMap, row, "Initial Email Date", new Date());
      setProgressValue_(sheet, headerMap, row, "Initial Email Sent", true);
      setProgressValue_(sheet, headerMap, row, "Current Status", "Initial email sent");
      return;
    }

    if (
      reminderDate &&
      today >= startOfDay_(reminderDate) &&
      !asBoolean_(rowValues["Reminder Sent"])
    ) {
      sendReminderEmail_(account, rowValues);
      setProgressValue_(sheet, headerMap, row, "Reminder Sent", true);
      setProgressValue_(sheet, headerMap, row, "Current Status", "Reminder sent");
      return;
    }

    if (suspensionDate && today >= startOfDay_(suspensionDate)) {
      if (!asBoolean_(rowValues["Groups Removed"])) {
        removeUserFromAllDirectGroups_(account);
        setProgressValue_(sheet, headerMap, row, "Groups Removed", true);
      }

      if (user.suspended) {
        setProgressValue_(sheet, headerMap, row, "Account Suspended", true);
        setProgressValue_(sheet, headerMap, row, "Current Status", "Suspended");
      } else if (!asBoolean_(rowValues["Account Suspended"])) {
        suspendUser_(account);
        setProgressValue_(sheet, headerMap, row, "Account Suspended", true);
        setProgressValue_(sheet, headerMap, row, "Current Status", "Suspended");
      }
    }
  } catch (err) {
    setProgressValue_(sheet, headerMap, row, "Current Status", "Error");
    setProgressValue_(sheet, headerMap, row, "Last Checked", new Date());
    setProgressValue_(sheet, headerMap, row, "Last Error", getErrorMessage_(err));
  }
}

function sendInitialEmail_(account, rowValues) {
  const suspensionDate = formatDateForEmail_(rowValues["Suspension Date"]);
  MailApp.sendEmail({
    to: account,
    cc: getCc_(account, rowValues),
    subject: "Your IACS account suspension date",
    htmlBody: [
      `<p>Hello ${escapeHtml_(rowValues["Staff Name"] || "")},</p>`,
      `<p>Your IACS account is scheduled to be suspended on <strong>${suspensionDate}</strong>.</p>`,
      "<p>Before that date, please make sure you have saved anything you need from your IACS account and transferred ownership of any shared school materials that should remain available.</p>",
      `<p>Please review these account exit instructions: <a href="${CONFIG.leavingAccountGuideUrl}">Leaving your IACS account</a>.</p>`,
      "<p>If you believe this date is incorrect, please contact the IACS technology team.</p>",
      "<p>Thank you!</p>",
      "<p>The IACS Tech Team</p>"
    ].join(""),
  });
}

function sendReminderEmail_(account, rowValues) {
  const suspensionDate = formatDateForEmail_(rowValues["Suspension Date"]);
  MailApp.sendEmail({
    to: account,
    cc: getCc_(account, rowValues),
    subject: "Reminder: your IACS account will be suspended soon",
    htmlBody: [
      `<p>Hello ${escapeHtml_(rowValues["Staff Name"] || "")},</p>`,
      `<p>This is a reminder that your IACS account is scheduled to be suspended on <strong>${suspensionDate}</strong>.</p>`,
      `<p>Please complete any needed file transfers or account cleanup before that date. The account exit instructions are here: <a href="${CONFIG.leavingAccountGuideUrl}">Leaving your IACS account</a>.</p>`,
    ].join(""),
  });
}

function notifyAccountReenabled_(sheet, headerMap, row, account) {
  const lastAlertSent = asDate_(sheet.getRange(row, headerMap["Last Alert Sent"]).getValue());
  if (lastAlertSent && startOfDay_(lastAlertSent).getTime() === startOfDay_(new Date()).getTime()) {
    return;
  }

  const statusCell = sheet.getRange(row, headerMap["Current Status"]);
  const editUrl = buildCellUrl_(sheet, statusCell);
  const message =
    `${account} was previously marked suspended, but the Google account is active again.\n\n` +
    `Review the suspension row here:\n${editUrl}`;

  statusCell.setNote(
    `Account was active again as of ${formatDateTime_(new Date())}. ` +
    "If this reactivation is expected, update the suspension dates or status fields."
  );
  setProgressValue_(sheet, headerMap, row, "Current Status", "Re-enabled after suspension");
  setProgressValue_(sheet, headerMap, row, "Last Alert Sent", new Date());

  if (CONFIG.adminNoticeEmails.length) {
    MailApp.sendEmail({
      to: CONFIG.adminNoticeEmails.join(","),
      subject: `Exit automation notice: ${account} is active again`,
      body: message,
    });
  }
}

function removeUserFromAllDirectGroups_(account) {
  const groups = listDirectGroups_(account);
  groups.forEach((groupEmail) => {
    AdminDirectory.Members.remove(groupEmail, account);
  });
}

function listDirectGroups_(account) {
  const groups = [];
  let pageToken;
  do {
    const response = AdminDirectory.Groups.list({
      userKey: account,
      pageToken: pageToken,
    });
    const groupList = response.groups || [];
    groupList.forEach((group) => groups.push(group.email));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return groups;
}

function suspendUser_(account) {
  AdminDirectory.Users.update({ suspended: true }, account);
}

function getUser_(account) {
  return AdminDirectory.Users.get(account);
}

function ensureProgressSheet_() {
  const spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(CONFIG.progressSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.progressSheetName);
  }

  const existingHeaders = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PROGRESS_HEADERS.length))
    .getValues()[0];
  PROGRESS_HEADERS.forEach((header, index) => {
    if (existingHeaders[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
  sheet.setFrozenRows(1);
  return sheet;
}

function appendProgressRow_(sheet, valuesByHeader) {
  const row = PROGRESS_HEADERS.map((header) =>
    Object.prototype.hasOwnProperty.call(valuesByHeader, header)
      ? valuesByHeader[header]
      : ""
  );
  sheet.appendRow(row);
}

function getProgressHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((map, header, index) => {
    if (header) map[String(header).trim()] = index + 1;
    return map;
  }, {});
}

function getProgressRowValues_(sheet, headerMap, row) {
  const values = {};
  Object.keys(headerMap).forEach((header) => {
    values[header] = sheet.getRange(row, headerMap[header]).getValue();
  });
  return values;
}

function setProgressValue_(sheet, headerMap, row, header, value) {
  if (!headerMap[header]) return;
  sheet.getRange(row, headerMap[header]).setValue(value);
}

function findProgressRowByAccount_(progressSheet, headerMap, account) {
  const accountColumn = headerMap["Account"];
  if (!accountColumn || progressSheet.getLastRow() < 2) {
    return 0;
  }

  const accountValues = progressSheet
    .getRange(2, accountColumn, progressSheet.getLastRow() - 1, 1)
    .getValues();

  for (let index = 0; index < accountValues.length; index++) {
    if (normalizeEmail_(accountValues[index][0]) === account) {
      return index + 2;
    }
  }
  return 0;
}

function getSourceHeaderInfo_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { headerRow: 0, headerMap: {} };

  const rowCount = Math.min(CONFIG.sourceHeaderSearchRows, sheet.getLastRow());
  if (rowCount < 1) return { headerRow: 0, headerMap: {} };

  const rows = sheet.getRange(1, 1, rowCount, lastColumn).getValues();
  let best = { headerRow: 0, headerMap: {}, score: 0 };

  rows.forEach((headers, index) => {
    const headerMap = buildSourceHeaderMap_(headers);
    const score = Object.keys(headerMap).length;
    const hasRequiredHeaders =
      headerMap[SOURCE_FIELDS.departingEmail] &&
      headerMap[SOURCE_FIELDS.expectedLastDay];
    if (hasRequiredHeaders && score > best.score) {
      best = { headerRow: index + 1, headerMap, score };
    }
  });

  return { headerRow: best.headerRow, headerMap: best.headerMap };
}

function buildSourceHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const field = getSourceFieldForHeader_(header);
    if (!field) return;
    map[field] = index + 1;
  });
  return map;
}

function getSourceFieldForHeader_(header) {
  const text = String(header || "").trim();
  if (!text) return "";

  for (const field in SOURCE_HEADER_MATCHERS) {
    if (SOURCE_HEADER_MATCHERS[field].some((matcher) => matcher.test(text))) {
      return field;
    }
  }
  return "";
}

function getSourceRowValues_(sheet, headerMap, row) {
  const values = {};
  Object.keys(headerMap).forEach((field) => {
    values[field] = sheet.getRange(row, headerMap[field]).getValue();
  });
  return values;
}

function installTrigger_(handlerFunction, eventType) {
  const spreadsheet = SpreadsheetApp.getActive();
  const alreadyInstalled = ScriptApp.getProjectTriggers().some(
    (trigger) =>
      trigger.getHandlerFunction() === handlerFunction &&
      trigger.getEventType() === eventType
  );
  if (alreadyInstalled) return;

  if (eventType === ScriptApp.EventType.ON_FORM_SUBMIT) {
    ScriptApp.newTrigger(handlerFunction).forSpreadsheet(spreadsheet).onFormSubmit().create();
  }
}

function installDailyTrigger_(handlerFunction, hour) {
  const alreadyInstalled = ScriptApp.getProjectTriggers().some(
    (trigger) =>
      trigger.getHandlerFunction() === handlerFunction &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK
  );
  if (alreadyInstalled) return;

  ScriptApp.newTrigger(handlerFunction)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

function wasReenabled_(rowValues, user) {
  return asBoolean_(rowValues["Account Suspended"]) && !user.suspended;
}

function asBoolean_(value) {
  if (value === true) return true;
  const text = String(value || "").trim().toLowerCase();
  return text === "true" || text === "yes" || text === "y";
}

function asDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function getNonNegativeInteger_(value) {
  const parsed = Number(value);
  if (!isFinite(parsed) || parsed < 0) return 0;
  return Math.floor(parsed);
}

function addDays_(date, days) {
  const result = startOfDay_(date);
  result.setDate(result.getDate() + days);
  return result;
}

function startOfDay_(date) {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function getCc_(account, rowValues) {
  const submitterEmail = normalizeEmail_(rowValues["Submitter Email"]);
  return submitterEmail && submitterEmail !== account ? submitterEmail : "";
}

function formatDateForEmail_(value) {
  const date = asDate_(value);
  if (!date) return "the scheduled suspension date";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM d, yyyy");
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM d, yyyy h:mm a");
}

function buildCellUrl_(sheet, range) {
  const spreadsheet = sheet.getParent();
  return `${spreadsheet.getUrl()}#gid=${sheet.getSheetId()}&range=${range.getA1Notation()}`;
}

function getErrorMessage_(err) {
  return err && err.message ? err.message : String(err);
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
