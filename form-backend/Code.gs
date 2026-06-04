/**
 * VCSAR Team 1 — Application form backend (Google Apps Script)
 * Receives POSTs from the Join page form and appends each one as a row
 * in your Google Sheet. Also emails a notification (optional).
 *
 * SETUP: see APPLICATION-FORM-SETUP.txt in this folder.
 */

// (Optional) email address to notify on each new application. Leave "" for none.
var NOTIFY_EMAIL = "admin@vcsar1.org";

// Column order — must match the field "name" attributes in the form.
var FIELDS = [
  "prefix", "first_name", "last_name", "email",
  "street", "apt", "city", "state", "zip", "phone",
  "age", "employer", "employer_callouts",
  "medical", "fitness", "wilderness",
  "other_experience", "why_join", "referral",
  "signature", "app_date"
];

var HEADERS = [
  "Timestamp", "Prefix", "First Name", "Last Name", "Email",
  "Street", "Apt/Suite", "City", "State", "ZIP", "Phone",
  "Age", "Employer (name & city)", "Employer allows callouts",
  "Medical training", "Fitness (1-10)", "Wilderness experience",
  "Other experience", "Why join", "Referral",
  "Signature", "Application date"
];

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};

    // Honeypot: silently accept bot submissions without recording them.
    if (p.company) {
      return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // Write header row once.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var row = [new Date()];
    FIELDS.forEach(function (f) { row.push(p[f] || ""); });
    sheet.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: "New VCSAR application: " + (p.first_name || "") + " " + (p.last_name || ""),
        body: FIELDS.map(function (f) { return f + ": " + (p[f] || ""); }).join("\n")
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you open the web-app URL in a browser to confirm it's live.
function doGet() {
  return ContentService.createTextOutput("VCSAR application endpoint is running.");
}
