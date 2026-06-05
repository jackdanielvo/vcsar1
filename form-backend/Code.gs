/**
 * VCSAR Team 1 — Application form backend (Google Apps Script)
 * Appends each submission to the Sheet and emails a formatted notification
 * with a link to the spreadsheet.
 *
 * SETUP: see APPLICATION-FORM-SETUP.txt in this folder.
 */

// Email address(es) to notify on each new application (comma-separated). Leave "" for none.
var NOTIFY_EMAIL = "officers@vcsar1.org, jack@voiceofjack.com";

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

// Friendly labels for the email.
var LABELS = {
  prefix: "Prefix", first_name: "First name", last_name: "Last name", email: "Email",
  street: "Street", apt: "Apt / Suite", city: "City", state: "State", zip: "ZIP", phone: "Phone",
  age: "Age", employer: "Employer (name & city)", employer_callouts: "Employer allows callouts",
  medical: "Medical training", fitness: "Fitness level (1–10)", wilderness: "Wilderness experience",
  other_experience: "Other experience", why_join: "Why they want to join", referral: "Referred by",
  signature: "Signature", app_date: "Application date"
};

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};

    // Honeypot: silently accept bot submissions without recording them.
    if (p.company) {
      return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

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
      var name = ((p.first_name || "") + " " + (p.last_name || "")).trim() || "New applicant";
      try {
        MailApp.sendEmail({
          to: NOTIFY_EMAIL,
          replyTo: p.email || NOTIFY_EMAIL,            // reply goes straight to the applicant
          subject: "New volunteer application — " + name,
          htmlBody: buildEmailHtml(p, ss.getUrl()),
          body: buildEmailText(p, ss.getUrl())          // plain-text fallback
        });
        console.log("Notification email sent to: " + NOTIFY_EMAIL);
      } catch (mailErr) {
        // Don't fail the whole request if email fails — but log why (see Executions log).
        console.error("EMAIL FAILED: " + mailErr + " | remaining MailApp quota: " + safeQuota_());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function esc_(v) {
  return String(v == null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function val_(p, key) {
  var v = p[key];
  return (v === undefined || v === null || String(v).trim() === "") ? "—" : esc_(v);
}

function buildEmailHtml(p, sheetUrl) {
  var ORANGE = "#ff5a1f", OLIVE = "#1f2416", INK = "#1b1b1b", MUTED = "#6c7160", LINE = "#e6e6dd";
  var fullName = [p.prefix, p.first_name, p.last_name].filter(function (x) { return x; }).join(" ") || "New applicant";

  function rows(keys) {
    return keys.map(function (k) {
      return '' +
        '<tr>' +
          '<td style="padding:10px 16px;border-bottom:1px solid ' + LINE + ';color:' + MUTED + ';font:13px Arial,sans-serif;width:40%;vertical-align:top">' + LABELS[k] + '</td>' +
          '<td style="padding:10px 16px;border-bottom:1px solid ' + LINE + ';color:' + INK + ';font:14px/1.5 Arial,sans-serif">' + val_(p, k) + '</td>' +
        '</tr>';
    }).join("");
  }
  function section(title, keys) {
    return '' +
      '<tr><td colspan="2" style="padding:22px 16px 6px;font:bold 12px Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;color:' + ORANGE + '">' + title + '</td></tr>' +
      rows(keys);
  }

  return '' +
  '<div style="background:#f3f4ee;padding:24px 0;font-family:Arial,sans-serif">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid ' + LINE + ';border-radius:8px;overflow:hidden">' +
      // Header
      '<tr><td style="background:' + OLIVE + ';padding:24px 24px">' +
        '<div style="color:#fff;font:bold 18px Arial,sans-serif;letter-spacing:.5px">VCSAR Team 1 <span style="color:' + ORANGE + '">· New Application</span></div>' +
        '<div style="color:#b9bda8;font:13px Arial,sans-serif;margin-top:4px">Ventura County Search &amp; Rescue · Fillmore</div>' +
      '</td></tr>' +
      // Applicant banner
      '<tr><td style="padding:20px 24px 4px">' +
        '<div style="font:bold 22px Arial,sans-serif;color:' + INK + '">' + esc_(fullName) + '</div>' +
        '<div style="font:13px Arial,sans-serif;color:' + MUTED + ';margin-top:4px">' +
          'Submitted ' + esc_(new Date().toLocaleString("en-US")) + '</div>' +
        (p.email ? '<div style="margin-top:6px"><a href="mailto:' + esc_(p.email) + '" style="color:' + ORANGE + ';font:14px Arial,sans-serif;text-decoration:none">' + esc_(p.email) + '</a>' +
          (p.phone ? '<span style="color:' + MUTED + '"> · ' + esc_(p.phone) + '</span>' : '') + '</div>' : '') +
      '</td></tr>' +
      // Detail tables
      '<tr><td style="padding:6px 8px 8px">' +
        '<table role="presentation" cellpadding="0" cellspacing="0" width="100%">' +
          section("Contact", ["email", "phone"]) +
          section("Address", ["street", "apt", "city", "state", "zip"]) +
          section("Eligibility & Work", ["age", "employer", "employer_callouts"]) +
          section("Experience", ["medical", "fitness", "wilderness", "other_experience", "why_join", "referral"]) +
          section("Signed", ["signature", "app_date"]) +
        '</table>' +
      '</td></tr>' +
      // CTA button to the spreadsheet
      '<tr><td style="padding:14px 24px 28px;text-align:center">' +
        '<a href="' + sheetUrl + '" style="display:inline-block;background:' + ORANGE + ';color:#150800;font:bold 14px Arial,sans-serif;text-decoration:none;padding:13px 26px;border-radius:4px;letter-spacing:.5px">Open the applications spreadsheet →</a>' +
        '<div style="font:12px Arial,sans-serif;color:' + MUTED + ';margin-top:14px">Reply to this email to respond directly to the applicant.</div>' +
      '</td></tr>' +
    '</table>' +
  '</div>';
}

function buildEmailText(p, sheetUrl) {
  var lines = ["New VCSAR Team 1 volunteer application", ""];
  FIELDS.forEach(function (k) {
    if (LABELS[k]) lines.push(LABELS[k] + ": " + (p[k] || "—"));
  });
  lines.push("", "Spreadsheet: " + sheetUrl);
  return lines.join("\n");
}

function safeQuota_() {
  try { return MailApp.getRemainingDailyQuota(); } catch (e) { return "unknown"; }
}

/**
 * Run this once from the editor (pick "testEmail" in the toolbar dropdown, then Run).
 * It forces Google to prompt for the "send email" permission and sends a test.
 * If the email arrives, sending works and real applications will email too.
 * If Run shows an error, that error is the reason emails aren't sending.
 */
function testEmail() {
  Logger.log("Remaining email quota: " + safeQuota_());
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: "VCSAR email test",
    htmlBody: "<p>If you received this, MailApp sending works. Quota left: " + safeQuota_() + "</p>"
  });
  Logger.log("Test email sent to: " + NOTIFY_EMAIL);
}

// Lets you open the web-app URL in a browser to confirm it's live.
function doGet() {
  return ContentService.createTextOutput("VCSAR application endpoint is running.");
}
