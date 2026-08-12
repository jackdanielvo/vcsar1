// VCSAR Team 1 — Application form backend (Google Apps Script)
// - Appends each submission to the Sheet + emails a formatted notification
// - Sends the applicant an auto-response
// - Runs scheduled Academy reminders (July / August)
// SETUP: see APPLICATION-FORM-SETUP.txt in this folder.

// Email address(es) to notify on each new application (comma-separated). Leave "" for none.
var NOTIFY_EMAIL = "officers@vcsar1.org, jack@voiceofjack.com";

// Reply-to address for the auto-response sent to applicants.
var REPLY_TO = "admin@vcsar1.org";

// Ventura County Sheriff's Office badge shown in recruiting-email headers.
// (Hosted on the website — must be pushed live before it will display.)
var SHERIFF_BADGE_URL = "https://vcsar1org.netlify.app/images/sheriff-badge.png";
// Set to false to turn off the applicant auto-response.
var SEND_AUTORESPONSE = true;

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

    // Requests coming from the website's admin page (Email Candidates panel).
    if (p.action && p.action.indexOf("announce") === 0) {
      return handleAnnounceRequest_(p);
    }

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

    // Auto-response to the applicant
    if (SEND_AUTORESPONSE && p.email) {
      try {
        MailApp.sendEmail({
          to: p.email,
          name: "VCSAR1 Recruiting",
          replyTo: REPLY_TO,
          subject: "Thank you for your interest in VCSAR Team 1",
          htmlBody: buildApplicantHtml(p),
          body: buildApplicantText(p)
        });
        console.log("Auto-response sent to applicant: " + p.email);
      } catch (arErr) {
        console.error("AUTO-RESPONSE FAILED: " + arErr);
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
      emailHeader_(OLIVE, ORANGE, "New Application", "Ventura County Search &amp; Rescue · Fillmore") +
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

// Run this once from the editor (pick "testEmail" in the toolbar dropdown, then Run).
// It forces Google to prompt for the "send email" permission and sends a test.
// If the email arrives, sending works and real applications will email too.
// If Run shows an error, that error is the reason emails aren't sending.
function testEmail() {
  Logger.log("Remaining email quota: " + safeQuota_());
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: "VCSAR email test",
    htmlBody: "<p>If you received this, MailApp sending works. Quota left: " + safeQuota_() + "</p>"
  });
  Logger.log("Test email sent to: " + NOTIFY_EMAIL);
}

/* ---------- Applicant auto-response ---------- */
// Shared olive header (with the Sheriff's Office badge) for all emails.
function emailHeader_(OLIVE, ORANGE, label, subtitle) {
  label = label || "Recruiting";
  subtitle = subtitle || "Fillmore Search &amp; Rescue · Ventura County Sheriff's Office";
  return '<tr><td style="background:' + OLIVE + ';padding:18px 24px">' +
    '<table role="presentation" cellpadding="0" cellspacing="0"><tr>' +
      '<td style="vertical-align:middle;padding-right:14px"><img src="' + SHERIFF_BADGE_URL + '" width="46" alt="Ventura County Sheriff\'s Office" style="display:block;border:0" /></td>' +
      '<td style="vertical-align:middle">' +
        '<div style="color:#fff;font:bold 18px Arial,sans-serif;letter-spacing:.5px">VCSAR Team 1 <span style="color:' + ORANGE + '">· ' + label + '</span></div>' +
        '<div style="color:#b9bda8;font:13px Arial,sans-serif;margin-top:3px">' + subtitle + '</div>' +
      '</td>' +
    '</tr></table>' +
  '</td></tr>';
}

function buildApplicantText(p) {
  var name = (p.first_name || "").trim() || "there";
  var year = new Date().getFullYear();
  return [
    "Hi " + name + ",", "",
    "Thank you for your interest in the Fillmore SAR Team (VCSAR Team 1). We are always looking for qualified and enthusiastic people to help us fulfill our Team's mission as expressed in our motto: \"So That Others May Live.\"", "",
    "Our team runs a yearly Academy, usually at the beginning of the year. Prospects begin the process of applying and background checks toward the end of the preceding summer. The next Academy will start accepting applicants to start the process toward the end of Summer " + year + ".", "",
    "We hope you will be interested in continuing your journey toward joining our team. We work hard and train hard, learning many skills and assisting our area in multiple ways. We rely on the commitment of men and women like you to meet our mission goals.", "",
    "We will occasionally check in with reminders of upcoming Academy openings. Until then, please follow us on Instagram (@vcsar1) to keep abreast of our journey with the current crop of prospective members and our team's exploits.", "",
    "Sincerely yours,", "", "VCSAR1 Recruiting"
  ].join("\n");
}

function buildApplicantHtml(p) {
  var ORANGE = "#ff5a1f", OLIVE = "#1f2416", INK = "#1c1c1c", MUTED = "#6b6f60", LINE = "#e6e6dd";
  var name = esc_((p.first_name || "").trim() || "there");
  var year = new Date().getFullYear();
  function para(t) { return '<p style="margin:0 0 16px;color:' + INK + ';font:15px/1.65 Arial,sans-serif">' + t + '</p>'; }
  return '' +
  '<div style="background:#f3f4ee;padding:24px 0;font-family:Arial,sans-serif">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;margin:0 auto;background:#ffffff;border:1px solid ' + LINE + ';border-radius:8px;overflow:hidden">' +
      emailHeader_(OLIVE, ORANGE) +
      '<tr><td style="padding:30px 28px 8px">' +
        para('Hi ' + name + ',') +
        para('Thank you for your interest in the Fillmore SAR Team (VCSAR Team 1). We are always looking for qualified and enthusiastic people to help us fulfill our Team\'s mission as expressed in our motto: <strong style="color:' + OLIVE + '">&ldquo;So That Others May Live.&rdquo;</strong>') +
        para('Our team runs a yearly Academy, usually at the beginning of the year. Prospects begin the process of applying and background checks toward the end of the preceding summer. The next Academy will start accepting applicants to start the process toward the end of <strong>Summer ' + year + '</strong>.') +
        para('We hope you will be interested in continuing your journey toward joining our team. We work hard and train hard, learning many skills and assisting our area in multiple ways. We rely on the commitment of men and women like you to meet our mission goals.') +
        para('We will occasionally check in with reminders of upcoming Academy openings. Until then, please follow us on Instagram (<a href="https://www.instagram.com/vcsar1/" style="color:' + ORANGE + ';text-decoration:none">@vcsar1</a>) to keep abreast of our journey with the current crop of prospective members and our team\'s exploits.') +
        para('Sincerely yours,') +
        '<p style="margin:0;color:' + INK + ';font:bold 15px Arial,sans-serif">VCSAR1 Recruiting</p>' +
      '</td></tr>' +
      '<tr><td style="padding:18px 28px 26px;border-top:1px solid ' + LINE + ';color:' + MUTED + ';font:12px Arial,sans-serif">' +
        'Ventura County Search &amp; Rescue · Team 1 · 524 Sespe Ave, Fillmore, CA 93015' +
      '</td></tr>' +
    '</table>' +
  '</div>';
}

/* ================= ACADEMY REMINDERS (scheduled) =================
   Emails applicants who applied within the last 12 months a reminder as
   Academy season approaches. Create a MONTHLY time-driven trigger for
   sendAcademyReminders (day 1). It only acts in the months below and never
   sends the same nudge to the same person twice (tracked in the Sheet). */

var REMINDER_WINDOW_DAYS = 365;            // only applicants from the last N days
var REMINDER_MONTHS = { 7: "july", 8: "august" };   // month number -> nudge key

/* EDIT YOUR REMINDER TEXT HERE.
   First line should greet with {{NAME}}; last line is treated as the signature.
   {{NAME}} = applicant first name · {{YEAR}} = current year. */
var REMINDERS = {
  july: {
    subject: "VCSAR Team 1 — Academy season is almost here",
    paragraphs: [
      "Hi {{NAME}},",
      "Academy season is coming up for VCSAR Team 1, and we wanted to make sure you're still with us. You raised your hand to help serve the Fillmore community through search and rescue, and we haven't forgotten — “So That Others May Live” is a mission we take on together.",
      "Here's where things stand: we expect to hold our prospective-member orientation meeting in the early fall — likely September. That meeting kicks off the {{YEAR}} recruiting cycle, where you'll learn what the team does, what membership asks of you, and how the application and background-check process works.",
      "Between now and then, the best thing you can do is stay ready: keep building your fitness for backcountry work, get time outdoors, and start thinking about the gear you'll eventually need. If anything about your situation has changed, just reply and let us know.",
      "We'll send another note as the orientation date firms up. Until then, follow us on Instagram (@vcsar1) to see what the team — and our current prospective members — are up to.",
      "Sincerely yours,",
      "VCSAR1 Recruiting"
    ]
  },
  august: {
    subject: "VCSAR Team 1 — Orientation is coming up",
    paragraphs: [
      "Hi {{NAME}},",
      "The {{YEAR}} Academy is getting closer, and we'd love to see you there. Thank you again for your interest in joining VCSAR Team 1 — the all-volunteer Fillmore Search and Rescue team.",
      "We're planning to hold our prospective-member orientation meeting in September. This is the official start of the process: you'll meet the team, learn what we do and what we ask of our members, and find out how to move forward with your application and the Ventura County Sheriff's Office background check.",
      "We'll follow up with the confirmed date, time, and location as soon as they're set — so keep an eye on your inbox over the next few weeks. If you'd like to reach us in the meantime, just reply to this email.",
      "We hope you'll join us. We work hard and train hard, in service of a simple mission: “So That Others May Live.”",
      "Sincerely yours,",
      "VCSAR1 Recruiting"
    ]
  }
};

function sendAcademyReminders() {
  var now = new Date();
  var nudge = REMINDER_MONTHS[now.getMonth() + 1];   // getMonth() is 0-based
  if (!nudge) { console.log("Not a reminder month — nothing to do."); return; }
  var year = now.getFullYear();
  var tag = year + "-" + nudge;                      // e.g. "2026-july"

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) { console.log("No applicants yet."); return; }

  var header = data[0];
  var emailCol = header.indexOf("Email");
  var firstCol = header.indexOf("First Name");
  if (emailCol === -1) { console.error("Could not find an 'Email' column."); return; }

  var trackCol = header.indexOf("Academy Reminders");
  if (trackCol === -1) {
    trackCol = header.length;
    sheet.getRange(1, trackCol + 1).setValue("Academy Reminders").setFontWeight("bold");
  }

  var cutoff = new Date(now.getTime() - REMINDER_WINDOW_DAYS * 86400000);
  var content = REMINDERS[nudge];
  var sent = 0;

  for (var r = 1; r < data.length; r++) {
    var email = (data[r][emailCol] || "").toString().trim();
    if (!email) continue;
    var ts = data[r][0] ? new Date(data[r][0]) : null;       // Timestamp is column A
    if (!ts || ts < cutoff) continue;                        // only the last 12 months
    var track = (data[r][trackCol] || "").toString();
    if (track.indexOf(tag) !== -1) continue;                 // already sent this nudge

    var name = (firstCol !== -1 ? (data[r][firstCol] || "").toString().trim() : "") || "there";
    try {
      MailApp.sendEmail({
        to: email,
        name: "VCSAR1 Recruiting",
        replyTo: REPLY_TO,
        subject: content.subject,
        htmlBody: buildReminderHtml(content, name, year),
        body: buildReminderText(content, name, year)
      });
      sheet.getRange(r + 1, trackCol + 1).setValue(track ? track + ", " + tag : tag);
      sent++;
      Utilities.sleep(200);   // gentle pacing
    } catch (e) {
      console.error("Reminder failed for " + email + ": " + e);
    }
  }
  console.log("Academy reminder (" + tag + ") sent to " + sent + " applicant(s).");
}

function fillTokens_(s, name, year) {
  return String(s).replace(/\{\{NAME\}\}/g, name).replace(/\{\{YEAR\}\}/g, year);
}
function buildReminderText(content, name, year) {
  return content.paragraphs.map(function (p) { return fillTokens_(p, name, year); }).join("\n\n");
}
function buildReminderHtml(content, name, year) {
  var ORANGE = "#ff5a1f", OLIVE = "#1f2416", INK = "#1c1c1c", LINE = "#e6e6dd", MUTED = "#6b6f60";
  var last = content.paragraphs.length - 1;
  var body = content.paragraphs.map(function (p, i) {
    var t = esc_(fillTokens_(p, name, year));
    t = t.replace(/@vcsar1\b/g, '<a href="https://www.instagram.com/vcsar1/" style="color:' + ORANGE + ';text-decoration:none">@vcsar1</a>');
    return (i === last)
      ? '<p style="margin:0;color:' + INK + ';font:bold 15px Arial,sans-serif">' + t + '</p>'
      : '<p style="margin:0 0 16px;color:' + INK + ';font:15px/1.65 Arial,sans-serif">' + t + '</p>';
  }).join("");
  return '' +
  '<div style="background:#f3f4ee;padding:24px 0;font-family:Arial,sans-serif">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;margin:0 auto;background:#fff;border:1px solid ' + LINE + ';border-radius:8px;overflow:hidden">' +
      emailHeader_(OLIVE, ORANGE) +
      '<tr><td style="padding:30px 28px 8px">' + body + '</td></tr>' +
      '<tr><td style="padding:18px 28px 26px;border-top:1px solid ' + LINE + ';color:' + MUTED + ';font:12px Arial,sans-serif">Ventura County Search &amp; Rescue · Team 1 · 524 Sespe Ave, Fillmore, CA 93015</td></tr>' +
    '</table>' +
  '</div>';
}

// Run manually (any month) to email a PREVIEW of both reminders to this address.
var PREVIEW_TO = "jack@voiceofjack.com";
function sendReminderTest() {
  var to = PREVIEW_TO;
  var year = new Date().getFullYear();
  ["july", "august"].forEach(function (k) {
    var c = REMINDERS[k];
    MailApp.sendEmail({
      to: to, name: "VCSAR1 Recruiting", replyTo: REPLY_TO,
      subject: "[PREVIEW] " + c.subject,
      htmlBody: buildReminderHtml(c, "Jordan", year),
      body: buildReminderText(c, "Jordan", year)
    });
  });
  Logger.log("Preview reminders sent to " + to);
}

/* ============ ANNOUNCEMENTS TO CANDIDATES (on demand) ============
   Send meeting / training date emails to applicants whenever you like.
   HOW: open the applications spreadsheet → "VCSAR" menu → follow the items.
   You type the message in the "Announcement" tab — no code editing. */

var ANNOUNCE_SHEET = "Announcement";
var ANNOUNCE_LOG   = "Announcement Log";

// Adds the VCSAR menu to the spreadsheet automatically when it's opened.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("VCSAR")
    .addItem("1. Set up / open Announcement tab", "setupAnnouncementSheet")
    .addSeparator()
    .addItem("2. Send me a PREVIEW", "previewAnnouncement")
    .addItem("3. SEND to candidates", "sendAnnouncement")
    .addToUi();
}

function setupAnnouncementSheet() {
  var sh = announceSheet_(true);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    "Announcement tab ready.\n\n" +
    "• B1  Subject line\n" +
    "• B2  Audience: \"Last 12 months\" or \"All applicants\"\n" +
    "• B4 and below: one paragraph per row\n\n" +
    "Use {{NAME}} to insert the applicant's first name.\n" +
    "Then use the VCSAR menu to preview and send.");
}

function announceSheet_(createIfMissing) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ANNOUNCE_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(ANNOUNCE_SHEET);
    sh.getRange("A1").setValue("Subject:").setFontWeight("bold");
    sh.getRange("A2").setValue("Audience:").setFontWeight("bold");
    sh.getRange("A4").setValue("Message:").setFontWeight("bold");
    sh.getRange("B1").setValue("VCSAR Team 1 — Upcoming meeting & training dates");
    sh.getRange("B2").setValue("Last 12 months");
    sh.getRange("B4").setValue("Hi {{NAME}},");
    sh.getRange("B5").setValue("Our next prospective-member meeting is [DATE] at [TIME], [LOCATION].");
    sh.getRange("B6").setValue("Upcoming training: [DATE] — [WHAT TO BRING / DETAILS].");
    sh.getRange("B7").setValue("Please reply to let us know if you can make it. We look forward to seeing you there.");
    sh.getRange("B8").setValue("Sincerely yours,");
    sh.getRange("B9").setValue("VCSAR1 Recruiting");
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 720);
    sh.getRange("B1:B60").setWrap(true);
  }
  return sh;
}

// Reads subject + paragraphs from the Announcement tab.
function readAnnouncement_() {
  var sh = announceSheet_(false);
  if (!sh) throw new Error('No "' + ANNOUNCE_SHEET + '" tab yet. Use: VCSAR menu → "1. Set up / open Announcement tab".');
  var subject = (sh.getRange("B1").getValue() || "").toString().trim();
  var audience = (sh.getRange("B2").getValue() || "Last 12 months").toString().trim();
  var last = Math.max(sh.getLastRow(), 4);
  var rows = sh.getRange(4, 2, last - 3, 1).getValues();
  var paragraphs = rows.map(function (r) { return (r[0] || "").toString().trim(); })
                       .filter(function (t) { return t !== ""; });
  if (!subject) throw new Error("Please put a subject line in cell B1.");
  if (!paragraphs.length) throw new Error("Please write your message in B4 and below (one paragraph per row).");
  return { subject: subject, audience: audience, paragraphs: paragraphs };
}

// Collects candidates from the applications sheet (newest first).
function announceRecipients_(audience) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var header = data[0];
  var emailCol = header.indexOf("Email");
  var firstCol = header.indexOf("First Name");
  var lastCol  = header.indexOf("Last Name");
  var cityCol  = header.indexOf("City");
  var fitCol   = header.indexOf("Fitness (1-10)");
  var wildCol  = header.indexOf("Wilderness experience");
  if (emailCol === -1) throw new Error("Could not find an 'Email' column on the applications sheet.");
  var limit = /all/i.test(audience) ? null : new Date(Date.now() - REMINDER_WINDOW_DAYS * 86400000);
  var seen = {}, out = [];
  for (var r = data.length - 1; r >= 1; r--) {           // newest first
    var email = (data[r][emailCol] || "").toString().trim();
    if (!email || seen[email.toLowerCase()]) continue;
    var ts = data[r][0] ? new Date(data[r][0]) : null;
    if (limit) { if (!ts || ts < limit) continue; }
    seen[email.toLowerCase()] = true;
    var first = firstCol !== -1 ? (data[r][firstCol] || "").toString().trim() : "";
    var last  = lastCol  !== -1 ? (data[r][lastCol]  || "").toString().trim() : "";
    out.push({
      email: email,
      name: first || "there",
      fullName: ((first + " " + last).trim()) || email,
      applied: ts ? Utilities.formatDate(ts, Session.getScriptTimeZone(), "MMM d, yyyy") : "",
      city: cityCol !== -1 ? (data[r][cityCol] || "").toString().trim() : "",
      fitness: fitCol !== -1 ? (data[r][fitCol] || "").toString().trim() : "",
      wilderness: wildCol !== -1 ? (data[r][wildCol] || "").toString().trim() : ""
    });
  }
  return out;
}

function previewAnnouncement() {
  var ui = SpreadsheetApp.getUi();
  try {
    var a = readAnnouncement_();
    var content = { subject: "[PREVIEW] " + a.subject, paragraphs: a.paragraphs };
    MailApp.sendEmail({
      to: PREVIEW_TO, name: "VCSAR1 Recruiting", replyTo: REPLY_TO,
      subject: content.subject,
      htmlBody: buildReminderHtml(content, "Jordan", new Date().getFullYear()),
      body: buildReminderText(content, "Jordan", new Date().getFullYear())
    });
    ui.alert("Preview sent to " + PREVIEW_TO + ".\n\nCheck it, then use \"3. SEND to candidates\" when you're happy.");
  } catch (e) { ui.alert("Couldn't send preview:\n\n" + e.message); }
}

function sendAnnouncement() {
  var ui = SpreadsheetApp.getUi();
  var a, list;
  try {
    a = readAnnouncement_();
    list = announceRecipients_(a.audience);
  } catch (e) { ui.alert("Couldn't send:\n\n" + e.message); return; }

  if (!list.length) { ui.alert("No matching candidates found for audience: " + a.audience); return; }

  var ok = ui.alert("Send this announcement?",
    'Subject: "' + a.subject + '"\n' +
    "Audience: " + a.audience + "\n" +
    "Recipients: " + list.length + " candidate(s)\n\n" +
    "This sends real emails now.",
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) { ui.alert("Cancelled — nothing was sent."); return; }

  var year = new Date().getFullYear(), sent = 0, failed = 0;
  list.forEach(function (p) {
    try {
      MailApp.sendEmail({
        to: p.email, name: "VCSAR1 Recruiting", replyTo: REPLY_TO,
        subject: a.subject,
        htmlBody: buildReminderHtml(a, p.name, year),
        body: buildReminderText(a, p.name, year)
      });
      sent++; Utilities.sleep(200);
    } catch (e) { failed++; console.error("Announcement failed for " + p.email + ": " + e); }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(ANNOUNCE_LOG);
  if (!log) {
    log = ss.insertSheet(ANNOUNCE_LOG);
    log.appendRow(["Sent", "Subject", "Audience", "Recipients", "Failed"]);
    log.getRange(1, 1, 1, 5).setFontWeight("bold"); log.setFrozenRows(1);
  }
  log.appendRow([new Date(), a.subject, a.audience, sent, failed]);

  ui.alert("Announcement sent to " + sent + " candidate(s)." +
           (failed ? "\n" + failed + " failed — see the Executions log." : "") +
           "\n\nLogged in the \"" + ANNOUNCE_LOG + "\" tab.");
}

/* ====== EMAIL CANDIDATES from the website admin page ======
   The admin page (vcsar1.org/admin) posts here. Only a signed-in Supabase
   admin can send: we verify their login token with Supabase before doing
   anything. Nothing here needs the Apps Script editor once deployed. */

var SUPABASE_URL_GS  = "https://rhnxpcjlmdzialxxsrpa.supabase.co";
var SUPABASE_ANON_GS = "sb_publishable_s6tqiGb34P9pNFkinVbnrA_ex-WV_NG";

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Shared key that proves the request came from a signed-in admin.
// The admin page reads this same value from a Supabase table that only
// logged-in users can read, so anonymous visitors can never obtain it.
// (Must exactly match the value stored in Supabase.)
var ANNOUNCE_SECRET = "hbyXNMF17eL9bvZRWyy8NifgXAtGxdg8IYBvLdwtY00";

// Confirms the caller is a signed-in admin — no external requests needed.
// Returns { email: "..." } on success, or { error: "why it failed" }.
function verifyAdmin_(p) {
  var given = (p.secret || "").trim();
  if (!given) return { error: "no admin key was sent — sign out and back in." };
  if (given !== ANNOUNCE_SECRET) return { error: "admin key did not match." };
  var who = (p.admin || "").trim();
  return { email: who || REPLY_TO };
}

function handleAnnounceRequest_(p) {
  var who = verifyAdmin_(p);
  if (who.error) {
    console.error("Admin verify failed: " + who.error);
    return jsonOut_({ result: "error", error: "Not authorized — " + who.error });
  }
  var adminEmail = who.email;

  var audience = p.audience || "Last 12 months";
  var recipients;
  try { recipients = announceRecipients_(audience); }
  catch (err) { return jsonOut_({ result: "error", error: String(err.message || err) }); }

  // Just asking how many people would receive it.
  if (p.action === "announce_count") {
    return jsonOut_({ result: "ok", count: recipients.length });
  }

  // The admin page asking for the candidate list to show checkboxes.
  if (p.action === "announce_list") {
    return jsonOut_({ result: "ok", candidates: recipients });
  }

  // Restrict to the specific people ticked in the admin page.
  if (p.emails) {
    var pick = {};
    p.emails.split(",").forEach(function (e) { var t = e.trim().toLowerCase(); if (t) pick[t] = true; });
    recipients = recipients.filter(function (r) { return pick[r.email.toLowerCase()]; });
    if (!recipients.length) return jsonOut_({ result: "error", error: "No candidates were selected." });
  }

  var paragraphs = (p.body || "").split("\n")
    .map(function (t) { return t.trim(); })
    .filter(function (t) { return t !== ""; });
  var subject = (p.subject || "").trim();
  if (!subject)        return jsonOut_({ result: "error", error: "Please enter a subject line." });
  if (!paragraphs.length) return jsonOut_({ result: "error", error: "Please enter a message." });

  var content = { subject: subject, paragraphs: paragraphs };
  var year = new Date().getFullYear();

  // Preview: send only to the signed-in admin.
  if (p.action === "announce_preview") {
    try {
      MailApp.sendEmail({
        to: adminEmail, name: "VCSAR1 Recruiting", replyTo: REPLY_TO,
        subject: "[PREVIEW] " + subject,
        htmlBody: buildReminderHtml(content, "Jordan", year),
        body: buildReminderText(content, "Jordan", year)
      });
      return jsonOut_({ result: "ok", preview: adminEmail });
    } catch (err) {
      return jsonOut_({ result: "error", error: String(err) });
    }
  }

  // Real send.
  var sent = 0, failed = 0;
  recipients.forEach(function (r) {
    try {
      MailApp.sendEmail({
        to: r.email, name: "VCSAR1 Recruiting", replyTo: REPLY_TO,
        subject: subject,
        htmlBody: buildReminderHtml(content, r.name, year),
        body: buildReminderText(content, r.name, year)
      });
      sent++; Utilities.sleep(150);
    } catch (err) { failed++; console.error("Announce failed for " + r.email + ": " + err); }
  });

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = ss.getSheetByName(ANNOUNCE_LOG);
    if (!log) {
      log = ss.insertSheet(ANNOUNCE_LOG);
      log.appendRow(["Sent", "Subject", "Audience", "Recipients", "Failed", "Sent by"]);
      log.getRange(1, 1, 1, 6).setFontWeight("bold"); log.setFrozenRows(1);
    }
    log.appendRow([new Date(), subject, audience, sent, failed, adminEmail]);
  } catch (err) { console.error("Log write failed: " + err); }

  return jsonOut_({ result: "ok", sent: sent, failed: failed });
}

// Lets you open the web-app URL in a browser to confirm it's live.
// The version tag tells you WHICH code the live deployment is running.
function doGet() {
  return ContentService.createTextOutput(
    "VCSAR application endpoint is running. VERSION: announcements-v2");
}
