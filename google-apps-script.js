// ═══════════════════════════════════════════════════════════
// RAIM — Google Apps Script (Waitlist + Agency + Email Sequence)
// ═══════════════════════════════════════════════════════════
//
// Handles both forms, saves to separate sheets, emails
// hello@theraim.com on every submission, and runs a 2-email
// scheduled sequence for waitlist signups.
//
// EMAIL SEQUENCE:
//   1. Welcome — sent immediately on signup (same content as announcement)
//   2. Announcement — broadcast to all existing registrants (22 May 2026, 11am Dubai)
//   3. Access Link — broadcast to all registrants (29 May 2026, 9am Dubai)
//
// SETUP:
//
// 1. Open your existing RAIM Google Sheet
// 2. Ensure the first tab is called "Waitlist"
// 3. Ensure a second tab called "Agency Enquiries" exists
//    - Row 1 headers: Timestamp | Name | Email | Phone | Agency | Website | Team Size
// 4. Waitlist tab columns F, G, H:
//    - F1: Welcome Sent (already populated) | G1: Announce Sent | H1: Access Sent
// 5. Go to Extensions > Apps Script
// 6. Replace ALL code with this file
// 7. Click Deploy > Manage deployments > Edit (pencil icon)
// 8. Set version to "New version" and click Deploy
//    (URL stays the same — no changes needed in HTML)
// 9. Run setupTriggers() once from the editor to schedule both emails.
//
// ═══════════════════════════════════════════════════════════

var NOTIFY_EMAIL = 'hello@theraim.com';

// Registration link for PS footers
var REGISTRATION_URL = 'https://www.theraim.com';

// Circle access link
var ACCESS_LINK = 'https://recruiter-ai-movement.circle.so/join?invitation_token=96fc8effaf42a969e842572670fe1cdad6e15fe3-c29c1df1-9ca6-4f79-abe4-478a656d80ac';

// Signature image hosted on theraim.com (no Drive dependency)
var SIGNATURE_URL = 'https://www.theraim.com/assets/email-signature.png';

// ── Column indices (0-based) for the Waitlist sheet ──────
var COL_TIMESTAMP      = 0; // A
var COL_NAME           = 1; // B
var COL_EMAIL          = 2; // C
var COL_PHONE          = 3; // D
var COL_COMPANY        = 4; // E
var COL_WELCOME_SENT   = 5; // F — original welcome email (already populated)
var COL_ANNOUNCE_SENT  = 6; // G — announcement broadcast (22 May)
var COL_ACCESS_SENT    = 7; // H — access link broadcast (29 May)


// ═══════════════════════════════════════════════════════════
// TRIGGER SETUP — Run once to schedule both emails
// ═══════════════════════════════════════════════════════════

function setupTriggers() {
  // Remove any existing time-based triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Email 1 (Announcement): 22 May 2026 at 11:00 AM Dubai (07:00 UTC)
  ScriptApp.newTrigger('sendAnnouncementBroadcast')
    .timeBased()
    .at(new Date('2026-05-22T07:00:00Z'))
    .create();

  // Email 2 (Access Link): 29 May 2026 at 9:00 AM Dubai (05:00 UTC)
  ScriptApp.newTrigger('sendAccessBroadcast')
    .timeBased()
    .at(new Date('2026-05-29T05:00:00Z'))
    .create();

  Logger.log('Triggers created. Announcement: 22 May 11am Dubai. Access: 29 May 9am Dubai.');
}


// ═══════════════════════════════════════════════════════════
// HTML EMAIL BUILDER
// ═══════════════════════════════════════════════════════════

function buildHtmlEmail(bodyHtml) {
  return '' +
    '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"></head>' +
    '<body style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:12pt;color:#1a1a1a;">' +
    bodyHtml +
    '<br>' +
    '<img src="' + SIGNATURE_URL + '" width="500" height="150" style="display:block;max-width:100%;height:auto;" alt="Guy Last - Founder, RAIM - Recruiter AI Movement">' +
    '</body></html>';
}

function line(text) {
  return text + '<br><br>';
}

function ps(text) {
  return '<span style="font-size:10pt;color:#888;">' + text + '</span>';
}

function sendHtmlEmail(to, subject, bodyHtml) {
  MailApp.sendEmail({
    to: to,
    subject: subject,
    htmlBody: buildHtmlEmail(bodyHtml),
    replyTo: NOTIFY_EMAIL,
    name: 'Guy Last'
  });
}


// ═══════════════════════════════════════════════════════════
// FORM HANDLERS
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter || {};
  }

  var type = data.type || 'waitlist';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === 'agency') {
    return handleAgency(ss, data);
  } else {
    return handleWaitlist(ss, data);
  }
}

function doGet(e) {
  var params = e.parameter || {};
  if (params.type || params.email) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var type = params.type || 'waitlist';
    if (type === 'agency') {
      return handleAgency(ss, params);
    } else {
      return handleWaitlist(ss, params);
    }
  }
  return ContentService.createTextOutput('RAIM endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Waitlist ──────────────────────────────────────────────

function handleWaitlist(ss, data) {
  var sheet = ss.getSheetByName('Waitlist') || ss.getSheets()[0];

  var name = data.name || '';
  var email = data.email || '';
  var phone = data.phone || '';
  var company = data.company || '';

  sheet.appendRow([
    new Date().toISOString(),
    name,
    email,
    phone,
    company,
    '', // F: Welcome Sent (legacy)
    '', // G: Announce Sent
    ''  // H: Access Sent
  ]);

  // Send welcome email to the registrant (same content as announcement)
  try {
    sendAnnouncementEmail(name, email);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, COL_WELCOME_SENT + 1).setValue(new Date().toISOString());
    sheet.getRange(lastRow, COL_ANNOUNCE_SENT + 1).setValue(new Date().toISOString());
  } catch (err) {
    // Don't fail the form submission if welcome email fails
  }

  // Notification to hello@theraim.com
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New RAIM Waitlist Signup: ' + (name || email),
      htmlBody:
        '<h2 style="color:#1a1a1a;font-family:sans-serif;">New Waitlist Signup</h2>' +
        '<table style="font-family:sans-serif;font-size:14px;border-collapse:collapse;">' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Name</td><td style="padding:6px 12px;">' + name + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Email</td><td style="padding:6px 12px;">' + email + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Phone</td><td style="padding:6px 12px;">' + phone + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Company</td><td style="padding:6px 12px;">' + company + '</td></tr>' +
        '</table>' +
        '<p style="font-size:12px;color:#888;margin-top:16px;">Source: theraim.com waitlist</p>'
    });
  } catch (err) {}

  return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Agency Enquiry ───────────────────────────────────────

function handleAgency(ss, data) {
  var sheet = ss.getSheetByName('Agency Enquiries');
  if (!sheet) {
    sheet = ss.insertSheet('Agency Enquiries');
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Agency', 'Website', 'Team Size']);
  }

  var name = data.name || '';
  var email = data.email || '';
  var phone = data.phone || '';
  var agency = data.agency || '';
  var website = data.website || '';
  var size = data.size || '';

  sheet.appendRow([
    new Date().toISOString(),
    name,
    email,
    phone,
    agency,
    website,
    size
  ]);

  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Agency Enquiry: ' + (agency || name),
      htmlBody:
        '<h2 style="color:#1a1a1a;font-family:sans-serif;">New Agency Enquiry</h2>' +
        '<table style="font-family:sans-serif;font-size:14px;border-collapse:collapse;">' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Name</td><td style="padding:6px 12px;">' + name + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Email</td><td style="padding:6px 12px;">' + email + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Phone</td><td style="padding:6px 12px;">' + phone + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Agency</td><td style="padding:6px 12px;">' + agency + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Website</td><td style="padding:6px 12px;">' + website + '</td></tr>' +
        '<tr><td style="padding:6px 12px;font-weight:bold;">Team Size</td><td style="padding:6px 12px;">' + size + '</td></tr>' +
        '</table>' +
        '<p style="font-size:12px;color:#888;margin-top:16px;">Source: theraim.com/agency</p>'
    });
  } catch (err) {}

  return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════
// EMAIL CONTENT
// ═══════════════════════════════════════════════════════════

// ── Email 1: Announcement / Welcome ─────────────────────
// Sent immediately to new signups AND broadcast to all
// existing registrants on 22 May at 11am Dubai.

function sendAnnouncementEmail(name, email) {
  var firstName = getFirstName(name);

  var subject = 'Next Friday';

  var body =
    line('Hey ' + firstName + ',') +
    line('Quick update.') +
    line('Next Friday (29th May) you\'ll receive an email from me with your personal access link to the Recruiter AI Movement platform. Everything opens that day.') +
    line('Here\'s what\'s waiting for you inside.') +
    line('The introduction and guidance to your AI operating system, full set up and ready to go. Recruitment AI skills you can run from day one. The RAIM community with other recruiters building alongside you. Two live office hours calls every week with me and Daniel, the founders. And a clear 30-day path designed to accelerate you into an AI-empowered recruiter.') +
    line('On Friday 29th May we\'re also hosting a live welcome call inside the platform.') +
    line('<b>3:00 PM Dubai / 12:00 PM UK / 6:00 AM Central US</b>') +
    line('We\'ll walk you through the platform, show you where everything lives, and demo one of the AI skills live so you can see exactly what this looks like in action. It\'s the fastest way to get oriented and start strong.') +
    line('That\'s all for now. Keep an eye on your inbox next Friday morning.') +
    line('See you inside.') +
    line('Guy') +
    ps('PS Know a recruiter who should be part of this? Send them to <a href="' + REGISTRATION_URL + '" style="color:#6E5845;">' + REGISTRATION_URL + '</a>');

  sendHtmlEmail(email, subject, body);
}

// ── Email 2: Access Link ────────────────────────────────
// Broadcast to all registrants on 29 May at 9am Dubai.

function sendAccessEmail(name, email) {
  var firstName = getFirstName(name);

  var subject = "You're in";

  var body =
    line('Hey ' + firstName + ',') +
    line('It\'s live. The Recruiter AI Movement platform is open.') +
    line('Here\'s your personal access link.') +
    line('<a href="' + ACCESS_LINK + '">' + ACCESS_LINK + '</a>') +
    line('Log in. Look around. Get set up.') +
    line('Today at <b>3:00 PM Dubai / 12:00 PM UK / 6:00 AM Central US</b> we\'re hosting a live welcome call inside the community. Daniel and I will walk you through the platform, show you how the AI skills work, and answer any questions you\'ve got. The event link is inside the platform once you log in.') +
    line('From today, you also have access to two live office hours calls every week with us. Bring questions, share wins, work through problems. This is how the community operates.') +
    line('Your first 30 days matter. The goal is simple. By the end of them, you are operating as an AI-empowered recruiter. The skills are there. The community is there. The founders are there. All you need to do is show up and start.') +
    line('Pick one task that takes too long on your desk right now. Sourcing a shortlist, screening CVs, writing a job ad. Run the skill for it today. That\'s your first move.') +
    line('Welcome to RAIM.') +
    line('Guy') +
    ps('PS Someone on your team needs to see this. Send them to <a href="' + REGISTRATION_URL + '" style="color:#6E5845;">' + REGISTRATION_URL + '</a>');

  sendHtmlEmail(email, subject, body);
}


// ═══════════════════════════════════════════════════════════
// BROADCAST FUNCTIONS (called by triggers)
// ═══════════════════════════════════════════════════════════

function sendAnnouncementBroadcast() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Waitlist') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var sent = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = row[COL_NAME] || '';
    var email = row[COL_EMAIL] || '';
    var alreadySent = row[COL_ANNOUNCE_SENT];

    if (!email || alreadySent) continue;

    try {
      sendAnnouncementEmail(name, email);
      sheet.getRange(i + 1, COL_ANNOUNCE_SENT + 1).setValue(new Date().toISOString());
      sent++;
      if (sent % 20 === 0) Utilities.sleep(1000);
    } catch (err) {
      Logger.log('Announcement failed for ' + email + ': ' + err.message);
    }
  }

  Logger.log('Announcement broadcast complete. Sent ' + sent + ' emails.');
}

function sendAccessBroadcast() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Waitlist') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var sent = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = row[COL_NAME] || '';
    var email = row[COL_EMAIL] || '';
    var alreadySent = row[COL_ACCESS_SENT];

    if (!email || alreadySent) continue;

    try {
      sendAccessEmail(name, email);
      sheet.getRange(i + 1, COL_ACCESS_SENT + 1).setValue(new Date().toISOString());
      sent++;
      if (sent % 20 === 0) Utilities.sleep(1000);
    } catch (err) {
      Logger.log('Access email failed for ' + email + ': ' + err.message);
    }
  }

  Logger.log('Access broadcast complete. Sent ' + sent + ' emails.');
}


// ═══════════════════════════════════════════════════════════
// TEST — Send test emails to yourself
// ═══════════════════════════════════════════════════════════

function testAnnouncementEmail() {
  sendAnnouncementEmail('Guy', 'hello@theraim.com');
  Logger.log('Test announcement email sent to hello@theraim.com');
}

function testAccessEmail() {
  sendAccessEmail('Guy Last', NOTIFY_EMAIL);
  Logger.log('Test access email sent to ' + NOTIFY_EMAIL);
}


// ═══════════════════════════════════════════════════════════
// ONE-TIME FIX — Backfill column G for manually sent emails
// ═══════════════════════════════════════════════════════════

function backfillAnnounceColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Waitlist') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var fixed = 0;

  for (var i = 1; i < data.length; i++) {
    var email = data[i][COL_EMAIL];
    var alreadySent = data[i][COL_ANNOUNCE_SENT];

    if (email && !alreadySent) {
      sheet.getRange(i + 1, COL_ANNOUNCE_SENT + 1).setValue('manual-' + new Date().toISOString());
      fixed++;
    }
  }

  Logger.log('Backfilled ' + fixed + ' rows in column G.');
}


// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function getFirstName(fullName) {
  if (!fullName) return 'there';
  var parts = fullName.trim().split(/\s+/);
  var first = parts[0];
  return first.charAt(0).toUpperCase() + first.slice(1);
}
