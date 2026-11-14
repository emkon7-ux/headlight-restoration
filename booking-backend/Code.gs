// Clear Beam booking backend: a Google Apps Script web app attached to a Google Sheet.
// The website reads taken times from it (GET) and saves new bookings to it (POST).
// A script lock around each booking stops two people from taking the same slot.

var TZ = "America/Los_Angeles";
var SLOTS = ["08:00", "10:00", "12:00", "14:00"]; // 2-hour jobs, 8 AM to 4 PM

// Optional text alert: your number at your carrier's email-to-text address, e.g.
// "8054023870@vtext.com" (Verizon) or "8054023870@tmomail.net" (T-Mobile). Leave "" for email only.
var TEXT_ALERT_ADDRESS = "";

// Run once from the Apps Script editor: creates the two tabs and asks for permissions.
function setup() {
  var ss = SpreadsheetApp.getActive();

  var bookings = ss.getSheetByName("Bookings") || ss.insertSheet("Bookings");
  if (bookings.getLastRow() === 0) {
    bookings.appendRow(["Submitted", "Date", "Time", "Name", "Phone", "Address", "Car make", "Car model"]);
  }
  bookings.getRange("B:C").setNumberFormat("@");
  bookings.setFrozenRows(1);

  var blocked = ss.getSheetByName("Blocked") || ss.insertSheet("Blocked");
  if (blocked.getLastRow() === 0) {
    blocked.appendRow(["Date (2026-11-21)", "Start time (8:00, 10:00, 12:00 or 14:00, or blank for the whole day)", "Note"]);
  }
  blocked.getRange("A:B").setNumberFormat("@");
  blocked.setFrozenRows(1);

  var first = ss.getSheetByName("Sheet1");
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(first);

  GmailApp.getUserLabelByName("Bookings") || GmailApp.createLabel("Bookings"); // also asks for Gmail permission now

  var hasTrigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "flagPending"; });
  if (!hasTrigger) ScriptApp.newTrigger("flagPending").timeBased().everyMinutes(5).create();
}

function doGet() {
  return json({ taken: takenSlots() });
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var error = validate(d);
  if (error) return json({ ok: false, error: error });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (takenSlots().indexOf(d.date + " " + d.time) !== -1) {
      return json({ ok: false, error: "taken" });
    }
    SpreadsheetApp.getActive().getSheetByName("Bookings")
      .appendRow([new Date(), d.date, d.time, d.name, d.phone, d.address, d.make, d.model]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  notify(d); // outside the lock so the next customer isn't kept waiting on email
  return json({ ok: true });
}

// Every booked or blocked slot from today on, as "2026-11-21 09:00"
function takenSlots() {
  var ss = SpreadsheetApp.getActive();
  var today = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  var taken = [];

  rows(ss.getSheetByName("Bookings"), 2, 2).forEach(function (r) {
    var date = normDate(r[0]), time = normTime(r[1]);
    if (date >= today && time) taken.push(date + " " + time);
  });

  rows(ss.getSheetByName("Blocked"), 1, 2).forEach(function (r) {
    var date = normDate(r[0]);
    if (!date || date < today) return;
    if (String(r[1]).trim() === "") {
      SLOTS.forEach(function (t) { taken.push(date + " " + t); });
    } else {
      var time = normTime(r[1]);
      if (time) taken.push(date + " " + time);
    }
  });

  return taken;
}

function rows(sheet, firstCol, numCols) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, firstCol, sheet.getLastRow() - 1, numCols).getValues();
}

function validate(d) {
  var fields = ["date", "time", "name", "phone", "address", "make", "model"];
  for (var i = 0; i < fields.length; i++) {
    var v = d[fields[i]];
    if (typeof v !== "string" || !v.trim() || v.length > 200) return "missing " + fields[i];
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return "bad date";
  var day = new Date(d.date + "T12:00:00Z").getUTCDay();
  if (day !== 0 && day !== 6) return "weekends only";
  if (SLOTS.indexOf(d.time) === -1) return "bad time";

  var now = new Date();
  var today = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  var hour = Utilities.formatDate(now, TZ, "HH:mm");
  if (d.date < today || (d.date === today && d.time <= hour)) return "past";
  return "";
}

function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  return String(v).trim();
}

// Accepts "9:00", "09:00", "9 AM", "1:00 PM" or a sheet time value; returns "09:00"
function normTime(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "HH:mm");
  var m = String(v).trim().toUpperCase().match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?$/);
  if (!m) return "";
  var h = Number(m[1]);
  if (m[3] === "PM" && h < 12) h += 12;
  if (m[3] === "AM" && h === 12) h = 0;
  return (h < 10 ? "0" : "") + h + ":" + (m[2] || "00");
}

// Emails the owner (the Google account that owns this sheet) a booking alert, then marks it
// Important, stars it and files it under a "Bookings" label in Gmail. Optionally texts too.
// A failed alert never undoes a booking that's already saved.
function notify(d) {
  var when = whenText(d);
  var ref = Utilities.getUuid().slice(0, 6).toUpperCase();
  var subject = "IMPORTANT: New booking " + when + " (" + d.name + ") #" + ref;
  var rows = [["When", when], ["Name", d.name], ["Phone", d.phone], ["Address", d.address], ["Car", d.make + " " + d.model]];

  var text = rows.map(function (r) { return r[0] + ": " + r[1]; }).join("\n") +
    "\n\nText the customer to confirm. This is an automatic alert; replies go nowhere.";
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:520px">' +
    '<div style="background:#d64545;color:#fff;font-weight:bold;padding:10px 14px;border-radius:8px 8px 0 0">' +
    "&#9888; NEW BOOKING, ACTION NEEDED</div>" +
    '<table style="width:100%;border:1px solid #ddd;border-top:0;border-collapse:collapse">' +
    rows.map(function (r) {
      return '<tr><td style="padding:8px 14px;color:#666;width:90px">' + r[0] +
        '</td><td style="padding:8px 14px;font-weight:bold">' + esc(r[1]) + "</td></tr>";
    }).join("") +
    '</table><p style="color:#666;font-size:13px">Text the customer to confirm. ' +
    "This is an automatic alert from your website; replies go nowhere.</p></div>";

  try {
    GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, text, {
      htmlBody: html,
      name: "Clear Beam Bookings (no reply)",
      replyTo: "no-reply@clearbeam.invalid"
    });
    Utilities.sleep(2000); // give Gmail a moment to index it; flagPending catches it later otherwise
    flagPending();
  } catch (err) {
    console.error("Booking alert email failed: " + err); // see Apps Script > Executions
  }

  if (TEXT_ALERT_ADDRESS) {
    try {
      GmailApp.sendEmail(TEXT_ALERT_ADDRESS, "New booking",
        when + "\n" + d.name + " " + d.phone + "\n" + d.address + "\n" + d.make + " " + d.model);
    } catch (err) {
      console.error("Text alert failed: " + err);
    }
  }
}

// Stars, marks Important and labels any booking alert not flagged yet. Runs after each booking and
// every 5 minutes (the trigger setup() creates), since Gmail can take a moment to index a new email.
function flagPending() {
  var label = GmailApp.getUserLabelByName("Bookings") || GmailApp.createLabel("Bookings");
  GmailApp.search('subject:"IMPORTANT: New booking" -label:Bookings newer_than:7d', 0, 20).forEach(function (thread) {
    thread.markImportant();
    thread.addLabel(label);
    thread.moveToInbox();
    thread.markUnread();
    thread.getMessages().forEach(function (m) { m.star(); });
  });
}

// Run from the Apps Script editor to send yourself a sample alert. Unlike a real booking it shows
// any error on screen, and it saves nothing to the sheet.
function testAlert() {
  var to = Session.getEffectiveUser().getEmail();
  if (!to) throw new Error("Google didn't give this script your email address. Run setup first and allow access.");
  console.log("Sending a test alert to " + to);
  GmailApp.sendEmail(to, "IMPORTANT: New booking TEST (ignore)", "Test alert from your booking script.", {
    htmlBody: '<div style="background:#d64545;color:#fff;font-weight:bold;padding:10px 14px">&#9888; TEST BOOKING ALERT</div>' +
      "<p>If you can read this, booking alerts work.</p>",
    name: "Clear Beam Bookings (no reply)"
  });
  Utilities.sleep(3000);
  flagPending();
  console.log("Sent. Check your inbox (and the Bookings label) for \"IMPORTANT: New booking TEST\".");
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// "Saturday, November 21, 10 AM – 12 PM"
function whenText(d) {
  var labels = { "08:00": "8 – 10 AM", "10:00": "10 AM – 12 PM", "12:00": "12 – 2 PM", "14:00": "2 – 4 PM" };
  var day = Utilities.formatDate(new Date(d.date + "T12:00:00Z"), "UTC", "EEEE, MMMM d");
  return day + ", " + labels[d.time];
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
