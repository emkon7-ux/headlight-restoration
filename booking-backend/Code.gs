// Clear Beam booking backend: a Google Apps Script web app attached to a Google Sheet.
// The website reads taken times from it (GET) and saves new bookings to it (POST).
// A script lock around each booking stops two people from taking the same slot.

var TZ = "America/Los_Angeles";

// Customer confirmation emails come from the Google account that owns this sheet. Google can only
// hide that address on Workspace accounts, so on a regular Gmail account confirmations are skipped
// to keep your address private. Set this to true if the sheet lives on a business Gmail you're
// happy for customers to see (e.g. a separate clearbeam...@gmail.com).
var SEND_FROM_THIS_ACCOUNT = false;
var SLOTS = ["08:00", "10:00", "12:00", "14:00"]; // 2-hour jobs, 8 AM to 4 PM

// Run once from the Apps Script editor: creates the two tabs and asks for permissions.
function setup() {
  var ss = SpreadsheetApp.getActive();

  var bookings = ss.getSheetByName("Bookings") || ss.insertSheet("Bookings");
  if (bookings.getLastRow() === 0) {
    bookings.appendRow(["Submitted", "Date", "Time", "Name", "Phone", "Address", "Car make", "Car model", "Email"]);
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

  MailApp.getRemainingDailyQuota(); // asks for email permission now, so booking emails work later
}

function doGet() {
  return json({ taken: takenSlots() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var error = validate(d);
    if (error) return json({ ok: false, error: error });

    if (takenSlots().indexOf(d.date + " " + d.time) !== -1) {
      return json({ ok: false, error: "taken" });
    }

    SpreadsheetApp.getActive().getSheetByName("Bookings")
      .appendRow([new Date(), d.date, d.time, d.name, d.phone, d.address, d.make, d.model, d.email || ""]);
    SpreadsheetApp.flush();
    var emailed = notify(d);
    return json({ ok: true, emailed: emailed });
  } finally {
    lock.releaseLock();
  }
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
  if (d.email && (typeof d.email !== "string" || d.email.length > 200 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email))) {
    return "bad email";
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

// Emails the owner about the booking, and the customer too if they gave an email.
// A failed email never undoes a booking that's already saved.
function notify(d) {
  var when = whenText(d);
  try {
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      "New booking: " + when,
      "When: " + when + "\nName: " + d.name + "\nPhone: " + d.phone + "\nEmail: " + (d.email || "none") +
      "\nAddress: " + d.address + "\nCar: " + d.make + " " + d.model
    );
  } catch (err) {}

  if (!d.email) return false;
  var message = {
    to: d.email,
    subject: "Your headlight restoration is booked: " + when,
    body: "Hi " + d.name + ",\n\nYou're booked for " + when + " at " + d.address + ", for your " +
      d.make + " " + d.model + ".\n\nI'll text you at " + d.phone + " to confirm.\n\nClear Beam Headlight Restoration",
    name: "Clear Beam Headlight Restoration"
  };
  if (!SEND_FROM_THIS_ACCOUNT) message.noReply = true; // Workspace only; throws on regular Gmail
  try {
    MailApp.sendEmail(message);
    return true;
  } catch (err) {
    return false;
  }
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
