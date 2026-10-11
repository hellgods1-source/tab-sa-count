var COUNTS_NAME = "counts";
var STAFF_NAME = "staff";
var CENTERS_NAME = "centers";
var COUNT_HEADERS = ["date", "center", "sa", "accept", "decline", "callback", "missed", "lastAt"];
var STAFF_HEADERS = ["center", "name", "active"];
var CENTER_HEADERS = ["label"];
var DEFAULT_CENTERS = ["용인", "동탄"];
var DEFAULT_STAFF = [
  ["용인", "이영주", 1],
  ["용인", "이현섭", 1],
  ["용인", "나정주", 1],
  ["동탄", "이한빈", 1],
  ["동탄", "이상택", 1],
  ["동탄", "박찬병", 1],
  ["동탄", "우용건", 1],
  ["동탄", "변제철", 1]
];
var FIELDS = { accept: 4, decline: 5, callback: 6, missed: 7 };
var LAST_AT_COL = 8;

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  var action = String(p.action || "");
  try {
    if (action === "inc") applyDelta(p, 1);
    else if (action === "undo") applyDelta(p, -1);
    else if (action === "staff_save") staffSave(p);
    else if (action === "center_add") centerAdd(p);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function countsSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(COUNTS_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COUNTS_NAME);
    sheet.getRange(1, 1, 1, COUNT_HEADERS.length).setValues([COUNT_HEADERS]);
    sheet.getRange("A:C").setNumberFormat("@");
    sheet.getRange("H:H").setNumberFormat("0");
    return sheet;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var h0 = String(header[0] || "").toLowerCase();
  var h1 = String(header[1] || "").toLowerCase();
  if (h0 === "date" && h1 === "sa") {
    sheet.insertColumnAfter(1);
    sheet.getRange(1, 2).setValue("center");
    var last = sheet.getLastRow();
    if (last >= 2) sheet.getRange(2, 2, last - 1, 1).setValue("동탄");
    sheet.getRange(1, 1, 1, COUNT_HEADERS.length).setValues([COUNT_HEADERS]);
    sheet.getRange("A:C").setNumberFormat("@");
    sheet.getRange("H:H").setNumberFormat("0");
  } else if (h0 !== "date") {
    sheet.getRange(1, 1, 1, COUNT_HEADERS.length).setValues([COUNT_HEADERS]);
  }
  return sheet;
}

function staffSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(STAFF_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STAFF_NAME);
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]);
    sheet.getRange(2, 1, DEFAULT_STAFF.length, 3).setValues(DEFAULT_STAFF);
    sheet.getRange("A:B").setNumberFormat("@");
    sheet.getRange("C:C").setNumberFormat("0");
    return sheet;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]);
  }
  return sheet;
}

function centersSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CENTERS_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CENTERS_NAME);
    sheet.getRange(1, 1, 1, CENTER_HEADERS.length).setValues([CENTER_HEADERS]);
    for (var i = 0; i < DEFAULT_CENTERS.length; i++) {
      sheet.getRange(i + 2, 1).setValue(DEFAULT_CENTERS[i]);
    }
    sheet.getRange("A:A").setNumberFormat("@");
    return sheet;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, CENTER_HEADERS.length).setValues([CENTER_HEADERS]);
  }
  return sheet;
}

function applyDelta(p, delta) {
  var dateKey = String(p.date || "").trim();
  var center = String(p.center || "").trim();
  var sa = String(p.sa || "").trim();
  var field = String(p.field || "").trim();
  if (!dateKey || !center || !sa || !FIELDS[field]) {
    throw new Error("date, center, sa, field 필요");
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = countsSheet();
    var last = sheet.getLastRow();
    var row = 0;
    if (last >= 2) {
      var values = sheet.getRange(2, 1, last - 1, 3).getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        var rowCenter = String(values[i][1] || "").trim() || "동탄";
        if (String(values[i][0]) === dateKey && rowCenter === center && String(values[i][2]) === sa) {
          row = i + 2;
          break;
        }
      }
    }
    if (!row) {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, COUNT_HEADERS.length).setValues([[dateKey, center, sa, 0, 0, 0, 0, 0]]);
    }
    var col = FIELDS[field];
    var cur = Number(sheet.getRange(row, col).getValue()) || 0;
    var next = cur + delta;
    if (next < 0) next = 0;
    sheet.getRange(row, col).setValue(next);
    sheet.getRange(row, LAST_AT_COL).setValue(new Date().getTime());
  } finally {
    lock.releaseLock();
  }
}

function staffSave(p) {
  var center = String(p.center || "").trim();
  var name = String(p.name || "").trim();
  var active = String(p.active) === "0" ? 0 : 1;
  if (!center || !name) throw new Error("center, name 필요");
  centerAdd({ label: center });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = staffSheet();
    var last = sheet.getLastRow();
    var row = 0;
    if (last >= 2) {
      var values = sheet.getRange(2, 1, last - 1, 2).getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === center && String(values[i][1]).trim() === name) {
          row = i + 2;
          break;
        }
      }
    }
    if (!row) {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, STAFF_HEADERS.length).setValues([[center, name, active]]);
    } else {
      sheet.getRange(row, 3).setValue(active);
    }
  } finally {
    lock.releaseLock();
  }
}

function centerAdd(p) {
  var label = String(p.label || p.center || "").trim();
  if (!label) throw new Error("label 필요");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = centersSheet();
    var last = sheet.getLastRow();
    if (last >= 2) {
      var values = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === label) return;
      }
    }
    sheet.getRange(sheet.getLastRow() + 1, 1).setValue(label);
  } finally {
    lock.releaseLock();
  }
}
