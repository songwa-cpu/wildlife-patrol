// ================================================================
//  WILDLIFE PATROL — Google Apps Script Backend
//  Paste this entire file into: script.google.com → New Project
//  Then: Deploy → New deployment → Web App
//  Execute as: Me | Who has access: Anyone
// ================================================================

// ── PASTE YOUR GOOGLE SHEET ID HERE ──────────────────────────────
//  Open your Google Sheet → look at the URL:
//  https://docs.google.com/spreadsheets/d/  ← SHEET_ID →  /edit
var SHEET_ID = "1Uknqx8ZLCCcp40VP_DWDE8H_DvumdgedoEDQ60ARm4Y";

// Column headers — matches your existing spreadsheet structure
var HEADERS = [
  "Record ID", "Date", "Time", "Patroller", "Location", "Patrol Type",
  "Category", "Species", "Count Seen", "Count Heard",
  "Latitude", "Longitude", "GPS Accuracy (m)",
  "Has Photo", "Has Audio", "Synced At"
];

// ================================================================
//  MAIN ENTRY POINT — handles POST from the patrol form
// ================================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Handle single record or batch array
    var records = Array.isArray(data) ? data : [data];
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var results = [];

    // Separate tracks from sightings
    var tracks = records.filter(function(r) { return r._type === "track"; });
    var sightings = records.filter(function(r) { return r._type !== "track"; });

    sightings.forEach(function(record) {
      var result = appendRecord(ss, record);
      results.push(result);
    });

    tracks.forEach(function(track) {
      var result = appendTrack(ss, track);
      results.push(result);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, count: results.length, results: results }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================================================================
//  APPEND ONE RECORD to the correct monthly sheet
// ================================================================
function appendRecord(ss, record) {
  var sheetName = getSheetName(record.date);
  var sheet = ss.getSheetByName(sheetName);

  // Create the sheet if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheetHeaders(sheet, sheetName);
  }

  // Check for duplicate (same record ID)
  var existing = findExisting(sheet, record.id);
  if (existing) {
    return { id: record.id, status: "duplicate_skipped" };
  }

  var row = [
    record.id,
    record.date,
    record.time,
    record.patroller || "",
    record.location || "",
    record.patrolType || "",
    record.category || "",
    record.animalName || record.species || "",
    record.countSeen || 0,
    record.countHeard || 0,
    record.lat || "",
    record.lng || "",
    record.gpsAcc ? Math.round(record.gpsAcc) : "",
    (record.photos && record.photos.length > 0) ? "Yes (" + record.photos.length + ")" : "No",
    record.audio ? "Yes" : "No",
    new Date().toISOString()
  ];

  sheet.appendRow(row);

  // Format the new row
  var lastRow = sheet.getLastRow();
  formatDataRow(sheet, lastRow, record.category);

  return { id: record.id, status: "saved", sheet: sheetName, row: lastRow };
}

// ================================================================
//  SHEET NAME — e.g. "2026 Jan-Mar", "2026 Apr-Jun"
// ================================================================
function getSheetName(dateStr) {
  if (!dateStr) return "Unassigned";
  var d = new Date(dateStr);
  var year = d.getFullYear();
  var month = d.getMonth(); // 0-indexed
  var quarters = ["Jan-Mar", "Apr-Jun", "Jul-Sep", "Oct-Dec"];
  var q = Math.floor(month / 3);
  return year + " " + quarters[q];
}

// ================================================================
//  SET UP HEADERS on a new sheet
// ================================================================
function setupSheetHeaders(sheet, sheetName) {
  // Title row
  sheet.getRange(1, 1, 1, HEADERS.length).merge();
  var titleCell = sheet.getRange(1, 1);
  titleCell.setValue("🌿 Wildlife Patrol Records — " + sheetName);
  titleCell.setBackground("#1b5e20");
  titleCell.setFontColor("#ffffff");
  titleCell.setFontWeight("bold");
  titleCell.setFontSize(13);
  titleCell.setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  // Header row
  var headerRange = sheet.getRange(2, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setBackground("#2e7d32");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  sheet.setRowHeight(2, 28);

  // Freeze top 2 rows
  sheet.setFrozenRows(2);

  // Column widths
  var widths = [160, 90, 60, 120, 100, 90, 130, 160, 80, 80, 100, 100, 90, 100, 80, 140];
  widths.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
}

// ================================================================
//  FORMAT a data row by category
// ================================================================
function formatDataRow(sheet, rowNum, category) {
  var colors = {
    "Rare & Endangered": "#fff8e1",
    "Other Mammals":     "#f1f8e9",
    "Poaching":          "#ffebee",
    "Lumber":            "#fff3e0",
    "Conflict":          "#e3f2fd"
  };
  var bg = colors[category] || "#ffffff";
  sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground(bg);
  sheet.setRowHeight(rowNum, 24);

  // Bold the species name (column 8)
  sheet.getRange(rowNum, 8).setFontWeight("bold");

  // Color the count cells if > 0
  var seenVal = sheet.getRange(rowNum, 9).getValue();
  var heardVal = sheet.getRange(rowNum, 10).getValue();
  if (seenVal > 0)  sheet.getRange(rowNum, 9).setBackground("#c8e6c9").setFontWeight("bold");
  if (heardVal > 0) sheet.getRange(rowNum, 10).setBackground("#b3e5fc").setFontWeight("bold");
}

// ================================================================
//  CHECK for duplicate record by ID
// ================================================================
function findExisting(sheet, recordId) {
  if (!recordId) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  var ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  return ids.some(function(r) { return r[0] === recordId; });
}

// ================================================================
//  APPEND PATROL TRACK to "Patrol Tracks" sheet
// ================================================================
var TRACK_HEADERS = [
  "Track ID", "Start Time", "End Time", "Duration (min)",
  "Points", "Distance (m)", "Start Lat", "Start Lng",
  "End Lat", "End Lng", "Synced At"
];

function appendTrack(ss, track) {
  var sheetName = "Patrol Tracks";
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupTrackHeaders(sheet);
  }

  // Check for duplicate
  var existing = findExistingTrack(sheet, track.id);
  if (existing) return { id: track.id, status: "duplicate_skipped" };

  var startDate = new Date(track.startTime);
  var endDate = new Date(track.endTime);
  var durationMin = Math.round((endDate - startDate) / 60000);

  var row = [
    track.id,
    track.startTime,
    track.endTime,
    durationMin,
    track.pointCount || 0,
    track.distance || 0,
    track.startLat || "",
    track.startLng || "",
    track.endLat || "",
    track.endLng || "",
    new Date().toISOString()
  ];

  sheet.appendRow(row);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, TRACK_HEADERS.length).setBackground("#e8f5e9");
  sheet.setRowHeight(lastRow, 24);

  // Bold the distance
  sheet.getRange(lastRow, 6).setFontWeight("bold");

  return { id: track.id, status: "saved", sheet: sheetName, row: lastRow };
}

function setupTrackHeaders(sheet) {
  sheet.getRange(1, 1, 1, TRACK_HEADERS.length).merge();
  var titleCell = sheet.getRange(1, 1);
  titleCell.setValue("🥾 Patrol Tracks");
  titleCell.setBackground("#0d47a1");
  titleCell.setFontColor("#ffffff");
  titleCell.setFontWeight("bold");
  titleCell.setFontSize(13);
  titleCell.setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  var headerRange = sheet.getRange(2, 1, 1, TRACK_HEADERS.length);
  headerRange.setValues([TRACK_HEADERS]);
  headerRange.setBackground("#1565c0");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);

  var widths = [180, 160, 160, 90, 70, 100, 100, 100, 100, 100, 160];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
}

function findExistingTrack(sheet, trackId) {
  if (!trackId) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  var ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  return ids.some(function(r) { return r[0] === trackId; });
}

// ================================================================
//  TEST FUNCTION — run this manually to verify setup
// ================================================================
function testSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("Sheet opened: " + ss.getName());
  Logger.log("Test OK — your Sheet ID is correct.");
}

// ================================================================
//  GET handler — health check + CORS support
// ================================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Patrol sync endpoint is live." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
//  DIRECT TEST — paste this URL in browser to confirm it's live:
//  https://script.google.com/macros/s/YOUR_ID/exec
//  You should see: {"status":"ok","message":"Patrol sync endpoint is live."}
// ================================================================
