// ============================================================
// ShiftTrack — Google Apps Script Backend v2
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Open script.google.com → New project → name it "ShiftTrack"
// 2. Paste this entire file into Code.gs
// 3. Run setupSheets() once from the Run menu to create all sheets
// 4. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL
// 6. In ShiftTrack HTML app → Admin → Export → Google Sheets sync
//    paste the URL and click Test connection
// ============================================================

// ── SHEET NAMES ─────────────────────────────────────────────
var SHEET_STAFF   = 'Staff';
var SHEET_EVENTS  = 'ClockEvents';
var SHEET_SHIFTS  = 'ShiftSummary';
var SHEET_AUDIT   = 'AuditLog';

// ── CORS HEADERS ────────────────────────────────────────────
function setCorsHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ENTRY POINT: GET ────────────────────────────────────────
// Used for connection test and data fetching
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action ? e.parameter.action : 'ping';

    if (action === 'ping') {
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', message: 'ShiftTrack API running', timestamp: new Date().toISOString() })
      ));
    }

    if (action === 'getStaff') {
      var staff = getAllStaff();
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok', data: staff })));
    }

    if (action === 'getEvents') {
      var events = getAllEvents();
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok', data: events })));
    }

    if (action === 'getShifts') {
      var shifts = getAllShifts();
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok', data: shifts })));
    }

    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'Unknown action: ' + action })
    ));

  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ));
  }
}

// ── ENTRY POINT: POST ───────────────────────────────────────
// Receives clock events from the HTML app
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || 'saveEvent';

    if (action === 'saveEvent') {
      var result = saveClockEvent(payload);
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify(result)));
    }

    if (action === 'saveBatch') {
      // Sync multiple queued events at once
      var results = [];
      (payload.events || []).forEach(function(ev) {
        results.push(saveClockEvent(ev));
      });
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', saved: results.length })
      ));
    }

    if (action === 'updateStaff') {
      updateStaffRecord(payload);
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok' })));
    }

    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'Unknown action' })
    ));

  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ));
  }
}

// ── SAVE CLOCK EVENT ────────────────────────────────────────
function saveClockEvent(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Log raw event to ClockEvents
  var evSheet = getOrCreateSheet(SHEET_EVENTS,
    ['Timestamp','Event','StaffID','Name','Role','Site','Date','Time','InTime','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','HoursWorked']
  );
  evSheet.appendRow([
    new Date().toISOString(),
    payload.event      || '',
    payload.staffId    || '',
    payload.name       || '',
    payload.role       || '',
    payload.site       || '',
    payload.date       || '',
    payload.time       || '',
    payload.inTime     || '',
    payload.lat        || '',
    payload.lng        || '',
    payload.dist       || '',
    payload.gps        || '',
    payload.device     || '',
    payload.hoursWorked || 0
  ]);

  // 2. If CLOCK_OUT: write a completed shift row to ShiftSummary
  if (payload.event === 'CLOCK_OUT') {
    var shiftSheet = getOrCreateSheet(SHEET_SHIFTS,
      ['StaffID','Name','Role','Site','Date','ClockIn','ClockOut','HoursWorked','GPS_Lat','GPS_Lng','Verification']
    );
    shiftSheet.appendRow([
      payload.staffId,
      payload.name,
      payload.role       || '',
      payload.site,
      payload.date,
      payload.inTime     || '',
      payload.time,
      parseFloat(payload.hoursWorked) || 0,
      payload.lat,
      payload.lng,
      payload.gps
    ]);

    // 3. Update staff running total in Staff sheet
    updateStaffHours(payload.staffId, parseFloat(payload.hoursWorked) || 0);
  }

  // 4. Write to AuditLog (every event, immutable record)
  var auditSheet = getOrCreateSheet(SHEET_AUDIT,
    ['LogTimestamp','Event','StaffID','Name','Site','Date','Time','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','IP']
  );
  auditSheet.appendRow([
    new Date().toISOString(),
    payload.event,
    payload.staffId,
    payload.name,
    payload.site,
    payload.date,
    payload.time,
    payload.lat,
    payload.lng,
    payload.dist,
    payload.gps,
    payload.device,
    'N/A' // IP not available in GAS client
  ]);

  return { status: 'ok', event: payload.event, staffId: payload.staffId };
}

// ── UPDATE STAFF HOURS ───────────────────────────────────────
function updateStaffHours(staffId, hoursToAdd) {
  var sh = getOrCreateSheet(SHEET_STAFF,
    ['StaffID','Name','Role','Site','HoursWorked','Shifts','LastUpdated']
  );
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(staffId).toUpperCase()) {
      var newHrs  = Math.round((parseFloat(data[i][4]) + hoursToAdd) * 100) / 100;
      var newShifts = parseInt(data[i][5]) + 1;
      sh.getRange(i + 1, 5).setValue(newHrs);
      sh.getRange(i + 1, 6).setValue(newShifts);
      sh.getRange(i + 1, 7).setValue(new Date().toISOString());
      return;
    }
  }
  // Staff not found — create a new row
  sh.appendRow([staffId, 'Unknown', '', '', hoursToAdd, 1, new Date().toISOString()]);
}

// ── UPDATE STAFF RECORD ──────────────────────────────────────
function updateStaffRecord(payload) {
  var sh = getOrCreateSheet(SHEET_STAFF,
    ['StaffID','Name','Role','Site','HoursWorked','Shifts','LastUpdated']
  );
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(payload.staffId).toUpperCase()) {
      if (payload.name)  sh.getRange(i + 1, 2).setValue(payload.name);
      if (payload.role)  sh.getRange(i + 1, 3).setValue(payload.role);
      if (payload.site)  sh.getRange(i + 1, 4).setValue(payload.site);
      sh.getRange(i + 1, 7).setValue(new Date().toISOString());
      return;
    }
  }
}

// ── READ HELPERS ─────────────────────────────────────────────
function getAllStaff() {
  var sh = getOrCreateSheet(SHEET_STAFF,
    ['StaffID','Name','Role','Site','HoursWorked','Shifts','LastUpdated']
  );
  var data = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id:      data[i][0],
      name:    data[i][1],
      role:    data[i][2],
      site:    data[i][3],
      hours:   parseFloat(data[i][4]) || 0,
      shifts:  parseInt(data[i][5])   || 0,
      updated: data[i][6] || ''
    });
  }
  return result;
}

function getAllEvents() {
  var sh = getOrCreateSheet(SHEET_EVENTS,
    ['Timestamp','Event','StaffID','Name','Role','Site','Date','Time','InTime','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','HoursWorked']
  );
  var data = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    result.push({
      timestamp:    data[i][0],
      event:        data[i][1],
      staffId:      data[i][2],
      name:         data[i][3],
      role:         data[i][4],
      site:         data[i][5],
      date:         data[i][6],
      time:         data[i][7],
      inTime:       data[i][8],
      lat:          data[i][9],
      lng:          data[i][10],
      dist:         data[i][11],
      gps:          data[i][12],
      device:       data[i][13],
      hoursWorked:  parseFloat(data[i][14]) || 0
    });
  }
  return result;
}

function getAllShifts() {
  var sh = getOrCreateSheet(SHEET_SHIFTS,
    ['StaffID','Name','Role','Site','Date','ClockIn','ClockOut','HoursWorked','GPS_Lat','GPS_Lng','Verification']
  );
  var data = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      staffId:     data[i][0],
      name:        data[i][1],
      role:        data[i][2],
      site:        data[i][3],
      date:        data[i][4],
      clockIn:     data[i][5],
      clockOut:    data[i][6],
      hoursWorked: parseFloat(data[i][7]) || 0,
      lat:         data[i][8],
      lng:         data[i][9],
      gps:         data[i][10]
    });
  }
  return result;
}

// ── SHEET HELPER ─────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    // Style the header row
    var hRange = sh.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight('bold')
          .setBackground('#1a237e')
          .setFontColor('#ffffff')
          .setFontSize(9);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

// ── ONE-TIME SETUP ───────────────────────────────────────────
// Run this once from the Apps Script editor: Run → setupSheets
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setName('ShiftTrack Data');

  // Create all sheets with headers
  getOrCreateSheet(SHEET_STAFF,  ['StaffID','Name','Role','Site','HoursWorked','Shifts','LastUpdated']);
  getOrCreateSheet(SHEET_EVENTS, ['Timestamp','Event','StaffID','Name','Role','Site','Date','Time','InTime','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','HoursWorked']);
  getOrCreateSheet(SHEET_SHIFTS, ['StaffID','Name','Role','Site','Date','ClockIn','ClockOut','HoursWorked','GPS_Lat','GPS_Lng','Verification']);
  getOrCreateSheet(SHEET_AUDIT,  ['LogTimestamp','Event','StaffID','Name','Site','Date','Time','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','IP']);

  // Seed Staff sheet with initial data
  var staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (staffSheet.getLastRow() <= 1) {
    var seed = [
      ['EMP-00421','Sipho Mokoena',  'Security','Sandton Office Park',   142.5, 18, new Date().toISOString()],
      ['EMP-00138','Lerato Dube',    'Cleaning','Rosebank Tower',        163.0, 22, new Date().toISOString()],
      ['EMP-00274','Thabo Nkosi',    'Security','Midrand Logistics Hub',  98.0, 13, new Date().toISOString()],
      ['EMP-00512','Aisha Patel',    'Admin',   'Sandton Office Park',    55.0,  7, new Date().toISOString()],
      ['EMP-00389','Nomsa Khumalo',  'Cleaning','Centurion Gate',        161.5, 21, new Date().toISOString()],
      ['EMP-00601','James Dlamini',  'Security','Pretoria North Depot',  120.0, 16, new Date().toISOString()],
      ['EMP-00712','Zanele Motha',   'Cleaning','Rosebank Tower',         88.5, 12, new Date().toISOString()],
      ['EMP-00834','Pieter van Wyk', 'Security','Centurion Gate',        175.0, 23, new Date().toISOString()],
    ];
    seed.forEach(function(row) { staffSheet.appendRow(row); });
    staffSheet.autoResizeColumns(1, 7);
  }

  // Remove default "Sheet1" if empty
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert(
    '✓ ShiftTrack setup complete!\n\n' +
    'Sheets created:\n' +
    '• Staff — staff records & running hours\n' +
    '• ClockEvents — every clock-in/out event\n' +
    '• ShiftSummary — completed shifts\n' +
    '• AuditLog — immutable audit trail\n\n' +
    'Next step: Deploy → New deployment → Web app\n' +
    'Then paste the URL into the ShiftTrack HTML app.'
  );
}

// ── MONTHLY RESET (run via time trigger on 1st of each month) ─
function monthlyReset() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) return;

  // Archive current month before resetting
  var now = new Date();
  var label = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  var archiveName = 'Archive_' + label;

  // Copy ShiftSummary to archive
  var shiftSheet = ss.getSheetByName(SHEET_SHIFTS);
  if (shiftSheet) {
    shiftSheet.copyTo(ss).setName(archiveName);
  }

  // Reset HoursWorked and Shifts for all staff
  var data = staffSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      staffSheet.getRange(i + 1, 5).setValue(0); // HoursWorked
      staffSheet.getRange(i + 1, 6).setValue(0); // Shifts
      staffSheet.getRange(i + 1, 7).setValue(new Date().toISOString());
    }
  }

  // Clear ShiftSummary for new month (keep headers)
  if (shiftSheet && shiftSheet.getLastRow() > 1) {
    shiftSheet.deleteRows(2, shiftSheet.getLastRow() - 1);
  }

  Logger.log('Monthly reset complete. Archive: ' + archiveName);
}

// ── SETUP MONTHLY TRIGGER ────────────────────────────────────
// Run this once to schedule automatic monthly resets
function setupMonthlyTrigger() {
  // Remove existing triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyReset') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create new trigger: 1st of every month at 00:01
  ScriptApp.newTrigger('monthlyReset')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
  SpreadsheetApp.getUi().alert('Monthly reset trigger created. Runs on the 1st of each month at midnight.');
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ShiftTrack')
    .addItem('Initial setup', 'setupSheets')
    .addItem('Setup monthly trigger', 'setupMonthlyTrigger')
    .addItem('Manual monthly reset', 'monthlyReset')
    .addItem('View deployment URL', 'showDeploymentUrl')
    .addToUi();
}

function showDeploymentUrl() {
  var url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().alert('Web App URL:\n\n' + url + '\n\nCopy this into ShiftTrack HTML → Admin → Sheets sync');
  } else {
    SpreadsheetApp.getUi().alert('No deployment found. Go to Deploy → New deployment to create one first.');
  }
}
