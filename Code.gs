// ============================================================
// Code.gs  —  ShiftTrack Google Apps Script backend
// ============================================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('ShiftTrack — Staff Clock-In')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── helpers ──────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

// ── Called from client: validate staff number ────────────────
function getStaffRecord(staffId) {
  var sh = getSheet('Staff');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === staffId.toUpperCase()) {
      return {
        found: true,
        id:     data[i][0],
        name:   data[i][1],
        role:   data[i][2],
        site:   data[i][3],
        hours:  Number(data[i][4]) || 0,
        shifts: Number(data[i][5]) || 0
      };
    }
  }
  return { found: false };
}

// ── Called from client: save a clock event ──────────────────
function saveClockEvent(payload) {
  var sh = getSheet('ClockEvents');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Event','StaffID','Name','Site','Date','Time','GPS_Lat','GPS_Lng','Dist_m','Verification','Device','Timestamp']);
  }
  sh.appendRow([
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
    new Date().toISOString()
  ]);

  // If clocking out, update the staff hours in Staff sheet
  if (payload.event === 'CLOCK_OUT') {
    updateStaffHours(payload.staffId, payload.hoursWorked);
  }
  return { ok: true };
}

function updateStaffHours(staffId, hoursToAdd) {
  var sh = getSheet('Staff');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === staffId.toUpperCase()) {
      var currentHrs = Number(data[i][4]) || 0;
      var newHrs = Math.round((currentHrs + hoursToAdd) * 100) / 100;
      sh.getRange(i + 1, 5).setValue(newHrs);
      // Increment shift count
      var currentShifts = Number(data[i][5]) || 0;
      sh.getRange(i + 1, 6).setValue(currentShifts + 1);
      return;
    }
  }
}

// ── Export: returns data to client as JSON for Google Sheets ─
function getExportData(type) {
  var staff = getAllStaff();
  var events = getAllClockEvents();
  return JSON.stringify({ staff: staff, events: events, type: type });
}

function getAllStaff() {
  var sh = getSheet('Staff');
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      rows.push({
        id: data[i][0], name: data[i][1], role: data[i][2],
        site: data[i][3], hours: Number(data[i][4]) || 0,
        shifts: Number(data[i][5]) || 0
      });
    }
  }
  return rows;
}

function getAllClockEvents() {
  var sh = getSheet('ClockEvents');
  if (sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      event: data[i][0], staffId: data[i][1], name: data[i][2],
      site: data[i][3], date: data[i][4], time: data[i][5],
      lat: data[i][6], lng: data[i][7], dist: data[i][8],
      gps: data[i][9], device: data[i][10]
    });
  }
  return rows;
}

// ── One-time setup: seed Staff sheet with sample data ────────
function setupStaffSheet() {
  var sh = getSheet('Staff');
  sh.clearContents();
  sh.appendRow(['StaffID','Name','Role','Site','HoursWorked','Shifts']);
  var seed = [
    ['EMP-00421','Sipho Mokoena',  'Security','Sandton Office Park',   142.5, 18],
    ['EMP-00138','Lerato Dube',    'Cleaning','Rosebank Tower',        163.0, 22],
    ['EMP-00274','Thabo Nkosi',    'Security','Midrand Logistics Hub',  98.0, 13],
    ['EMP-00512','Aisha Patel',    'Admin',   'Sandton Office Park',    55.0,  7],
    ['EMP-00389','Nomsa Khumalo',  'Cleaning','Centurion Gate',        161.5, 21],
    ['EMP-00601','James Dlamini',  'Security','Pretoria North Depot',  120.0, 16],
    ['EMP-00712','Zanele Motha',   'Cleaning','Rosebank Tower',         88.5, 12],
    ['EMP-00834','Pieter van Wyk', 'Security','Centurion Gate',        175.0, 23],
  ];
  seed.forEach(function(r){ sh.appendRow(r); });
  SpreadsheetApp.getUi().alert('Staff sheet seeded successfully.');
}

// ── Returns the URL of the backing spreadsheet ───────────────
function getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

// ── Creates a new export sheet and returns its URL ──────────
function createExportSheet(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var label = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var newSs = SpreadsheetApp.create('ShiftTrack Export — ' + type + ' — ' + label);

  var staffSh = getSheet('Staff');
  var staffData = staffSh.getDataRange().getValues();

  // Monthly Summary sheet
  var sumSh = newSs.getActiveSheet();
  sumSh.setName('Monthly Summary');
  sumSh.appendRow(['Staff No.','Full Name','Role','Site','Hours Worked','Remaining Hrs','% of Limit','Shifts','Status']);
  for (var i = 1; i < staffData.length; i++) {
    if (!staffData[i][0]) continue;
    var hrs = Number(staffData[i][4]) || 0;
    var pct = (hrs / 180 * 100).toFixed(1) + '%';
    var status = hrs >= 180 ? 'BLOCKED' : hrs >= 160 ? 'Near limit' : 'On track';
    sumSh.appendRow([staffData[i][0],staffData[i][1],staffData[i][2],staffData[i][3],hrs,(180-hrs).toFixed(1),pct,staffData[i][5],status]);
  }
  sumSh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  sumSh.setFrozenRows(1);
  sumSh.autoResizeColumns(1,9);

  // Clock Events sheet (if not limit-only)
  if (type !== 'limit') {
    var evSh = newSs.insertSheet('Shift Log');
    var evData = getSheet('ClockEvents').getDataRange().getValues();
    evData.forEach(function(row,idx){
      if (type === 'gps' && idx > 0 && row[9] !== 'GPS Drift') return;
      evSh.appendRow(row);
    });
    evSh.getRange(1,1,1,evData[0].length).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    evSh.setFrozenRows(1);
    evSh.autoResizeColumns(1,evData[0]?evData[0].length:10);
  }

  // Audit log (admin only)
  if (type === 'audit' || type === 'full') {
    var auditSh = newSs.insertSheet('Audit Log');
    var auRows = getSheet('ClockEvents').getDataRange().getValues();
    auRows.forEach(function(row){auditSh.appendRow(row);});
    auditSh.getRange(1,1,1,auRows[0]?auRows[0].length:12).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    auditSh.setFrozenRows(1);
  }

  return newSs.getUrl();
}
