/**
 * OSOGBO RCC – Mobile Dashboard JSON API
 * ---------------------------------------
 * Add this as a SECOND file in the SAME Apps Script project as
 * incident_report_form_automation.gs (Apps Script editor > Files > +
 * > Script > name it "DashboardApi"). It shares that project's
 * spreadsheet binding, so no extra setup is needed.
 *
 * DEPLOY:
 *   Deploy > New deployment > Type: Web app
 *     Execute as: Me
 *     Who has access: Anyone
 *   Deploy, then copy the Web app URL - that's what goes into the
 *   dashboard's index.html as CONFIG.API_URL.
 *
 * NOTE ON ACCESS: "Anyone" means anyone with the URL can read the
 * aggregated dashboard stats (no raw spreadsheet access, just the
 * summary JSON below). If you'd rather restrict that, say so and we
 * can look at tightening it - it just adds a bit more setup.
 *
 * Optional: call the URL with ?month=YYYY-MM (e.g. ?month=2026-04) to
 * pull a past month instead of the current one.
 */

var INCIDENT_FIELD_KEYS = [
  'sno', 'date', 'timeOut', 'location', 'type', 'description',
  'actionTaken', 'affectedStations', 'equipment', 'loadInterruptedMW',
  'potentialImpact', 'rootCause', 'correctiveAction', 'dateOfClosure',
  'timeIn', 'status', 'remarks'
];
var INCIDENT_DATA_START_ROW = 3;

var MAXMIN_FIELD_KEYS = [
  'date', 'maxMW', 'maxTime', 'maxFreq', 'maxVoltage',
  'minMW', 'minTime', 'minFreq', 'minVoltage'
];
var MAXMIN_DATA_START_ROW = 5;

function doGet(e) {
  try {
    var monthParam = e && e.parameter && e.parameter.month; // "YYYY-MM"
    var targetDate = monthParam ? parseYearMonth_(monthParam) : new Date();
    var data = buildDashboardData_(targetDate);
    return jsonOutput_(data);
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

function buildDashboardData_(targetDate) {
  var monthNames = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  var monthAbbr = monthNames[targetDate.getMonth()];
  var year = targetDate.getFullYear().toString();
  var monthLabel = Utilities.formatDate(targetDate, Session.getScriptTimeZone() || 'Etc/UTC', 'MMMM yyyy');

  var incidentSheet = findExistingMonthSheet_(monthAbbr, year, 'INCIDENT REPORT');
  var maxMinSheet = findExistingMonthSheet_(monthAbbr, year, 'DAILY MAX/MIN');

  var incidents = incidentSheet ? readIncidentRows_(incidentSheet) : [];
  var loadTrend = maxMinSheet ? readMaxMinRows_(maxMinSheet) : [];

  return {
    month: monthLabel,
    generatedAt: new Date().toISOString(),
    incidentSheetFound: !!incidentSheet,
    maxMinSheetFound: !!maxMinSheet,
    incidentStats: summarizeIncidents_(incidents),
    recentIncidents: incidents.slice(-10).reverse(),
    loadTrend: loadTrend,
    latestReading: loadTrend.length ? loadTrend[loadTrend.length - 1] : null
  };
}

function findExistingMonthSheet_(monthAbbr, year, suffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var canonicalName = monthAbbr + ' ' + year + ' - ' + suffix;

  var sheet = ss.getSheetByName(canonicalName);
  if (sheet) return sheet;

  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var name = all[i].getName().toUpperCase();
    if (name.indexOf(year) !== -1 && name.indexOf(monthAbbr) !== -1 && name.indexOf(suffix) !== -1) {
      return all[i];
    }
  }
  return null;
}

function readIncidentRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < INCIDENT_DATA_START_ROW) return [];

  var values = sheet.getRange(
    INCIDENT_DATA_START_ROW, 1,
    lastRow - INCIDENT_DATA_START_ROW + 1,
    INCIDENT_FIELD_KEYS.length
  ).getValues();

  var rows = [];
  values.forEach(function (r) {
    if (!r[1]) return; // skip rows with no incident date
    var obj = {};
    INCIDENT_FIELD_KEYS.forEach(function (key, i) {
      obj[key] = normalizeCell_(r[i], key);
    });
    rows.push(obj);
  });
  return rows;
}

function readMaxMinRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < MAXMIN_DATA_START_ROW) return [];

  var values = sheet.getRange(
    MAXMIN_DATA_START_ROW, 1,
    lastRow - MAXMIN_DATA_START_ROW + 1,
    MAXMIN_FIELD_KEYS.length
  ).getValues();

  var rows = [];
  values.forEach(function (r) {
    if (!r[0]) return; // skip blank rows
    var obj = {};
    MAXMIN_FIELD_KEYS.forEach(function (key, i) {
      obj[key] = normalizeCell_(r[i], key);
    });
    rows.push(obj);
  });
  return rows;
}

function normalizeCell_(value, key) {
  if (value instanceof Date) {
    var fmt = (key === 'date') ? 'd-MMM' : 'HH:mm';
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Etc/UTC', fmt);
  }
  if (typeof value === 'string') return value.trim();
  return value;
}

function summarizeIncidents_(incidents) {
  var byType = { FORCED: 0, PLANNED: 0, EMERGENCY: 0, URGENT: 0, OTHER: 0 };
  var totalMW = 0;
  var openCount = 0;

  incidents.forEach(function (inc) {
    var t = (inc.type || '').toString().trim().toUpperCase();
    if (byType.hasOwnProperty(t)) byType[t]++; else byType.OTHER++;

    var mw = parseFloat(inc.loadInterruptedMW);
    if (!isNaN(mw)) totalMW += mw;

    var hasClosure = inc.dateOfClosure && inc.dateOfClosure.toString().trim() !== '';
    var status = (inc.status || '').toString().toUpperCase();
    if (!hasClosure || status.indexOf('OUT OF SERVICE') !== -1) openCount++;
  });

  return {
    total: incidents.length,
    byType: byType,
    totalMWInterrupted: Math.round(totalMW * 10) / 10,
    openCount: openCount
  };
}

function parseYearMonth_(s) {
  var parts = s.split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, 1);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
