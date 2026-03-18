function doGet(e) {
  var sheetName = 'sheet1';

  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("No active spreadsheet found. Is this script bound to a sheet?");
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      error: "Could not access spreadsheet. Ensure this is a Container-Bound script.",
      details: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];

  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  var headers = data[0];
  var rows = data.slice(1);

  // --- Server-side date filtering ---
  // Accept startDate and endDate as YYYY-MM-DD query params
  // Default: last 30 days (avoids sending entire sheet history)
  var params = e && e.parameter ? e.parameter : {};
  var now = new Date();

  var filterStart = null;
  var filterEnd = null;

  if (params.startDate) {
    var sp = params.startDate.split('-');
    if (sp.length === 3) filterStart = new Date(parseInt(sp[0]), parseInt(sp[1]) - 1, parseInt(sp[2]));
  }
  if (params.endDate) {
    var ep = params.endDate.split('-');
    if (ep.length === 3) {
      filterEnd = new Date(parseInt(ep[0]), parseInt(ep[1]) - 1, parseInt(ep[2]));
      // Include the full end day
      filterEnd.setHours(23, 59, 59, 999);
    }
  }

  // If no date params given, default to TODAY only
  if (!filterStart && !filterEnd) {
    filterStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
    filterEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  }

  // Find which column is the Date column
  var dateColIndex = -1;
  for (var h = 0; h < headers.length; h++) {
    var hdr = String(headers[h]).toLowerCase().trim();
    if (hdr === 'date' || hdr === 'timestamp') {
      dateColIndex = h;
      break;
    }
  }

  var result = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];

    // Apply server-side date filter if we found a date column
    if (dateColIndex >= 0 && (filterStart || filterEnd)) {
      var rawDate = row[dateColIndex];
      var rowDate = null;

      if (rawDate instanceof Date) {
        rowDate = rawDate;
      } else if (typeof rawDate === 'string' && rawDate.trim() !== '') {
        rowDate = new Date(rawDate);
      } else if (typeof rawDate === 'number') {
        // Google Sheets serial number — skip server filter, send row
        rowDate = null;
      }

      if (rowDate && !isNaN(rowDate.getTime())) {
        var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
        if (filterStart && rd < filterStart) continue;
        if (filterEnd) {
          var rdEnd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate(), 23, 59, 59);
          if (rdEnd > filterEnd) continue;
        }
      }
      // If rowDate is null / unparseable, include the row (don't silently drop it)
    }

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        val = val.toISOString();
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
