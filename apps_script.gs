function doGet(e) {
  var sheetName = 'sheet1';
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet found.");
  } catch (error) {
    return createErrorResponse("Spreadsheet Access Error", error.toString());
  }

  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  // Optimization: Fetch only the last 2000 rows to ensure fast response.
  // 2000 rows should easily cover Today & Yesterday plus several weeks.
  var MAX_FETCH = 2000; 
  var startRow = Math.max(2, lastRow - MAX_FETCH + 1);
  var numRows = lastRow - startRow + 1;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  var params = e && e.parameter ? e.parameter : {};
  var startDate = params.startDate ? new Date(params.startDate) : null;
  var endDate = params.endDate ? new Date(params.endDate) : null;
  
  // Identify Date column index
  var dateIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h === 'date' || h === 'timestamp') { dateIdx = i; break; }
  }

  var result = [];
  // Process rows from bottom to top (most recent first)
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var rowDateValue = dateIdx >= 0 ? row[dateIdx] : null;
    var rowDate = (rowDateValue instanceof Date) ? rowDateValue : (rowDateValue ? new Date(rowDateValue) : null);

    // Apply date filtering if parameters provided
    if (startDate && rowDate && rowDate < startDate) continue;
    if (endDate && rowDate && rowDate > endDate) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        obj[headers[j]] = val.toISOString();
      } else {
        obj[headers[j]] = val;
      }
    }
    result.push(obj);
    
    // Safety break
    if (result.length >= 1000) break;
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function createErrorResponse(msg, details) {
  return ContentService.createTextOutput(JSON.stringify({ error: msg, details: details }))
    .setMimeType(ContentService.MimeType.JSON);
}
