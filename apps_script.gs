function doGet(e) {
  var sheetName = 'sheet1';
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet found.");
  } catch (error) {
    return createErrorResponse("Spreadsheet Access Error", error.toString());
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // If 'sheet1' doesn't exist, try common variations or fallback to first sheet
    sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  // Increased limit to 5000 rows for better historical visibility
  var MAX_FETCH = 5000; 
  var startRow = Math.max(2, lastRow - MAX_FETCH + 1);
  var numRows = lastRow - startRow + 1;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  var params = e && e.parameter ? e.parameter : {};
  var startDate = params.startDate ? parseDateParam(params.startDate) : null;
  var endDate = params.endDate ? parseDateParam(params.endDate) : null;
  
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
    var rowDate = null;
    
    if (rowDateValue instanceof Date) {
      rowDate = rowDateValue;
    } else if (rowDateValue) {
      rowDate = new Date(rowDateValue);
    }

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
    
    // Removed the 1000 limit to allow for larger data sets to be filtered on frontend
    // Safety break at 5000 (same as fetch)
    if (result.length >= MAX_FETCH) break;
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseDateParam(str) {
  if (!str) return null;
  var d = new Date(str);
  if (isNaN(d.getTime())) return null;
  // Normalize to start of day for comparison
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


function createErrorResponse(msg, details) {
  return ContentService.createTextOutput(JSON.stringify({ error: msg, details: details }))
    .setMimeType(ContentService.MimeType.JSON);
}
