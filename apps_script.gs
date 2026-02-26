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
  
  var result = rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      var val = row[index];
      // Convert Date objects to ISO strings for consistent JSON parsing
      if (val instanceof Date) {
        val = val.toISOString();
      }
      obj[header] = val;
    });
    return obj;
  });
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
