function doGet(e) {
  // If ?api=true is passed, return JSON data for GitHub Pages to fetch
  if (e && e.parameter && e.parameter.api === 'true') {
    const data = getSheetData();
    return ContentService.createTextOutput(data)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Otherwise, return the HTML interface for Apps Script
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Equipment Directory')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheetData() {
  // === CONFIGURATION ===
  // 1. Replace with your actual Spreadsheet ID (from the URL of your Google Sheet)
  const sheetId = '1CJkTy5Z9XvT-Wq67M201tAN2RVxS8FUA4JeNItww8F0'; 
  
  // 2. Replace with the name of the tab in your Google Sheet (e.g., 'Sheet1')
  const sheetName = 'Form Responses 1'; 
  // =====================
  
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
        return { error: 'Sheet "' + sheetName + '" not found.' };
    }
    
    // Get all data from the sheet as strings to prevent google.script.run silent failures (e.g. from Date objects)
    const data = sheet.getDataRange().getDisplayValues();
    
    if (data.length === 0) {
        return { data: [] };
    }
    
    // Assume the first row contains column headers
    const headers = data[0];
    const rows = data.slice(1);
    
    // Convert rows into an array of objects based on headers
    const result = rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        // Ensure header is a string and trim whitespace
        const cleanHeader = String(header).trim();
        obj[cleanHeader] = row[index];
      });
      return obj;
    });
    
    return JSON.stringify({ data: result });
  } catch (e) {
    return JSON.stringify({ error: e.toString() + " - Make sure your Spreadsheet ID is correct and the Apps Script has permissions." });
  }
}
