// Config.gs
function setConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'SHEET_ID': 'YOUR_SHEET_ID_HERE',
    'SHEET_TRANSACTIONS': 'Bot Transactions',
    'SHEET_DAILY_REPORT' : 'Dialy Report'
    
  });
}

// MainScript.gs
function processSheetData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_ID');
  const sheetName = scriptProperties.getProperty('SHEET_NAME');
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  Logger.log(sheet.getRange('A1').getValue()); // ตัวอย่างการอ่านข้อมูล
}