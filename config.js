// Config.js
function setSheetConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = {
    'SHEET_ID': 'YOUR_SHEET_ID_HERE', // แทนที่ด้วย Sheet ID จริง
    'BOT_TRANSACTIONS_SHEET': 'Bot Transactions',
    'DAILY_REPORT_SHEET': 'Daily Report',
    'SUMMARY_SHEET': 'Summary of All',
    'BUDGET_SHEET': 'Budget',
    'LOG_SHEET_NAME': 'TaxLogs'
  };
  // เพิ่ม CONFIG.MONTHS_THAI
  const CONFIG = {
    MONTHS_THAI: [
      "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
      "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
    ]
  };
  scriptProperties.setProperties(properties); // บันทึกคุณสมบัติ
  return { ...properties, CONFIG }; // ส่งคืนทั้ง properties และ CONFIG
}

function processAllSheets() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  
  // ดึงชื่อ sheet จาก Script Properties
  const botSheet = spreadsheet.getSheetByName(scriptProperties.getProperty('BOT_TRANSACTIONS_SHEET'));
  const dailySheet = spreadsheet.getSheetByName(scriptProperties.getProperty('DAILY_REPORT_SHEET'));
  const summarySheet = spreadsheet.getSheetByName(scriptProperties.getProperty('SUMMARY_SHEET'));
  const budgetSheet = spreadsheet.getSheetByName(scriptProperties.getProperty('BUDGET_SHEET'));
  
  // ตัวอย่าง: อ่านข้อมูลจากแต่ละหน้า
  Logger.log('Bot Transactions A1: ' + botSheet.getRange('A1').getValue());
  Logger.log('Daily Report A1: ' + dailySheet.getRange('A1').getValue());
  Logger.log('Summary A1: ' + summarySheet.getRange('A1').getValue());
  Logger.log('Budget A1: ' + budgetSheet.getRange('A1').getValue());
}