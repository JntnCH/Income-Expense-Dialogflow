
// Insert ID sheet into this function

/*
   ฟังก์ชันสำหรับการตั้งค่า SHEET_ID ใน Script Properties
   วิธีใช้: แก้ไข ID ในฟังก์ชัน setSheetId() แล้วรันฟังก์ชันนี้หนึ่งครั้ง
 */

// ฟังก์ชันสำหรับตั้งค่า SHEET_ID

function setSheetId() {
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', 'YOUR_SHEET_ID_HERE');
  return "✅ ตั้งค่า SHEET_ID เรียบร้อย";
}

// ดึงค่าจาก Sctipt Properties
function getSheetId() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    throw new Error('SHEET_ID is not set. Please run the setSheetId() function to configure it.');
  }
  return sheetId;
}
// ฟังก์ชันสำหรับลบการตั้งค่า SHEET_ID
function clearSheetId() {
  PropertiesService.getScriptProperties().deleteProperty('SHEET_ID');
  return "✅ ลบการตั้งค่า SHEET_ID เรียบร้อย";
}