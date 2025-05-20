/*
   ฟังก์ชันสำหรับการตั้งค่า SHEET_ID ใน Script Properties
   วิธีใช้: แก้ไข ID ในฟังก์ชัน setSheetId() แล้วรันฟังก์ชันนี้หนึ่งครั้ง
 */

// ฟังก์ชันสำหรับตั้งค่า SHEET_ID
function setSheetId() {
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', 'YOUR_SHEET_ID');
  return "✅ ตั้งค่า SHEET_ID เรียบร้อย";
}

// ฟังก์ชันสำหรับดึงค่า SHEET_ID
function getSheetId() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    return "❌ ยังไม่ได้ตั้งค่า SHEET_ID กรุณารันฟังก์ชัน setSheetId() ก่อน";
  }
  return "SHEET_ID ปัจจุบัน: " + sheetId;
}
