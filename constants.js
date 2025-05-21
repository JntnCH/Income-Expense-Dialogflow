// ใส่ SHEET_ID ของคุณ ทึ่ฟังก์ชั่นนี้
function setSheetId() {
  PropertiesService.getScriptProperties.setProperty('SHEET_ID', 'YOUR_SHEET_ID_HERE');
  return "✅ ตั้งค่า SHEET เรียบร้อยแล้ว";
}
// ดึงค่าจาก Sctipt Properties
function getSheetId() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    throw new Error('กรุณาตั้งค่า SHEET_ID ก่อนใช้งาน โดยรันฟังก์ชั่น setSheetId()');
  }
  return sheetId;
}
// ฟังก์ชันสำหรับลบการตั้งค่า SHEET_ID
function clearSheetId() {
  PropertiesService.getScriptProperties().deleteProperty('SHEET_ID');
  return "✅ ลบการตั้งค่า SHEET_ID เรียบร้อย";
}