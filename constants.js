function getSheetId() { 
 const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
 if (!sheetId) {
  throw new Error ('กรุณาตั้งค่า SHEET_ID ก่อนใช้งาน โดยรันฟังก์ชั่น setSheetId()');
 }
 return sheetId;
}
function setSheetId() {
 PropertiesService.getScriptProperties.setProperty('SHEET_ID','YOUR_SHEET_ID');
 return "✅ ตั้งค่า SHEET เรียบร้อยแล้ว";
}