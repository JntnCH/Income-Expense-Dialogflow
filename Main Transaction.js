//ตัวแปรส่วนกลางสำหรับเก็บ SHEET_ID
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');

// ตัวแปรส่วนกลางสำหรับเก็บเดือนย่อ
const monthsThai = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

function formatNumberWithIntl(number, locale, options) 
{ 
   try { 
      return new Intl.NumberFormat(locale, options).format(number);
      }catch (e) {
         return number; 
         
      } 
 }

/*
ฟังก์ชันสำหรับบันทึกยอดคงเหลือรายเดือน
@param {Sheet} sheet - ชีตที่ใช้บันทึกข้อมูล
@param {Sheet} dailyReportSheet - ชีตรายงานประจำวัน 
*/
function recordMonthlyBalance(sheet, dailyReportSheet) { 
   const today = new Date(); 
   const balance = dailyReportSheet.getRange("C10").getValue();
   
   const category = "ยอดคงเหลือรายเดือน"; // หมวดหมู่ข้อมูล
   const detail = "บันทึกยอดคงเหลือ ณ สิ้นเดือน"; // รายละเอียดข้อมูล
   
   const formattedDate = today.getDate() + " " + monthsThai[today.getMonth()] + " " + today.getFullYear();

sheet.appendRow([ 
   formattedDate,  // วันที่
   "balance",      // ประเภทข้อมูล 
   balance,        // จำนวนเงิน 
   detail,         // รายละเอียด 
   category,       // หมวดหมู่ 
   "System"        // แหล่งที่มา 
   ]); 
}
/*
ฟังก์ชันหลักที่ใช้รับข้อมูลจาก Dialogflow และบันทึกลง Google Sheets
@param {object} e - ข้อมูลที่ได้รับจาก Dialogflow
@returns {TextOutput} คำตอบที่ส่งกลับไปยัง Dialogflow 
*/
function doPost(e) { 
   const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Bot Transactions"); 
   const dailyReportSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Daily report");
   const data = JSON.parse(e.postData.contents); 
   const intent = data.queryResult.intent.displayName; 
   const params = data.queryResult.parameters; 
   const source = data.originalDetectIntentRequest.source; 
   let responseText = "";

if (intent === "บันทึกรายจ่าย" || intent === "บันทึกรายรับ") 
{ 
   let category = "-"; // หมวดหมู่ของรายการ
   const amount = Number(params.amount); //จำนวนเงิน 
   let detail = "-"; // รายละเอียดของรายการ 
   const type = (intent === "บันทึกรายจ่าย") ? "expense" : "income"; // ประเภทของรายการ (รายรับ/รายจ่าย)

if (type === "expense") {
     category = params["Expense-category"]; // หมวดหมู่รายจ่าย
     detail = params["Expense-categoryoriginal"]; // รายละเอียดหมวดหมู่รายจ่าย
 } else {
     category = params["Income_category"]; // หมวดหมู่รายรับ
     detail = params["Income_categoryoriginal"]; // รายละเอียดหมวดหมู่รายรับ
 }
    const date = new Date();
    const formattedDate = date.getDate() + " " + monthsThai[date.getMonth()] + " " + date.getFullYear();

 // บันทึกข้อมูลลงชีต
 sheet.appendRow([
     formattedDate, // วันที่
     type,         // ประเภท (รายรับ/รายจ่าย)
     amount,       // จำนวนเงิน
     detail,       // รายละเอียด
     category,     // หมวดหมู่
     source        // แหล่งที่มา (LINE, Telegram ฯลฯ)
 ]);

 responseText = `📝 ฉันบันทึก ${detail}\n️🎟️ หมวดหมู่ ${category}\n\n💷 จำนวน ${amount} บาท\n\n✅ ให้คุณเรียบร้อยแล้ว `;

} else if (intent === "เช็คยอด") { 
   try { 
     const dailyReportValues = dailyReportSheet.getRange("C4:C10").getValues();
     const dailyIncome = formatNumberWithIntl(Number(dailyReportValues[0][0]), 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || 0;
     const dailyExpense = formatNumberWithIntl(Number(dailyReportValues[1][0]), 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || 0;
     const monthlyIncome = formatNumberWithIntl(Number(dailyReportValues[3][0]), 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || 0;
     const monthlyExpense = formatNumberWithIntl(Number(dailyReportValues[4][0]), 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || 0;
     const balance = formatNumberWithIntl(Number(dailyReportValues[6][0]), 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || 0;

     const today = new Date();
     const todayFormatted = Utilities.formatDate(today, "Asia/Bangkok", "dd MMM yyyy");

     const allTransactions = sheet.getDataRange().getValues();

     // ค้นหารายการที่เป็นของวันนี้
     const todayTransactions = allTransactions
         .filter((row, index) => index !== 0 && Utilities.formatDate(new Date(row[0]), "Asia/Bangkok", "dd MMM yyyy") === todayFormatted)
         .map(row => `${row[3]} ${row[2]} บาท`);

     responseText = `🗓️ ยอดประจำวันที่ ${todayFormatted}\n` +
                    `\n📝 รายการวันนี้\n\n` + (todayTransactions.length > 0 ? todayTransactions.join("\n") : "ไม่มีรายการ") +
                    `\n\n💸 รายรับวันนี้  ${dailyIncome} บาท\n` +
                    `️🛍️ รายจ่ายวันนี้ ${dailyExpense} บาท\n\n` +
                    `💰 รายรับเดือนนี้  ${monthlyIncome} บาท\n` +
                    `📑 รายจ่ายเดือนนี้  ${monthlyExpense} บาท\n\n` +
                    `🪙 ยอดคงเหลือ ${balance} บาท`;
 } 
 catch (error) 
 {
     responseText = `เกิดข้อผิดพลาด: ${error.message}`;
  }

}

const today = new Date(); 
if (today.getDate() === 1) { 
    recordMonthlyBalance(sheet, dailyReportSheet);
}

  return ContentService.createTextOutput(JSON.stringify({ 
   "fulfillmentText": responseText 
   })).setMimeType(ContentService.MimeType.JSON);
}
