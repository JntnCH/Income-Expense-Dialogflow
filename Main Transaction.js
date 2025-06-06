function formatNumberWithIntl(number, locale, options) {
  try {
    return new Intl.NumberFormat(locale, options).format(number);
  } catch (e) {
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
  const category = "ยอดคงเหลือรายเดือน";
  const detail = "บันทึกยอดคงเหลือ ณ สิ้นเดือน";
  const formattedDate = today.getDate() + " " + CONFIG.MONTHS_THAI[today.getMonth()] + " " + today.getFullYear();
  
  sheet.appendRow([
    formattedDate,
    "balance",
    balance,
    detail,
    category,
    "System"
  ]);
}

/*
ฟังก์ชันหลักที่ใช้รับข้อมูลจาก Dialogflow และบันทึกลง Google Sheets
@param {object} e - ข้อมูลที่ได้รับจาก Dialogflow
@returns {TextOutput} คำตอบที่ส่งกลับไปยัง Dialogflow 
*/
function doPost(e) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(BOT_TRANSACTIONS_SHEET);
  const dailyReportSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DAILY_REPORT_SHEET);
  const data = JSON.parse(e.postData.contents);
  const intent = data.queryResult.intent.displayName;
  const params = data.queryResult.parameters;
  const source = data.originalDetectIntentRequest.source;
  let responseText = "";
  
  if (intent === "บันทึกรายจ่าย" || intent === "บันทึกรายรับ") {
    let category = "-";
    const amount = Number(params.amount);
    let detail = "-";
    const type = (intent === "บันทึกรายจ่าย") ? "expense" : "income";
    
    if (type === "expense") {
      category = params["Expense-category"];
      detail = params["Expense-categoryoriginal"];
    } else {
      category = params["Income_category"];
      detail = params["Income_categoryoriginal"];
    }
    const date = new Date();
    const formattedDate = date.getDate() + " " + CONFIG.MONTHS_THAI[date.getMonth()] + " " + date.getFullYear();
    // การบันทึก
    sheet.appendRow([
      formattedDate, // วันที่
      type,          // ประเภท รายรับ/รายจ่าย
      amount,        // จำนวนเงิน
      detail,        // รายละเอียด
      category,      // หมวดหมู่
      source         // แหล่งที่มา
    ]);
    
    // ปรับการตอบกลับเป็น HTML สำหรับ Telegram
    if (source === "telegram") {
      responseText = `<b>📝 ฉันบันทึก</b>\n` +
        `<b>รายการ:</b> ${detail}\n` +
        `<b>หมวดหมู่:</b> ${category}\n` +
        `<b>จำนวน:</b> ${formatNumberWithIntl(amount, 'th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท\n` +
        `<b>สถานะ:</b> ✅ สำเร็จ`;
    } else {
      responseText = `📝 ฉันบันทึก ${detail}\n️🎟️ หมวดหมู่ ${category}\n\n💷 จำนวน ${amount} บาท\n\n✅ ให้คุณเรียบร้อยแล้ว `;
    }
    
  } else if (intent === "เช็คยอด") {
    try {
      const dailyReportValues = dailyReportSheet.getRange("C4:C10").getValues();
      const formatOptions = { minimumFractionDigits: 2, maximumFractionDigits: 2 };
      const formatter = new Intl.NumberFormat('th-TH', formatOptions);
      const formatValue = (value) => {
        const num = Number(value);
        return isNaN(num) ? '0.00' : formatter.format(num);
      };
      
      const [dailyIncome, dailyExpense, , monthlyIncome, monthlyExpense, , balance] =
      dailyReportValues.map(row => formatValue(row[0]));
      
      const today = new Date();
      const todayFormatted = Utilities.formatDate(today, "Asia/Bangkok", "dd MMM yyyy");
      
      const allTransactions = sheet.getDataRange().getValues();
      const todayTransactions = allTransactions
        .filter((row, index) => index !== 0 && Utilities.formatDate(new Date(row[0]), "Asia/Bangkok", "dd MMM yyyy") === todayFormatted)
        .map(row => `${row[3]} ${formatValue(row[2])} บาท`);
      
      // ปรับการตอบกลับเป็น HTML สำหรับ Telegram และใช้ <ol> สำหรับรายการวันนี้
      if (source === "telegram") {
        const transactionList = todayTransactions.length > 0 ?
          `<ol>\n${todayTransactions.map(item => `<li>${item}</li>`).join('\n')}\n</ol>` :
          "ไม่มีรายการ";
        responseText = `<b>🗓️ ยอดประจำวันที่ ${todayFormatted}</b>\n\n` +
          `<b>📝 รายการวันนี้</b>\n${transactionList}\n\n` +
          `<b>💸 รายรับวันนี้:</b> ${dailyIncome} บาท\n` +
          `<b>🛍️ รายจ่ายวันนี้:</b> ${dailyExpense} บาท\n\n` +
          `<b>💰 รายรับเดือนนี้:</b> ${monthlyIncome} บาท\n` +
          `<b>📑 รายจ่ายเดือนนี้:</b> ${monthlyExpense} บาท\n\n` +
          `<b>🪙 ยอดคงเหลือ:</b> ${balance} บาท`;
      } else {
        responseText = `🗓️ ยอดประจำวันที่ ${todayFormatted}\n` +
          `\n📝 รายการวันนี้\n\n` + (todayTransactions.length > 0 ? todayTransactions.join("\n") : "ไม่มีรายการ") +
          `\n\n💸 รายรับวันนี้  ${dailyIncome} บาท\n` +
          `️🛍️ รายจ่ายวันนี้ ${dailyExpense} บาท\n\n` +
          `💰 รายรับเดือนนี้  ${monthlyIncome} บาท\n` +
          `📑 รายจ่ายเดือนนี้  ${monthlyExpense} บาท\n\n` +
          `🪙 ยอดคงเหลือ ${balance} บาท`;
      }
    } catch (error) {
      responseText = source === "telegram" ?
        `<b>เกิดข้อผิดพลาด:</b> ${error.message}` :
        `เกิดข้อผิดพลาด: ${error.message}`;
    }
  }
  
  const today = new Date();
  if (today.getDate() === 1) {
    recordMonthlyBalance(sheet, dailyReportSheet);
  }
  
  // ปรับ JSON response สำหรับ Telegram เพื่อระบุ parse_mode: HTML
  const responseJson = source === "telegram" ?
    {
      "fulfillmentText": responseText,
      "fulfillmentMessages": [{
        "text": { "text": [responseText] },
        "platform": "TELEGRAM"
      }],
      "payload": {
        "telegram": {
          "text": responseText,
          "parse_mode": "HTML"
        }
      }
    } :
    { "fulfillmentText": responseText };
  
  return ContentService.createTextOutput(JSON.stringify(responseJson))
    .setMimeType(ContentService.MimeType.JSON);
}