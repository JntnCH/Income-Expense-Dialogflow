function calculateTax(e) {
  try {
    const params = e.parameters || {};
    const year = params.year ? parseInt(params.year) : new Date().getFullYear();
    const hasSpouse = params.hasSpouse === "true";
    const numChildren = parseInt(params.numChildren) || 0;
    const numChildrenEducation = parseInt(params.numChildrenEducation) || 0;
    const numParents = parseInt(params.numParents) || 0;
    const numDisabled = parseInt(params.numDisabled) || 0;
    const socialSecurity = parseFloat(params.socialSecurity) || 0;
    const providentFund = parseFloat(params.providentFund) || 0;
    const homeLoanInterest = parseFloat(params.homeLoanInterest) || 0;
    const donation = parseFloat(params.donation) || 0;
    const specialDonation = parseFloat(params.specialDonation) || 0;
    const lifeInsurance = parseFloat(params.lifeInsurance) || 0;
    const healthInsurance = parseFloat(params.healthInsurance) || 0;
    const parentHealthInsurance = parseFloat(params.parentHealthInsurance) || 0;
    const shopDeeMeeChen = parseFloat(params.shopDeeMeeChen) || 0;
    const maternity = parseFloat(params.maternity) || 0;
    
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName("Bot Transactions");
    const data = sheet.getDataRange().getValues();
    
    let totalIncome = 0;
    let totalExpense = 0;
    
    data.forEach(row => {
      const date = new Date(row[0]); // A: วันที่
      const type = row[1]; // B: ประเภท
      const amount = parseFloat(row[2] || 0); // C: จำนวนเงิน
      
      if (isNaN(amount) || !type || date.getFullYear() !== year) return;
      
      if (type === "Income") {
        totalIncome += amount;
      } else if (type === "Expense") {
        totalExpense += amount;
      }
    });
    
    let totalDeductions = PERSONAL_DEDUCTION;
    if (hasSpouse) totalDeductions += SPOUSE_DEDUCTION;
    totalDeductions += numChildren * CHILD_DEDUCTION;
    totalDeductions += numChildrenEducation * (CHILD_EDUCATION_DEDUCTION - CHILD_DEDUCTION);
    totalDeductions += Math.min(numParents, 4) * PARENT_DEDUCTION;
    totalDeductions += numDisabled * DISABLED_DEDUCTION;
    totalDeductions += Math.min(socialSecurity, SOCIAL_SECURITY_MAX);
    totalDeductions += Math.min(providentFund, PROVIDENT_FUND_MAX, totalIncome * 0.15);
    totalDeductions += Math.min(homeLoanInterest, HOME_LOAN_INTEREST_MAX);
    totalDeductions += Math.min(lifeInsurance, LIFE_INSURANCE_MAX);
    totalDeductions += Math.min(healthInsurance, HEALTH_INSURANCE_MAX);
    totalDeductions += Math.min(parentHealthInsurance, PARENT_HEALTH_INSURANCE_MAX);
    totalDeductions += Math.min(shopDeeMeeChen, SHOP_DEE_MEE_CHEN_MAX);
    totalDeductions += Math.min(maternity, MATERNITY_DEDUCTION_MAX);
    
    let taxableIncome = totalIncome - totalExpense - totalDeductions;
    if (taxableIncome < 0) taxableIncome = 0;
    
    const donationDeduction = Math.min(donation, taxableIncome * 0.1);
    const specialDonationDeduction = specialDonation * 1.5;
    taxableIncome -= (donationDeduction + specialDonationDeduction);
    if (taxableIncome < 0) taxableIncome = 0;
    
    let totalTax = 0;
    for (const bracket of TAX_BRACKETS) {
      if (taxableIncome > bracket.min) {
        const taxableInBracket = Math.min(taxableIncome, bracket.max) - bracket.min;
        totalTax += taxableInBracket * bracket.rate;
      }
    }
    
    const response = {
      fulfillmentText: `สำหรับปี ${year}:\n` +
        `รายรับทั้งหมด: ${totalIncome.toLocaleString()} บาท\n` +
        `รายจ่ายทั้งหมด: ${totalExpense.toLocaleString()} บาท\n` +
        `ค่าลดหย่อนทั้งหมด: ${totalDeductions.toLocaleString()} บาท\n` +
        `เงินบริจาค: ${(donationDeduction + specialDonationDeduction).toLocaleString()} บาท\n` +
        `รายได้สุทธิที่ต้องเสียภาษี: ${taxableIncome.toLocaleString()} บาท\n` +
        `ภาษีที่ต้องชำระ: ${totalTax.toLocaleString()} บาท`
    };
    
    return response;
    
  } catch (error) {
    Logger.log("Error: " + error);
    return {
      fulfillmentText: "เกิดข้อผิดพลาดในการคำนวณภาษี กรุณาลองใหม่"
    };
  }
}