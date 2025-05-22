// Tax Calculation.gs

// การกำหนดค่า
const CONFIG = {
  SHEET_NAME: 'Bot Transactions',
  LOG_SHEET_NAME: 'TaxLogs',
  COLUMNS: {
    DATE: 0,
    TYPE: 1,
    AMOUNT: 2
  },
  MIN_YEAR: 2000,
  DEFAULT_YEAR: new Date().getFullYear()
};

// ข้อความสำหรับ localization
const MESSAGES = {
  th: {
    SUCCESS: 'สำหรับปี %s:\nรายรับ: %s บาท\nรายจ่าย: %s บาท\nลดหย่อน: %s บาท\nบริจาค: %s บาท\nรายได้สุทธิที่ต้องเสียภาษี: %s บาท\nภาษี: %s บาท',
    ERROR: 'เกิดข้อผิดพลาด: %s',
    INVALID_YEAR: 'กรุณาระบุปีระหว่าง %s ถึง %s',
    INVALID_PARAM: 'ข้อมูล %s ไม่ถูกต้อง',
    NO_DATA: 'ไม่พบข้อมูลในชีต %s สำหรับปี %s',
    CONFIG_ERROR: 'การตั้งค่า %s ไม่ถูกต้อง'
  },
  en: {
    SUCCESS: 'For year %s:\nIncome: %s THB\nExpenses: %s THB\nDeductions: %s THB\nDonations: %s THB\nTaxable Income: %s THB\nTax: %s THB',
    ERROR: 'Error: %s',
    INVALID_YEAR: 'Please specify a year between %s and %s',
    INVALID_PARAM: 'Parameter %s is invalid',
    NO_DATA: 'No data found in sheet %s for year %s',
    CONFIG_ERROR: 'Configuration %s is invalid'
  }
};

// อัตราภาษีเงินได้บุคคลธรรมดา (ปี 2568)
const TAX_BRACKETS = [
  { min: 0, max: 150000, rate: 0 },
  { min: 150001, max: 300000, rate: 0.05 },
  { min: 300001, max: 500000, rate: 0.10 },
  { min: 500001, max: 750000, rate: 0.15 },
  { min: 750001, max: 1000000, rate: 0.20 },
  { min: 1000001, max: 2000000, rate: 0.25 },
  { min: 2000001, max: 5000000, rate: 0.30 },
  { min: 5000001, max: Infinity, rate: 0.35 }
];

// ค่าลดหย่อน
const DEDUCTIONS = {
  PERSONAL: 60000,
  SPOUSE: 60000,
  CHILD: 30000,
  CHILD_EDUCATION: 60000,
  PARENT: 30000,
  DISABLED: 60000,
  SOCIAL_SECURITY_MAX: 100000,
  PROVIDENT_FUND_MAX: 500000,
  HOME_LOAN_INTEREST_MAX: 100000,
  LIFE_INSURANCE_MAX: 100000,
  HEALTH_INSURANCE_MAX: 25000,
  PARENT_HEALTH_INSURANCE_MAX: 15000,
  SHOP_DEE_MEE_CHEN_MAX: 40000,
  MATERNITY_MAX: 60000
};

// ฟังก์ชันย่อย: การจัดการข้อความ
function getMessage(lang, key, ...args) {
  const messages = MESSAGES[lang] || MESSAGES.th;
  return Utilities.formatString(messages[key], ...args);
}

// อัพเดทใหม่! ลบการประกาศ CONFIG ที่ซ้ำซ้อน และใช้ CONFIG จากไฟล์กลาง
function calculateTax() {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const taxLogSheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEETS.TAX_LOGS);
}
// ฟังก์ชันย่อย: ตรวจสอบพารามิเตอร์
function validateParameters(params) {
  const errors = [];
  const year = parseInt(params.year) || CONFIG.DEFAULT_YEAR;
  const currentYear = new Date().getFullYear();

  if (year < CONFIG.MIN_YEAR || year > currentYear + 1) {
    errors.push(getMessage('th', 'INVALID_YEAR', CONFIG.MIN_YEAR, currentYear + 1));
  }

  const numberParams = ['numChildren', 'numChildrenEducation', 'numParents', 'numDisabled'];
  numberParams.forEach(param => {
    const value = parseInt(params[param]) || 0;
    if (value < 0) errors.push(getMessage('th', 'INVALID_PARAM', param));
  });

  const floatParams = ['socialSecurity', 'providentFund', 'homeLoanInterest', 'donation', 'specialDonation', 'lifeInsurance', 'healthInsurance', 'parentHealthInsurance', 'shopDeeMeeChen', 'maternity'];
  floatParams.forEach(param => {
    const value = parseFloat(params[param]) || 0;
    if (value < 0 || isNaN(value)) errors.push(getMessage('th', 'INVALID_PARAM', param));
  });

  if (errors.length > 0) {
    throw new Error(errors.join('\n'));
  }

  return {
    year,
    hasSpouse: params.hasSpouse === 'true',
    numChildren: parseInt(params.numChildren) || 0,
    numChildrenEducation: parseInt(params.numChildrenEducation) || 0,
    numParents: parseInt(params.numParents) || 0,
    numDisabled: parseInt(params.numDisabled) || 0,
    socialSecurity: parseFloat(params.socialSecurity) || 0,
    providentFund: parseFloat(params.providentFund) || 0,
    homeLoanInterest: parseFloat(params.homeLoanInterest) || 0,
    donation: parseFloat(params.donation) || 0,
    specialDonation: parseFloat(params.specialDonation) || 0,
    lifeInsurance: parseFloat(params.lifeInsurance) || 0,
    healthInsurance: parseFloat(params.healthInsurance) || 0,
    parentHealthInsurance: parseFloat(params.parentHealthInsurance) || 0,
    shopDeeMeeChen: parseFloat(params.shopDeeMeeChen) || 0,
    maternity: parseFloat(params.maternity) || 0
  };
}

// ฟังก์ชันย่อย: ดึงข้อมูลจาก sheet
function getSheetData(spreadsheet, year) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS_TRANSACTIONS);
  if (!sheet || sheet.getLastRow() <= 1) {
    throw new Error(getMessage('th', 'NO_DATA', CONFIG.SHEETS_TRANSACTIONS, year));
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(...Object.values(CONFIG.COLUMNS)) + 1).getValues();
  return data.filter(row => new Date(row[CONFIG.COLUMNS.DATE]).getFullYear() === year);
}

// ฟังก์ชันย่อย: คำนวณรายรับ-รายจ่าย
function calculateIncomeExpense(data) {
  let totalIncome = 0, totalExpense = 0;
  data.forEach(row => {
    const amount = parseFloat(row[CONFIG.COLUMNS.AMOUNT] || 0);
    if (row[CONFIG.COLUMNS.TYPE] === 'Income') totalIncome += amount;
    else if (row[CONFIG.COLUMNS.TYPE] === 'Expense') totalExpense += amount;
  });
  return { totalIncome, totalExpense };
}

// ฟังก์ชันย่อย: คำนวณค่าลดหย่อน
function calculateDeductions(params, totalIncome) {
  let totalDeductions = DEDUCTIONS.PERSONAL;

  if (params.hasSpouse) totalDeductions += DEDUCTIONS.SPOUSE;
  totalDeductions += params.numChildren * DEDUCTIONS.CHILD;
  totalDeductions += params.numChildrenEducation * (DEDUCTIONS.CHILD_EDUCATION - DEDUCTIONS.CHILD);
  totalDeductions += Math.min(params.numParents, 4) * DEDUCTIONS.PARENT;
  totalDeductions += params.numDisabled * DEDUCTIONS.DISABLED;
  totalDeductions += Math.min(params.socialSecurity, DEDUCTIONS.SOCIAL_SECURITY_MAX);
  totalDeductions += Math.min(params.providentFund, DEDUCTIONS.PROVIDENT_FUND_MAX, totalIncome * 0.15);
  totalDeductions += Math.min(params.homeLoanInterest, DEDUCTIONS.HOME_LOAN_INTEREST_MAX);
  totalDeductions += Math.min(params.lifeInsurance, DEDUCTIONS.LIFE_INSURANCE_MAX);
  totalDeductions += Math.min(params.healthInsurance, DEDUCTIONS.HEALTH_INSURANCE_MAX);
  totalDeductions += Math.min(params.parentHealthInsurance, DEDUCTIONS.PARENT_HEALTH_INSURANCE_MAX);
  totalDeductions += Math.min(params.shopDeeMeeChen, DEDUCTIONS.SHOP_DEE_MEE_CHEN_MAX);
  totalDeductions += Math.min(params.maternity, DEDUCTIONS.MATERNITY_MAX);

  return { totalDeductions, donation: params.donation, specialDonation: params.specialDonation };
}

// ฟังก์ชันย่อย: คำนวณภาษี
function calculateTaxAmount(taxableIncome) {
  let totalTax = 0;
  for (const bracket of TAX_BRACKETS) {
    if (taxableIncome > bracket.min) {
      const taxableInBracket = Math.min(taxableIncome, bracket.max) - bracket.min;
      totalTax += taxableInBracket * bracket.rate;
    }
  }
  return totalTax;
}

// ฟังก์ชันย่อย: บันทึก log
function logAction(spreadsheet, userId, action, result) {
  const logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME) || spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);
  logSheet.appendRow([new Date(), userId || 'unknown', action, JSON.stringify(result)]);
}

// ฟังก์ชันหลัก
function calculateTax(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอ 10 วินาที
    validateConfig();
    const params = validateParameters(e.parameters || {});
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const data = getSheetData(spreadsheet, params.year);

    const { totalIncome, totalExpense } = calculateIncomeExpense(data);
    const { totalDeductions, donation, specialDonation } = calculateDeductions(params, totalIncome);

    let taxableIncome = totalIncome - totalExpense - totalDeductions;
    if (taxableIncome < 0) taxableIncome = 0;

    const donationDeduction = Math.min(donation, taxableIncome * 0.1);
    const specialDonationDeduction = specialDonation * 1.5;
    taxableIncome -= (donationDeduction + specialDonationDeduction);
    if (taxableIncome < 0) taxableIncome = 0;

    const totalTax = calculateTaxAmount(taxableIncome);

    const response = {
      fulfillmentText: getMessage('th', 'SUCCESS',
        params.year,
        totalIncome.toLocaleString(),
        totalExpense.toLocaleString(),
        totalDeductions.toLocaleString(),
        (donationDeduction + specialDonationDeduction).toLocaleString(),
        taxableIncome.toLocaleString(),
        totalTax.toLocaleString()
      )
    };

    logAction(spreadsheet, e.userId, 'calculateTax', response);
    return response;

  } catch (error) {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    logAction(spreadsheet, e.userId, 'calculateTax_error', error.message);
    return {
      fulfillmentText: getMessage('th', 'ERROR', error.message)
    };
  } finally {
    lock.releaseLock();
  }
}

// ฟังก์ชันทดสอบ
function testCalculateTax() {
  const testCases = [
    {
      name: 'Normal case',
      params: {
        year: '2025',
        hasSpouse: 'true',
        numChildren: '2',
        numChildrenEducation: '1',
        numParents: '2',
        numDisabled: '0',
        socialSecurity: '50000',
        providentFund: '100000',
        homeLoanInterest: '80000',
        donation: '20000',
        specialDonation: '10000',
        lifeInsurance: '50000',
        healthInsurance: '20000',
        parentHealthInsurance: '10000',
        shopDeeMeeChen: '30000',
        maternity: '50000'
      }
    },
    {
      name: 'Invalid year',
      params: { year: '1800' }
    },
    {
      name: 'Negative children',
      params: { year: '2025', numChildren: '-1' }
    }
  ];

  testCases.forEach(test => {
    Logger.log(`Running test: ${test.name}`);
    const result = calculateTax({ parameters: test.params, userId: 'test_user' });
    Logger.log(JSON.stringify(result, null, 2));
  });
}