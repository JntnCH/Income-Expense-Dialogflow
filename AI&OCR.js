 /**
 * บอทวิเคราะห์เอกสารอัตโนมัติ
 * รองรับการวิเคราะห์ไฟล์ PDF และรูปภาพด้วย OCR และ AI
 * สนับสนุนบริการ AI และ OCR หลากหลาย: OpenAI, Anthropic, Google, Mistral
  
  ขั้นตอนการใช้งาน:
  1. ตั้งค่า SHEET_ID ด้วยฟังก์ชัน setSheetId()
  2. ตั้งค่า API Keys ด้วยฟังก์ชัน setupAPIKeys()
  3. สร้าง Config Sheet ด้วยฟังก์ชัน setupConfigSheet()
  4. ตั้งค่า OCR และ AI ตามต้องการด้วยฟังก์ชัน setOCRConfig() และ setAIConfig()
  5. เผยแพร่เป็น Web App และเชื่อมต่อกับ Dialogflow
 */


// ============= การตั้งค่า API Keys =============


//  ฟังก์ชันสำหรับตั้งค่า API Keys ทั้งหมด
//  รันฟังก์ชันนี้หนึ่งครั้งเพื่อตั้งค่า
function setupAPIKeys() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // แทนที่ด้วย API Keys ของคุณ
  scriptProperties.setProperty("OPENAI_API_KEY", "sk-xxxxxxxxxxxxxxxxxxxxxxxx");
  scriptProperties.setProperty("ANTHROPIC_API_KEY", "sk-ant-xxxxxxxxxxxxxxxxxxxxxxxx");
  scriptProperties.setProperty("GEMINI_API_KEY", "AIzaSyCuDEdFYSkmvd3qIXyclfdagHUSn-HEEkU");
  scriptProperties.setProperty("MISTRAL_API_KEY", "xxxxxxxxxxxxxxxxxxxxxxxx");
  scriptProperties.setProperty("VISION_API_KEY", "AIzaxxxxxxxxxxxxxxxxxxxxxxxx");
  
  return "✅ ตั้งค่า API Keys เรียบร้อย";
}

/**
 * ฟังก์ชันสำหรับตรวจสอบ API Keys ที่ตั้งค่าไว้
 */
function checkAPIKeys() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const keys = {
    "OPENAI_API_KEY": scriptProperties.getProperty("OPENAI_API_KEY") ? "✅ ตั้งค่าแล้ว" : "❌ ยังไม่ได้ตั้งค่า",
    "ANTHROPIC_API_KEY": scriptProperties.getProperty("ANTHROPIC_API_KEY") ? "✅ ตั้งค่าแล้ว" : "❌ ยังไม่ได้ตั้งค่า",
    "GEMINI_API_KEY": scriptProperties.getProperty("GEMINI_API_KEY") ? "✅ ตั้งค่าแล้ว" : "❌ ยังไม่ได้ตั้งค่า",
    "MISTRAL_API_KEY": scriptProperties.getProperty("MISTRAL_API_KEY") ? "✅ ตั้งค่าแล้ว" : "❌ ยังไม่ได้ตั้งค่า",
    "VISION_API_KEY": scriptProperties.getProperty("VISION_API_KEY") ? "✅ ตั้งค่าแล้ว" : "❌ ยังไม่ได้ตั้งค่า"
  };
  
  return keys;
}

// ============= การตั้งค่า Config Sheet =============

/**
  ฟังก์ชันสำหรับสร้าง Config Sheet
  รันฟังก์ชันนี้หนึ่งครั้งเพื่อสร้าง Sheet สำหรับการตั้งค่า
 */
function setupConfigSheet() {
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!SHEET_ID) {
    return "❌ กรุณาตั้งค่า SHEET_ID ก่อน โดยรันฟังก์ชัน setSheetId()";
  }
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // สร้าง Sheet สำหรับบันทึกการทำรายการ
    let transactionSheet = ss.getSheetByName("Bot Transactions");
    if (!transactionSheet) {
      transactionSheet = ss.insertSheet("Bot Transactions");
      transactionSheet.appendRow(["Timestamp", "Action", "File URL", "Source"]);
    }
    
    // สร้าง Sheet สำหรับการตั้งค่า
    let configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
      configSheet = ss.insertSheet("Config");
      configSheet.appendRow(["type", "config"]);
      
      // เพิ่มค่าเริ่มต้น
      configSheet.appendRow(["ocr", JSON.stringify({
        language: "th",
        engine: "vision"
      })]);
      
      configSheet.appendRow(["ai", JSON.stringify({
        model: "gpt-4",
        prompt: "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}"
      })]);
    }
    
    return "✅ สร้าง Config Sheet เรียบร้อย";
  } catch (error) {
    return "❌ เกิดข้อผิดพลาดในการสร้าง Config Sheet: " + error;
  }
}

/**
 * ฟังก์ชันสำหรับตั้งค่า OCR
 * @param {string} language - ภาษาที่ต้องการ OCR เช่น "th", "en"
 * @param {string} engine - engine ที่ต้องการใช้ เช่น "vision", "gemini", "gpt4", "claude"
 */
function setOCRConfig(language, engine) {
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!SHEET_ID) {
    return "❌ กรุณาตั้งค่า SHEET_ID ก่อน โดยรันฟังก์ชัน setSheetId()";
  }
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName("Config");
    
    if (!configSheet) {
      return "❌ ไม่พบ Config Sheet กรุณารันฟังก์ชัน setupConfigSheet() ก่อน";
    }
    
    const ocrConfig = {
      language: language || "th",
      engine: engine || "vision"
    };
    
    // อัปเดตการตั้งค่า OCR
    updateConfig(configSheet, "ocr", ocrConfig);
    
    return `✅ ตั้งค่า OCR เรียบร้อย: ภาษา=${language}, Engine=${engine}`;
  } catch (error) {
    return "❌ เกิดข้อผิดพลาดในการตั้งค่า OCR: " + error;
  }
}

/**
 * ฟังก์ชันสำหรับตั้งค่า AI
 * @param {string} model - โมเดล AI ที่ต้องการใช้
 * @param {string} prompt - prompt ที่ต้องการใช้
 */
function setAIConfig(model, prompt) {
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!SHEET_ID) {
    return "❌ กรุณาตั้งค่า SHEET_ID ก่อน โดยรันฟังก์ชัน setSheetId()";
  }
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName("Config");
    
    if (!configSheet) {
      return "❌ ไม่พบ Config Sheet กรุณารันฟังก์ชัน setupConfigSheet() ก่อน";
    }
    
    const aiConfig = {
      model: model || "gpt-4",
      prompt: prompt || "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}"
    };
    
    // อัปเดตการตั้งค่า AI
    updateConfig(configSheet, "ai", aiConfig);
    
    return `✅ ตั้งค่า AI เรียบร้อย: Model=${model}`;
  } catch (error) {
    return "❌ เกิดข้อผิดพลาดในการตั้งค่า AI: " + error;
  }
}

/**
 * ฟังก์ชันสำหรับดูการตั้งค่าปัจจุบัน
 */
function viewCurrentConfig() {
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!SHEET_ID) {
    return "❌ กรุณาตั้งค่า SHEET_ID ก่อน โดยรันฟังก์ชัน setSheetId()";
  }
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName("Config");
    
    if (!configSheet) {
      return "❌ ไม่พบ Config Sheet กรุณารันฟังก์ชัน setupConfigSheet() ก่อน";
    }
    
    const config = getConfigFromSheet(configSheet);
    
    return JSON.stringify(config, null, 2);
  } catch (error) {
    return "❌ เกิดข้อผิดพลาดในการดูการตั้งค่า: " + error;
  }
}

// ============= ข้อมูลเพิ่มเติม =============

/**
 * ข้อมูล AI models ที่รองรับ
 */
function supportedAIModels() {
  return {
    "OpenAI": [
      "gpt-4", 
      "gpt-4o", 
      "gpt-4o-mini"
    ],
    "Anthropic": [
      "claude-3-opus-20240229", 
      "claude-3-sonnet-20240229", 
      "claude-3-haiku-20240307"
    ],
    "Google": [
      "gemini-2.0-flash", 
      "gemini-2.0-flash-lite", 
      "gemini-1.5-pro"
    ],
    "Mistral": [
      "mistral-large", 
      "mistral-medium", 
      "mistral-small"
    ]
  };
}

/**
 * ข้อมูล OCR engines ที่รองรับ
 */
function supportedOCREngines() {
  return {
    "vision": "Google Vision API (ค่าเริ่มต้น)",
    "gemini": "Gemini 1.5 Flash OCR (ราคาถูกที่สุด)",
    "gpt4": "GPT-4o Vision OCR (แม่นยำสูง)",
    "claude": "Claude 3 Opus OCR (แม่นยำสูงสุด)"
  };
}

/**
 * ข้อมูลราคาโดยประมาณ
 */
function pricingInfo() {
  return {
    "OCR": {
      "gemini": "$0.00005 ต่อหน้า (ถูกที่สุด)",
      "vision": "$0.0015 ต่อหน้า",
      "gpt4": "$0.005 ต่อหน้า",
      "claude": "$0.015 ต่อหน้า (แพงที่สุด)"
    },
    "AI": {
      "gemini-2.0-flash-lite": "$0.075 ต่อ 1M tokens input (ถูกที่สุด)",
      "gemini-2.0-flash": "$0.35 ต่อ 1M tokens input",
      "gemini-1.5-pro": "$3.50 ต่อ 1M tokens input",
      "gpt-4o-mini": "$0.15 ต่อ 1M tokens input",
      "gpt-4o": "$5.00 ต่อ 1M tokens input",
      "claude-3-haiku": "$0.25 ต่อ 1M tokens input",
      "claude-3-sonnet": "$3.00 ต่อ 1M tokens input",
      "claude-3-opus": "$15.00 ต่อ 1M tokens input (แพงที่สุด)"
    }
  };
}

// ============= ฟังก์ชั่นหลัก =============

/**
 * ฟังก์ชันหลักสำหรับรับคำขอ POST จาก Dialogflow
 */
function doPost(e) {
    // ใช้ SHEET_ID จาก Script Properties แทนการฝังในโค้ด
    const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!SHEET_ID) {
        return ContentService.createTextOutput(JSON.stringify({
            fulfillmentText: "❌ กรุณาตั้งค่า SHEET_ID ก่อนใช้งาน โดยรันฟังก์ชัน setSheetId()"
        })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Bot Transactions");
    const configSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Config");

    // อ่านการตั้งค่าจาก Config Sheet
    const config = getConfigFromSheet(configSheet);

    let data;
    try {
        data = JSON.parse(e.postData.contents);
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            fulfillmentText: "❌ ข้อมูลที่ส่งมาไม่ถูกต้อง"
        })).setMimeType(ContentService.MimeType.JSON);
    }

    const intent = data.queryResult.intent.displayName;
    const params = data.queryResult.parameters;
    const source = data.originalDetectIntentRequest.source;
    let responseText = "";

    // ตรวจสอบ intent ที่รับเข้ามา
    if (intent === "วิเคราะห์ไฟล์") {
        const fileUrl = params.fileUrl; // รับ URL ไฟล์ PDF/รูปภาพจาก Dialogflow

        if (!fileUrl) {
            return ContentService.createTextOutput(JSON.stringify({
                fulfillmentText: "❌ กรุณาส่งไฟล์ PDF หรือรูปภาพ"
            })).setMimeType(ContentService.MimeType.JSON);
        }

        // บันทึกข้อมูลการทำรายการ
        const timestamp = new Date();
        sheet.appendRow([timestamp, "วิเคราะห์ไฟล์", fileUrl, source]);

        // ดึงข้อความจากไฟล์
        const extractedText = extractTextFromFile(fileUrl, config);

        if (!extractedText) {
            return ContentService.createTextOutput(JSON.stringify({
                fulfillmentText: "❌ ไม่สามารถอ่านข้อความจากไฟล์ได้ กรุณาตรวจสอบไฟล์หรือลองใหม่อีกครั้ง"
            })).setMimeType(ContentService.MimeType.JSON);
        }

        // สร้าง prompt สำหรับ AI
        const prompt = config.aiConfig.prompt.replace("{text}", extractedText);

        // เรียกใช้ AI วิเคราะห์ข้อความ
        responseText = callAIService(prompt, config);

    } else if (intent === "ตั้งค่า_OCR") {
        // ตั้งค่า OCR
        const language = params.language || "th";
        const engine = params.engine || "vision";

        const ocrConfig = {
            language: language,
            engine: engine
        };

        // บันทึกการตั้งค่าลงใน Config Sheet
        updateConfig(configSheet, "ocr", ocrConfig);

        responseText = `✅ ตั้งค่า OCR เรียบร้อย:\nภาษา: ${language}\nEngine: ${engine}`;

    } else if (intent === "ตั้งค่า_AI") {
        // ตั้งค่า AI
        const model = params.model || "gpt-4";
        const prompt = params.prompt || "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}";

        const aiConfig = {
            model: model,
            prompt: prompt
        };

        // บันทึกการตั้งค่าลงใน Config Sheet
        updateConfig(configSheet, "ai", aiConfig);

        responseText = `✅ ตั้งค่า AI เรียบร้อย:\nโมเดล: ${model}\nPrompt: ${prompt}`;

    } else if (intent === "ดูการตั้งค่า") {
        // แสดงการตั้งค่าปัจจุบัน
        responseText = `📝 การตั้งค่าปัจจุบัน:\n\n` +
            `🔍 OCR:\n` +
            `- ภาษา: ${config.ocrConfig.language}\n` +
            `- Engine: ${config.ocrConfig.engine}\n\n` +
            `🤖 AI:\n` +
            `- โมเดล: ${config.aiConfig.model}\n` +
            `- Prompt: ${config.aiConfig.prompt}`;
    } else {
        responseText = "❓ ไม่เข้าใจคำสั่ง กรุณาลองใหม่อีกครั้ง";
    }

    return ContentService.createTextOutput(JSON.stringify({
        fulfillmentText: responseText
    })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ฟังก์ชันอ่านการตั้งค่าจาก Sheet
 */
function getConfigFromSheet(sheet) {
    const config = {
        ocrConfig: {
            language: "th",
            engine: "vision"  // ค่าเริ่มต้น
        },
        aiConfig: {
            model: "gpt-4",  // ค่าเริ่มต้น
            prompt: "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}"
        }
    };

    try {
        const data = sheet.getDataRange().getValues(); // อ่านข้อมูลจากชีต
        for (let i = 1; i < data.length; i++) {        // ข้ามหัวตาราง เริ่มจากแถวที่ 2 
            const row = data[i];
            if (row[0] === "ocr" && row[1]) {
                try {
                    // แปลง JSON string เป็น object และรวมกับค่าเริ่มต้น
                    const ocrSettings = JSON.parse(row[1]);
                    config.ocrConfig = {...config.ocrConfig, ...ocrSettings};
                } catch (error) {
                    Logger.log("Error parsing OCR config JSON: " + error);
                }
            } else if (row[0] === "ai" && row[1]) {
                try {
                    // แปลง JSON string เป็น object และรวมกับค่าเริ่มต้น
                    const aiSettings = JSON.parse(row[1]);
                    config.aiConfig = {...config.aiConfig, ...aiSettings};
                } catch (error) {
                    Logger.log("Error parsing AI config JSON: " + error);
                }
            }
        }
    } catch (e) {
        // ถ้าอ่านไม่ได้ ใช้ค่าเริ่มต้น
        Logger.log("Error reading config: " + e);
    }

    return config;
}

/**
 * ฟังก์ชันอัปเดตการตั้งค่าใน Sheet
 */
function updateConfig(sheet, type, configObj) {
    try {
        const data = sheet.getDataRange().getValues();
        let found = false;

        // ตรวจสอบว่ามีการตั้งค่าประเภทนี้อยู่แล้วหรือไม่
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === type) {
                // อัปเดตการตั้งค่าที่มีอยู่
                sheet.getRange(i + 1, 2).setValue(JSON.stringify(configObj));
                found = true;
                break;
            }
        }

        // ถ้าไม่พบการตั้งค่า ให้เพิ่มใหม่
        if (!found) {
            sheet.appendRow([type, JSON.stringify(configObj)]);
        }

        return true;
    } catch (error) {
        Logger.log("Error updating config: " + error);
        return false;
    }
}

// ============= ฟังก์ชัน OCR =============

/**
 * ฟังก์ชันดึงข้อความจาก PDF/รูปภาพโดยใช้ Google Drive API
 */
function extractTextFromFile(fileUrl, config) {
    try {
        const fileId = fileUrl.split("/d/")[1].split("/")[0]; // ดึง ID จาก URL
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();

        // ใช้การตั้งค่า OCR จาก config
        const ocrConfig = config.ocrConfig || {
            language: "th",
            engine: "vision"
        };

        if (mimeType.includes("image")) {
            return extractTextFromImage(file, ocrConfig);
        } else if (mimeType.includes("pdf")) {
            return extractTextFromPDF(file, ocrConfig);
        } else {
            return null;
        }
    } catch (error) {
        Logger.log("Error extracting text: " + error);
        return null;
    }
}

/**
 * ใช้ Google Drive API แปลง PDF เป็นข้อความ
 */
function extractTextFromPDF(file, ocrConfig) {
    // เลือก engine ตามการตั้งค่า
    const engine = ocrConfig.engine || "vision";
    
    if (engine === "vision" || engine === "document-ai") {
        const resource = {
            title: file.getName(),
            mimeType: file.getMimeType()
        };

        const options = {
            ocr: true,
            ocrLanguage: ocrConfig.language || "th"
        };

        try {
            const text = Drive.Files.insert(resource, file.getBlob(), options).title;
            return text || null;
        } catch (error) {
            Logger.log("PDF OCR error: " + error);
            return null;
        }
    } else if (engine === "gemini") {
        return geminiOCR(file, ocrConfig);
    } else if (engine === "gpt4") {
        return gpt4VisionOCR(file, ocrConfig);
    } else if (engine === "claude") {
        return claudeOCR(file, ocrConfig);
    } else {
        return "❌ ไม่รองรับ OCR engine: " + engine + " สำหรับไฟล์ PDF";
    }
}

/**
 * ใช้ Google Cloud Vision API อ่านข้อความจากรูปภาพ
 */
function extractTextFromImage(file, ocrConfig) {
    const engine = ocrConfig.engine || "vision";
    
    if (engine === "vision") {
        const apiKey = PropertiesService.getScriptProperties().getProperty("VISION_API_KEY");
        const visionUrl = "https://vision.googleapis.com/v1/images:annotate?key=" + apiKey;

        const imageBlob = file.getBlob();
        const base64Image = Utilities.base64Encode(imageBlob.getBytes());

        const payload = {
            requests: [{
                image: { content: base64Image },
                features: [{ type: "TEXT_DETECTION" }],
                imageContext: {
                    languageHints: [ocrConfig.language || "th"]
                }
            }]
        };

        const options = {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload)
        };

        try {
            const response = UrlFetchApp.fetch(visionUrl, options);
            const json = JSON.parse(response.getContentText());
            return json.responses?.[0]?.fullTextAnnotation?.text || null;
        } catch (error) {
            Logger.log("Vision API error: " + error);
            return null;
        }
    } else if (engine === "document-ai") {
        // ทางเลือกเพิ่มเติม: ใช้ Document AI (ต้องตั้งค่าเพิ่มเติม)
        const docaiKey = PropertiesService.getScriptProperties().getProperty("DOCUMENT_AI_KEY");
        const docaiProcessor = PropertiesService.getScriptProperties().getProperty("DOCUMENT_AI_PROCESSOR");

        // เพิ่มโค้ดเรียก Document AI API ตรงนี้
        // ...

        return "ใช้ Document AI: ยังไม่ได้ตั้งค่า";
    } else if (engine === "gemini") {
        return geminiOCR(file, ocrConfig);
    } else if (engine === "gpt4") {
        return gpt4VisionOCR(file, ocrConfig);
    } else if (engine === "claude") {
        return claudeOCR(file, ocrConfig);
    } else {
        return "❌ ไม่รองรับ OCR engine: " + engine;
    }
}

/**
 * ฟังก์ชันสำหรับ Gemini OCR
 */
function geminiOCR(file, ocrConfig) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) return "❌ ไม่พบ Gemini API Key";
    
    // ใช้ Gemini 1.5 Flash สำหรับ OCR (ราคาถูกที่สุด)
    const model = "gemini-1.5-flash";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    const imageBlob = file.getBlob();
    const base64Image = Utilities.base64Encode(imageBlob.getBytes());
    
    const language = ocrConfig.language || "th";
    const systemPrompt = `คุณเป็นระบบ OCR ที่แม่นยำ กรุณาอ่านข้อความทั้งหมดจากภาพนี้ในภาษา${language === "th" ? "ไทย" : language} และส่งคืนเฉพาะข้อความที่อ่านได้ โดยรักษารูปแบบการจัดวางให้ใกล้เคียงกับต้นฉบับ`;
    
    const payload = {
        contents: [
            {
                parts: [
                    { text: systemPrompt },
                    {
                        inline_data: {
                            mime_type: file.getMimeType(),
                            data: base64Image
                        }
                    }
                ]
            }
        ],
        generationConfig: {
            maxOutputTokens: 2048
        }
    };
    
    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Gemini OCR";
    } catch (error) {
        Logger.log("Gemini OCR error: " + error);
        return "🚨 เกิดข้อผิดพลาด Gemini OCR: " + error.message;
    }
}

/**
 * ฟังก์ชันสำหรับ GPT-4 Vision OCR
 */
function gpt4VisionOCR(file, ocrConfig) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if (!apiKey) return "❌ ไม่พบ OpenAI API Key";
    
    // ใช้ GPT-4o สำหรับ OCR
    const model = "gpt-4o";
    const url = "https://api.openai.com/v1/chat/completions";
    
    const imageBlob = file.getBlob();
    const base64Image = Utilities.base64Encode(imageBlob.getBytes());
    
    const language = ocrConfig.language || "th";
    const systemPrompt = `คุณเป็นระบบ OCR ที่แม่นยำ กรุณาอ่านข้อความทั้งหมดจากภาพนี้ในภาษา${language === "th" ? "ไทย" : language} และส่งคืนเฉพาะข้อความที่อ่านได้ โดยรักษารูปแบบการจัดวางให้ใกล้เคียงกับต้นฉบับ`;
    
    const payload = {
        model: model,
        messages: [
            {
                role: "system",
                content: systemPrompt
            },
            {
                role: "user",
                content: [
                    {
                        type: "image_url",
                        image_url: {
                            url: `data:${file.getMimeType()};base64,${base64Image}`
                        }
                    }
                ]
            }
        ],
        max_tokens: 1024
    };
    
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "Authorization": "Bearer " + apiKey,
            "Content-Type": "application/json"
        },
        payload: JSON.stringify(payload)
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.choices?.[0]?.message?.content || "❌ ไม่พบผลลัพธ์จาก GPT-4 Vision OCR";
    } catch (error) {
        Logger.log("GPT-4 Vision OCR error: " + error);
        return "🚨 เกิดข้อผิดพลาด GPT-4 Vision OCR: " + error.message;
    }
}

/**
 * ฟังก์ชันสำหรับ Claude OCR
 */
function claudeOCR(file, ocrConfig) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
    if (!apiKey) return "❌ ไม่พบ Anthropic API Key";
    
    // ใช้ Claude 3 Opus สำหรับ OCR (แม่นยำสูงสุด)
    const model = "claude-3-opus-20240229";
    const url = "https://api.anthropic.com/v1/messages";
    
    const imageBlob = file.getBlob();
    const base64Image = Utilities.base64Encode(imageBlob.getBytes());
    
    const language = ocrConfig.language || "th";
    const systemPrompt = `คุณเป็นระบบ OCR ที่แม่นยำ กรุณาอ่านข้อความทั้งหมดจากภาพนี้ในภาษา${language === "th" ? "ไทย" : language} และส่งคืนเฉพาะข้อความที่อ่านได้ โดยรักษารูปแบบการจัดวางให้ใกล้เคียงกับต้นฉบับ`;
    
    const payload = {
        model: model,
        messages: [
            {
                role: "user",
                content: [
                    {
                        type: "text",
                        text: systemPrompt
                    },
                    {
                        type: "image",
                        source: {
                            type: "base64",
                            media_type: file.getMimeType(),
                            data: base64Image
                        }
                    }
                ]
            }
        ],
        max_tokens: 1024
    };
    
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        },
        payload: JSON.stringify(payload)
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.content?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Claude OCR";
    } catch (error) {
        Logger.log("Claude OCR error: " + error);
        return "🚨 เกิดข้อผิดพลาด Claude OCR: " + error.message;
    }
}

// ============= ฟังก์ชัน AI =============

/**
 * ฟังก์ชันเรียก AI วิเคราะห์ข้อมูล
 */
function callAIService(prompt, config) {
    // ใช้การตั้งค่า AI จาก config
    const aiConfig = config.aiConfig || {
        model: "gpt-4",
        prompt: "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}"
    };
    
    const model = aiConfig.model || "gpt-4";
    
    // เลือก AI service ตามโมเดล
    if (model.includes("gpt")) {
        return callOpenAI(model, prompt);
    } else if (model.includes("claude")) {
        return callAnthropic(model, prompt);
    } else if (model.includes("gemini")) {
        return callGemini(model, prompt);
    } else if (model.includes("mistral")) {
        return callMistral(model, prompt);
    } else {
        return "❌ ไม่รองรับ AI model: " + model;
    }
}

/**
 * ฟังก์ชันเรียกใช้ OpenAI API
 */
function callOpenAI(model, prompt) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if (!apiKey) return "❌ ไม่พบ OpenAI API Key";
    
    const url = "https://api.openai.com/v1/chat/completions";
    
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "Authorization": "Bearer " + apiKey,
            "Content-Type": "application/json"
        },
        payload: JSON.stringify({
            model: model,
            messages: [
                { role: "system", content: "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน" },
                { role: "user", content: prompt }
            ],
            max_tokens: 1024
        })
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.choices?.[0]?.message?.content || "❌ ไม่พบผลลัพธ์จาก OpenAI";
    } catch (error) {
        Logger.log("OpenAI error: " + error);
        return "🚨 เกิดข้อผิดพลาด OpenAI: " + error.message;
    }
}

/**
 * ฟังก์ชันเรียกใช้ Anthropic API
 */
function callAnthropic(model, prompt) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
    if (!apiKey) return "❌ ไม่พบ Anthropic API Key";
    
    const url = "https://api.anthropic.com/v1/messages";
    
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        },
        payload: JSON.stringify({
            model: model,
            system: "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน",
            messages: [
                { role: "user", content: prompt }
            ],
            max_tokens: 1024
        })
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.content?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Anthropic";
    } catch (error) {
        Logger.log("Anthropic error: " + error);
        return "🚨 เกิดข้อผิดพลาด Anthropic: " + error.message;
    }
}

/**
 * ฟังก์ชันเรียกใช้ Google Gemini API
 */
function callGemini(model, prompt) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) return "❌ ไม่พบ Gemini API Key";
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
            contents: [
                {
                    parts: [
                        { text: "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน" }
                    ],
                    role: "system"
                },
                {
                    parts: [
                        { text: prompt }
                    ],
                    role: "user"
                }
            ],
            generationConfig: {
                maxOutputTokens: 1024
            }
        })
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Gemini";
    } catch (error) {
        Logger.log("Gemini error: " + error);
        return "🚨 เกิดข้อผิดพลาด Gemini: " + error.message;
    }
}

/**
 * ฟังก์ชันเรียกใช้ Mistral API
 */
function callMistral(model, prompt) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("MISTRAL_API_KEY");
    if (!apiKey) return "❌ ไม่พบ Mistral API Key";
    
    const url = "https://api.mistral.ai/v1/chat/completions";
    
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "Authorization": "Bearer " + apiKey,
            "Content-Type": "application/json"
        },
        payload: JSON.stringify({
            model: model,
            messages: [
                { role: "system", content: "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน" },
                { role: "user", content: prompt }
            ],
            max_tokens: 1024
        })
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.choices?.[0]?.message?.content || "❌ ไม่พบผลลัพธ์จาก Mistral";
    } catch (error) {
        Logger.log("Mistral error: " + error);
        return "🚨 เกิดข้อผิดพลาด Mistral: " + error.message;
    }
}

// ============= ตัวอย่างการใช้งาน =============

/**
 * ตัวอย่างการใช้งานทั้งหมด
 * ฟังก์ชันนี้แสดงตัวอย่างการใช้งานทั้งระบบ
 */
function exampleUsage() {
  // 1. ตรวจสอบ SHEET_ID
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!SHEET_ID) {
    return "❌ กรุณาตั้งค่า SHEET_ID ก่อน โดยรันฟังก์ชัน setSheetId()";
  }
  
  // 2. เปิด Sheet และอ่านการตั้งค่า
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName("Config");
  const config = getConfigFromSheet(configSheet);
  
  // 3. สมมติว่ามีไฟล์ URL
  const fileUrl = "https://drive.google.com/file/d/1AbCdEfGhIjKlMnOpQrStUvWxYz/view";
  
  // 4. ดึงข้อความจากไฟล์ด้วย OCR
  // ในสถานการณ์จริง คุณจะใช้ extractTextFromFile(fileUrl, config)
  const extractedText = "นี่คือข้อความที่สกัดได้จากไฟล์ด้วย OCR engine: " + config.ocrConfig.engine;
  
  // 5. สร้าง prompt สำหรับ AI
  const prompt = config.aiConfig.prompt.replace("{text}", extractedText);
  
  // 6. เรียกใช้ AI วิเคราะห์ข้อความ
  // ในสถานการณ์จริง คุณจะใช้ callAIService(prompt, config)
  const aiResponse = "นี่คือผลการวิเคราะห์จาก AI model: " + config.aiConfig.model;
  
  // 7. บันทึกข้อมูลการทำรายการ
  const transactionSheet = ss.getSheetByName("Bot Transactions");
  const timestamp = new Date();
  transactionSheet.appendRow([timestamp, "วิเคราะห์ไฟล์", fileUrl, "example"]);
  
  return {
    config: config,
    extractedText: extractedText,
    prompt: prompt,
    aiResponse: aiResponse
  };
}
