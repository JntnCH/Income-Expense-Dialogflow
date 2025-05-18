function doPost(e) {
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

        // 1. ดึงข้อความจาก PDF/รูปภาพ
        const extractedText = extractTextFromFile(fileUrl, config.ocrConfig);

        if (!extractedText) {
            return ContentService.createTextOutput(JSON.stringify({
                fulfillmentText: "❌ ไม่สามารถอ่านข้อมูลจากไฟล์ได้"
            })).setMimeType(ContentService.MimeType.JSON);
        }

        // 2. ให้ AI วิเคราะห์ข้อความ
        const aiAnalysis = callAIService(
            config.aiConfig.model,
            config.aiConfig.prompt.replace("{text}", extractedText)
        );

        // 3. บันทึกลง Google Sheets
        const date = new Date();
        const formattedDate = Utilities.formatDate(date, "Asia/Bangkok", "dd MMM yyyy");
        sheet.appendRow([formattedDate, "analyzed", "-", aiAnalysis, "AI Analysis", source]);

        responseText = `✅ ฉันได้วิเคราะห์ไฟล์ให้แล้ว!\n📄 ข้อมูลที่ดึงมา: ${extractedText}\n🧠 ผลลัพธ์ AI: ${aiAnalysis}`;

    } else if (intent === "ตั้งค่า_OCR") {
        // ตั้งค่า OCR จากแชท
        const ocrLanguage = params.ocrLanguage || "th";
        const ocrEngine = params.ocrEngine || "vision";

        updateConfig(configSheet, "ocr", {
            language: ocrLanguage,
            engine: ocrEngine
        });

        responseText = `✅ ตั้งค่า OCR เรียบร้อย!\n🔤 ภาษา: ${ocrLanguage}\n🔧 Engine: ${ocrEngine}`;

    } else if (intent === "ตั้งค่า_AI") {
        // ตั้งค่า AI จากแชท
        const aiModel = params.aiModel || "gpt-4";
        const aiPrompt = params.aiPrompt || "วิเคราะห์ข้อมูลการเงินจากข้อความนี้:\n\n{text}";

        updateConfig(configSheet, "ai", {
            model: aiModel,
            prompt: aiPrompt
        });

        responseText = `✅ ตั้งค่า AI เรียบร้อย!\n🤖 โมเดล: ${aiModel}\n📝 พรอมพ์: ${aiPrompt}`;

    } else if (intent === "ดูการตั้งค่า") {
        // แสดงการตั้งค่าปัจจุบัน
        responseText = `🔍 การตั้งค่าปัจจุบัน:\n\n` +
                       `📋 OCR:\n- ภาษา: ${config.ocrConfig.language}\n- Engine: ${config.ocrConfig.engine}\n\n` +
                       `🤖 AI:\n- โมเดล: ${config.aiConfig.model}\n- พรอมพ์: ${config.aiConfig.prompt}`;

    } else {
        responseText = "❌ คำสั่งนี้ไม่รองรับ";
    }

    return ContentService.createTextOutput(JSON.stringify({
        fulfillmentText: responseText
    })).setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันอ่านการตั้งค่าจาก Sheet
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

// ฟังก์ชันสำหรับเรียกใช้ AI ตามโมเดลที่กำหนด
function callAI(prompt, config) {
    // ตรวจสอบโมเดลที่ตั้งค่าไว้
    const model = config.aiConfig.model || "gpt-4"; // ค่าเริ่มต้นถ้าไม่ได้ระบุ
    const apiKey = getApiKey(model);
    
    // เรียกใช้ฟังก์ชันตามโมเดลที่เลือก
    if (model.includes("gpt")) {
        return callOpenAI(model, prompt, apiKey);
    } else if (model.includes("claude")) {
        return callAnthropic(model, prompt, apiKey);
    } else if (model.includes("gemini")) {
        return callGemini(model, prompt, apiKey);
    } else if (model.includes("mistral")) {
        return callMistral(model, prompt, apiKey);
    } else {
        return "❌ ไม่รองรับโมเดล AI ที่ระบุ";
    }
}

// ฟังก์ชันสำหรับดึง API key ตามโมเดล
function getApiKey(model) {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    if (model.includes("gpt")) {
        return scriptProperties.getProperty("OPENAI_API_KEY");
    } else if (model.includes("claude")) {
        return scriptProperties.getProperty("ANTHROPIC_API_KEY");
    } else if (model.includes("gemini")) {
        return scriptProperties.getProperty("GEMINI_API_KEY");
    } else if (model.includes("mistral")) {
        return scriptProperties.getProperty("MISTRAL_API_KEY");
    }
    
    return null;
}

// ฟังก์ชันสำหรับทำ OCR ตาม engine ที่กำหนด
function performOCR(fileUrl, config) {
    const engine = config.ocrConfig.engine || "vision"; // ค่าเริ่มต้นถ้าไม่ได้ระบุ
    const language = config.ocrConfig.language || "th"; // ค่าเริ่มต้นถ้าไม่ได้ระบุ
    
    // เรียกใช้ OCR ตาม engine ที่เลือก
    if (engine === "vision") {
        return googleVisionOCR(fileUrl, language);
    } else if (engine === "gemini") {
        return geminiOCR(fileUrl);
    } else if (engine === "gpt4") {
        return gpt4VisionOCR(fileUrl);
    } else if (engine === "claude") {
        return claudeOCR(fileUrl);
    } else {
        return "❌ ไม่รองรับ OCR engine ที่ระบุ";
    }
}

// 📌 ฟังก์ชันดึงข้อความจาก PDF/รูปภาพโดยใช้ Google Drive API
function extractTextFromFile(fileUrl, ocrConfig) {
    try {
        const fileId = fileUrl.split("/d/")[1].split("/")[0]; // ดึง ID จาก URL
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();

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

// 📌 ใช้ Google Drive API แปลง PDF เป็นข้อความ
function extractTextFromPDF(file, ocrConfig) {
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
}

// 📌 ใช้ Google Cloud Vision API อ่านข้อความจากรูปภาพ
function extractTextFromImage(file, ocrConfig) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("VISION_API_KEY");

    // เลือกตัวเลือก OCR ตามการตั้งค่า
    if (ocrConfig.engine === "vision") {
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
    } else if (ocrConfig.engine === "document-ai") {
        // ทางเลือกเพิ่มเติม: ใช้ Document AI (ต้องตั้งค่าเพิ่มเติม)
        const docaiKey = PropertiesService.getScriptProperties().getProperty("DOCUMENT_AI_KEY");
        const docaiProcessor = PropertiesService.getScriptProperties().getProperty("DOCUMENT_AI_PROCESSOR");

        // เพิ่มโค้ดเรียก Document AI API ตรงนี้
        // ...

        return "ใช้ Document AI: ยังไม่ได้ตั้งค่า";
    }

    return null;
}

// 📌 ฟังก์ชันเรียก AI วิเคราะห์ข้อมูล
function callAIService(model, prompt) {
    const aiService = PropertiesService.getScriptProperties().getProperty("AI_SERVICE");
    const apiKey = PropertiesService.getScriptProperties().getProperty("AI_API_KEY");

    if (!apiKey) return '❌ API Key หาย ตรวจสอบการตั้งค่า';

    // เลือก AI service ตามการตั้งค่า
    if (aiService === "openai" || aiService === "auto") {
        return callOpenAI(model, prompt, apiKey);
    } else if (aiService === "anthropic") {
        return callAnthropic(model, prompt, apiKey);
    } else if (aiService === "gemini") {
        return callGemini(model, prompt, apiKey);
    } else {
        return "❌ ไม่รองรับ AI Service: " + aiService;
    }
}

function callOpenAI(model, prompt, apiKey) {
    const url = "https://api.openai.com/v1/chat/completions";

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'Authorization': 'Bearer ' + apiKey,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
            model: model,
            messages: [
                { role: "system", content: "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน" },
                { role: "user", content: prompt }
            ],
            max_tokens: 500
        })
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());

        return jsonResponse.choices?.[0]?.message?.content || "❌ ไม่พบผลลัพธ์จาก OpenAI";
    } catch (error) {
        return '🚨 เกิดข้อผิดพลาด OpenAI: ' + error.message;
    }
}

function callAnthropic(model, prompt, apiKey) {
    const url = "https://api.anthropic.com/v1/messages";

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01',
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
            model: model,
            messages: [
                { role: "user", content: prompt }
            ],
            max_tokens: 500
        })
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());

        return jsonResponse.content?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Anthropic";
    } catch (error) {
        return '🚨 เกิดข้อผิดพลาด Anthropic: ' + error.message;
    }
}

function callGemini(model, prompt, apiKey) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
            contents: [
                {
                    parts: [
                        { text: prompt }
                    ]
                }
            ],
            generationConfig: {
                maxOutputTokens: 500
            }
        })
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());

        return jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text || "❌ ไม่พบผลลัพธ์จาก Gemini";
    } catch (error) {
        return '🚨 เกิดข้อผิดพลาด Gemini: ' + error.message;
    }
}