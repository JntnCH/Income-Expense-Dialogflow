/**
 * บอทวิเคราะห์เอกสารอัตโนมัติ
 * รองรับการวิเคราะห์ไฟล์ PDF และรูปภาพด้วย OCR และ AI
 * สนับสนุนบริการ AI และ OCR หลากหลาย: OpenAI, Anthropic, Google, Mistral
 */

// ค่าคงที่
const SHEET_ID = "YOUR_SHEET_ID"; // แทนที่ด้วย ID ของ Google Sheet ของคุณ

/**
 * ฟังก์ชันหลักสำหรับรับคำขอ POST จาก Dialogflow
 */
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

/**
 * ฟังก์ชันเรียก AI วิเคราะห์ข้อมูล
 */
function callAIService(prompt, config) {
    // ใช้การตั้งค่า AI จาก config
    const aiConfig = config.aiConfig || {
        model: "gpt-4",
        prompt: "วิเคราะห์ข้อมูลจากข้อความนี้:\n\n{text}"
    };
    
    const model = aiConfig.model || "gpt-4";
    
    // ดึง API key ตามโมเดลที่เลือก
    const apiKey = getApiKey(model);
    if (!apiKey) return '❌ ไม่พบ API Key สำหรับโมเดล ' + model;
    
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
        return "❌ ไม่รองรับโมเดล AI: " + model;
    }
}

/**
 * ฟังก์ชันสำหรับดึง API key ตามโมเดล
 */
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

/**
 * ฟังก์ชันเรียก OpenAI
 */
function callOpenAI(model, prompt, apiKey) {
    const url = "https://api.openai.com/v1/chat/completions";
    
    // ปรับ system prompt ตามโมเดล
    let systemPrompt = "คุณเป็น AI ที่เชี่ยวชาญในการวิเคราะห์ข้อมูลการเงิน";
    
    // ปรับ max_tokens ตามโมเดล
    let maxTokens = 500;
    if (model.includes("gpt-4")) {
        maxTokens = 1000; // เพิ่ม token สำหรับโมเดลที่มีความสามารถสูง
    }

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
                { role: "system", content: systemPrompt },
                { role: "user", content: prompt }
            ],
            max_tokens: maxTokens
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

/**
 * ฟังก์ชันเรียก Anthropic
 */
function callAnthropic(model, prompt, apiKey) {
    const url = "https://api.anthropic.com/v1/messages";
    
    // ปรับ max_tokens ตามโมเดล
    let maxTokens = 500;
    if (model.includes("opus")) {
        maxTokens = 1000; // เพิ่ม token สำหรับโมเดลที่มีความสามารถสูง
    }

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
            max_tokens: maxTokens
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

/**
 * ฟังก์ชันเรียก Gemini
 */
function callGemini(model, prompt, apiKey) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    // ปรับ max_tokens ตามโมเดล
    let maxTokens = 500;
    if (model.includes("pro")) {
        maxTokens = 1000; // เพิ่ม token สำหรับโมเดลที่มีความสามารถสูง
    }

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
                maxOutputTokens: maxTokens
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

/**
 * ฟังก์ชันเรียก Mistral
 */
function callMistral(model, prompt, apiKey) {
    const url = "https://api.mistral.ai/v1/chat/completions";
    
    // ปรับ max_tokens ตามโมเดล
    let maxTokens = 500;
    if (model.includes("large")) {
        maxTokens = 1000; // เพิ่ม token สำหรับโมเดลที่มีความสามารถสูง
    }

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
                { role: "user", content: prompt }
            ],
            max_tokens: maxTokens
        })
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const jsonResponse = JSON.parse(response.getContentText());

        return jsonResponse.choices?.[0]?.message?.content || "❌ ไม่พบผลลัพธ์จาก Mistral";
    } catch (error) {
        return '🚨 เกิดข้อผิดพลาด Mistral: ' + error.message;
    }
}

/**
 * ฟังก์ชันสำหรับตั้งค่า API Keys
 * ใช้สำหรับตั้งค่า API Keys ครั้งแรก
 */
function setupAPIKeys() {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // ตั้งค่า API Keys (แทนที่ด้วย API Keys ของคุณ)
    scriptProperties.setProperty("OPENAI_API_KEY", "YOUR_OPENAI_API_KEY");
    scriptProperties.setProperty("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY");
    scriptProperties.setProperty("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY");
    scriptProperties.setProperty("MISTRAL_API_KEY", "YOUR_MISTRAL_API_KEY");
    scriptProperties.setProperty("VISION_API_KEY", "YOUR_VISION_API_KEY");
    
    return "✅ ตั้งค่า API Keys เรียบร้อย";
}

/**
 * ฟังก์ชันสำหรับสร้าง Sheet ครั้งแรก
 */
function setupSheets() {
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
        
        return "✅ สร้าง Sheets เรียบร้อย";
    } catch (error) {
        return "❌ เกิดข้อผิดพลาดในการสร้าง Sheets: " + error;
    }
}

/**
 * ฟังก์ชันสำหรับทดสอบการทำงาน
 */
function testBot() {
    // ทดสอบการอ่านการตั้งค่า
    const configSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Config");
    const config = getConfigFromSheet(configSheet);
    
    return JSON.stringify(config, null, 2);
}
