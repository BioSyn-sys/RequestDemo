/**
 * Google Apps Script Web App Backend for Medical Device Trial Forms
 * 
 * ระบบหลังบ้านรวม (Google Apps Script) สำหรับบันทึกข้อมูลแบบฟอร์มเครื่องมือแพทย์ทั้ง 3 ฟอร์ม
 * พร้อมการอัปโหลดไฟล์ และระบบแจ้งเตือนผ่าน Email และ Telegram Bot
 * ปลอดภัยสูง: ปกป้องการส่งและดึงข้อมูลด้วยระบบตรวจสอบคีย์ลับ (API Key Check)
 */

// ==========================================
// ส่วนกำหนดค่าปรับแต่ง (Configuration)
// ==========================================
const MAIN_FOLDER_ID = "1XSwGTD3SjK5WP0lURXA4F_8bE6ybeaTw"; // โฟลเดอร์หลัก Google Drive สำหรับเก็บเอกสารทั้งหมด
const ADMIN_EMAIL = "your_email@domain.com";             // เปลี่ยนเป็นอีเมลของแอดมินหรือหน่วยงานวิศวกรรมการแพทย์ (BME)

// คีย์ลับป้องกันระบบหลังบ้าน (ห้ามให้ผู้อื่นทราบ)
const API_KEY = "BioSyn-Secure-Key-2026"; // ⚠️ คีย์ความปลอดภัย ป้องกันการเจาะเข้าระบบหลังบ้านโดยไม่ได้รับอนุญาต

// ตั้งค่า Telegram Bot
const TELEGRAM_BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN_HERE"; // เปลี่ยนเป็น Token ของ Telegram Bot ของคุณ
const TELEGRAM_CHAT_ID = "YOUR_TELEGRAM_CHAT_ID_HERE";     // เปลี่ยนเป็น Chat ID หรือ Group ID ที่ต้องการรับการแจ้งเตือน

// ==========================================
// ฟังก์ชันส่งแจ้งเตือนผ่าน Telegram Bot
// ==========================================
function sendTelegramNotification(text) {
  if (!TELEGRAM_BOT_TOKEN || TELEGRAM_BOT_TOKEN === "YOUR_TELEGRAM_BOT_TOKEN_HERE" || !TELEGRAM_CHAT_ID) {
    Logger.log("ไม่ได้ตั้งค่า Telegram Credentials หรือใช้ค่าเริ่มต้น");
    return;
  }
  const url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
  const payload = {
    "chat_id": TELEGRAM_CHAT_ID,
    "text": text
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  try {
    UrlFetchApp.fetch(url, options);
  } catch (err) {
    Logger.log("ไม่สามารถส่งข้อความไปที่ Telegram ได้: " + err.toString());
  }
}

// ==========================================
// ฟังก์ชันหลักรับข้อมูล GET Request (ดึงข้อมูลสำหรับ Dashboard)
// ==========================================
function doGet(e) {
  try {
    const apiKey = e.parameter.apiKey;
    
    // ตรวจสอบความถูกต้องของ API Key ในคำขอ GET
    if (!apiKey || apiKey !== API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'ปฏิเสธการเชื่อมต่อ: ไม่มีสิทธิ์การเข้าถึงข้อมูลระบบ (Unauthorized Access)'
      })).setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', '*');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // ดึงข้อมูลในรูปแบบ JSON Array จากแต่ละแผ่นงาน
    const requestsData = getSheetDataAsJson(spreadsheet, 'requests_db');
    const evaluationsData = getSheetDataAsJson(spreadsheet, 'evaluations_db');
    const pacsData = getSheetDataAsJson(spreadsheet, 'pacs_disclosures_db');
    const uploadsData = getSheetDataAsJson(spreadsheet, 'additional_uploads_db');
    
    const result = {
      status: 'success',
      timestamp: new Date().toISOString(),
      requests: requestsData,
      evaluations: evaluationsData,
      pacs: pacsData,
      additional_uploads: uploadsData
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
}

// ฟังก์ชันดึงข้อมูลจากแต่ละชีตแปลงเป็น JSON Array
function getSheetDataAsJson(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];
  
  const headers = values[0];
  const jsonArray = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[j]] = val;
    }
    jsonArray.push(obj);
  }
  return jsonArray;
}

// ==========================================
// ฟังก์ชันหลักรับข้อมูล POST Request
// ==========================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // ตรวจสอบความถูกต้องของ API Key ในการยื่นส่งแบบฟอร์ม POST
    if (!data.apiKey || data.apiKey !== API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'ปฏิเสธการเชื่อมต่อ: สิทธิ์เข้าใช้งานไม่ถูกต้อง (Unauthorized Request)'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const formType = data.formType; // 'request', 'evaluation', 'pacs' หรือ 'additional_upload'
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (formType === 'request') {
      return handleRequestForm(data, spreadsheet);
    } else if (formType === 'evaluation') {
      return handleEvaluationForm(data, spreadsheet);
    } else if (formType === 'pacs') {
      return handlePacsForm(data, spreadsheet);
    } else if (formType === 'additional_upload') {
      return handleAdditionalUpload(data, spreadsheet);
    } else {
      throw new Error('ไม่พบประเภทฟอร์มที่ถูกต้อง (formType)');
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// ฟังก์ชันตัวช่วยดึงรหัสวันเวลาในรูปแบบภาษาไทย (วัน เดือน ปี เวลา)
// ==========================================
function getThaiFormattedDateTime() {
  const now = new Date();
  const thaiYear = now.getFullYear() + 543; // ปี พ.ศ.
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${day}-${month}-${thaiYear}_${hours}-${minutes}-${seconds}`;
}

// ==========================================
// 1. จัดการฟอร์มยื่นคำขออนุญาตนำเครื่องมือแพทย์เข้า (Form 1)
// ==========================================
function handleRequestForm(data, spreadsheet) {
  const sheet = getOrCreateSheet(spreadsheet, 'requests_db');
  
  // 1.1 สร้างโฟลเดอร์เก็บข้อมูลรายโครงการ [ID] [CompanyName] [วัน เดือน ปี เวลา]
  let projectFolder = null;
  if (MAIN_FOLDER_ID) {
    try {
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      const formattedDate = getThaiFormattedDateTime();
      const folderName = `${data.submissionId} ${data.companyName} (${formattedDate})`;
      projectFolder = mainFolder.createFolder(folderName);
    } catch (fError) {
      Logger.log("ไม่สามารถสร้างโฟลเดอร์ Google Drive ได้: " + fError.toString());
    }
  }

  // 1.2 วนลูปตรวจสอบเอกสารแนบ 9 รายการ
  const fileUrls = {};
  const uploadedFileNames = [];
  
  for (let i = 1; i <= 9; i++) {
    const fileData = data[`file${i}`];
    if (fileData && fileData.base64 && projectFolder) {
      try {
        const decodedBytes = Utilities.base64Decode(fileData.base64.split(',')[1]);
        const blob = Utilities.newBlob(decodedBytes, fileData.type || 'application/pdf', fileData.name);
        const newFile = projectFolder.createFile(blob);
        fileUrls[`doc${i}_url`] = newFile.getUrl();
        uploadedFileNames.push(fileData.name);
      } catch (uploadError) {
        fileUrls[`doc${i}_url`] = "อัปโหลดล้มเหลว: " + uploadError.toString();
      }
    } else {
      fileUrls[`doc${i}_url`] = fileData ? "ล้มเหลว (ไม่ได้กำหนดสิทธิ์โฟลเดอร์)" : "";
    }
  }

  // 1.3 สร้างบทสรุปสั้นแต่ครบถ้วน (Summary Text)
  const summaryText = 
    `📝 สรุปคำขออนุญาตนำเครื่องมือแพทย์มาทดลองใช้\n` +
    `-----------------------------------------\n` +
    `• รหัสคำขออ้างอิง: ${data.submissionId}\n` +
    `• บริษัท: ${data.companyName} (ผู้ติดต่อ: ${data.repName} โทร: ${data.repPhone})\n` +
    `• เครื่องมือแพทย์: ${data.deviceName} ยี่ห้อ ${data.deviceBrand} รุ่น ${data.deviceModel} (ผลิตจาก: ${data.deviceCountry})\n` +
    `• แผนกทดลองใช้: ${data.userDepartment} (ผู้ประเมินหลัก: ${data.staffInCharge})\n` +
    `• ระยะเวลา: ${data.testDuration} วัน (ระหว่าง ${data.testStart} ถึง ${data.testEnd})\n` +
    `• อุปกรณ์นำเข้าประกอบด้วย: ${data.itemsList}\n` +
    `• เอกสารแนบที่ได้รับแล้ว: ${uploadedFileNames.length > 0 ? uploadedFileNames.join(', ') : 'ไม่มีเอกสารแนบ'}\n` +
    `-----------------------------------------`;

  // 1.4 บันทึกแถวข้อมูลลง requests_db
  sheet.appendRow([
    new Date(),
    data.submissionId,
    data.companyName,
    data.companyAddress,
    data.repName,
    data.repPhone,
    data.deviceName,
    data.deviceBrand,
    data.deviceModel,
    data.deviceCountry,
    data.deviceSpecs,
    data.devicePurpose,
    data.refHospital,
    data.userDepartment,
    data.staffInCharge,
    data.itemsList,
    data.testDuration,
    data.testStart,
    data.testEnd,
    fileUrls.doc1_url || "",
    fileUrls.doc2_url || "",
    fileUrls.doc3_url || "",
    fileUrls.doc4_url || "",
    fileUrls.doc5_url || "",
    fileUrls.doc6_url || "",
    fileUrls.doc7_url || "",
    fileUrls.doc8_url || "",
    fileUrls.doc9_url || "",
    summaryText
  ]);

  // 1.5 ส่งแจ้งเตือนแอดมินทางอีเมล
  if (ADMIN_EMAIL && ADMIN_EMAIL !== "your_email@domain.com") {
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: `[คำขอใหม่] ขออนุญาตนำเครื่องมือแพทย์ทดลองใช้ - ${data.companyName} (${data.submissionId})`,
        body: summaryText + `\n\nสามารถตรวจสอบรายการใน Google Sheet`
      });
    } catch (eError) {
      Logger.log("ไม่สามารถส่งอีเมลได้: " + eError.toString());
    }
  }

  // 1.6 ส่งแจ้งเตือนทาง Telegram
  sendTelegramNotification(summaryText);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'บันทึกคำขอและสร้างโฟลเดอร์เก็บเอกสารสำเร็จ',
    submissionId: data.submissionId,
    summary: summaryText
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. จัดการข้อมูลผลประเมินการใช้งานเครื่องมือแพทย์ (Form 2)
// ==========================================
function handleEvaluationForm(data, spreadsheet) {
  const sheet = getOrCreateSheet(spreadsheet, 'evaluations_db');
  
  sheet.appendRow([
    new Date(),
    data.evaluationId,
    data.companyName,
    data.deviceName,
    data.deviceBrand,
    data.deviceModel,
    data.deviceCountry,
    data.userDepartment,
    data.staffInCharge,
    data.score_q1,
    data.score_q2,
    data.score_q3,
    data.score_q4,
    data.score_q5,
    data.score_q6,
    data.score_q7,
    data.pros,
    data.cons,
    data.compareBrand,
    data.compareModel,
    data.compareDetails,
    data.suggestions
  ]);

  const avgScore = ((data.score_q1 + data.score_q2 + data.score_q3 + data.score_q4 + data.score_q5 + data.score_q6 + data.score_q7) / 7).toFixed(2);
  const evalMsg = `📊 ผลประเมินเครื่องมือแพทย์ใหม่\n` +
                  `-----------------------------------------\n` +
                  `• เครื่องมือ: ${data.deviceName} (ยี่ห้อ/รุ่น: ${data.deviceBrand} / ${data.deviceModel})\n` +
                  `• แผนกประเมิน: ${data.userDepartment}\n` +
                  `• ผู้ประเมินหลัก: ${data.staffInCharge}\n` +
                  `• คะแนนเฉลี่ยความพึงพอใจ: ${avgScore} / 5.00 คะแนน\n` +
                  `• ข้อเสนอแนะเพิ่มเติม: ${data.suggestions || 'ไม่มี'}\n` +
                  `-----------------------------------------`;

  // 2.1 ส่งแจ้งเตือนทางอีเมลแอดมิน
  if (ADMIN_EMAIL && ADMIN_EMAIL !== "your_email@domain.com") {
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: `[ผลประเมินใหม่] แบบประเมินเครื่องมือแพทย์ - ${data.deviceName}`,
        body: evalMsg
      });
    } catch (eError) {
      Logger.log("ไม่สามารถส่งอีเมลประเมินได้: " + eError.toString());
    }
  }

  // 2.2 ส่งแจ้งเตือนทาง Telegram
  sendTelegramNotification(evalMsg);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'บันทึกคะแนนการประเมินสำเร็จ'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 3. จัดการหนังสือขออนุญาตเชื่อมระบบและเปิดเผยข้อมูล PACS (Form 3)
// ==========================================
function handlePacsForm(data, spreadsheet) {
  const sheet = getOrCreateSheet(spreadsheet, 'pacs_disclosures_db');
  
  sheet.appendRow([
    new Date(),
    data.pacsId,
    data.requestDate,
    data.subjectTarget,
    data.dearClient,
    data.bodyTarget,
    data.additionalNotes,
    data.opt1 ? "✓" : "",
    data.opt2 ? "✓" : "",
    data.opt3 ? "✓" : "",
    data.opt4 ? "✓" : "",
    data.opt5 ? "✓" : "",
    data.opt6 ? "✓" : "",
    data.opt6Text || ""
  ]);

  const pacsMsg = `🩺 คำขอเชื่อมต่อระบบ PACS ใหม่\n` +
                  `-----------------------------------------\n` +
                  `• รหัสอ้างอิง: ${data.pacsId}\n` +
                  `• วันที่แจ้งขอ: ${data.requestDate}\n` +
                  `• ถึง: ${data.dearClient}\n` +
                  `• เรื่อง: ${data.subjectTarget}\n` +
                  `• รายละเอียดเพิ่มเติม: ${data.additionalNotes || 'ไม่มี'}\n` +
                  `-----------------------------------------`;

  // 3.1 ส่งแจ้งเตือนทาง Telegram
  sendTelegramNotification(pacsMsg);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'บันทึกหนังสือขอสิทธิ์เชื่อม PACS สำเร็จ'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 4. จัดการอัปโหลดเอกสารเพิ่มเติมหลังการยื่นขอเสร็จสิ้น (Additional Upload)
// ==========================================
function handleAdditionalUpload(data, spreadsheet) {
  const sheet = getOrCreateSheet(spreadsheet, 'additional_uploads_db');
  const formattedDate = getThaiFormattedDateTime();
  const searchId = data.submissionId.trim();
  
  let targetFolder = null;
  
  // 4.1 ค้นหาโฟลเดอร์เดิมใน Google Drive ที่มีรหัส SubmissionID นี้ประกอบอยู่
  if (MAIN_FOLDER_ID) {
    try {
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      const subFolders = mainFolder.getFolders();
      while (subFolders.hasNext()) {
        const folder = subFolders.next();
        if (folder.getName().indexOf(searchId) !== -1) {
          targetFolder = folder;
          break;
        }
      }
      
      if (!targetFolder) {
        targetFolder = mainFolder.createFolder(`${searchId} Additional_Upload (${formattedDate})`);
      }
    } catch (fError) {
      Logger.log("ไม่สามารถเข้าถึง Drive ได้: " + fError.toString());
    }
  }

  // 4.2 วนลูปบันทึกไฟล์แนบเพิ่ม (ส่งมาสูงสุด 3 ไฟล์)
  const fileUrls = {};
  const uploadedFileNames = [];
  
  for (let i = 1; i <= 3; i++) {
    const fileData = data[`file${i}`];
    if (fileData && fileData.base64 && targetFolder) {
      try {
        const decodedBytes = Utilities.base64Decode(fileData.base64.split(',')[1]);
        const blob = Utilities.newBlob(decodedBytes, fileData.type || 'application/pdf', fileData.name);
        const newFile = targetFolder.createFile(blob);
        fileUrls[`file${i}_url`] = newFile.getUrl();
        fileUrls[`file${i}_name`] = fileData.name;
        uploadedFileNames.push(fileData.name);
      } catch (uploadError) {
        fileUrls[`file${i}_url`] = "อัปโหลดล้มเหลว: " + uploadError.toString();
        fileUrls[`file${i}_name`] = fileData.name;
      }
    } else {
      fileUrls[`file${i}_url`] = "";
      fileUrls[`file${i}_name`] = "";
    }
  }

  // 4.3 บันทึกข้อมูลประวัติการอัปโหลดลงในชีต
  sheet.appendRow([
    new Date(),
    searchId,
    data.note || "ไม่มีบันทึกข้อความเสริม",
    fileUrls.file1_name || "",
    fileUrls.file1_url || "",
    fileUrls.file2_name || "",
    fileUrls.file2_url || "",
    fileUrls.file3_name || "",
    fileUrls.file3_url || ""
  ]);

  const uploadMsg = `📤 เอกสารเพิ่มเติมใหม่\n` +
                    `-----------------------------------------\n` +
                    `• รหัสคำขอเดิม: ${searchId}\n` +
                    `• บันทึกย่อ: ${data.note || '-'}\n` +
                    `• วันที่ดำเนินการ: ${formattedDate}\n` +
                    `• ไฟล์ที่แนบเสริม: ${uploadedFileNames.join(', ') || 'ไม่มี'}\n` +
                    `-----------------------------------------`;

  // 4.4 ส่งแจ้งเตือนแอดมินทางอีเมล
  if (ADMIN_EMAIL && ADMIN_EMAIL !== "your_email@domain.com") {
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: `[เอกสารเพิ่มเติม] รหัสโครงการ ${searchId} มีการอัปโหลดไฟล์เพิ่มเข้ามา`,
        body: uploadMsg
      });
    } catch (eError) {
      Logger.log("ไม่สามารถส่งอีเมลเพิ่มเติมได้: " + eError.toString());
    }
  }

  // 4.5 ส่งแจ้งเตือนทาง Telegram
  sendTelegramNotification(uploadMsg);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'อัปโหลดไฟล์เพิ่มเติมและแนบลงโฟลเดอร์โครงการเดิมเสร็จสมบูรณ์',
    folderName: targetFolder ? targetFolder.getName() : ""
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// ฟังก์ชันตัวช่วยดึงหรือสร้างแผ่นงานใน Google Sheets
// ==========================================
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    
    // ตั้งค่าหัวข้อคอลัมน์ (Headers) อัตโนมัติในแถวที่ 1
    let headers = [];
    if (sheetName === 'requests_db') {
      headers = [
        'Timestamp', 'SubmissionID', 'CompanyName', 'CompanyAddress', 'RepName', 'RepPhone', 
        'DeviceName', 'DeviceBrand', 'DeviceModel', 'DeviceCountry', 'DeviceSpecs', 'DevicePurpose', 
        'RefHospital', 'UserDepartment', 'StaffInCharge', 'ItemsList', 'DurationDays', 'StartDate', 
        'EndDate', 'Doc1_URL', 'Doc2_URL', 'Doc3_URL', 'Doc4_URL', 'Doc5_URL', 'Doc6_URL', 
        'Doc7_URL', 'Doc8_URL', 'Doc9_URL', 'Summary'
      ];
    } else if (sheetName === 'evaluations_db') {
      headers = [
        'Timestamp', 'EvaluationID', 'CompanyName', 'DeviceName', 'DeviceBrand', 'DeviceModel', 
        'DeviceCountry', 'UserDepartment', 'StaffInCharge', 'Score_Q1', 'Score_Q2', 'Score_Q3', 
        'Score_Q4', 'Score_Q5', 'Score_Q6', 'Score_Q7', 'Pros', 'Cons', 'CompareBrand', 
        'CompareModel', 'CompareDetails', 'Suggestions'
      ];
    } else if (sheetName === 'pacs_disclosures_db') {
      headers = [
        'Timestamp', 'PACS_ID', 'RequestDate', 'SubjectTarget', 'DearClient', 'BodyTarget', 
        'AdditionalNotes', 'Opt1_DICOM', 'Opt2_Worklist', 'Opt3_PatientInfo', 'Opt4_Report', 
        'Opt5_PACS_Access', 'Opt6_DICOM_Outside', 'Opt6Text'
      ];
    } else if (sheetName === 'additional_uploads_db') {
      headers = [
        'Timestamp', 'SubmissionID', 'UploadNote', 
        'File1_Name', 'File1_URL', 'File2_Name', 'File2_URL', 'File3_Name', 'File3_URL'
      ];
    }
    
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}
