// =====================================================================
// CytoFlow v1.5.1 AI Core - Multi-Database Architecture (SaaS Ready)
// =====================================================================

// --- DATABASE CONFIGURATION ---
// นำ ID ของ Google Sheets ทั้ง 4 ไฟล์มาใส่ที่นี่
const DB_FILES = {
  SYSTEM: 'XXX',      // [Sheets]: Master_Config
  USER: 'XXX',        // [Sheets]: User_Config
  REF: 'XXX',         // [Sheets]: Reference_Data_Config
  LOG: 'XXX'          // [Sheets]: Log_Config
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CytoFlow 2026 (v1.5.1 AI Core)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function formatDateVal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function formatDateTimeVal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  }
  return String(val);
}

// Helper: สำหรับเปรียบเทียบเพื่อทำ Audit Log
function safeString(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

function getDiffString(oldArr, newArr, headers) {
  let diffs = [];
  for (let i = 0; i < oldArr.length; i++) {
     let oldVal = safeString(oldArr[i]);
     let newVal = safeString(newArr[i]);
     
     // Remove leading tick for clean logging
     if(oldVal.startsWith("'")) oldVal = oldVal.substring(1);
     if(newVal.startsWith("'")) newVal = newVal.substring(1);

     if (oldVal !== newVal) {
        diffs.push(`[${headers[i]}]: '${oldVal}' -> '${newVal}'`);
     }
  }
  return diffs.length > 0 ? diffs.join(' | ') : 'No data changed';
}

// Helper: ดึง Sheet ข้อมูลคนไข้ตามปี
function getDbSheet(year) {
  const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
  const configSheet = sysSS.getSheetByName('Year');
  if(!configSheet) throw new Error("ไม่พบแท็บ Year ในไฟล์ Master_Config");
  
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(year)) {
      return SpreadsheetApp.openById(data[i][1]).getSheetByName('Data');
    }
  }
  throw new Error("ไม่พบ Config สำหรับปี: " + year);
}

// --- LOGGING (แยกไฟล์เพื่อความปลอดภัยระดับ PDPA) ---
function logSystem(action, detail, username) {
  try {
    const logSS = SpreadsheetApp.openById(DB_FILES.LOG);
    let sheet = logSS.getSheetByName('System_Logs');
    if (!sheet) {
      sheet = logSS.insertSheet('System_Logs');
      sheet.appendRow(['Timestamp', 'User', 'Action', 'Detail']);
    }
    sheet.appendRow([new Date(), username, action, detail]);
  } catch(e) { console.log("Log Sys Error: " + e); }
}

function logData(year, action, detail, username, cytoNo) {
  try {
    const logSS = SpreadsheetApp.openById(DB_FILES.LOG);
    const sheetName = 'Log_' + year;
    let sheet = logSS.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = logSS.insertSheet(sheetName);
      sheet.appendRow(['Timestamp', 'User', 'Action', 'CytoNo', 'Detail']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), username, action, cytoNo, detail]);
  } catch(e) { 
    console.log("Log Data Error: " + e); 
  }
}

function getCurrentYear() {
  const today = new Date();
  return today.getFullYear() + 543; // ปี พ.ศ. ปัจจุบัน
}

// ให้ Frontend เรียกใช้บันทึก Log ได้
function apiFrontendLog(action, detail, username) {
  logSystem(action, detail, username);
  return { status: 'success' };
}

// --- API: AUDIT TRAIL LOGS (v1.5.1 AI Core) ---
// ดึงประวัติการแก้ไขเพื่อแสดงผลในหน้าอ่านอย่างเดียว
function apiGetEditLogs(year, cytoNo) {
  try {
    const logSS = SpreadsheetApp.openById(DB_FILES.LOG);
    const sheetName = 'Log_' + year;
    const sheet = logSS.getSheetByName(sheetName);
    if (!sheet) return { status: 'success', regEdits: [], reportEdits: [], initialReport: null };

    const data = sheet.getDataRange().getValues();
    let regEdits = [];
    let reportEdits = [];
    let initialReport = null;

    // เริ่มจากแถว 1 (ข้าม Header)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]) === String(cytoNo)) {
        const timestamp = formatDateTimeVal(data[i][0]);
        const user = data[i][1];
        const action = data[i][2];
        
        if (action === "Edit Registration") {
          regEdits.push({ timestamp: timestamp, user: user });
        } else if (action === "Submit Report") {
          initialReport = { timestamp: timestamp, user: user };
        } else if (action === "Edit Report") {
          reportEdits.push({ timestamp: timestamp, user: user });
        }
      }
    }
    
    return { 
      status: 'success', 
      regEdits: regEdits, 
      reportEdits: reportEdits,
      initialReport: initialReport
    };
  } catch (e) {
    return { status: 'error', message: 'Failed to fetch logs: ' + e.message };
  }
}

// --- API: STICKER CONFIG ---
function apiGetStickerConfig() {
  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    
    let config = {
      width: 50, height: 25, autoPrintCount: 2, manualPrintCount: 1, barcodeHeight: 30, barcodeWidth: 1.5,
      layoutJSON: JSON.stringify({
        cyto: { x: 25, y: 4, size: 12, rot: 0, align: 'center', visible: true, bold: true, font: 'Sarabun' },
        name: { x: 25, y: 9, size: 9, rot: 0, align: 'center', visible: true, bold: false, font: 'Sarabun' },
        age:  { x: 25, y: 13, size: 9, rot: 0, align: 'center', visible: true, bold: false, font: 'Sarabun' },
        spec: { x: 25, y: 17, size: 9, rot: 0, align: 'center', visible: true, bold: false, font: 'Sarabun' },
        unit: { x: 25, y: 21, size: 9, rot: 0, align: 'center', visible: true, bold: false, font: 'Sarabun' },
        bar:  { x: 25, y: 24, rot: 0, align: 'center', visible: true, width: 1.5 },
        qrcode: { x: 25, y: 35, size: 50, rot: 0, align: 'center', visible: false }
      })
    };

    if (!sheet) {
      sheet = sysSS.insertSheet('Sticker_Config');
      sheet.appendRow(['Key', 'Value', 'Description']);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f1f5f9");
      const descriptions = { width: "ความกว้างของสติ๊กเกอร์ (mm)", height: "ความสูงของสติ๊กเกอร์ (mm)", autoPrintCount: "จำนวนแผ่นที่จะพิมพ์อัตโนมัติ", manualPrintCount: "จำนวนแผ่นเมื่อกดพิมพ์เอง", barcodeHeight: "ความสูงบาร์โค้ด", barcodeWidth: "ความกว้างบาร์โค้ด", layoutJSON: "พิกัดและดีไซน์" };
      for (let key in config) sheet.appendRow([key, config[key], descriptions[key]]);
      sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 300);
    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (config.hasOwnProperty(data[i][0])) {
          let val = data[i][1];
          if (val === 'true' || val === true) config[data[i][0]] = true;
          else if (val === 'false' || val === false) config[data[i][0]] = false;
          else if (data[i][0] === 'layoutJSON') config[data[i][0]] = String(val);
          else config[data[i][0]] = Number(val) || val;
        }
      }
    }
    return { status: 'success', config: config };
  } catch (e) { return { status: 'error', message: 'Get Sticker Config Error: ' + e.message }; }
}

function apiSaveStickerConfig(newConfig, username) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    if (!sheet) return { status: 'error', message: 'Sticker_Config sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let key in newConfig) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(newConfig[key]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, newConfig[key], "Auto-generated field"]);
    }
    logSystem("Update Config", "Admin updated Sticker Configuration Layout", username);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: 'Save Sticker Config Error: ' + e.message }; } 
  finally { lock.releaseLock(); }
}

// --- API: GET MASTER DATA ---
function apiGetMasterData() {
  try {
    const refSS = SpreadsheetApp.openById(DB_FILES.REF);
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    
    var getColA = function(sheetName, targetSS) {
      var sheet = targetSS.getSheetByName(sheetName);
      if(!sheet) return [];
      var data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
      var res = [];
      for(var i=0; i<data.length; i++) {
        if(data[i][0] && data[i][0].toString().trim() !== "") res.push(data[i][0].toString().trim());
      }
      return res;
    };

    var getPrefixes = function() {
      var sheet = refSS.getSheetByName("Prefix");
      if(!sheet) return [];
      var data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
      var res = [];
      for(var i=0; i<data.length; i++) {
        if(data[i][0] && data[i][0].toString().trim() !== "") {
          res.push({ prefix: data[i][0].toString().trim(), sex: (data[i][1] ? data[i][1].toString().trim() : 'หญิง') });
        }
      }
      return res;
    };

    const adequacySheet = refSS.getSheetByName('SPECIMEN ADEQUACY');
    const sqSheet = refSS.getSheetByName('200 Squamous Cell');
    const glSheet = refSS.getSheetByName('200 Glandular Cell');
    const cytoTechSheet = userSS.getSheetByName('Cytotechnologist');
    const pathoSheet = userSS.getSheetByName('Pathologist');
    
    let adequacyMaster = []; let cytoTechs = []; let pathos = [];
    let masterSquamous = []; let masterGlandular = []; 

    if (adequacySheet) {
      const adData = adequacySheet.getRange(2, 1, Math.max(1, adequacySheet.getLastRow() - 1), 2).getValues();
      adequacyMaster = adData.map(r => ({ group: String(r[0]).trim(), text: String(r[1]).trim() })).filter(x => x.text);
    }
    if (sqSheet) {
      const sqData = sqSheet.getRange(2, 1, Math.max(1, sqSheet.getLastRow() - 1), 3).getValues();
      masterSquamous = sqData.map(r => ({ main: String(r[0]).trim(), detail1: String(r[1]).trim(), detail2: String(r[2]).trim() })).filter(x => x.main || x.detail1);
    }
    if (glSheet) {
      const glData = glSheet.getRange(2, 1, Math.max(1, glSheet.getLastRow() - 1), 2).getValues();
      masterGlandular = glData.map(r => ({ main: String(r[0]).trim(), detail1: String(r[1]).trim() })).filter(x => x.main || x.detail1);
    }
    if (cytoTechSheet) {
      const ctData = cytoTechSheet.getRange(2, 1, Math.max(1, cytoTechSheet.getLastRow() - 1), 2).getValues();
      cytoTechs = ctData.map(r => (String(r[0]).trim() + " " + String(r[1]).trim()).trim()).filter(Boolean);
    }
    if (pathoSheet) {
      const ptData = pathoSheet.getRange(2, 1, Math.max(1, pathoSheet.getLastRow() - 1), 2).getValues();
      pathos = ptData.map(r => (String(r[0]).trim() + " " + String(r[1]).trim()).trim()).filter(Boolean);
    }
    
    return { 
      status: 'success', 
      units: getColA("Sampling_Unit", refSS), 
      districts: getColA("District", refSS), 
      adequacyMaster: adequacyMaster, 
      cytoTechs: cytoTechs, 
      pathos: pathos,
      masterSquamous: masterSquamous,
      masterGlandular: masterGlandular,
      masterCat300: getColA("300 OTHER", refSS),
      masterOrganism: getColA("100 Organism", refSS),
      masterNonNeo: getColA("100 Other non neoplastic", refSS),
      masterContras: getColA("Contraception", refSS),
      masterPrevTx: getColA("Previous treatment", refSS),
      prefixes: getPrefixes(),
      masterSpecimenType: getColA("Specimen Type", refSS),
      masterComment: getColA("Comment", refSS)
    };
  } catch (e) { return { status: 'error', message: 'Master Data Error: ' + e.message }; }
}

function apiGetNextCytoNo(year) {
  try {
    const sheet = getDbSheet(year);
    const yearPrefix = year.toString().substring(2);
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue().toString();
      if (lastId.startsWith(yearPrefix)) {
        const numPart = parseInt(lastId.substring(2)); 
        if (!isNaN(numPart)) nextNum = numPart + 1;
      }
    }
    return { status: 'success', cytoNoPreview: yearPrefix + nextNum.toString().padStart(4, '0') };
  } catch (e) { return { status: 'error', message: e.message }; }
}

// --- API: MPI / PREVIOUS RESULTS ---
// สำหรับค้นหาประวัติทั้งหมด (History Table)
function apiGetPatientHistory(cid, username) {
  if (!cid || String(cid).trim() === "") {
    return { status: 'error', message: 'กรุณาระบุเลขประจำตัวประชาชน (CID) เพื่อค้นหาประวัติ' };
  }
  
  const searchCid = String(cid).trim();
  let historyData = [];
  
  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
    const configSheet = sysSS.getSheetByName('Year');
    if (!configSheet) throw new Error("ไม่พบแท็บ Year ในไฟล์ Master_Config");

    const yearData = configSheet.getDataRange().getValues();
    
    for (let i = 1; i < yearData.length; i++) {
      const year = String(yearData[i][0]);
      const sheetId = yearData[i][1];
      if (!sheetId) continue;

      try {
        const dbSS = SpreadsheetApp.openById(sheetId);
        const dataSheet = dbSS.getSheetByName('Data');
        if (!dataSheet) continue;

        const lastRow = dataSheet.getLastRow();
        if (lastRow < 2) continue;

        const data = dataSheet.getRange(2, 1, lastRow - 1, 42).getValues();

        for (let r = 0; r < data.length; r++) {
          const rowCid = String(data[r][2]).trim();
          if (rowCid === searchCid) {
            historyData.push({
              dbYear: year, rowId: r + 2, cytoNo: String(data[r][0]), hn: String(data[r][1]), cid: rowCid,
              prefix: data[r][3], fname: data[r][4], lname: data[r][5], age: data[r][6], sex: data[r][7],
              specimenDate: formatDateVal(data[r][8]), recDate: formatDateVal(data[r][9]), unit: data[r][10] ? String(data[r][10]) : "ไม่ระบุ",
              district: data[r][11], hcode: data[r][12], coordinator: data[r][13], phone: data[r][14],
              para: data[r][15], last: data[r][16], lmp: formatDateVal(data[r][17]), contraception: data[r][18], prevTx: data[r][19],
              clinFind: data[r][20], clinDx: data[r][21], lastPap: data[r][22], method: data[r][23], registerName: data[r][24],
              regTimestamp: formatDateTimeVal(data[r][25]), adequacy: data[r][26], adequacyDetail: data[r][27], additional: data[r][28], 
              organism: data[r][29] ? String(data[r][29]) : "", nonNeo: data[r][30] ? String(data[r][30]) : "",
              squamousMain: data[r][31] ? String(data[r][31]) : "", squamousSub: data[r][32] ? String(data[r][32]) : "",
              glandularMain: data[r][33] ? String(data[r][33]) : "", glandularSub: data[r][34] ? String(data[r][34]) : "",
              cat300: data[r][35], comment: data[r][36], cytoName: data[r][37], cytoDateTime: String(data[r][38]), 
              pathoName: data[r][39], pathoDateTime: String(data[r][40]), status: data[r][41] ? String(data[r][41]) : "Pending"
            });
          }
        }
      } catch (err) {
        console.log(`[History MPI] Error reading DB for year ${year}: ${err.message}`);
      }
    }

    historyData.sort((a, b) => {
      const dateA = new Date(a.recDate || a.specimenDate || 0);
      const dateB = new Date(b.recDate || b.specimenDate || 0);
      return dateB - dateA; 
    });

    return { status: 'success', data: historyData };
  } catch (e) { return { status: 'error', message: 'History System Error: ' + e.message }; }
}

// --- API: CHECK PATIENT MPI (v1.5.1 AI Core) ---
// ค้นหาประวัติว่าเคยมีในระบบไหม เพื่อ Auto-fill (ปรับปรุงให้ค้นหาด้วย CID เท่านั้น ตามหลัก PDPA และความแม่นยำ)
function apiCheckPatientMPI(searchType, searchValue) {
  // หมายเหตุ: ตัวแปร searchType ถูกเก็บไว้เพื่อไม่ให้กระทบกับการส่งค่ามาจากฝั่ง Frontend เดิม 
  // แต่โค้ดภายในจะสนใจเฉพาะการตรวจจับ CID เท่านั้น
  if (!searchValue || String(searchValue).trim() === "") return { status: 'not_found' };
  
  const searchStr = String(searchValue).trim();
  let latestPatient = null;
  let latestDate = 0;

  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
    const configSheet = sysSS.getSheetByName('Year');
    if (!configSheet) throw new Error("Missing Year Config");

    const yearData = configSheet.getDataRange().getValues();
    
    // วนลูปย้อนหลังจากปีล่าสุด เพื่อหาข้อมูลที่อัปเดตที่สุดให้เร็วขึ้น
    for (let i = yearData.length - 1; i >= 1; i--) {
      const sheetId = yearData[i][1];
      if (!sheetId) continue;

      try {
        const dbSS = SpreadsheetApp.openById(sheetId);
        const dataSheet = dbSS.getSheetByName('Data');
        if (!dataSheet) continue;

        const lastRow = dataSheet.getLastRow();
        if (lastRow < 2) continue;

        const data = dataSheet.getRange(2, 1, lastRow - 1, 10).getValues(); // ดึงแค่วันที่และข้อมูลส่วนตัวพอเพื่อความเร็ว

        for (let r = data.length - 1; r >= 0; r--) {
          const hn = String(data[r][1]).trim();
          const cid = String(data[r][2]).trim();
          
          // ตรวจสอบการ Match จาก CID อย่างเดียว ตาม Requirement
          if (cid === searchStr) {
            const dateVal = new Date(data[r][9] || data[r][8] || 0).getTime(); // recDate or specimenDate
            // หากเจอข้อมูลที่ใหม่กว่า ให้ทับค่าเดิม
            if (dateVal >= latestDate) {
              latestDate = dateVal;
              latestPatient = {
                prefix: data[r][3],
                fname: data[r][4],
                lname: data[r][5],
                hn: hn,
                cid: cid
              };
            }
          }
        }
        // ถ้าเจอในระดับปีล่าสุดแล้ว ให้หยุดหา (Optimized Performance)
        if(latestPatient) break; 
      } catch (err) { console.log("MPI Scan error: " + err); }
    }

    if (latestPatient) {
      return { status: 'found', patient: latestPatient };
    } else {
      return { status: 'not_found' };
    }

  } catch (e) {
    return { status: 'error', message: e.message };
  }
}


// --- API: LOGIN (v1.5.1 AI Core) ---
function apiLoginStep1(username, password) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    const sheet = userSS.getSheetByName('Users');
    if(!sheet) return { status: 'error', message: 'ไม่พบฐานข้อมูล Users' };
    
    const data = sheet.getDataRange().getValues();
    let logoUrl = "https://drive.google.com/thumbnail?id=1ztvazUTKvglF0vRo4leaDe38olLfXQWs&sz=w200";
    try {
      const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
      const logoSheet = sysSS.getSheetByName('App_Logo');
      if (logoSheet && logoSheet.getLastRow() > 1) logoUrl = logoSheet.getRange(2, 2).getValue();
    } catch(e) { console.log("Logo fetch error: " + e); }

    let userFound = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(username) && String(data[i][1]) == String(password)) {
        userFound = true;
        const email = data[i][5];
        if (!email || email === "") {
          logSystem("Login Failed", "Account missing email", username);
          return { status: 'error', message: 'Account นี้ยังไม่ระบุ Email ในระบบ' };
        }
        
        // --- OTP RATE LIMITING LOGIC (v1.5.1) ---
        const cache = CacheService.getScriptCache();
        const otpCacheKey = "OTP_" + username;
        const timeCacheKey = "OTP_TIME_" + username;
        
        const existingOtp = cache.get(otpCacheKey);
        const reqTime = cache.get(timeCacheKey);
        const maskedEmail = email.replace(/^(.)(.*)(.@.*)$/, "$1***$3");

        // ถ้ามี OTP เดิม และเคยขอไว้
        if (existingOtp && reqTime) {
          let elapsedSeconds = Math.floor((new Date().getTime() - parseInt(reqTime)) / 1000);
          
          if (elapsedSeconds < 60) {
            // ยังไม่ถึง 1 นาที ห้ามส่งเมล์ใหม่ ให้ใช้ OTP เดิมไปก่อน
            return { 
              status: 'otp_cooldown', 
              message: 'กรุณากรอกรหัส OTP เดิมที่ถูกส่งไปยัง ' + maskedEmail, 
              systemLogo: logoUrl,
              cooldown: 60 - elapsedSeconds // ส่งเวลาที่ต้องรอเพื่อขอใหม่
            };
          }
        }

        // กรณีเกิน 1 นาที หรือยังไม่มี OTP ค้างในระบบ ให้ส่งใหม่ได้
        const otp = Math.floor(100000 + Math.random() * 900000).toString();
        
        // เก็บ OTP ไว้ 5 นาที (300 วินาที)
        cache.put(otpCacheKey, otp, 300);
        // เก็บเวลาที่ขอ เพื่อใช้คำนวณ Cooldown (1 นาที)
        cache.put(timeCacheKey, new Date().getTime().toString(), 300);

        // --- NEW OTP EMAIL TEMPLATE (GitHub Style with PDPA) ---
        const htmlTemplate = `
        <div style="font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif; color: #24292f; max-width: 600px; margin: 0 auto; padding: 20px;">
            <div style="text-align: center; margin-bottom: 24px;">
                <img src="${logoUrl}" alt="CytoFlow Logo" style="width: 64px; height: 64px; border-radius: 50%; object-fit: contain; vertical-align: middle;">
                <span style="font-size: 32px; font-weight: 600; color: #24292f; vertical-align: middle; margin-left: 12px; letter-spacing: -0.5px; display: inline-block;">CytoFlow</span>
            </div>
            <h2 style="font-size: 24px; font-weight: 400; text-align: center; margin-bottom: 24px; color: #24292f;">
                กรุณายืนยันตัวตนของคุณ, <strong>${username}</strong>
            </h2>
            <div style="background-color: #ffffff; border: 1px solid #d0d7de; border-radius: 6px; padding: 24px;">
                <p style="margin-top: 0; margin-bottom: 16px; font-size: 14px;">นี่คือรหัส OTP สำหรับยืนยันการเข้าสู่ระบบของคุณ:</p>
                <div style="text-align: center; font-size: 32px; font-family: ui-monospace,SFMono-Regular,monospace; letter-spacing: 8px; color: #24292f; margin: 24px 0;">
                    ${otp}
                </div>
                <p style="font-size: 14px; margin-bottom: 16px;">
                    รหัสนี้มีอายุการใช้งาน <strong>5 นาที</strong> และสามารถใช้ได้เพียงครั้งเดียว
                </p>
                <p style="font-size: 14px; margin-bottom: 16px;">
                    <strong>โปรดอย่าแชร์รหัสนี้กับใคร:</strong> ทีมงานจะไม่ขอรหัสผ่านหรือ OTP ของคุณทางโทรศัพท์หรืออีเมล<br>
                    เด็ดขาด (มาตรการรักษาความปลอดภัยข้อมูลส่วนบุคคล - PDPA)
                </p>
                <p style="font-size: 14px; margin-bottom: 0;">
                    ขอบคุณ,<br>ทีมงาน CytoFlow
                </p>
            </div>
            <div style="margin-top: 32px; font-size: 12px; color: #6e7781; text-align: center; line-height: 1.5;">
                คุณได้รับอีเมลฉบับนี้เนื่องจากมีการร้องขอรหัสยืนยันสำหรับบัญชีผู้ใช้ CytoFlow ของคุณ หากคุณไม่ได้เป็นผู้ร้องขอ<br>
                โปรดเพิกเฉยต่ออีเมลฉบับนี้ หรือแจ้งผู้ดูแลระบบทันที
            </div>
        </div>
        `;

        try { 
          MailApp.sendEmail({ 
            to: email, 
            subject: "รหัส OTP สำหรับเข้าสู่ระบบ CytoFlow", 
            htmlBody: htmlTemplate,
            name: "CytoFlow" 
          }); 
          logSystem("OTP Requested", "New OTP generated and sent", username);
        } 
        catch (mailErr) { 
          logSystem("OTP Error", "Failed to send OTP email: " + mailErr.message, username);
          return { status: 'error', message: 'ส่งอีเมล OTP ไม่สำเร็จ: ' + mailErr.message }; 
        }

        return { 
          status: 'otp_required', 
          message: 'ส่งรหัส OTP ใหม่ไปยัง ' + maskedEmail + ' เรียบร้อยแล้ว', 
          systemLogo: logoUrl,
          cooldown: 60
        };
      }
    }
    
    logSystem("Login Failed", "Invalid username or password", username);
    return { status: 'error', message: 'Username หรือ Password ไม่ถูกต้อง' };
  } catch (e) { 
    return { status: 'error', message: 'System Error: ' + e.message }; 
  }
}

function apiVerifyOtp(username, inputOtp) {
  try {
    const cache = CacheService.getScriptCache();
    const storedOtp = cache.get("OTP_" + username);
    if (storedOtp && storedOtp === inputOtp) {
      cache.remove("OTP_" + username);
      cache.remove("OTP_TIME_" + username);
      
      const userSS = SpreadsheetApp.openById(DB_FILES.USER);
      const sheet = userSS.getSheetByName('Users');
      const data = sheet.getDataRange().getValues();
      let userData = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) == String(username)) {
          userData = { name: data[i][2], position: data[i][3], role: data[i][4], image: data[i][6] || "", username: username };
          break;
        }
      }
      
      const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
      const configSheet = sysSS.getSheetByName('Year');
      const years = configSheet.getDataRange().getValues().slice(1).map(r => String(r[0]));
      
      logSystem("Login Success", "Successfully verified OTP", username);
      return { status: 'success', user: userData, years: years, currentYear: getCurrentYear() };
    } else { 
      logSystem("Login Failed", "Invalid or Expired OTP", username);
      return { status: 'error', message: 'รหัส OTP ไม่ถูกต้อง หรือหมดอายุ' }; 
    }
  } catch (e) { return { status: 'error', message: 'Verify Error: ' + e.message }; }
}

function apiVerifyPassword(username, password) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    const sheet = userSS.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(username) && String(data[i][1]) == String(password)) {
        logSystem("Unlock Screen", "Successfully unlocked screen", username);
        return { status: 'success' };
      }
    }
    logSystem("Unlock Failed", "Invalid password during screen unlock", username);
    return { status: 'error', message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (e) { return { status: 'error', message: 'System Error: ' + e.message }; }
}

// --- API: DASHBOARD ---
function apiGetDashboardData(year) {
  try {
    const sheet = getDbSheet(year);
    const lastRow = sheet.getLastRow();
    let stats = { total: 0, reported: 0, abnormal: 0 };
    let patients = [];

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 42).getValues();
      data.forEach((r, index) => {
        const status = r[41] ? String(r[41]) : "Pending";
        
        stats.total++;
        if (status === "Reported") stats.reported++;

        patients.push({
          rowId: index + 2, cytoNo: String(r[0]), hn: String(r[1]), cid: String(r[2]),
          prefix: r[3], fname: r[4], lname: r[5], age: r[6], sex: r[7],
          specimenDate: formatDateVal(r[8]), recDate: formatDateVal(r[9]),
          unit: r[10] ? String(r[10]) : "ไม่ระบุ", district: r[11], hcode: r[12], coordinator: r[13], phone: r[14],
          para: r[15], last: r[16], lmp: formatDateVal(r[17]), contraception: r[18], prevTx: r[19],
          clinFind: r[20], clinDx: r[21], lastPap: r[22], method: r[23], registerName: r[24],
          regTimestamp: formatDateTimeVal(r[25]),
          adequacy: r[26], adequacyDetail: r[27], additional: r[28], 
          organism: r[29] ? String(r[29]) : "",
          nonNeo: r[30] ? String(r[30]) : "",
          squamousMain: r[31] ? String(r[31]) : "",
          squamousSub: r[32] ? String(r[32]) : "",
          glandularMain: r[33] ? String(r[33]) : "",
          glandularSub: r[34] ? String(r[34]) : "",
          cat300: r[35], comment: r[36], 
          cytoName: r[37], cytoDateTime: String(r[38]), 
          pathoName: r[39], pathoDateTime: String(r[40]), 
          status: status 
        });
      });
    }
    return { status: 'success', stats: stats, patients: patients.reverse() };
  } catch (e) { return { status: 'error', message: "Data Error: " + e.message }; }
}

// --- API: REGISTER ---
function apiRegisterSample(form, year, username) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const sheet = getDbSheet(year);
    const yearPrefix = year.toString().substring(2);
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue().toString();
      if (lastId.startsWith(yearPrefix)) { const numPart = parseInt(lastId.substring(2)); if (!isNaN(numPart)) nextNum = numPart + 1; }
    }
    const cytoNo = yearPrefix + nextNum.toString().padStart(4, '0');
    const phoneStr = form.phone ? "'" + form.phone : ""; 
    
    // Server-side Sanitization: จำกัดให้อายุเป็นตัวเลขและไม่เกิน 3 หลัก
    const safeAge = form.age ? String(form.age).replace(/\D/g, '').substring(0, 3) : "";

    let record = [
      cytoNo, String(form.hn), String(form.cid), form.prefix, form.fname, form.lname, safeAge, form.sex,
      form.specimenDate, form.receivedDate, form.unit, form.district, form.hcode, form.coordinator, phoneStr,
      form.para, form.last, form.lmp, form.contraception, form.prevTx, form.clinFind, form.clinDx,
      form.lastPap, form.method, form.registerName
    ]; 
    record.push(new Date()); 
    for(let i = 0; i < 15; i++) { record.push(""); }
    record.push("Pending");

    sheet.appendRow(record);
    logData(year, "Register", "Created new sample", username, cytoNo);
    return { status: 'success', cytoNo: cytoNo };
  } catch (e) { return { status: 'error', message: "Save Failed: " + e.message }; }
  finally { lock.releaseLock(); }
}

// --- API: UPDATE ---
function apiUpdateSample(form, year, rowId, username) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const sheet = getDbSheet(year); 
    const rowIndex = parseInt(rowId);
    if (rowIndex > sheet.getLastRow()) return { status: 'error', message: 'Row not found' };
    
    // 1. Get Old Data for Diff
    const oldData = sheet.getRange(rowIndex, 2, 1, 24).getValues()[0];
    
    // 2. Prepare New Data
    const phoneStr = form.phone ? "'" + form.phone : ""; 
    const safeAge = form.age ? String(form.age).replace(/\D/g, '').substring(0, 3) : "";

    const newRecord = [
      String(form.hn), String(form.cid), form.prefix, form.fname, form.lname, safeAge, form.sex,
      form.specimenDate, form.receivedDate, form.unit, form.district, form.hcode, form.coordinator, phoneStr,
      form.para, form.last, form.lmp, form.contraception, form.prevTx, form.clinFind, form.clinDx,
      form.lastPap, form.method, form.registerName
    ];
    
    // 3. Generate Diff String
    const headers = ['HN', 'CID', 'Prefix', 'Fname', 'Lname', 'Age', 'Sex', 'SpecimenDate', 'RecDate', 'Unit', 'District', 'HCode', 'Coordinator', 'Phone', 'Para', 'Last', 'LMP', 'Contraception', 'PrevTx', 'ClinFind', 'ClinDx', 'LastPap', 'SpecimenType', 'RegName'];
    const diffLog = getDiffString(oldData, newRecord, headers);
    
    // 4. Update Database
    sheet.getRange(rowIndex, 2, 1, 24).setValues([newRecord]); 
    const cytoNo = sheet.getRange(rowIndex, 1).getValue();
    
    // 5. Save Log
    logData(year, "Edit Registration", diffLog, username, cytoNo);
    
    return { status: 'success', cytoNo: cytoNo };
  } catch (e) { return { status: 'error', message: "Update Failed: " + e.message }; }
  finally { lock.releaseLock(); }
}

// --- API: REPORT ---
function apiSubmitReport(form, year, username) {
  try {
    const sheet = getDbSheet(year);
    const row = parseInt(form.rowId);
    
    const oldData = sheet.getRange(row, 27, 1, 16).getValues()[0]; 
    
    const cytoDT = form.cytoDateTime ? "'" + form.cytoDateTime : "";
    const pathoDT = form.pathoDateTime ? "'" + form.pathoDateTime : "";
    
    const newRecord = [
      form.adequacy, form.adequacyDetail, form.additional, form.organism, form.nonNeo,
      form.squamousMain, form.squamousSub, form.glandularMain, form.glandularSub,
      form.cat300, form.comment, form.cytoName, cytoDT, form.pathoName, pathoDT, "Reported"
    ];
    
    sheet.getRange(row, 27, 1, 16).setValues([newRecord]); 

    const cytoNo = sheet.getRange(row, 1).getValue();
    const headers = ['Adequacy', 'AdqDetail', 'Additional', 'Organism', 'NonNeo', 'SqMain', 'SqSub', 'GlMain', 'GlSub', 'Cat300', 'Comment', 'CytoName', 'CytoDT', 'PathoName', 'PathoDT', 'Status'];
    
    if (form.isEdit) {
      const diffLog = getDiffString(oldData, newRecord, headers);
      logData(year, "Edit Report", diffLog, username, cytoNo);
    } else {
      logData(year, "Submit Report", "Initial Report Submitted", username, cytoNo);
    }
    
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.message }; }
}

// --- API: PROFILE IMAGE & LOGO ---
function apiSaveProfileImage(username, base64Data) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER); 
    const sheet = userSS.getSheetByName('Users'); 
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1; let oldFileUrl = "";
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(username)) { rowIndex = i + 1; oldFileUrl = data[i][6]; break; } }
    if (rowIndex === -1) return { status: 'error', message: 'User not found' };
    if (oldFileUrl && oldFileUrl.includes("drive.google.com")) { try { const idMatch = oldFileUrl.match(/id=([^&]+)/); if (idMatch && idMatch[1]) DriveApp.getFileById(idMatch[1]).setTrashed(true); } catch (e) {} }
    
    const folderName = "CytoFlow_Profiles"; const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, `profile_${username}_${Date.now()}.jpg`); const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s400`; sheet.getRange(rowIndex, 7).setValue(fileUrl);
    
    logSystem("Change Profile Pic", "Updated profile image", username); return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiChangePassword(username, newPassword) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER); 
    const sheet = userSS.getSheetByName('Users'); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(username)) { sheet.getRange(i + 1, 2).setValue(newPassword); logSystem("Change Password", "Updated password", username); return { status: 'success' }; } }
    return { status: 'error', message: 'User not found' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveSystemLogo(base64Data, username) {
  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM); 
    let sheet = sysSS.getSheetByName('App_Logo');
    if (!sheet) { sheet = sysSS.insertSheet('App_Logo'); sheet.appendRow(['Name', 'Url']); }
    
    const folderName = "CytoFlow_Logo"; const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg")) ext = "jpg";
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7)); const blob = Utilities.newBlob(bytes, contentType, `app_logo_${Date.now()}.${ext}`);
    const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s1000`;
    
    if (sheet.getLastRow() < 2) sheet.appendRow(['MainLogo', fileUrl]); else sheet.getRange(2, 2).setValue(fileUrl);
    logSystem("Change Logo", "Updated system logo", username); return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}
