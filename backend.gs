/**
 * Elevator Maintenance Tracking - Backend
 * 
 * Column Order:
 * A: Timestamp (0), B: Employee (1), C: SiteID (2), D: Name (3), E: Project # (4),
 * F: Elevator (5) - Position 6
 * G: Type (6), H: Lat (7), I: Lng (8), J: Photo (9), K: Note (10), L: Duration (11),
 * M: Maps Link (12)
 */
const EMAIL_RECIPIENT = "geldyyevm@gmail.com";
const SAFETY_MINUTES = 60;
const MANUAL_URL = 'https://script.google.com/macros/s/AKfycbxI1LxlYGcwe41RDV9zzwQ78KdCc-BODlWjLKg_2ydx6wPSY9n74ZGIc3j3bDo-QpKv/exec'; 

const FOLDER_ID = '1ybLlBkK0DeXqVDuQdJRpG7dpyPuSSZI3'; 
const SHEET_NAME = 'Logs';
const SITES_SHEET_NAME = 'Sites'; 

let sitesCache = null;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Tracking System')
    .addItem('Generate QR Sticker', 'showQR')
    .addItem('Check Safety Alerts Now', 'checkSafetyAlerts')
    .addItem('Clear All Highlights', 'clearAllHighlights')
    .addSeparator()
    .addItem('Setup Sheet Headers (Fix Table)', 'setupSheet')
    .addItem('Migrate Old Data (Sync Columns)', 'migrateData')
    .addToUi();
}

function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  
  // 1. SIMPLE RESET FOR F & G (Just clear, don't delete to avoid shifting other columns)
  try {
    sheet.getRange("F:G").clearDataValidations();
    sheet.getRange("F:G").clearFormat();
    sheet.getRange("F:G").setNumberFormat("@");
  } catch(e) {}

  // 2. ATTEMPT TO WIPE REMAINING A:Z
  try { sheet.getRange("A:Z").clearDataValidations(); } catch(e) {}
  try { sheet.clearConditionalFormatRules(); } catch(e) {}

  // 3. SET ALL HEADERS INDIVIDUALLY TO BE 100% SURE
  const headers = ["Timestamp", "Employee Name", "Site ID", "Project Name", "Project #", "Elevator", "Type", "Lat", "Lng", "Photo", "Note", "Duration", "Maps Link"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
  
  // 4. APPLY FORMATS
  sheet.getRange("H:I").setNumberFormat("0.000000"); // Lat/Lng
  
  ui.alert("🔧 FIXED: Table Structure Reset.\n- Headers synchronized.\n- Formats applied.\n\nNow try submitting from the app again.");
}


function showQR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  
  if (sheet.getName() !== SITES_SHEET_NAME || row === 1) {
    SpreadsheetApp.getUi().alert("Please select a site in '" + SITES_SHEET_NAME + "'");
    return;
  }
  
  const data = sheet.getRange(row, 1, 1, 3).getValues()[0];
  const html = HtmlService.createHtmlOutput(getQrContent(data[0], data[1], data[2]))
    .setWidth(600).setHeight(700).setTitle('QR Stickers');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doGet(e) {
  const siteId = e.parameter.site || "";
  const elevator = e.parameter.el || "";
  const siteDetails = getSiteDetails(siteId);
  const serverConfig = JSON.stringify({
    siteId: siteId, siteFound: siteDetails.found,
    projectName: siteDetails.name, projectNumber: siteDetails.number,
    selectedElevator: elevator, lat: siteDetails.lat, lng: siteDetails.lng,
    safetyMinutes: SAFETY_MINUTES
  });
  const template = HtmlService.createTemplateFromFile('index');
  template.serverConfig = serverConfig;
  return template.evaluate().setTitle('MMAC Tracker').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processSubmission(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const rangeStart = Math.max(1, lastRow - 500);
    const recentLogs = lastRow > 0 ? sheet.getRange(rangeStart, 1, lastRow - rangeStart + 1, sheet.getLastColumn()).getValues() : [];

    const siteDetails = getSiteDetails(data.siteId);
    if (!siteDetails.found) return { success: false, error: "Invalid Site" };

    // 1. AUTO-CAPTURE Coordinates if missing in Sites Sheet
    const hasSiteCoords = siteDetails.lat && siteDetails.lng;
    if (!hasSiteCoords) {
      try {
        const sitesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SITES_SHEET_NAME);
        sitesSheet.getRange(siteDetails.row, 4, 1, 2).setValues([[data.lat, data.lng]]);
        siteDetails.lat = data.lat; 
        siteDetails.lng = data.lng;
      } catch (err) { console.warn("Auto-save failed: " + err); }
    } else {
      const distance = getDistance(Number(data.lat || 0), Number(data.lng || 0), Number(siteDetails.lat || 0), Number(siteDetails.lng || 0));
      if (distance > 0.3) return { success: false, error: "Too far from site" };
    }

    // 2. SESSION LOGIC (Restore missing variables)
    let finalType = data.type;
    let lastDevice = null;
    let durationVal = '';

    if (!finalType || finalType === 'CHECK' || finalType === 'OUT') {
      const smart = getSmartInfoFromData(recentLogs, data.employee, data.siteId, data.elevator);
      if (!finalType || finalType === 'CHECK') finalType = smart.type;
      lastDevice = smart.lastDevice;
      if (finalType === 'OUT' && smart.lastTime) {
        durationVal = Math.floor((new Date().getTime() - smart.lastTime.getTime()) / 60000);
      }
    }

    const photoUrl = data.photo && data.photo.length > 50 ? savePhoto(data.photo, `${data.siteId}_${finalType}.jpg`) : 'No photo';
    const finalNote = data.device + (finalType === 'OUT' && lastDevice && lastDevice !== data.device ? ' [⚠️ Device Mismatch]' : '');

    // ABSOLUTE TYPE ENFORCEMENT TO PREVENT "DATA VALIDATION" FAILURES
    const payload = [
      new Date(),                             // A: Timestamp
      String(data.employee),                  // B: Employee
      String(data.siteId),                    // C: Site ID
      String(siteDetails.name),               // D: Project Name
      String(data.projectNumber || siteDetails.number), // E: Project #
      Number(data.elevator || 1),             // F: Elevator (Numeric)
      String(finalType),                      // G: Type (Text)
      Number(data.lat || 0),                  // H: Lat (Numeric)
      Number(data.lng || 0),                  // I: Lng (Numeric)
      String(photoUrl),                       // J: Photo
      String(finalNote),                      // K: Note
      durationVal !== '' ? Number(durationVal) : '', // L: Duration (Numeric / Empty)
      `https://www.google.com/maps?q=${data.lat},${data.lng}` // M: Maps
    ];

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, payload.length).setValues([payload]);

    // 3. CLEAR PREVIOUS HIGHLIGHT IF CHECKING OUT
    if (finalType === 'OUT') {
      try {
        const smart = getSmartInfoFromData(recentLogs, data.employee, data.siteId, data.elevator);
        if (smart.row) {
          const rowToClear = smart.row;
          sheet.getRange(rowToClear, 1, 1, sheet.getLastColumn()).setBackground(null);
        }
      } catch (e) { console.warn("Highlight clear failed: " + e); }
    }

    return { success: true, startTime: payload[0].getTime(), employee: data.employee };
  } catch (e) { 
    console.error("Submission error:", e);
    return { success: false, error: e.toString() }; 
  }
}

/** 
 * Automatically shifts data if it was recorded in the old column format.
 * Detects if "IN/OUT" is in the new Elevator column instead of the Type column.
 */
function migrateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const range = sheet.getDataRange();
  const data = range.getValues();
  let headers = data[0];
  let rowsModified = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const colF = String(row[5] || "").toUpperCase();
    // If Column F contains IN/OUT, it means Type was at index 5 (Position F)
    // In new order, Type is at index 6 (Position G)
    if (colF === "IN" || colF === "OUT") {
      const newRow = [...row];
      newRow.splice(5, 0, 1); // Insert "1" as default elevator at index 5
      const targetRange = sheet.getRange(i + 1, 1, 1, Math.min(newRow.length, 26));
      targetRange.clearDataValidations();
      targetRange.setValues([newRow.slice(0, 26)]);
      rowsModified++;
    }
  }

  if (rowsModified > 0) {
    SpreadsheetApp.getUi().alert(`✅ Migrated ${rowsModified} rows. Elevator 1 inserted and columns shifted.`);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ No misaligned data detected.");
  }
}

function savePhoto(base64, name) {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1] || base64), 'image/jpeg', name);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getSiteDetails(siteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SITES_SHEET_NAME);
  if (!sheet) return { found: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === siteId.toLowerCase()) {
      return { found: true, name: data[i][1], number: data[i][2], lat: data[i][3], lng: data[i][4], row: i + 1 };
    }
  }
  return { found: false };
}

function getSessionInfo(siteId, device, elevator) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return { active: false };
  const data = sheet.getDataRange().getValues();
  const searchSite = siteId.toLowerCase(), searchElevator = (elevator || "1").toString().toLowerCase(), searchDev = device.toString();

  for (let i = data.length - 1; i >= 0; i--) {
    const rowSite = (data[i][2] || "").toString().toLowerCase();
    const rowElevator = (data[i][5] || "1").toString().toLowerCase();
    if (rowSite === searchSite && rowElevator === searchElevator) {
      const rowDev = (data[i][10] || "").toString();
      if (rowDev.includes(searchDev) || searchDev.includes(rowDev)) {
        if (data[i][6] === 'IN') {
          return { active: true, employee: data[i][1], startTime: new Date(data[i][0]).getTime() };
        }
        return { active: false };
      }
    }
  }
  return { active: false };
}

function getDistance(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2-lat1)*Math.PI/180, dLon = (lon2-lon1)*Math.PI/180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function getQrContent(siteId, name, number) {
  const baseUrl = MANUAL_URL || ScriptApp.getService().getUrl();
  const match = (number || "").toString().match(/-(\d+)$/);
  const count = match ? parseInt(match[1]) : 1;
  let html = `<style>
    body { font-family: 'Inter', sans-serif; background: #f8fafc; padding: 40px; }
    .sticker { background: white; border: 2px solid #0f172a; border-radius: 20px; padding: 40px; margin: 0 auto 40px; max-width: 480px; text-align: center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    h2 { margin: 0; color: #0f172a; font-size: 1.8rem; }
    p { color: #64748b; margin: 8px 0 24px; font-weight: 500; }
    img { width: 280px; height: 280px; border: 1px solid #e2e8f0; padding: 12px; border-radius: 16px; margin-bottom: 24px; }
    .btns { display: flex; gap: 12px; justify-content: center; }
    .btn { padding: 12px 24px; border-radius: 10px; font-weight: 700; cursor: pointer; border: none; text-decoration: none; font-size: 0.9rem; }
    .btn-print { background: #0f172a; color: white; }
    .btn-dl { background: #2563eb; color: white; }
    @media print { .btn-dl { display: none; } .sticker { box-shadow: none; border: 2px solid #000; } }
  </style>`;
  
  for (let i = 1; i <= count; i++) {
    const url = `${baseUrl}?site=${siteId}${count > 1 ? '&el='+i : ''}`;
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(url)}`;
    html += `<div class="sticker">
      <h2>${name}</h2>
      <p>${number} (LIFT ${i})</p>
      <img src="${qrUrl}">
      <div class="btns">
        <button class="btn btn-print" onclick="window.print()">PRINT LABEL</button>
        <a href="${qrUrl}" target="_blank" download="${siteId}_QR_L${i}.png" class="btn btn-dl">DOWNLOAD PNG</a>
      </div>
    </div>`;
  }
  return html;
}

function getSmartInfoFromData(data, worker, siteId, elevator) {
  const w = (worker || "").toString().toLowerCase(), s = (siteId || "").toString().toLowerCase(), e = (elevator || "1").toString().toLowerCase();
  const rangeStart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getLastRow() - 500;
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][1].toString().toLowerCase() === w && data[i][2].toString().toLowerCase() === s && (data[i][5] || "1").toString().toLowerCase() === e) {
      const actualRow = Math.max(1, rangeStart) + i;
      if (data[i][6] === 'IN') return { type: 'OUT', lastDevice: data[i][10], lastTime: new Date(data[i][0]), row: actualRow };
      return { type: 'IN', row: actualRow };
    }
  }
  return { type: 'IN' };
}

function checkSafetyAlerts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date().getTime();
  const active = {};
  for (let i = 1; i < data.length; i++) {
    const key = `${data[i][1]}|${data[i][2]}|${data[i][5]}`;
    active[key] = { 
      type: data[i][6], 
      time: new Date(data[i][0]), 
      row: i+1, 
      note: data[i][10], 
      siteId: data[i][2],
      siteName: data[i][3], 
      employee: data[i][1], 
      elevator: data[i][5] 
    };
  }
  for (const k in active) {
    const s = active[k];
    if (s.type === 'IN' && (now - s.time.getTime()) > (SAFETY_MINUTES * 60000)) {
      if (!String(s.note).includes('[Alert]')) {
        const durationMins = Math.floor((now - s.time.getTime()) / 60000);
        const emailBody = `⚠️ SAFETY ALERT: WORKER OVERDUE\n` +
                          `----------------------------------\n` +
                          `Employee: ${s.employee}\n` +
                          `Site: ${s.siteName} (${s.siteId})\n` +
                          `Elevator: ${s.elevator}\n` +
                          `Entered At: ${s.time.toLocaleString()}\n` +
                          `Current Duration: ${durationMins} minutes\n` +
                          `----------------------------------\n` +
                          `This worker has exceeded the safety limit of ${SAFETY_MINUTES} minutes.`;
                          
        MailApp.sendEmail(EMAIL_RECIPIENT, `Safety Alert: ${s.employee} @ ${s.siteName}`, emailBody);
        sheet.getRange(s.row, 11).setValue(String(s.note) + " [Alert Sent]");
      }
      // Always highlight if overdue, even if alert was already sent previously
      sheet.getRange(s.row, 1, 1, sheet.getLastColumn()).setBackground("#fee2e2");
    }
  }
}

function clearAllHighlights() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).setBackground(null);
}
