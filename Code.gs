// =====================================================
//  NBSC OJT Weekly Report System — Code.gs
//  Google Apps Script Backend
//  Northern Bukidnon State College | Form 9
// =====================================================
//
//  SETUP INSTRUCTIONS:
//  1. Go to script.google.com → New Project
//  2. Replace all code with this file
//  3. Update SHEET_ID and FOLDER_ID below
//  4. Deploy → New Deployment → Web App
//     • Execute as: Me
//     • Who has access: Anyone
//  5. Copy the Web App URL into the frontend sidebar
// =====================================================

// ── CONFIGURATION — UPDATE THESE ──────────────────
const SHEET_ID  = 'YOUR_GOOGLE_SHEET_ID_HERE';   // From Sheet URL: /spreadsheets/d/SHEET_ID/
const FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';   // From Drive URL: /drive/folders/FOLDER_ID
const SHEET_NAME = 'OJT Reports';                 // Tab name inside your Google Sheet
// ──────────────────────────────────────────────────

// ── HEADERS (must match column order in Sheet) ─────
const HEADERS = [
  'Timestamp', 'Name', 'Company', 'Week',
  'Objectives', 'Activities', 'Reflections',
  'Image URLs', 'Sig_Trainee', 'Sig_Supervisor', 'Sig_Coordinator'
];

// ── CORS HEADERS ───────────────────────────────────
function setCors(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── ENTRY POINTS ───────────────────────────────────
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || '';

  let result;
  try {
    if (action === 'getReport') {
      result = getReport(params.name, params.week, params.company);
    } else if (action === 'listReports') {
      result = listReports();
    } else {
      result = { status: 'ok', message: 'NBSC OJT Report API is running.' };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  const output = ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return setCors(output);
}

function doPost(e) {
  let body, result;
  try {
    body   = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if (action === 'saveReport') {
      result = saveReport(body);
    } else if (action === 'uploadFile') {
      result = uploadFile(body);
    } else {
      result = { status: 'error', message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  const output = ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return setCors(output);
}

// ── SAVE REPORT ────────────────────────────────────
function saveReport(data) {
  const sheet = getOrCreateSheet();

  // Check if record exists (same Name + Week + Company)
  const allData   = sheet.getDataRange().getValues();
  const nameCol   = HEADERS.indexOf('Name');
  const weekCol   = HEADERS.indexOf('Week');
  const companyCol= HEADERS.indexOf('Company');
  let   rowIdx    = -1;

  for (let i = 1; i < allData.length; i++) {
    if (
      String(allData[i][nameCol]).toLowerCase()    === String(data.name).toLowerCase() &&
      String(allData[i][weekCol]).toLowerCase()    === String(data.week).toLowerCase() &&
      String(allData[i][companyCol]).toLowerCase() === String(data.company).toLowerCase()
    ) {
      rowIdx = i + 1; // 1-indexed
      break;
    }
  }

  const row = [
    new Date().toISOString(),
    data.name            || '',
    data.company         || '',
    data.week            || '',
    data.objectives      || '',
    data.activities      || '',
    data.reflections     || '',
    data.imageUrls       || '',
    data.sig_trainee     || '',
    data.sig_supervisor  || '',
    data.sig_coordinator || ''
  ];

  if (rowIdx > 0) {
    // Update existing row
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    // Append new row
    sheet.appendRow(row);
  }

  // Auto-resize columns
  try { sheet.autoResizeColumns(1, HEADERS.length); } catch(e) {}

  return { status: 'ok', message: 'Report saved successfully.', rowUpdated: rowIdx > 0 };
}

// ── GET REPORT ─────────────────────────────────────
function getReport(name, week, company) {
  if (!name || !week) return { status: 'error', message: 'Name and Week are required.' };

  const sheet   = getOrCreateSheet();
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];

  for (let i = 1; i < allData.length; i++) {
    const row     = allData[i];
    const rowName = String(row[headers.indexOf('Name')]).toLowerCase().trim();
    const rowWeek = String(row[headers.indexOf('Week')]).toLowerCase().trim();
    const rowComp = String(row[headers.indexOf('Company')]).toLowerCase().trim();

    const nameMatch    = rowName === String(name).toLowerCase().trim();
    const weekMatch    = rowWeek === String(week).toLowerCase().trim();
    const companyMatch = !company || rowComp === String(company).toLowerCase().trim();

    if (nameMatch && weekMatch && companyMatch) {
      const data = {};
      headers.forEach((h, idx) => { data[h.toLowerCase().replace(/ /g,'_')] = row[idx]; });
      return { status: 'ok', data };
    }
  }

  return { status: 'not_found', message: 'No report found for the given Name and Week.' };
}

// ── LIST REPORTS ───────────────────────────────────
function listReports() {
  const sheet   = getOrCreateSheet();
  const allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) return { status: 'ok', reports: [] };

  const headers = allData[0];
  const reports = allData.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return { status: 'ok', reports };
}

// ── UPLOAD FILE TO GOOGLE DRIVE ────────────────────
function uploadFile(data) {
  if (!data.fileName || !data.data) {
    return { status: 'error', message: 'fileName and data are required.' };
  }

  const folder   = DriveApp.getFolderById(FOLDER_ID);
  const decoded  = Utilities.base64Decode(data.data);
  const blob     = Utilities.newBlob(decoded, data.mimeType || 'application/octet-stream', data.fileName);
  const file     = folder.createFile(blob);

  // Make file publicly readable
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId  = file.getId();
  const fileUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

  return { status: 'ok', fileId, fileUrl, fileName: data.fileName };
}

// ── HELPERS ────────────────────────────────────────
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1a4731');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

// ── OPTIONAL: Run this manually once to initialize ─
function initializeSheet() {
  const sheet = getOrCreateSheet();
  Logger.log('Sheet initialized: ' + sheet.getName());
  Logger.log('Sheet URL: ' + SpreadsheetApp.openById(SHEET_ID).getUrl());
}
