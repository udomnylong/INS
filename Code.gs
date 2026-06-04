/**
 * ================================================================
 *  Construction DMS — Google Apps Script (GAS)
 * ================================================================
 *  របៀបប្រើប្រាស់:
 *  1. បើក Google Sheet (Sheet DC)
 *  2. Extensions > Apps Script
 *  3. Copy Code នេះទៅ Editor
 *  4. Run: setupSheets()  ← បង្កើត Tabs + Headers + Sample Data
 *  5. Deploy > New Deployment > Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  6. Copy Web App URL → ដាក់ក្នុង HTML (CFG.apiUrl)
 * ================================================================
 */

// ────────────────────────────────────────────────────────────────
//  CONFIGURATION
// ────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = '1zOXTjsaJ4t909Qbw24KGO9JetaOePea_1m1e0TVEhUk';
const SECRET_TOKEN   = 'DMS_SECRET_2026'; // ប្ដូរ Token នេះ!

// Tab Headers Definition
const TAB_CONFIG = {
  User: {
    headers: ['Username','Password','FullName','Role','Email','Status'],
    color:   '#1a73e8',
    sample: [
      ['admin',  'admin123',  'Admin User',  'Administrator', 'admin@dms.com',  'Active'],
      ['sokha',  'sokha123',  'Sokha Chan',  'Inspector',     'sokha@dms.com',  'Active'],
      ['dara',   'dara123',   'Dara Meng',   'Engineer',      'dara@dms.com',   'Active'],
      ['ratha',  'ratha123',  'Ratha Ly',    'Site Manager',  'ratha@dms.com',  'Active'],
      ['vanna',  'vanna123',  'Vanna Keo',   'QC Officer',    'vanna@dms.com',  'Active'],
    ]
  },
  DTF: {
    headers: ['Code','Date','ProjectName','SubProject','Transmittal N°','DocumentTitle','Description','Document N°','TypeDocument','SubmitBy','ReceivedBy','Status','Remark'],
    color:   '#34a853',
    sample: []
  },
  Project: {
    headers: ['ProjectID','ProjectCode','ProjectName','Location','StartDate','FinishDate','Status'],
    color:   '#fbbc04',
    sample: [
      ['PRJ-001','ANJ-PH1','Angkor National Residence PH1','Siem Reap','2025-01-01','2027-12-31','Active'],
      ['PRJ-002','PPV','Phnom Penh Villa Project','Phnom Penh','2025-06-01','2026-12-31','Active'],
    ]
  },
  Scope: {
    headers: ['ProjectID','ProjectName','LOA','TypeLOA','Block','HouseType','HouseNo','Amount'],
    color:   '#ea4335',
    sample: [
      ['PRJ-001','Angkor National Residence PH1','LOA-001','Standard','Block A','VT1','H001','85000'],
      ['PRJ-001','Angkor National Residence PH1','LOA-002','Variation Order(+)','Block A','VT2','H002','92000'],
      ['PRJ-001','Angkor National Residence PH1','LOA-003','Standard','Block B','VT1','H003','85000'],
      ['PRJ-002','Phnom Penh Villa Project','LOA-004','Standard','Block C','VT3','H001','120000'],
    ]
  },
  IPT: {
    headers: ['Date','ProjectName','CodeITP','ContactPerson','MainBlock','SubBlock','HouseNo','Location','IssuedDate','Subject','WorkType','Description','PreparedBy','CheckBy','ApprovedBy'],
    color:   '#9c27b0',
    sample: []
  },
  INS: {
    headers: ['DateInput','Project','CodeINS','ContactPerson','MainBlock','SubBlock','HouseNo','Location','IssuedDate','Subject','Remark','ITPNo'],
    color:   '#00897b',
    sample: []
  },
  CheckList: {
    headers: ['TypeOfWork','CodeINS','DescriptionOfWork'],
    color:   '#0288d1',
    sample: [
      ['Foundation','INS-0001','Check rebar alignment and spacing per drawing'],
      ['Foundation','INS-0001','Verify concrete cover ≥ 50mm'],
      ['Foundation','INS-0001','Check formwork alignment plumb ±5mm'],
      ['Rebar','INS-0002','Bar size correct per structural drawing'],
      ['Rebar','INS-0002','Lap length adequate (40d minimum)'],
      ['Rebar','INS-0002','Links/stirrups correct spacing'],
      ['Concrete','INS-0003','Slump test 100-150mm conducted'],
      ['Concrete','INS-0003','Cube samples taken (3 sets minimum)'],
      ['Concrete','INS-0003','Curing started within 2 hours'],
      ['Formwork','INS-0004','Formwork clean and properly oiled'],
      ['Formwork','INS-0004','Props and bracing adequate'],
      ['Formwork','INS-0004','No gaps greater than 3mm'],
    ]
  },
  DocumentType: {
    headers: ['DocumentCode','DocumentType'],
    color:   '#5d4037',
    sample: [
      ['DT-001','Architectural Drawing'],['DT-002','Structural Drawing'],['DT-003','M&E Drawing'],
      ['DT-004','Shop Drawing'],['DT-005','As-Built Drawing'],['DT-006','Technical Specification'],
      ['DT-007','Method Statement'],['DT-008','Inspection Request'],['DT-009','Material Submittal'],
      ['DT-010','RFI - Request for Information'],['DT-011','Transmittal Letter'],['DT-012','Site Instruction'],
    ]
  },
};

// ────────────────────────────────────────────────────────────────
//  SETUP — Run this ONCE to initialize all Sheets
// ────────────────────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let created = 0, updated = 0;

  Object.entries(TAB_CONFIG).forEach(([tabName, cfg]) => {
    let sheet = ss.getSheetByName(tabName);

    // Create tab if not exists
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      created++;
      Logger.log('✅ Created tab: ' + tabName);
    } else {
      updated++;
      Logger.log('🔄 Found tab: ' + tabName);
    }

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, cfg.headers.length);
    headerRange.setValues([cfg.headers]);
    headerRange.setBackground(cfg.color);
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);

    // Freeze header row
    sheet.setFrozenRows(1);

    // Auto-resize columns
    sheet.autoResizeColumns(1, cfg.headers.length);

    // Add sample data if sheet is empty (no data rows)
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1 && cfg.sample && cfg.sample.length > 0) {
      sheet.getRange(2, 1, cfg.sample.length, cfg.headers.length).setValues(cfg.sample);
      Logger.log('  📝 Added ' + cfg.sample.length + ' sample rows to: ' + tabName);
    }

    // Apply alternating row colors for data
    if (lastRow > 1) {
      applyAlternatingColors(sheet, cfg.color);
    }
  });

  // Remove default "Sheet1" if empty
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
    Logger.log('🗑 Removed empty Sheet1');
  }

  // Show summary
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '✅ Setup Complete!',
    `Created: ${created} tabs\nUpdated: ${updated} tabs\n\nTabs: ${Object.keys(TAB_CONFIG).join(', ')}\n\nNext step: Deploy > Web App`,
    ui.ButtonSet.OK
  );
}

function applyAlternatingColors(sheet, accentColor) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (let i = 2; i <= lastRow; i++) {
    const color = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
  }
}

// ────────────────────────────────────────────────────────────────
//  WEB APP — doGet (Read)
// ────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'read';
  const tab    = e.parameter.tab    || '';
  const token  = e.parameter.token  || '';

  try {
    if (action === 'login') {
      return handleLogin(e.parameter.username || '', e.parameter.password || '');
    }
    if (action === 'read') {
      return handleRead(tab);
    }
    if (action === 'codes') {
      return handleGetCodes();
    }
    return jsonResp({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResp({ ok: false, error: err.message });
  }
}

// ────────────────────────────────────────────────────────────────
//  WEB APP — doPost (Write / Update / Delete)
// ────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const token  = body.token  || '';
    const tab    = body.tab    || '';

    // Verify token
    if (token !== SECRET_TOKEN) {
      return jsonResp({ ok: false, error: 'Unauthorized — Invalid token' });
    }

    if (action === 'append')   return handleAppend(tab, body.data);
    if (action === 'update')   return handleUpdate(tab, body.row, body.data);
    if (action === 'delete')   return handleDelete(tab, body.row);
    if (action === 'login')    return handleLogin(body.username, body.password);
    if (action === 'printDTF') return handlePrintDTF(body.data || {});
    if (action === 'printITP') return handlePrintITP(body.data || {});

    return jsonResp({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResp({ ok: false, error: err.message });
  }
}

// ────────────────────────────────────────────────────────────────
//  HANDLERS
// ────────────────────────────────────────────────────────────────

/** LOGIN: validate user from tab[User] */
function handleLogin(username, password) {
  if (!username || !password) {
    return jsonResp({ ok: false, error: 'Username/Password required' });
  }
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TAB_CONFIG.User ? 'User' : 'Users');
  if (!sheet) return jsonResp({ ok: false, error: 'User tab not found' });

  const rows = sheet.getDataRange().getValues();
  // headers: [0]Username [1]Password [2]FullName [3]Role [4]Email [5]Status
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] === username && r[1] === password && (r[5] || 'Active') === 'Active') {
      return jsonResp({
        ok:       true,
        user: {
          username: r[0],
          fullName: r[2] || r[0],
          role:     r[3] || 'User',
          email:    r[4] || '',
          status:   r[5] || 'Active',
        }
      });
    }
  }
  return jsonResp({ ok: false, error: 'Invalid credentials' });
}

/** READ: get all rows from a tab */
function handleRead(tabName) {
  if (!tabName) return jsonResp({ ok: false, error: 'tab required' });
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResp({ ok: false, error: 'Tab not found: ' + tabName });

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0] || [];
  const tz      = Session.getScriptTimeZone();
  const data    = [];

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue; // skip empty rows
    const obj = { _row: i + 1 };
    headers.forEach((h, ci) => {
      const val = rows[i][ci];
      if (val instanceof Date) {
        // Format Date → yyyy-MM-dd
        obj[h] = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      } else if (val !== undefined && val !== null) {
        obj[h] = val.toString();
      } else {
        obj[h] = '';
      }
    });
    data.push(obj);
  }

  return jsonResp({ ok: true, tab: tabName, count: data.length, data });
}

/** APPEND: add a new row */
function handleAppend(tabName, rowData) {
  if (!tabName || !rowData) return jsonResp({ ok: false, error: 'tab and data required' });
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResp({ ok: false, error: 'Tab not found: ' + tabName });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowData[h] !== undefined ? rowData[h] : '');

  sheet.appendRow(row);

  // Apply alternating color to new row
  const newRow = sheet.getLastRow();
  const color  = newRow % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(newRow, 1, 1, headers.length).setBackground(color);

  return jsonResp({ ok: true, action: 'appended', row: newRow });
}

/** UPDATE: update an existing row */
function handleUpdate(tabName, rowIndex, rowData) {
  if (!tabName || !rowIndex || !rowData) return jsonResp({ ok: false, error: 'tab, row, data required' });
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResp({ ok: false, error: 'Tab not found: ' + tabName });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowData[h] !== undefined ? rowData[h] : '');

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);

  return jsonResp({ ok: true, action: 'updated', row: rowIndex });
}

/** DELETE: delete a row */
function handleDelete(tabName, rowIndex) {
  if (!tabName || !rowIndex) return jsonResp({ ok: false, error: 'tab and row required' });
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResp({ ok: false, error: 'Tab not found: ' + tabName });

  sheet.deleteRow(parseInt(rowIndex));
  return jsonResp({ ok: true, action: 'deleted', row: rowIndex });
}

/** GET CODES: return next auto-codes for DTF, ITP, INS */
function handleGetCodes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function getNextCode(tabName, colIndex, prefix, padLen) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() <= 1) return prefix + String(1).padStart(padLen, '0');
    const values = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getValues();
    let max = 0;
    values.forEach(r => {
      if (!r[0]) return;
      const n = parseInt(r[0].toString().split('-').pop()) || parseInt(r[0]) || 0;
      if (n > max) max = n;
    });
    return prefix + String(max + 1).padStart(padLen, '0');
  }

  return jsonResp({
    ok:   true,
    codes: {
      dtf: getNextCode('DTF',  1, '',                    4),
      itp: getNextCode('IPT',  3, 'HE-ANJ-PH1-IPT-LHB-', 4),
      ins: getNextCode('INS',  3, 'INS-',                 4),
    }
  });
}

// ────────────────────────────────────────────────────────────────
//  UTILITY
// ────────────────────────────────────────────────────────────────
function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────────
//  TRIGGERS — Auto-format on edit
// ────────────────────────────────────────────────────────────────
function onEdit(e) {
  const sheet    = e.range.getSheet();
  const row      = e.range.getRow();
  const sheetName = sheet.getName();

  if (row <= 1) return; // skip header

  // Auto-generate Code for DTF
  if (sheetName === 'DTF' && e.range.getColumn() === 1 && !e.value) {
    const nextCode = getNextDTFCode();
    e.range.setValue(nextCode);
  }

  // Auto-set Date if empty
  if (['DTF','IPT','INS'].includes(sheetName) && row > 1) {
    const dateCol = sheetName === 'INS' ? 1 : (sheetName === 'DTF' ? 2 : 1);
    const dateCell = sheet.getRange(row, dateCol);
    if (!dateCell.getValue()) {
      dateCell.setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    }
  }

  // Apply alternating row color
  const color = row % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(color);
}

function getNextDTFCode() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DTF');
  if (!sheet || sheet.getLastRow() <= 1) return '0001';
  const codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const max   = codes.reduce((m, c) => Math.max(m, parseInt(c) || 0), 0);
  return String(max + 1).padStart(4, '0');
}

// ────────────────────────────────────────────────────────────────
//  MENU — Custom menu in Google Sheets
// ────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Construction DMS')
    .addItem('⚙ Setup All Sheets',     'setupSheets')
    .addSeparator()
    .addItem('📊 Export DTF to PDF',    'exportDtfPdf')
    .addItem('🔄 Refresh All Formats',  'refreshFormats')
    .addSeparator()
    .addItem('👥 Manage Users',         'openUserManager')
    .addItem('📋 View Web App URL',     'showWebAppUrl')
    .addToUi();
}

// ────────────────────────────────────────────────────────────────
//  EXPORT — Generate PDF of current sheet
// ────────────────────────────────────────────────────────────────
function exportDtfPdf() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = ss.getSheetByName('DTF');
  const sheetId = sheet.getSheetId();
  const ssId    = ss.getId();

  const pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?`
    + `exportFormat=pdf&format=pdf`
    + `&gid=${sheetId}`
    + `&size=A4&portrait=false&fitw=true`
    + `&sheetnames=false&printtitle=false&pagenumbers=false`
    + `&gridlines=false&fzr=true`;

  const ui = SpreadsheetApp.getUi();
  ui.alert('📄 DTF PDF', 'Open this URL to download PDF:\n\n' + pdfUrl, ui.ButtonSet.OK);
}

// ────────────────────────────────────────────────────────────────
//  REFRESH FORMATS
// ────────────────────────────────────────────────────────────────
function refreshFormats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.entries(TAB_CONFIG).forEach(([tabName, cfg]) => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    // Re-style header
    const hr = sheet.getRange(1, 1, 1, cfg.headers.length);
    hr.setBackground(cfg.color);
    hr.setFontColor('#ffffff');
    hr.setFontWeight('bold');
    hr.setFontSize(11);
    // Alternating rows
    applyAlternatingColors(sheet, cfg.color);
    sheet.autoResizeColumns(1, cfg.headers.length);
  });
  SpreadsheetApp.getUi().alert('✅ Formats refreshed!');
}

// ────────────────────────────────────────────────────────────────
//  USER MANAGER
// ────────────────────────────────────────────────────────────────
function openUserManager() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('User');
  if (!sheet) { SpreadsheetApp.getUi().alert('User tab not found. Run Setup first.'); return; }
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    '👥 User Manager',
    'Tab[User] is now active.\n\nColumns:\n- Username (ចូល)\n- Password (លេខសម្ងាត់)\n- FullName (ឈ្មោះពេញ)\n- Role (តួនាទី)\n- Email (អ៊ីម៉ែល)\n- Status: Active / Inactive',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ────────────────────────────────────────────────────────────────
//  PRINT DTF FORM — copy DTF_Form tab, fill data, export PDF
// ────────────────────────────────────────────────────────────────
function handlePrintDTF(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const templateSheet = ss.getSheetByName('DTF_Form');
    if (!templateSheet) return jsonResp({ ok: false, error: 'DTF_Form tab រកមិនឃើញ' });

    // Copy template → temp sheet
    const tempSheet = templateSheet.copyTo(ss);
    const tempName  = '_DTF_' + Date.now();
    tempSheet.setName(tempName);
    ss.setActiveSheet(tempSheet);
    ss.moveActiveSheet(ss.getNumSheets()); // move to last

    // ── Parse date (frontend sends DD-MMM-YYYY via fmtDateDMY) ────
    const MONTHS_S = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    let yy = '', mm = '', dd = '', dateStr = '';
    if (data.date) {
      const parts = String(data.date).split('-');
      if (parts.length === 3 && MONTHS_S.indexOf(parts[1]) >= 0) {
        // DD-MMM-YYYY format from frontend
        dd      = parts[0].padStart(2, '0');
        mm      = String(MONTHS_S.indexOf(parts[1]) + 1).padStart(2, '0');
        yy      = parts[2].slice(-2);
        dateStr = data.date; // keep DD-MMM-YYYY for template
      } else {
        // Fallback: ISO yyyy-MM-dd
        const dt = new Date(data.date);
        if (!isNaN(dt.getTime())) {
          dd      = String(dt.getDate()).padStart(2, '0');
          mm      = String(dt.getMonth() + 1).padStart(2, '0');
          yy      = String(dt.getFullYear()).slice(-2);
          dateStr = dd + '-' + MONTHS_S[dt.getMonth()] + '-' + dt.getFullYear();
        }
      }
    }

    // ── Fill main fields ─────────────────────────────────────────
    const set = function(cell, val) {
      if (val !== undefined && val !== null && val !== '') {
        try { tempSheet.getRange(cell).setValue(val); } catch(e) {}
      }
    };
    set('B6', data.project);
    set('D6', data.subProject);
    try {
      tempSheet.getRange('D6').setHorizontalAlignment('center').setVerticalAlignment('middle');
    } catch(e) {}
    set('J6', data.transmittalNo);
    set('L6', data.docTitle);
    set('R6', dateStr);
    set('M9', yy);
    set('M10', mm);
    set('M11', dd);

    // ── Description + Item row 1 ─────────────────────────────────
    set('B16', '1');
    set('C16', data.description);
    set('J16', data.docNo);

    // ── Submit By / Remark ───────────────────────────────────────
    set('C48', data.submitBy);
    set('G48', data.remark);

    // ── Status ───────────────────────────────────────────────────
    // set('P48', data.status); // Adjust cell to match your template

    // ── Type Document checkboxes ─────────────────────────────────
    // cell → matching typeDoc value(s) in our system
    const chkMap = {
      'B42': ['Approval'],
      'G42': ['Signature & Return'],
      'K42': ['Approved Document Submission'],
      'B43': ['Comment'],
      'G43': ['Site Instruction'],
      'B44': ['Construction'],
      'B45': ['Replace Drawing'],
      'G44': ['Other: For Action and Request'],
      'G45': ['Request'],
    };
    const td = (data.typeDoc || '').trim();
    Object.entries(chkMap).forEach(function(entry) {
      const cell = entry[0], types = entry[1];
      // Explicitly set every checkbox: TRUE for match, FALSE for all others
      try { tempSheet.getRange(cell).setValue(types.indexOf(td) >= 0); } catch(e) {}
    });

    SpreadsheetApp.flush();

    // ── Export sheet as PDF via Drive URL ────────────────────────
    const gid   = tempSheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const url   = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
      '/export?format=pdf&gid=' + gid +
      '&size=A4&portrait=true&fitw=true' +
      '&top_margin=0.1969&bottom_margin=0.1969&left_margin=0.1969&right_margin=0.3937' +
      '&gridlines=false&printtitle=false&sheetnames=false' +
      '&pagenumbers=false&bg=false&attachment=true';

    const resp    = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const pdfB64  = Utilities.base64Encode(resp.getContent());

    // ── Clean up temp sheet ──────────────────────────────────────
    ss.deleteSheet(tempSheet);

    return jsonResp({ ok: true, pdf: pdfB64, filename: (data.docNo || data.code || 'DTF') + '.pdf' });

  } catch(err) {
    return jsonResp({ ok: false, error: err.toString() });
  }
}

// ────────────────────────────────────────────────────────────────
//  PRINT ITP FORM  (ITP_FORM tab → PDF)
// ────────────────────────────────────────────────────────────────
function handlePrintITP(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const templateSheet = ss.getSheetByName('ITP_FORM');
    if (!templateSheet) return jsonResp({ ok: false, error: 'ITP_FORM tab រកមិនឃើញ' });

    // Copy template → temp sheet
    const tempSheet = templateSheet.copyTo(ss);
    const tempName  = '_ITP_' + Date.now();
    tempSheet.setName(tempName);
    ss.setActiveSheet(tempSheet);
    ss.moveActiveSheet(ss.getNumSheets());

    // ── Parse dates ──────────────────────────────────────────────
    const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    let issuedStr = '', dateB36 = '';
    if (data.issuedDate) {
      const dt = new Date(data.issuedDate);
      if (!isNaN(dt)) {
        const dd  = String(dt.getDate()).padStart(2, '0');
        const mon = MONTHS[dt.getMonth()];
        const yyyy = dt.getFullYear();
        issuedStr = dd + ' ' + mon + ' ' + yyyy;   // DD MMM YYYY
      }
    }
    if (data.date) {
      const dt = new Date(data.date);
      if (!isNaN(dt)) {
        const dd  = String(dt.getDate()).padStart(2, '0');
        const mon = MONTHS[dt.getMonth()];
        const yyyy = dt.getFullYear();
        dateB36 = 'Date : ' + dd + ' ' + mon + ' ' + yyyy;
      }
    }

    // ── Helper ───────────────────────────────────────────────────
    const set = function(cell, val) {
      if (val !== undefined && val !== null && val !== '') {
        try { tempSheet.getRange(cell).setValue(val); } catch(e) {}
      }
    };

    // ── Header fields ────────────────────────────────────────────
    set('E3', data.projectName);          // Project
    set('E4', data.codeITP);             // CODE ITP
    set('D6', data.contactPerson);       // Contact Person
    set('D7', data.mainBlock);           // Main Block
    set('D8', data.subBlock);            // Sub Block
    set('H8', data.houseNo);             // House N°
    set('D9', data.location);            // Location
    set('D10', issuedStr);               // Issued Date
    set('B36', dateB36);                 // Date : DD MMM YYYY

    // ── D11: Subject label ───────────────────────────────────────
    if (data.subject) {
      set('D11', 'Request for Inspection and test plan of ' + data.subject);
    }

    // ── Work Type checkboxes C14-C17 ─────────────────────────────
    const wt = data.workType || '';
    const chkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    ['B14','B15','B16','B17'].forEach(function(cell) {
      try { tempSheet.getRange(cell).setDataValidation(chkRule); } catch(e) {}
    });
    try { tempSheet.getRange('B14').setValue(wt === 'Structure Work'); } catch(e) {}
    try { tempSheet.getRange('B15').setValue(wt === 'Architecture Work'); } catch(e) {}
    try { tempSheet.getRange('B16').setValue(wt === 'MEP Work'); } catch(e) {}
    try { tempSheet.getRange('B17').setValue(wt === 'Other'); } catch(e) {}

    // ── Description → C21:C26 (split by newline, max 6 rows) ────
    const lines = (data.description || '').split('\n').slice(0, 6);
    for (let i = 0; i < 6; i++) {
      try { tempSheet.getRange('C' + (21 + i)).setValue(lines[i] || ''); } catch(e) {}
    }

    // ── Prepared / Check / Approved By ──────────────────────────
    set('B33', 'Name : ' + (data.preparedBy || ''));
    set('F33', 'Name : ' + (data.checkBy || ''));
    set('H33', 'Name : ' + (data.approvedBy || 'Mr. Seng Nora'));

    SpreadsheetApp.flush();

    // ── Export as PDF ────────────────────────────────────────────
    const gid   = tempSheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const url   = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
      '/export?format=pdf&gid=' + gid +
      '&size=A4&portrait=true&fitw=true' +
      '&top_margin=0.1969&bottom_margin=0.1969&left_margin=0.1969&right_margin=0.3937' +
      '&gridlines=false&printtitle=false&sheetnames=false' +
      '&pagenumbers=false&bg=false&attachment=true';

    const resp   = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const pdfB64 = Utilities.base64Encode(resp.getContent());

    ss.deleteSheet(tempSheet);

    return jsonResp({ ok: true, pdf: pdfB64, filename: (data.codeITP || 'ITP') + '.pdf' });

  } catch(err) {
    return jsonResp({ ok: false, error: err.toString() });
  }
}

// ────────────────────────────────────────────────────────────────
//  SHOW WEB APP URL
// ────────────────────────────────────────────────────────────────
function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui  = SpreadsheetApp.getUi();
  if (url) {
    ui.alert('🌐 Web App URL', 'Copy URL below and paste in HTML (CFG.apiUrl):\n\n' + url, ui.ButtonSet.OK);
  } else {
    ui.alert('⚠ Not Deployed', 'Please deploy as Web App first:\nDeploy > New Deployment > Web App', ui.ButtonSet.OK);
  }
}

// ────────────────────────────────────────────────────────────────
//  DATA VALIDATION — Add dropdown validations to sheets
// ────────────────────────────────────────────────────────────────
function addDataValidations() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // DTF — TypeDocument dropdown (col I) and Status dropdown (col L)
  // New header order: A=Code B=Date C=ProjectName D=SubProject E=Transmittal N°
  //                   F=DocumentTitle G=Description H=Document N° I=TypeDocument
  //                   J=SubmitBy K=ReceivedBy L=Status M=Remark
  const dtfSheet = ss.getSheetByName('DTF');
  if (dtfSheet) {
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Approval','Signature & Return','Approved Document Submission',
        'Comment','Site Instruction','Construction','Replace Drawing',
        'Request','Other: For Action and Request'
      ], true).build();
    dtfSheet.getRange('I2:I1000').setDataValidation(typeRule);

    const dtfStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending','Approved','Rejected','On Hold'], true).build();
    dtfSheet.getRange('L2:L1000').setDataValidation(dtfStatusRule);
  }

  // Project — Status dropdown (col G after adding ProjectCode)
  const projSheet = ss.getSheetByName('Project');
  if (projSheet) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active','Completed','On Hold','Cancelled'], true).build();
    projSheet.getRange('G2:G1000').setDataValidation(statusRule);
  }

  // Scope — TypeLOA dropdown
  const scopeSheet = ss.getSheetByName('Scope');
  if (scopeSheet) {
    const loaRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Variation Order(+)','Variation Order(-)','Standard'], true).build();
    scopeSheet.getRange('D2:D1000').setDataValidation(loaRule);
  }

  // User — Status dropdown
  const userSheet = ss.getSheetByName('User');
  if (userSheet) {
    const userStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active','Inactive'], true).build();
    userSheet.getRange('F2:F1000').setDataValidation(userStatusRule);
  }

  SpreadsheetApp.getUi().alert('✅ Data Validations added!');
}
