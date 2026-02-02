/*****************************************
 * CASH BOOK SYSTEM - Code.gs created by Dev.@shah_md_Wasi
 ****************************************/

const CONFIG = {
  cashFormHtmlFile: 'managerDashboard', // Updated to point to new Dashboard

  // UPDATED SALES SOURCE (xternal)
  sales: {
    spreadsheetId: '1jda6NtDS8QQo9D-v2tN7t5Rxk6NvHFsDDORP1qdtbVo',
    sheetName: 'Sales Data',
    headerRow: 1,          // data starts row 2
    dateColLetter: 'A',    // Date column A (yyyy-mm-dd)
    amountColLetter: 'L',  // Amount column L
    minDateKey: '2026-01-01'
  },

  ledger: {
    sheetName: 'Ledger',

    headerRow: 2,
    headers: {
      date: 'Date',
      header: 'Headers',
      description: 'Description',
      in: 'In',
      out: 'Out',
      handoverTo: 'Handover To',
      balance: 'Balance',
      classification: 'classification'
    },
    dateNumberFormat: 'yyyy-mm-dd'
  },

  settings: {
    sheetName: 'Settings',
    headerRow: 1,
    colHeader: 1,           // A
    colClassification: 2,   // B
    colInternalId: 3        // C
  },

  // PAYROLL CONFIG
  payroll: {
    masterSheet: 'Staff_Master',
    attendanceSheet: 'Attendance',
    headerSalary: 'Staff Salary',
    headerAdvance: 'Staff Advance',
    otRateMultiplier: 1.5   // Default OT multiplier if not specified per staff
  },

  accountsSummarySheetName: 'Accounts_Summary',
  importLogSheetName: 'Sales_Import_Log',

  ghorooaLiabilityRate: 0.65,
  internalIds: {
    dailySales: '701',
    ghorooaPayment: '702'
  }
};

/** ========= MENU ========= */
function onOpen() {
  try { initializeSheets(); } catch (e) { Logger.log(e); }

  SpreadsheetApp.getUi()
    .createMenu('Mod.geek')
    .addItem('Manager Dashboard', 'openManagerDashboard')
    .addSeparator()
    .addItem('Process Daily Sales', 'processDailySales')
    .addSeparator()
    .addItem('Diagnostics', 'diagnoseSetup')
    .addSeparator()
    .addItem('Refresh Accounts Summary', 'updateAccountsSummary')
    .addToUi();

  // POS Sync menu (from pos_sync.gs)
  try { posSyncOnOpen_(); } catch (e) {}
}

function openManagerDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('managerDashboard')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manager Dashboard');
}

/** ========= UTILITIES ========= */
function normalizeHeader_(s) {
  return (s ?? '')
    .toString()
    .replace(/\u00A0/g, ' ')
    .replace(/[\t\r\n]+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function colLetterToIndex_(letter) {
  const s = (letter || '').toString().trim().toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n; // 1-based
}

function extractStaffIdTag_(text) {
  // Accepts: "ID:12 Name", "ID=12", "STAFF_ID:12", "staff_id=12"
  const s = String(text || "");
  const m = s.match(/(?:^|\b)(?:ID|STAFF_ID)\s*[:=]\s*(\d+)\b/i);
  return m ? m[1] : null;
}

function parseYMD_(s) {
  const m = (s || '').toString().trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(y, mo - 1, d, 12, 0, 0);
  return isNaN(dt) ? null : dt;
}

// dd/mm/yyyy parser (Sales source)
function parseDMY_(v) {
  if (v instanceof Date && !isNaN(v)) return v;

  const s = (v ?? '').toString().trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (!m) return null;

  const dd = Number(m[1]);
  const mm = Number(m[2]);
  let yyyy = Number(m[3]);

  if (!Number.isFinite(dd) || !Number.isFinite(mm) || !Number.isFinite(yyyy)) return null;
  if (yyyy < 100) yyyy = 2000 + yyyy;

  const d = new Date(yyyy, mm - 1, dd, 12, 0, 0);
  return isNaN(d) ? null : d;
}

function toDateKey_(dateObj, tz) {
 return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
}
 
// yyyy-mm-dd -> Google Sheets date serial (timezone-proof)
function dateKeyToSerial_(key) {
 const m = (key || '').toString().trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
 if (!m) return null;
 const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
 // Google Sheets "day 0" is 1899-12-30
 const serial = (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
 return Number.isFinite(serial) ? serial : null;
}


function parseAmount_(v) {
  if (typeof v === 'number') return v;
  const s = (v ?? '').toString().trim();
  if (!s) return NaN;
  const cleaned = s.replace(/,/g, '').replace(/[^\d.\-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureImportLog_() {
  const sh = ensureSheet_(CONFIG.importLogSheetName);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 7).setValues([[
      'Timestamp', 'Status', 'From Date', 'Days Imported', 'Total Amount', 'First Day', 'Last Day'
    ]]).setFontWeight('bold');
  }
  return sh;
}

function logImport_(status, fromDateKey, daysImported, totalAmount, firstDay, lastDay) {
  try {
    const sh = ensureImportLog_();
    sh.appendRow([
      new Date(),
      status,
      fromDateKey || '',
      daysImported || 0,
      totalAmount || 0,
      firstDay || '',
      lastDay || ''
    ]);
  } catch (e) {
    Logger.log('Log failed: ' + e);
  }
}

function getHeaderMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  const row = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  row.forEach((h, i) => {
    const k = normalizeHeader_(h);
    if (k) map[k] = i + 1;
  });
  return map;
}

function getLedgerCols_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName(CONFIG.ledger.sheetName);
  if (!ledger) throw new Error(`Ledger sheet not found: ${CONFIG.ledger.sheetName}`);

  const hm = getHeaderMap_(ledger, CONFIG.ledger.headerRow);
  const h = CONFIG.ledger.headers;

  const cols = {
    sheet: ledger,
    headerRow: CONFIG.ledger.headerRow,
    date: hm[normalizeHeader_(h.date)],
    header: hm[normalizeHeader_(h.header)],
    description: hm[normalizeHeader_(h.description)],
    in: hm[normalizeHeader_(h.in)],
    out: hm[normalizeHeader_(h.out)],
    handoverTo: hm[normalizeHeader_(h.handoverTo)],
    balance: hm[normalizeHeader_(h.balance)],
    classification: hm[normalizeHeader_(h.classification)]
  };

  const missing = Object.keys(cols).filter(k => !['sheet', 'headerRow'].includes(k) && !cols[k]);
  if (missing.length) {
    throw new Error(
      `Ledger header mismatch. Missing: ${missing.join(', ')}.\n` +
      `Expected in row ${CONFIG.ledger.headerRow}: ${Object.values(CONFIG.ledger.headers).join(' | ')}`
    );
  }
  return cols;
}

/** ========= SETTINGS ========= */
function getHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.settings.sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= CONFIG.settings.headerRow) return [];

  const values = sh.getRange(CONFIG.settings.headerRow + 1, 1, lastRow - CONFIG.settings.headerRow, 2).getValues();
  return values
    .map(r => ({
      header: (r[0] || '').toString().trim(),
      classification: (r[1] || '').toString().trim()
    }))
    .filter(x => x.header && x.classification);
}

function getInternalIdsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.settings.sheetName);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (lastRow <= CONFIG.settings.headerRow) return {};

  const data = sh.getRange(CONFIG.settings.headerRow + 1, 1, lastRow - CONFIG.settings.headerRow, 3).getValues();
  const map = {};
  data.forEach(r => {
    const headerName = (r[CONFIG.settings.colHeader - 1] || '').toString().trim();
    const internalId = (r[CONFIG.settings.colInternalId - 1] || '').toString().trim();
    if (headerName && internalId) map[internalId] = headerName;
  });
  return map;
}

/** ========= CORE: Manual Entry ========= */
function saveBulkEntries(entries) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another process is running. Try again in 30 seconds.');
  }

  try {
    const cols = getLedgerCols_();
    const ledger = cols.sheet;
    const settings = getHeaders();

    const classificationMap = settings.reduce((map, obj) => {
        map[obj.header] = obj.classification;
        return map;
    }, {});

    let currentBalance = 0;
    const ledgerLastRow = ledger.getLastRow();
    if (ledgerLastRow > cols.headerRow) {
      currentBalance = Number(ledger.getRange(ledgerLastRow, cols.balance).getValue()) || 0;
    }

    const tableMinCol = Math.min(cols.date, cols.header, cols.description, cols.in, cols.out, cols.handoverTo, cols.balance, cols.classification);
    const tableMaxCol = Math.max(cols.date, cols.header, cols.description, cols.in, cols.out, cols.handoverTo, cols.balance, cols.classification);
    const tableWidth = tableMaxCol - tableMinCol + 1;

    const rowsToWrite = [];

    entries.forEach(entry => {
      const { date, header, description, handoverTo, in: inVal, out: outVal } = entry;
      
      // Validation
      if (!date) throw new Error("Entry missing Date");
      if (!header) throw new Error("Entry missing Header");
      if (isNaN(inVal) || isNaN(outVal)) throw new Error("Entry has invalid amount");

      currentBalance = currentBalance + inVal - outVal;
      
      const classification = classificationMap[header] || 'General';

      const row = new Array(tableWidth).fill('');
      row[cols.date - tableMinCol] = new Date(date);
      row[cols.header - tableMinCol] = header;
      row[cols.description - tableMinCol] = description;
      row[cols.in - tableMinCol] = inVal;
      row[cols.out - tableMinCol] = outVal;
      row[cols.handoverTo - tableMinCol] = handoverTo;
      row[cols.balance - tableMinCol] = currentBalance;
      row[cols.classification - tableMinCol] = classification;
      
      rowsToWrite.push(row);
    });

    if (rowsToWrite.length > 0) {
      const startRow = ledger.getLastRow() + 1;
      ledger.getRange(startRow, tableMinCol, rowsToWrite.length, tableWidth).setValues(rowsToWrite);
      ledger.getRange(startRow, cols.date, rowsToWrite.length, 1).setNumberFormat(CONFIG.ledger.dateNumberFormat);
    }
    
    updateAccountsSummary();
    
    return { success: true, message: `${rowsToWrite.length} entries saved.` };

  } catch(e) {
    Logger.log(e.stack || e);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/***********************
 * EVENT LOGER + PENDING
 * Truth source: Ledger (via saveBulkEntries) + Pending_Log (append-only)
 ***********************/

// ===== Fixed names (your spec) =====
const PENDING_LOG_SHEET = 'Pending_Log';
const SALES_SHEET_NAME  = 'Sales Data';
const SALES_INVOICE_COL = 4;   // D
const SALES_AMOUNT_COL  = 12;  // L (final_due_amount)

// Ledger headers (must exist in Settings for proper classification)
const H_DUE_CREATE = 'Unpaid - (Due)';
const H_DUE_SETTLE = 'Due Received';
const H_LOAN_TAKEN = 'Receiving Debt';
const H_LOAN_REPAID = 'Paying off Debt';
const H_LOAN_GIVEN = 'Advancing credit';
const H_LOAN_RCVD  = 'Credit settled';

// Types
const T_CUST = 'CustomerDue';
const T_LT   = 'LoanTaken';
const T_LG   = 'LoanGiven';

// Events
const E_CREATE = 'CREATE';
const E_SETTLE = 'SETTLE';

// ===== Date helpers (stable for Asia/Dhaka) =====
function dateKey_(d) {
  const tz = Session.getScriptTimeZone();
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return '';
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
}

// IMPORTANT: ledger date is created with +06:00 to avoid UTC shift confusion
function ledgerDate_(dateKey) {
  const s = String(dateKey || '').trim();
  if (!s) return new Date();
  return new Date(`${s}T00:00:00+06:00`);
}

function nowIso_() {
  return new Date().toISOString();
}

// ===== Pending_Log ensure (append-only) =====
function ensurePendingLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PENDING_LOG_SHEET);
  if (!sh) sh = ss.insertSheet(PENDING_LOG_SHEET);

  const headers = ['date_key','type','event','ref','party','amount','note','created_at_iso'];
  const lastRow = sh.getLastRow();

  if (lastRow === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  } else {
    const cur = sh.getRange(1,1,1,headers.length).getValues()[0].map(x => String(x||'').trim());
    if (cur.join('|') !== headers.join('|')) {
      throw new Error(`Pending_Log header mismatch. Refusing to continue.`);
    }
  }
  return sh;
}

function readPendingLog_() {
  const sh = ensurePendingLog_();
  const last = sh.getLastRow();
  if (last <= 1) return [];
  const values = sh.getRange(2,1,last-1,8).getValues();
  return values.map(r => ({
    date_key: String(r[0]||'').trim(),
    type: String(r[1]||'').trim(),
    event: String(r[2]||'').trim(),
    ref: String(r[3]||'').trim(),
    party: String(r[4]||'').trim(),
    amount: Number(r[5]||0) || 0,
    note: String(r[6]||'').trim(),
    created_at_iso: String(r[7]||'').trim(),
  }));
}

function appendPendingLog_(rows) {
  const sh = ensurePendingLog_();
  if (!rows || !rows.length) return;
  const start = sh.getLastRow() + 1;
  sh.getRange(start,1,rows.length,8).setValues(rows);
}

// ===== Sales lookup (invoice -> final_due_amount) =====
function getDueAmountByInvoice_(invoiceNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) throw new Error(`Sales sheet not found: ${SALES_SHEET_NAME}`);

  const last = sh.getLastRow();
  if (last < 2) throw new Error('Sales Data empty');

  const inv = String(invoiceNo || '').trim();
  if (!inv) throw new Error('InvoiceNo required');

  const width = SALES_AMOUNT_COL - SALES_INVOICE_COL + 1; // D..L
  const values = sh.getRange(2, SALES_INVOICE_COL, last-1, width).getValues();

  for (let i=0; i<values.length; i++) {
    const rowInv = String(values[i][0] || '').trim(); // D
    if (rowInv === inv) {
      const amt = Number(values[i][width-1] || 0);    // L
      if (!isFinite(amt)) throw new Error(`Invalid amount in Sales Data for invoice ${inv}`);
      return amt;
    }
  }
  throw new Error(`Invoice not found in Sales Data (col D): ${inv}`);
}

// ===== Open/Remaining calculation =====
function hasCreate_(rows, type, ref) {
  return rows.some(r => r.type === type && r.event === E_CREATE && r.ref === ref);
}

function remainingByRef_(rows, type, ref) {
  let created = 0, settled = 0;
  let party = '';
  let createdDate = '';

  rows.forEach(r => {
    if (r.type !== type || r.ref !== ref) return;
    if (!party && r.party) party = r.party;
    if (!createdDate && r.event === E_CREATE) createdDate = r.date_key;
    if (r.event === E_CREATE) created += r.amount;
    if (r.event === E_SETTLE) settled += r.amount;
  });

  return { created, settled, remaining: created - settled, party, createdDate };
}

// ===== LoanId sequencing =====
function nextSeq_(key) {
  const p = PropertiesService.getDocumentProperties();
  const cur = Number(p.getProperty(key) || '0') || 0;
  const next = cur + 1;
  p.setProperty(key, String(next));
  return next;
}

function makeLoanId_(prefix, dateKey) {
  const ymd = String(dateKey || '').replace(/-/g,'');
  const seqKey = `SEQ_${prefix}_${ymd}`;
  const n = nextSeq_(seqKey);
  return `${prefix}-${ymd}-${String(n).padStart(4,'0')}`;
}

// =====================================================
// PUBLIC API — Event Loger (CREATE)
// =====================================================

function createCustomerDue(date, invoiceNo, customerName) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const inv = String(invoiceNo||'').trim();
  const party = String(customerName||'').trim();
  if (!inv) return { success:false, message:'InvoiceNo required' };
  if (!party) return { success:false, message:'Customer name required' };

  const logRows = readPendingLog_();
  if (hasCreate_(logRows, T_CUST, inv)) {
    return { success:false, message:`Already exists: CustomerDue CREATE for invoice ${inv}` };
  }

  const amt = getDueAmountByInvoice_(inv);
  if (!(amt > 0)) return { success:false, message:`Invalid due amount for invoice ${inv}: ${amt}` };

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_DUE_CREATE,
    description: `DUE_CREATE | INV:${inv}`,
    handoverTo: party,
    in: 0,
    out: amt
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_CUST, E_CREATE, inv, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`CustomerDue created: ${inv} amount=${amt}` };
}

function createLoanTaken(date, lenderName, amount) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const party = String(lenderName||'').trim();
  const amt = Number(amount||0);
  if (!party) return { success:false, message:'Lender name required' };
  if (!isFinite(amt) || amt <= 0) return { success:false, message:'Invalid amount' };

  const loanId = makeLoanId_('LT', dk);

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_LOAN_TAKEN,
    description: `LOAN_TAKEN | ID:${loanId}`,
    handoverTo: party,
    in: amt,
    out: 0
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_LT, E_CREATE, loanId, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`LoanTaken created: ${loanId} amount=${amt}`, loanId };
}

function createLoanGiven(date, borrowerName, amount) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const party = String(borrowerName||'').trim();
  const amt = Number(amount||0);
  if (!party) return { success:false, message:'Borrower name required' };
  if (!isFinite(amt) || amt <= 0) return { success:false, message:'Invalid amount' };

  const loanId = makeLoanId_('LG', dk);

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_LOAN_GIVEN,
    description: `LOAN_GIVEN | ID:${loanId}`,
    handoverTo: party,
    in: 0,
    out: amt
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_LG, E_CREATE, loanId, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`LoanGiven created: ${loanId} amount=${amt}`, loanId };
}

// =====================================================
// PUBLIC API — Pending (OPEN LIST)
// =====================================================

function listOpenCustomerDues() {
  const rows = readPendingLog_();
  const refs = Array.from(new Set(rows.filter(r => r.type===T_CUST).map(r => r.ref).filter(Boolean)));

  const items = refs.map(ref => {
    const x = remainingByRef_(rows, T_CUST, ref);
    return { ref, party:x.party, created:x.created, settled:x.settled, remaining:x.remaining, createdDate:x.createdDate };
  }).filter(x => x.remaining > 0);

  items.sort((a,b) => String(a.createdDate).localeCompare(String(b.createdDate)));
  return { success:true, items };
}

function listOpenLoanTaken() {
  const rows = readPendingLog_();
  const refs = Array.from(new Set(rows.filter(r => r.type===T_LT).map(r => r.ref).filter(Boolean)));

  const items = refs.map(ref => {
    const x = remainingByRef_(rows, T_LT, ref);
    return { ref, party:x.party, created:x.created, settled:x.settled, remaining:x.remaining, createdDate:x.createdDate };
  }).filter(x => x.remaining > 0);

  items.sort((a,b) => String(a.createdDate).localeCompare(String(b.createdDate)));
  return { success:true, items };
}

function listOpenLoanGiven() {
  const rows = readPendingLog_();
  const refs = Array.from(new Set(rows.filter(r => r.type===T_LG).map(r => r.ref).filter(Boolean)));

  const items = refs.map(ref => {
    const x = remainingByRef_(rows, T_LG, ref);
    return { ref, party:x.party, created:x.created, settled:x.settled, remaining:x.remaining, createdDate:x.createdDate };
  }).filter(x => x.remaining > 0);

  items.sort((a,b) => String(a.createdDate).localeCompare(String(b.createdDate)));
  return { success:true, items };
}

// =====================================================
// PUBLIC API — Pending (SETTLE)
// =====================================================

function settleCustomerDue(date, invoiceNo, amount) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const inv = String(invoiceNo||'').trim();
  const amt = Number(amount||0);
  if (!inv) return { success:false, message:'InvoiceNo required' };
  if (!isFinite(amt) || amt <= 0) return { success:false, message:'Invalid amount' };

  const rows = readPendingLog_();
  const x = remainingByRef_(rows, T_CUST, inv);
  if (!(x.created > 0)) return { success:false, message:`No CREATE found for invoice ${inv}` };
  if (amt > x.remaining) return { success:false, message:`Amount exceeds remaining. Remaining=${x.remaining}` };

  const party = x.party || '';

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_DUE_SETTLE,
    description: `DUE_SETTLE | INV:${inv}`,
    handoverTo: party,
    in: amt,
    out: 0
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_CUST, E_SETTLE, inv, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`Due settled: ${inv} amount=${amt}` };
}

function settleLoanTaken(date, loanId, amount) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const ref = String(loanId||'').trim();
  const amt = Number(amount||0);
  if (!ref) return { success:false, message:'LoanId required' };
  if (!isFinite(amt) || amt <= 0) return { success:false, message:'Invalid amount' };

  const rows = readPendingLog_();
  const x = remainingByRef_(rows, T_LT, ref);
  if (!(x.created > 0)) return { success:false, message:`No CREATE found for LoanTaken ${ref}` };
  if (amt > x.remaining) return { success:false, message:`Amount exceeds remaining. Remaining=${x.remaining}` };

  const party = x.party || '';

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_LOAN_REPAID,
    description: `LOAN_REPAID | ID:${ref}`,
    handoverTo: party,
    in: 0,
    out: amt
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_LT, E_SETTLE, ref, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`Debt repaid: ${ref} amount=${amt}` };
}

function settleLoanGiven(date, loanId, amount) {
  const dk = dateKey_(date);
  if (!dk) return { success:false, message:'Invalid date' };

  const ref = String(loanId||'').trim();
  const amt = Number(amount||0);
  if (!ref) return { success:false, message:'LoanId required' };
  if (!isFinite(amt) || amt <= 0) return { success:false, message:'Invalid amount' };

  const rows = readPendingLog_();
  const x = remainingByRef_(rows, T_LG, ref);
  if (!(x.created > 0)) return { success:false, message:`No CREATE found for LoanGiven ${ref}` };
  if (amt > x.remaining) return { success:false, message:`Amount exceeds remaining. Remaining=${x.remaining}` };

  const party = x.party || '';

  const led = saveBulkEntries([{
    date: ledgerDate_(dk),
    header: H_LOAN_RCVD,
    description: `LOAN_RCVD | ID:${ref}`,
    handoverTo: party,
    in: amt,
    out: 0
  }]);
  if (!led || !led.success) return { success:false, message:`Ledger failed: ${led && led.message ? led.message : 'unknown'}` };

  appendPendingLog_([[
    dk, T_LG, E_SETTLE, ref, party, amt, '', nowIso_()
  ]]);

  return { success:true, message:`Credit settled: ${ref} amount=${amt}` };
}

/** ========= CORE: Process Daily Sales (UPDATED SOURCE) ========= */
function processDailySales() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert('Another process is running. Try again in 30 seconds.');
    return;
  }

  const fromDateKey = CONFIG.sales.minDateKey; // Automatically set from config

  try {
    const cols = getLedgerCols_();
    const ledger = cols.sheet;

    const internalIdsMap = getInternalIdsMap();
    const dailySalesHeader = internalIdsMap[CONFIG.internalIds.dailySales];
    if (!dailySalesHeader) throw new Error(`Settings missing Internal ID ${CONFIG.internalIds.dailySales} (Daily Sales).`);

    // processed dates set
    const processed = new Set();
    const ledgerLastRow = ledger.getLastRow();
    if (ledgerLastRow > cols.headerRow) {
      const minCol = Math.min(cols.date, cols.header);
      const maxCol = Math.max(cols.date, cols.header);
      const width = maxCol - minCol + 1;

      const data = ledger.getRange(cols.headerRow + 1, minCol, ledgerLastRow - cols.headerRow, width).getValues();
      const dateOffset = cols.date - minCol;
      const headerOffset = cols.header - minCol;

      data.forEach(r => {
        const hdr = (r[headerOffset] || '').toString().trim();
        if (hdr !== dailySalesHeader) return;

        const cell = r[dateOffset];
        const d = (cell instanceof Date && !isNaN(cell)) ? cell : parseYMD_(cell);
        if (!d) return;

        processed.add(toDateKey_(d, tz));
      });
    }

    // external sales sheet
    const salesSS = SpreadsheetApp.openById(CONFIG.sales.spreadsheetId);
    const salesSheet = salesSS.getSheetByName(CONFIG.sales.sheetName);
    if (!salesSheet) throw new Error(`Sales sheet not found: ${CONFIG.sales.sheetName}`);

    const lastSalesRow = salesSheet.getLastRow();
    if (lastSalesRow <= CONFIG.sales.headerRow) {
      ui.alert('No sales rows to process.');
      logImport_('NO_DATA', fromDateKey, 0, 0, '', '');
      return;
    }

    const numRows = lastSalesRow - CONFIG.sales.headerRow;
    const dateCol = colLetterToIndex_(CONFIG.sales.dateColLetter);
    const amtCol = colLetterToIndex_(CONFIG.sales.amountColLetter);

    const dateRange = salesSheet.getRange(CONFIG.sales.headerRow + 1, dateCol, numRows, 1);
    const amtRange  = salesSheet.getRange(CONFIG.sales.headerRow + 1, amtCol,  numRows, 1);

    const dateVals = dateRange.getValues();
    const dateDisp = dateRange.getDisplayValues();
    const amtVals  = amtRange.getValues();
    const amtDisp  = amtRange.getDisplayValues();

      const dailyTotals = {};
      for (let i = 0; i < numRows; i++) {
       let d = null;
       const dv = dateVals[i][0];
       if (dv instanceof Date && !isNaN(dv)) d = dv;
       if (!d) d = parseYMD_(dateDisp[i][0]);
       if (!d) d = parseDMY_(dateDisp[i][0]);
       if (!d) continue;
 
       const rawKey = (dateDisp[i][0] || '').toString().trim();
       const key = (/^\d{4}-\d{2}-\d{2}$/.test(rawKey)) ? rawKey : toDateKey_(d, tz);
       if (key < fromDateKey) continue;
       if (processed.has(key)) continue;

      let amount = NaN;
      const av = amtVals[i][0];
      if (typeof av === 'number') amount = av;
      if (isNaN(amount)) amount = parseAmount_(amtDisp[i][0]);

      if (!isNaN(amount) && amount > 0) {
        dailyTotals[key] = (dailyTotals[key] || 0) + amount;
      }
    }

    const keys = Object.keys(dailyTotals).sort();
    if (!keys.length) {
      ui.alert(`No new sales found from ${fromDateKey}.`);
      logImport_('NO_NEW', fromDateKey, 0, 0, '', '');
      return;
    }

    // current balance from last row only
    let currentBalance = 0;
    if (ledgerLastRow > cols.headerRow) {
      currentBalance = Number(ledger.getRange(ledgerLastRow, cols.balance).getValue()) || 0;
    }

    const tableMinCol = Math.min(cols.date, cols.header, cols.description, cols.in, cols.out, cols.handoverTo, cols.balance, cols.classification);
    const tableMaxCol = Math.max(cols.date, cols.header, cols.description, cols.in, cols.out, cols.handoverTo, cols.balance, cols.classification);
    const tableWidth = tableMaxCol - tableMinCol + 1;

    const rowsToWrite = [];
    let totalImported = 0;

    keys.forEach(key => {
      const dSerial = dateKeyToSerial_(key);
      const amt = dailyTotals[key];
      totalImported += amt;
      currentBalance += amt;

      const row = new Array(tableWidth).fill('');
      row[cols.date - tableMinCol] = (dSerial != null) ? dSerial : parseYMD_(key);
      row[cols.header - tableMinCol] = dailySalesHeader;
      row[cols.description - tableMinCol] = `Total daily sales for ${key}`;
      row[cols.in - tableMinCol] = amt;
      row[cols.out - tableMinCol] = 0;
      row[cols.handoverTo - tableMinCol] = '';
      row[cols.balance - tableMinCol] = currentBalance;
      row[cols.classification - tableMinCol] = 'Income';

      rowsToWrite.push(row);
    });

    const startRow = ledger.getLastRow() + 1;
    ledger.getRange(startRow, tableMinCol, rowsToWrite.length, tableWidth).setValues(rowsToWrite);
    ledger.getRange(startRow, cols.date, rowsToWrite.length, 1).setNumberFormat(CONFIG.ledger.dateNumberFormat);

    updateAccountsSummary();
    logImport_('OK', fromDateKey, rowsToWrite.length, totalImported, keys[0], keys[keys.length - 1]);
    ui.alert(`Imported ${rowsToWrite.length} day(s). Total: ${totalImported}`);

  } catch (e) {
    Logger.log(e.stack || e);
    logImport_('ERROR: ' + e.message, fromDateKey, 0, 0, '', '');
    ui.alert(`processDailySales failed:\n${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/** ========= Accounts Summary ========= */
function updateAccountsSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let summary = ss.getSheetByName(CONFIG.accountsSummarySheetName);
    if (!summary) summary = ss.insertSheet(CONFIG.accountsSummarySheetName);
    summary.clear();

    summary.getRange('A1:D1')
      .setValues([['Account', 'Total In (Credit)', 'Total Out (Debit)', 'Balance']])
      .setFontWeight('bold')
      .setBackground('#f0f0f0');

    const cols = getLedgerCols_();
    const ledger = cols.sheet;

    const lastRow = ledger.getLastRow();
    if (lastRow <= cols.headerRow) return;

    const minCol = Math.min(cols.header, cols.in, cols.out);
    const maxCol = Math.max(cols.header, cols.in, cols.out);
    const width = maxCol - minCol + 1;

    const data = ledger.getRange(cols.headerRow + 1, minCol, lastRow - cols.headerRow, width).getValues();
    const headerOffset = cols.header - minCol;
    const inOffset = cols.in - minCol;
    const outOffset = cols.out - minCol;

    const totals = {};
    data.forEach(r => {
      const h = (r[headerOffset] || '').toString().trim();
      if (!h) return;
      const IN = Number(r[inOffset]) || 0;
      const OUT = Number(r[outOffset]) || 0;
      if (!totals[h]) totals[h] = { in: 0, out: 0 };
      totals[h].in += IN;
      totals[h].out += OUT;
    });

    const internalIdsMap = getInternalIdsMap();
    const dailySalesHeader = internalIdsMap[CONFIG.internalIds.dailySales] || null;
    const ghorooaPaymentHeader = internalIdsMap[CONFIG.internalIds.ghorooaPayment] || null;

    const rows = Object.keys(totals).map(h => [h, totals[h].in, totals[h].out, totals[h].in - totals[h].out]);

    const totalDailySalesIn = (dailySalesHeader && totals[dailySalesHeader]) ? totals[dailySalesHeader].in : 0;
    const totalGhorooaPaymentsOut = (ghorooaPaymentHeader && totals[ghorooaPaymentHeader]) ? totals[ghorooaPaymentHeader].out : 0;

    const liabilityIn = totalDailySalesIn * CONFIG.ghorooaLiabilityRate;
    const liabilityBal = liabilityIn - totalGhorooaPaymentsOut;
    rows.push(['Ghorooa Liability', liabilityIn, totalGhorooaPaymentsOut, liabilityBal]);

    if (rows.length) {
      summary.getRange(2, 1, rows.length, 4).setValues(rows);
      summary.getRange(2, 2, rows.length, 3).setNumberFormat('#,##0.00');
      summary.autoResizeColumns(1, 4);
    }
  } catch (e) {
    Logger.log(e.stack || e);
  }
}

/** ========= Diagnostics ========= */
function diagnoseSetup() {
  SpreadsheetApp.getUi().alert('Diagnostics OK (base).');
}

/** ========= Initialize (non-destructive) ========= */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(CONFIG.ledger.sheetName)) ss.insertSheet(CONFIG.ledger.sheetName);

  let settings = ss.getSheetByName(CONFIG.settings.sheetName);
  if (!settings) settings = ss.insertSheet(CONFIG.settings.sheetName);

  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 3).setValues([['Headers', 'Classification', 'Internal ID']]);
  }

  ensureImportLog_();
  ensurePayrollSetup_();
}

/** ========= PAYROLL SYSTEM ========= */

function ensurePayrollSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(CONFIG.settings.sheetName);
  const lastRow = settings.getLastRow();
  
  // Helper to check if header exists
  const headerExists = (name) => {
    if (lastRow <= 1) return false;
    const headers = settings.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    return headers.map(h => String(h).trim()).includes(name);
  };

  // Add Payroll Headers to Settings if missing
  const payrollHeaders = [
    { name: CONFIG.payroll.headerSalary, class: 'Expense' },
    { name: CONFIG.payroll.headerAdvance, class: 'Liability' }
  ];

  payrollHeaders.forEach(h => {
    if (!headerExists(h.name)) {
      settings.appendRow([h.name, h.class, '']);
    }
  });

  // Create Staff Master Sheet
  let staff = ss.getSheetByName(CONFIG.payroll.masterSheet);
  if (!staff) {
    staff = ss.insertSheet(CONFIG.payroll.masterSheet);
    staff.getRange(1, 1, 1, 5).setValues([['ID', 'Name', 'Type', 'Rate', 'OT_Rate']])
          .setFontWeight('bold');
  }

  // Create Attendance Sheet
  let att = ss.getSheetByName(CONFIG.payroll.attendanceSheet);
  if (!att) {
    att = ss.insertSheet(CONFIG.payroll.attendanceSheet);
    att.getRange(1, 1, 1, 6).setValues([['Date', 'ID', 'Name', 'Status', 'OT_Hours', 'Note']])
       .setFontWeight('bold');
  }
}

function getStaffList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.payroll.masterSheet);
  if (!sh || sh.getLastRow() <= 1) return [];
  
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return data.map(r => ({
    id: String(r[0]).trim(),
    name: String(r[1]).trim(),
    type: String(r[2]).trim(),
    rate: Number(r[3]) || 0,
    otRate: Number(r[4]) || 0
  }));
}

function saveAttendance(record) {
  // record: { date, id, status, otHours, fine, note }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.payroll.attendanceSheet);
  
  // Get Staff Name
  const staff = getStaffList();
  const sObj = staff.find(s => s.id === record.id);
  const name = sObj ? sObj.name : 'Unknown';

  sh.appendRow([
    new Date(record.date),
    record.id,
    name,
    record.status,
    Number(record.otHours) || 0,
    Number(record.fine) || 0,
    record.note || ''
  ]);

  return { success: true };
}

function calculatePayrollData(id, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  
  // 1. Get Staff Details
  const allStaff = getStaffList();
  const staff = allStaff.find(s => s.id === id);
  if (!staff) throw new Error('Staff ID not found');

  // 2. Get Attendance in Range
  const attSh = ss.getSheetByName(CONFIG.payroll.attendanceSheet);
  const lastRow = attSh.getLastRow();
  if (lastRow <= 1) throw new Error('No attendance records found');
  
  const attData = attSh.getRange(2, 1, lastRow - 1, 7).getValues();
  
  let daysPresent = 0;
  let otHoursTotal = 0;
  let fineTotal = 0;


  attData.forEach(r => {
    const rDate = r[0];
    const rId = String(r[1]).trim();
    const rStatus = String(r[3]).trim().toLowerCase();
    const rOt = Number(r[4]) || 0;
    const rFine = Number(r[5]) || 0;

    if (rId !== id) return;
    
    let d = (rDate instanceof Date) ? rDate : parseYMD_(rDate);
    if (!d) return;

    const key = toDateKey_(d, tz);
    if (key >= startDate && key <= endDate) {
      if (rStatus === 'present' || rStatus === 'p' || rStatus === 'ot') {
        daysPresent++;
      }
      otHoursTotal += rOt;
      
    }
  });

  // 3. Calculate Salary
   // RULE: Rate is per-day for BOTH Daily and Monthly staff.
   // 1 Present = 1 Rate
   const baseSalary = daysPresent * staff.rate;

  const otAmount = otHoursTotal * (staff.otRate || staff.rate * 1.5); // Use specific OT rate or default 1.5x
  const grossSalary = baseSalary + otAmount;

  // 4. Calculate Advance from Ledger
  const cols = getLedgerCols_();
  const ledger = cols.sheet;
  const ledLastRow = ledger.getLastRow();
  let totalAdvance = 0;

  if (ledLastRow > cols.headerRow) {
    const ledData = ledger.getRange(cols.headerRow + 1, 1, ledLastRow - cols.headerRow, ledger.getLastColumn()).getValues();
    const hIdx = cols.header - 1;
    const descIdx = cols.description - 1;
    const outIdx = cols.out - 1;
    const handIdx = cols.handoverTo - 1;

    ledData.forEach(r => {
      const header = normalizeHeader_(r[hIdx]);
      const desc = String(r[descIdx] || '').toLowerCase();
      const out = Number(r[outIdx]) || 0;

       if (header === normalizeHeader_(CONFIG.payroll.headerAdvance)) {
       const hand = String(r[handIdx] || '').trim();
       const dsc = String(r[descIdx] || '').trim();
       const tagId = extractStaffIdTag_(hand) || extractStaffIdTag_(dsc);

       if (tagId) {
       if (String(tagId) === String(staff.id)) totalAdvance += out;
      } else {
       // fallback for old rows
       const nameLow = staff.name.toLowerCase();
       if (hand.toLowerCase().includes(nameLow) || dsc.toLowerCase().includes(nameLow)) totalAdvance += out;
      }
     }
    });
  }

  const netPay = grossSalary - totalAdvance - fineTotal;
  
  return {
    staffName: staff.name,
    baseSalary,
    otHours: otHoursTotal,
    otAmount,
    grossSalary,
    totalAdvance,
    fineTotal,
    netPay,
    daysWorked: daysPresent
  };
}

function processStaffPayment(payload) {
  // payload: { id, startDate, endDate, netPay, description }
  const data = calculatePayrollData(payload.id, payload.startDate, payload.endDate);
  
  // Validation Alert Logic
  if (data.totalAdvance > 2000) {
    // We can't throw an error if they want to proceed, but we return a warning flag.
    // For now, we log it. The UI handles the alert before calling this.
  }
  
  if (data.netPay < 0) {
    return { success: false, message: 'Error: Net Payable is negative. Cannot process.' };
  }

  const entry = {
    date: new Date(), // Pay day
    header: CONFIG.payroll.headerSalary,
    description: `Salary Pay - ${data.staffName} (${payload.startDate} to ${payload.endDate})`,
    in: 0,
    out: data.netPay,
    handoverTo: 'ID:' + payload.id + ' ' + data.staffName

  };

  const result = saveBulkEntries([entry]);
  return result;
}
// ===============================
// Append Inventory Slip Snapshot to Return_log (once per date)
// Source rate: Product list (Purchase name -> Price)
// ===============================
function appendReturnLog_(reportDate, rows) {
  const ss = SpreadsheetApp.openById(CONFIG.sales.spreadsheetId);

  const dateKey = normalizeDateKeyServer_(reportDate);
  if (!dateKey) return { status: 'error', message: 'Invalid reportDate' };

  // Ensure Return_log exists with headers
  const sh = ensureSheetWithHeader_(ss, 'Return_log', [
    'report_date', 'item_name', 'rate', 'returnqty', 'total', 'created_at'
  ]);

  const last = sh.getLastRow();
  if (last >= 2) {
    const tz = Session.getScriptTimeZone();
    const existingDates = sh.getRange(2, 1, last - 1, 1).getValues().flat();

    const existingKeys = new Set(existingDates.map(v => {
      if (Object.prototype.toString.call(v) === '[object Date]') {
        return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      const s = String(v || '').trim();
      // if someone stored as '2026-01-28 00:00:00' etc
      const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
      return m ? m[1] : s;
    }));

    if (existingKeys.has(dateKey)) {
      return { status: 'exists' };
    }
  }
// Build rate map from Product list:
  // Purchase name = column C, Price = column G
  const prod = ss.getSheetByName('Product list');
  if (!prod) return { status: 'error', message: 'Product list sheet not found' };

  const prodLast = prod.getLastRow();
  if (prodLast < 2) return { status: 'error', message: 'Product list empty' };

  // আমরা Column C (3) থেকে শুরু করে Column G (7) পর্যন্ত ডাটা নিচ্ছি। 
  // কলাম সংখ্যা ৫টি (C, D, E, F, G)
  const prodVals = prod.getRange(2, 3, prodLast - 1, 5).getValues();
  
  // চেক করার জন্য নিচের লাইনটি দেখুন (Execution Log এ দেখা যাবে)
  // console.log("First Row Data:", prodVals[0]); 

  const rateMap = new Map();
  prodVals.forEach((r, index) => {
    const purchaseName = (r[0] || '').toString().trim(); // Column C
    const price = Number(r[4] || 0); // r[4] মানে ৫ নম্বর পজিশন বা Column G
    
    if (purchaseName) {
      rateMap.set(purchaseName, price);
      // ছোট একটি ডিবাগ লগ (প্রথম ৫টি আইটেম চেক করার জন্য)
      if(index < 5) console.log("Item: " + purchaseName + " | Price taken from G: " + price);
    }
  });

  // Validate rows & prepare append
  // Dedup items for the day: sum remain_qty by item_name (normalized)
  const agg = new Map(); // key -> { name, qty }
  (rows || []).forEach(obj => {
    const name = (obj.item_name || '').toString().trim();
    if (!name) return;
    const key = invNormName_ ? invNormName_(name) : name.toLowerCase(); // if you have invNormName_ in project, it'll use it
    const qty = Number(obj.remain_qty || 0);
    if (!agg.has(key)) agg.set(key, { name, qty: 0 });
    agg.get(key).qty += qty;
  });

  // overwrite rows to aggregated list
  rows = Array.from(agg.values()).map(x => ({ item_name: x.name, remain_qty: x.qty }));

  const missing = [];
  const out = [];
  const now = new Date();

  (rows || []).forEach(obj => {
    const name = (obj.item_name || '').toString().trim();
    const qty = Number(obj.remain_qty || 0);

    if (!name) return; // skip blank
    if (!rateMap.has(name)) {
      missing.push(name);
      return;
    }

    const rate = Number(rateMap.get(name) || 0);
    const total = rate * qty;

    out.push([dateKey, name, rate, qty, total, now]);
  });

  if (missing.length) {
    // block printing to preserve invariants (no partial dump)
    return { status: 'missing_rate', missing_items: unique_(missing).slice(0, 50) };
  }

  if (!out.length) return { status: 'error', message: 'No rows to append' };

  sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  return { status: 'ok', appended: out.length, report_date: dateKey };
}

// ---- helpers ----
function ensureSheetWithHeader_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const isHeaderMissing = headers.some((h, i) => String(firstRow[i] || '').trim() !== h);
  if (isHeaderMissing) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function normalizeDateKeyServer_(input) {
  // Accept YYYY-MM-DD from HTML date input OR Date object
  if (!input) return '';
  if (Object.prototype.toString.call(input) === '[object Date]') {
    return Utilities.formatDate(input, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(input).trim();
  // If already yyyy-mm-dd keep it
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Try parse
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function unique_(arr) {
  return Array.from(new Set(arr.map(s => String(s))));
}

function getPayrollDashboardViewData(dateKey, mode) {
  // mode: 'daily' or 'monthly'
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  
  const dateObj = parseYMD_(dateKey);
  if (!dateObj) throw new Error("Invalid Date");

  // 1. Define Range
  let startKey, endKey;
  
  if (mode === 'daily') {
    startKey = dateKey;
    endKey = dateKey;
  } else if (mode === 'monthly') {
    const y = dateObj.getFullYear();
    const m = dateObj.getMonth() + 1;
    const start = new Date(y, m - 1, 1, 12, 0, 0);
    const end = new Date(y, m, 0, 12, 0, 0); // Last day of month
    startKey = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
    endKey = Utilities.formatDate(end, tz, 'yyyy-MM-dd');
  }

  // 2. Get Staff & Attendance
  const staffList = getStaffList();
  const attSheet = ss.getSheetByName(CONFIG.payroll.attendanceSheet);
  const attLastRow = attSheet ? attSheet.getLastRow() : 0;
  
  // Map Attendance: ID -> { days, ot, fine }
  const attMap = {}; 
  if (attLastRow > 1) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 7).getValues();
    attData.forEach(r => {
      const d = (r[0] instanceof Date) ? r[0] : parseYMD_(r[0]);
      if (!d) return;
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      
      if (k >= startKey && k <= endKey) {
        const id = String(r[1]);
        const status = String(r[3]).toLowerCase();
        const ot = Number(r[4]) || 0;
        const fine = Number(r[5]) || 0;
        if (!attMap[id]) attMap[id] = { days: 0, ot: 0, fine: 0 };
        
        if (status === 'present' || status === 'p' || status === 'ot') {
          attMap[id].days++;
        }
        attMap[id].ot += ot;
        attMap[id].fine += fine;
      }
    });
  }

  // 3. Get Ledger (Paid & Advance)
  const cols = getLedgerCols_();
  const ledger = cols.sheet;
  const ledLastRow = ledger.getLastRow();
  const ledMap = {}; // ID -> { paid, advance }

  if (ledLastRow > cols.headerRow) {
    const data = ledger.getRange(cols.headerRow + 1, 1, ledLastRow - cols.headerRow, ledger.getLastColumn()).getValues();
    const hIdx = cols.header - 1;
    const handIdx = cols.handoverTo - 1;
    const outIdx = cols.out - 1;
    const dateIdx = cols.date - 1;
    const descIdx = cols.description - 1;               

    data.forEach(r => {
      const d = (r[dateIdx] instanceof Date) ? r[dateIdx] : parseYMD_(r[dateIdx]);
      if (!d) return;
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      
      if (k >= startKey && k <= endKey) {
        const header = normalizeHeader_(r[hIdx]);
        const handover = String(r[handIdx] || '').trim();
        const amt = Number(r[outIdx]) || 0;
        
        if (amt === 0) return;

        // Find Staff ID by Name (Reverse lookup)
        const desc = String(r[descIdx] || '').trim();

        // 1) ID-first match (deterministic)
        let staffId = extractStaffIdTag_(handover) || extractStaffIdTag_(desc);

        // 2) Fallback: old name-substring match (backward compatibility)
        let staff = null;
        if (staffId) {
         staff = staffList.find(s => String(s.id) === String(staffId));
        } else {
          const hLow = handover.toLowerCase();
          const dLow = desc.toLowerCase();
          staff = staffList.find(s => hLow.includes(s.name.toLowerCase()) || dLow.includes(s.name.toLowerCase()));
          }

         if (staff) {
         if (!ledMap[staff.id]) ledMap[staff.id] = { paid: 0, advance: 0 };

         if (header === normalizeHeader_(CONFIG.payroll.headerSalary)) {
           ledMap[staff.id].paid += amt;
          } else if (header === normalizeHeader_(CONFIG.payroll.headerAdvance)) {
           ledMap[staff.id].advance += amt;
          }
        }      

      }
    });
  }

  // 4. Construct Result
  const result = staffList.map(s => {
    const att = attMap[s.id] || { days: 0, ot: 0 };
    const fin = ledMap[s.id] || { paid: 0, advance: 0 };
    
   const baseSalary = att.days * s.rate; // per-day rate for all staff types

    const otAmt = att.ot * (s.otRate || s.rate * 1.5);
    const gross = baseSalary + otAmt;
    const due = gross - fin.paid - fin.advance - (att.fine || 0);

    return {
      id: s.id,
      name: s.name,
      type: s.type,
      rate: s.rate,
      daysWorked: att.days,
      otHours: att.ot,
      fineTotal: att.fine || 0,
      grossSalary: gross,
      paid: fin.paid,
      advance: fin.advance,
      balanceDue: due
    };
  });

  return { data: result, start: startKey, end: endKey };
}

function processManualStaffPayment(payload) {
  // payload: { staffId, date, amount, note }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "System busy. Try again." };

  try {
    const allStaff = getStaffList();
    const staff = allStaff.find(s => String(s.id) === String(payload.staffId));
    if (!staff) throw new Error("Staff ID not found");

    const amount = Number(payload.amount);
    if (!Number.isFinite(amount) || amount <= 0) throw new Error("Invalid Amount");

    // Build dateKey (yyyy-MM-dd) deterministically
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    let dateKey = null;
    if (payload.date instanceof Date && !isNaN(payload.date)) {
      dateKey = Utilities.formatDate(payload.date, tz, "yyyy-MM-dd");
    } else {
      const raw = (payload.date || "").toString().trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) dateKey = raw;
    }
    if (!dateKey) throw new Error("Invalid payment date (need yyyy-mm-dd)");

    // Warning Check: total advances already taken (name-based, existing behavior)
    const cols = getLedgerCols_();
    const ledger = cols.sheet;
    const lastRow = ledger.getLastRow();
    let totalAdvance = 0;

    if (lastRow > cols.headerRow) {
      const data = ledger
        .getRange(cols.headerRow + 1, 1, lastRow - cols.headerRow, ledger.getLastColumn())
        .getValues();

      const hIdx = cols.header - 1;
      const handoverIdx = cols.handoverTo - 1;
      const outIdx = cols.out - 1;

      data.forEach(r => {
        const header = normalizeHeader_(r[hIdx]);
        const handoverName = (r[handoverIdx] || "").toString().toLowerCase();
        const out = Number(r[outIdx]) || 0;

        if (header === normalizeHeader_(CONFIG.payroll.headerAdvance)) {
          if (handoverName.includes(staff.name.toLowerCase())) totalAdvance += out;
        }
      });
    }

    let warningMsg = "";
    if (totalAdvance > 2000) warningMsg = `ALERT: Staff has taken ${totalAdvance} in advances.`;

    // Get current due using the same engine as Payroll tab (monthly mode)
    const view = getPayrollDashboardViewData(dateKey, "monthly");
    const row = (view && view.data) ? view.data.find(r => String(r.id) === String(payload.staffId)) : null;
    if (!row) throw new Error("Payroll row not found for this staff/date");

    const dueNow = Number(row.balanceDue) || 0;

    // Split payment into Salary vs Advance (Overpay)
    const salaryOut = Math.min(amount, Math.max(0, dueNow));
    const advanceOut = amount - salaryOut;

    // Write date without day-shift (noon local)
    const payDate = parseYMD_(dateKey);
    if (!payDate) throw new Error("Invalid payment dateKey");
    payDate.setHours(12, 0, 0, 0);

    const note = (payload.note || "Salary Pay (Dashboard)").toString();

    const entries = [];
    if (salaryOut > 0) {
      entries.push({
        date: payDate,
        header: CONFIG.payroll.headerSalary, // Staff Salary
        description: note,
        in: 0,
        out: salaryOut,
        handoverTo: 'ID:' + staff.id + ' ' + staff.name

      });
    }
    if (advanceOut > 0) {
      entries.push({
        date: payDate,
        header: CONFIG.payroll.headerAdvance, // Staff Advance
        description: `Advance (Overpay) | ${note}`,
        in: 0,
        out: advanceOut,
        handoverTo: 'ID:' + staff.id + ' ' + staff.name

      });
    }

    const result = saveBulkEntries(entries);
    if (result && result.success) result.warning = warningMsg;
    return result;

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getStaffMonthlyReport(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const cols = getLedgerCols_();
  const ledger = cols.sheet;

  const y = Number(year);
  const m = Number(month);
  if (!y || !m) throw new Error("Invalid Year/Month");

  const startDate = new Date(y, m - 1, 1, 12, 0, 0);
  const endDate = new Date(y, m, 0, 12, 0, 0);
  const startKey = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  const endKey = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');

  const lastRow = ledger.getLastRow();
  if (lastRow <= cols.headerRow) return { data: [] };

  // Fetch relevant columns: Date, Header, HandoverTo, Out
  const data = ledger.getRange(cols.headerRow + 1, 1, lastRow - cols.headerRow, ledger.getLastColumn()).getValues();
  
  const dateIdx = cols.date - 1;
  const headerIdx = cols.header - 1;
  const handoverIdx = cols.handoverTo - 1;
  const outIdx = cols.out - 1;

  const stats = {}; // Key: Staff Name, Value: { salary: 0, advance: 0 }

  data.forEach(r => {
    const d = r[dateIdx];
    if (!d) return;
    
    let dObj = (d instanceof Date) ? d : parseYMD_(d);
    if (!dObj) return;

    const key = Utilities.formatDate(dObj, tz, 'yyyy-MM-dd');
    if (key < startKey || key > endKey) return;

    const header = normalizeHeader_(r[headerIdx]);
    const isSalary  = header === normalizeHeader_(CONFIG.payroll.headerSalary);
    const isAdvance = header === normalizeHeader_(CONFIG.payroll.headerAdvance);
    if (!isSalary && !isAdvance) return;

    const staffName = (r[handoverIdx] || '').toString().trim();
    const amount = Number(r[outIdx]) || 0;
    if (!staffName || amount === 0) return;

    if (!stats[staffName]) stats[staffName] = { salary: 0, advance: 0 };
    if (isSalary) stats[staffName].salary += amount;
    if (isAdvance) stats[staffName].advance += amount;
  });

  // Convert to array
  const result = Object.keys(stats).map(name => {
    const s = stats[name];
    return {
      name: name,
      paid: s.salary,
      advance: s.advance,
      total: s.salary + s.advance
    };
  }).sort((a, b) => b.total - a.total); // Sort by highest payout

  return { data: result, period: `${y}-${String(m).padStart(2, '0')}` };
}

function getHeaderDetailReport(headerName, startDateKey, endDateKey) {
  const cols = getLedgerCols_();
  const ledger = cols.sheet;
  const lastRow = ledger.getLastRow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const targetHeaderRaw = (headerName || '').toString();
  const targetHeaderTrim = targetHeaderRaw.trim();
  const targetHeaderNorm = normalizeHeader_(targetHeaderTrim);

  const startKey = (startDateKey instanceof Date)
    ? Utilities.formatDate(startDateKey, tz, 'yyyy-MM-dd')
    : (startDateKey || '').toString().trim();

  const endKey = (endDateKey instanceof Date)
    ? Utilities.formatDate(endDateKey, tz, 'yyyy-MM-dd')
    : (endDateKey || '').toString().trim();

  const debug = {
    tz: tz,
    input: { headerNameRaw: targetHeaderRaw, headerNameTrim: targetHeaderTrim, headerNameNorm: targetHeaderNorm, startKey: startKey, endKey: endKey },
    cols: { headerRow: cols.headerRow, date: cols.date, header: cols.header, description: cols.description, in: cols.in, out: cols.out, handoverTo: cols.handoverTo, balance: cols.balance },
    counts: { totalRows: 0, dateMissing: 0, dateIsDate: 0, dateParseFail: 0, outOfRange: 0, headerMismatch: 0, matched: 0 },
    samples: []
  };

  if (lastRow <= cols.headerRow) {
    debug.note = 'No data rows under headerRow';
    Logger.log(JSON.stringify(debug));
    console.log(debug);
    return { header: headerName, transactions: [], start: startKey, end: endKey, debug: debug };
  }

  const data = ledger.getRange(
    cols.headerRow + 1,
    1,
    lastRow - cols.headerRow,
    ledger.getLastColumn()
  ).getValues();

  const transactions = [];

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    debug.counts.totalRows++;

    const dVal = r[cols.date - 1];
    if (!dVal) {
      debug.counts.dateMissing++;
      if (debug.samples.length < 8) debug.samples.push({ i: i + 1, reason: 'dateMissing', dVal: dVal });
      continue;
    }

    let dObj = null;
    if (dVal instanceof Date) {
      debug.counts.dateIsDate++;
      dObj = dVal;
    } else {
      dObj = parseYMD_(dVal);
      if (!dObj) {
        debug.counts.dateParseFail++;
        if (debug.samples.length < 8) debug.samples.push({ i: i + 1, reason: 'dateParseFail', dVal: String(dVal) });
        continue;
      }
    }

    const key = Utilities.formatDate(dObj, tz, 'yyyy-MM-dd');

    if (key < startKey || key > endKey) {
      debug.counts.outOfRange++;
      if (debug.samples.length < 8) debug.samples.push({ i: i + 1, reason: 'outOfRange', key: key, startKey: startKey, endKey: endKey });
      continue;
    }

    const rowHeaderRaw = (r[cols.header - 1] || '').toString();
    const rowHeaderNorm = normalizeHeader_(rowHeaderRaw);

    if (rowHeaderNorm !== targetHeaderNorm) {
      debug.counts.headerMismatch++;
      if (debug.samples.length < 8) debug.samples.push({ i: i + 1, reason: 'headerMismatch', key: key, rowHeaderRaw: rowHeaderRaw, rowHeaderNorm: rowHeaderNorm, targetHeaderNorm: targetHeaderNorm });
      continue;
    }

    debug.counts.matched++;

    transactions.push({
      dateStr: key,
      description: (r[cols.description - 1] || '').toString().trim(),
      handover: (r[cols.handoverTo - 1] || '').toString().trim(),
      in: Number(r[cols.in - 1]) || 0,
      out: Number(r[cols.out - 1]) || 0,
      balance: Number(r[cols.balance - 1]) || 0
    });

  }

  Logger.log(JSON.stringify(debug));
  console.log(debug);

  return { header: headerName, transactions: transactions, start: startKey, end: endKey, debug: debug };
}

function saveBulkAttendance(records) {
  // records: [{ staffId, date, status, otHours }]
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "System busy." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attSheet = ss.getSheetByName(CONFIG.payroll.attendanceSheet);
    if (!attSheet) return { success: false, message: `Attendance sheet not found: ${CONFIG.payroll.attendanceSheet}` };

    if (!Array.isArray(records) || records.length === 0) {
      return { success: false, message: 'No attendance records received.' };
    }

    // --- Day-level lock (NO duplicates) ---
    // Rule: if any attendance already exists for the date, block the entire save.
    // This must be strict to prevent double-entry corruption.
    const tz = ss.getSpreadsheetTimeZone();
    const firstDateObj = (records[0].date instanceof Date)
      ? records[0].date
      : (parseYMD_(records[0].date) || parseDMY_(records[0].date));
    if (!firstDateObj) return { success: false, message: `Invalid date: ${records[0].date}` };
    const dateKey = toDateKey_(firstDateObj, tz);

    // Ensure all rows are for the same day (defensive)
    for (const r of records) {
      const dObj = (r.date instanceof Date) ? r.date : (parseYMD_(r.date) || parseDMY_(r.date));
      if (!dObj) return { success: false, message: `Invalid date in payload: ${r.date}` };
      if (toDateKey_(dObj, tz) !== dateKey) {
        return { success: false, message: `Mixed dates in one save call. Expected ${dateKey}.` };
      }
    }

    // Check existing Attendance sheet for the same dateKey
    const lastRow = attSheet.getLastRow();
    if (lastRow >= 2) {
      const existingDates = attSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingDates.length; i++) {
        const v = existingDates[i][0];
        if (!v) continue;
        const dObj = (v instanceof Date) ? v : (parseYMD_(v) || parseDMY_(v));
        if (!dObj) continue;
        if (toDateKey_(dObj, tz) === dateKey) {
          // IMPORTANT: Keep this phrase for search/debug.
          return { success: false, message: `Already saved attendance for ${dateKey}.` };
        }
      }
    }
    const staffList = getStaffList();
    
    // Map ID -> Name for fast lookup
    const staffMap = {};
    staffList.forEach(s => staffMap[s.id] = s.name);

    const rows = [];

    records.forEach(r => {
      const name = staffMap[r.staffId] || "Unknown";
      const dObj = (r.date instanceof Date) ? r.date : (parseYMD_(r.date) || parseDMY_(r.date));
      rows.push([
        dObj,
        r.staffId,
        name,
        r.status,
        Number(r.otHours) || 0,
        Number(r.fine) || 0,
        "" // Note
      ]);
    });

    if (rows.length > 0) {
      attSheet.getRange(attSheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    }

    return { success: true, message: `Saved ${rows.length} records for ${dateKey}.` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** ========= ATTENDANCE MONTHLY REPORT (READ-ONLY) =========
 * Source of truth: Attendance sheet
 * Columns: Date | ID | Name | Status | OT_Hours | Fine | Note
 * Status expected: Present / Absent (case-insensitive)
 */

// Returns months that exist in Attendance sheet, as [{year, month, key, label}]
function getAttendanceAvailableMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheetName = (CONFIG && CONFIG.payroll && CONFIG.payroll.attendanceSheet) ? CONFIG.payroll.attendanceSheet : 'Attendance';
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const dateVals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const seen = new Set();

  dateVals.forEach(r => {
    const v = r[0];
    let d = null;
    if (v instanceof Date && !isNaN(v)) d = v;
    else d = parseYMD_(v) || parseDMY_(v);
    if (!d) return;
    const key = Utilities.formatDate(d, tz, 'yyyy-MM');
    seen.add(key);
  });

  const months = Array.from(seen).sort();
  return months.map(k => {
    const parts = k.split('-');
    return { year: Number(parts[0]), month: Number(parts[1]), key: k, label: k };
  });
}

// Builds a monthly attendance aggregation for the given year/month.
function getAttendanceMonthlyReport(year, month) {
  const y = Number(year);
  const m = Number(month);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) {
    throw new Error('Invalid year/month');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheetName = (CONFIG && CONFIG.payroll && CONFIG.payroll.attendanceSheet) ? CONFIG.payroll.attendanceSheet : 'Attendance';
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Attendance sheet not found: ' + sheetName);

  const periodKey = y + '-' + String(m).padStart(2, '0');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      period: periodKey,
      totals: { staffCount: 0, recordCount: 0, presentDays: 0, absentDays: 0, otHours: 0, fine: 0 },
      byStaff: [],
      warnings: { unknownStatus: 0 }
    };
  }

  // Read columns A:G (7 cols)
  const data = sh.getRange(2, 1, lastRow - 1, 7).getValues();

  const byId = {}; // id -> agg
  let recordCount = 0;
  let presentDays = 0;
  let absentDays = 0;
  let otHours = 0;
  let fine = 0;
  let unknownStatus = 0;

  data.forEach(r => {
    const dVal = r[0];
    const id = (r[1] || '').toString().trim();
    if (!id) return;

    let dObj = null;
    if (dVal instanceof Date && !isNaN(dVal)) dObj = dVal;
    else dObj = parseYMD_(dVal) || parseDMY_(dVal);
    if (!dObj) return;

    const ym = Utilities.formatDate(dObj, tz, 'yyyy-MM');
    if (ym !== periodKey) return;

    recordCount++;

    const name = (r[2] || '').toString().trim();
    const statusRaw = (r[3] || '').toString().trim();
    const status = statusRaw.toLowerCase();

    const ot = Number(r[4] || 0) || 0;
    const fn = Number(r[5] || 0) || 0;

    if (!byId[id]) {
      byId[id] = { id: id, name: name, present: 0, absent: 0, otHours: 0, fine: 0 };
    } else if (name && !byId[id].name) {
      byId[id].name = name;
    }

    // Status counts (no assumptions beyond Present/Absent)
    if (status === 'present') {
      byId[id].present += 1;
      presentDays += 1;
    } else if (status === 'absent') {
      byId[id].absent += 1;
      absentDays += 1;
    } else {
      unknownStatus += 1;
    }

    byId[id].otHours += ot;
    byId[id].fine += fn;

    otHours += ot;
    fine += fn;
  });

  const rows = Object.keys(byId).map(k => byId[k]);

  // Sort by ID (numeric if possible, else lexicographic)
  rows.sort((a, b) => {
    const an = Number(a.id), bn = Number(b.id);
    const aNum = Number.isFinite(an) && String(an) === a.id;
    const bNum = Number.isFinite(bn) && String(bn) === b.id;
    if (aNum && bNum) return an - bn;
    return a.id.localeCompare(b.id);
  });

  return {
    period: periodKey,
    totals: {
      staffCount: rows.length,
      recordCount: recordCount,
      presentDays: presentDays,
      absentDays: absentDays,
      otHours: otHours,
      fine: fine
    },
    byStaff: rows,
    warnings: { unknownStatus: unknownStatus }
  };
}

/***********************
 * REPORTING BACKEND  *
 * Ledger read-only    *
 ***********************/
function getAvailableMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const cols = getLedgerCols_();
  const ledger = cols.sheet;

  const lastRow = ledger.getLastRow();
  if (lastRow <= cols.headerRow) return [];

  const dateVals = ledger.getRange(cols.headerRow + 1, cols.date, lastRow - cols.headerRow, 1).getValues();
  const seen = new Set();

  dateVals.forEach(r => {
    const v = r[0];
    let d = null;
    if (v instanceof Date && !isNaN(v)) d = v;
    else d = parseYMD_(v);
    if (!d) return;
    const key = Utilities.formatDate(d, tz, 'yyyy-MM');
    seen.add(key);
  });

  const months = Array.from(seen).sort();
  return months.map(k => {
    const [yy, mm] = k.split('-');
    return { year: Number(yy), month: Number(mm), key: k, label: k };
  });
}

function getMonthlyReportData(year, month) {
  const y = Number(year);
  const m = Number(month);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) {
    throw new Error('Invalid year/month');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const start = new Date(y, m - 1, 1, 12, 0, 0);
  const end = new Date(y, m, 0, 12, 0, 0);

  return buildLedgerReport_(toDateKey_(start, tz), toDateKey_(end, tz));
}

function getCustomReportData(startDateKey, endDateKey) {
  const start = (startDateKey || '').toString().trim();
  const end = (endDateKey || '').toString().trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
    throw new Error('Invalid date format. Use YYYY-MM-DD.');
  }
  if (start > end) throw new Error('Start date must be <= end date.');

  return buildLedgerReport_(start, end);
}

function buildLedgerReport_(startKey, endKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const cols = getLedgerCols_();
  const ledger = cols.sheet;

  const lastRow = ledger.getLastRow();
  const result = {
    startDate: startKey,
    endDate: endKey,
    totals: { income: 0, expense: 0, net: 0, txCount: 0 },
    byHeader: []
  };

  if (lastRow <= cols.headerRow) return result;

  const minCol = Math.min(cols.date, cols.header, cols.in, cols.out);
  const maxCol = Math.max(cols.date, cols.header, cols.in, cols.out);
  const width = maxCol - minCol + 1;

  const values = ledger.getRange(cols.headerRow + 1, minCol, lastRow - cols.headerRow, width).getValues();
  const dateOff = cols.date - minCol;
  const headerOff = cols.header - minCol;
  const inOff = cols.in - minCol;
  const outOff = cols.out - minCol;

  let totalIn = 0, totalOut = 0, tx = 0;
  const by = {};

  values.forEach(r => {
    const dv = r[dateOff];
    let d = null;
    if (dv instanceof Date && !isNaN(dv)) d = dv;
    else d = parseYMD_(dv);
    if (!d) return;

    const key = toDateKey_(d, tz);
    if (key < startKey || key > endKey) return;

    const IN = Number(r[inOff]) || 0;
    const OUT = Number(r[outOff]) || 0;
    if (IN === 0 && OUT === 0) return;

    const header = ((r[headerOff] || '').toString().trim()) || '(Blank Header)';

    totalIn += IN;
    totalOut += OUT;
    tx++;

    if (!by[header]) by[header] = { in: 0, out: 0, count: 0 };
    by[header].in += IN;
    by[header].out += OUT;
    by[header].count++;
  });

  result.totals.income = totalIn;
  result.totals.expense = totalOut;
  result.totals.net = totalIn - totalOut;
  result.totals.txCount = tx;

  result.byHeader = Object.keys(by).map(h => {
    const IN = by[h].in;
    const OUT = by[h].out;
    return { header: h, income: IN, expense: OUT, net: IN - OUT, txCount: by[h].count };
  });

  return result;
}
// Public wrapper (so google.script.run always sees it)
function appendReturnLog(reportDate, rows) {
  return appendReturnLog_(reportDate, rows);
}


