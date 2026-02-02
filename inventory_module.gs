/**
 * Inventory + Chalan Module (Additive)
 * Production rules:
 * - Append-only logs: Chalan_Entries, Inventory_AutoTopup_Log
 * - Output sheet Inventory_Daily_Slip is regenerated (safe to overwrite)
 * - Does NOT modify Ledger / POS raw sheets
 */

const INV = {
  tzFallback: 'Asia/Dhaka',
  sheets: {
    productList: 'Product list',
    chalanItems: 'Chalan_Items',
    chalanEntries: 'Chalan_Entries',
    autoTopupLog: 'Inventory_AutoTopup_Log',
    dailySlip: 'Inventory_Daily_Slip',
    posProduct: 'Q-Sales Data'
  },
  headers: {
    chalanItems: ['sl', 'purchase_name', 'rate', 'active'],
    chalanEntries: ['entry_id','entry_type','report_date','purchase_name','rate','qty_in','line_total','created_at','created_by'],
    autoTopup: ['run_id','report_date','item_name','in_before','out','auto_topup_added','in_after','rate','total','created_at'],
    slipTable: ['Name','IN','OUT','REMAIN']
  },
  entryTypes: {
    PLANNED: 'PLANNED',
    TOPUP: 'TOPUP'
  }
};

/** ============= PUBLIC UI ============= */



/** ============= SETUP / SHEETS ============= */

function invEnsureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create module-owned sheets if missing (do NOT auto-create user-owned or POS-owned sources)
  [
    INV.sheets.chalanItems,
    INV.sheets.chalanEntries,
    INV.sheets.autoTopupLog,
    INV.sheets.dailySlip
  ].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // Required sources (must already exist)
  const prod = ss.getSheetByName(INV.sheets.productList);
  if (!prod) throw new Error(`Product list sheet not found: ${INV.sheets.productList}`);

  const posProd = ss.getSheetByName(INV.sheets.posProduct);
  if (!posProd) throw new Error(`POS product sheet not found: ${INV.sheets.posProduct}. Run POS Sync first.`);

  // Chalan_Items schema
  const items = ss.getSheetByName(INV.sheets.chalanItems);
  invEnsureHeaderRow_(items, INV.headers.chalanItems);

  // Chalan_Entries schema (append-only)
  const entries = ss.getSheetByName(INV.sheets.chalanEntries);
  invEnsureHeaderRow_(entries, INV.headers.chalanEntries);
  invSetTextFormat_(entries, 'C:C'); // report_date text

  // AutoTopup log schema (append-only)
  const autoLog = ss.getSheetByName(INV.sheets.autoTopupLog);
  invMigrateAutoTopupLogSchema_(autoLog);
  invEnsureHeaderRow_(autoLog, INV.headers.autoTopup);
  invSetTextFormat_(autoLog, 'B:B'); // report_date text

  // Daily slip output schema (regenerated)
  const slip = ss.getSheetByName(INV.sheets.dailySlip);
  if (slip.getLastRow() === 0) {
    slip.getRange(1, 1).setValue('DATE');
    slip.getRange(1, 2).setValue('');
    slip.getRange(3, 1, 1, INV.headers.slipTable.length).setValues([INV.headers.slipTable]).setFontWeight('bold');
  }

  // POS product sheet: must exist for OUT. If missing, show clear error when generating.
}

function invEnsureHeaderRow_(sheet, expectedHeader) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]).setFontWeight('bold');
    return;
  }

  // If header is blank row 1, set it. If mismatch, STOP to prevent corruption.
  const existing = sheet.getRange(1, 1, 1, expectedHeader.length).getValues()[0];
  const existingJoined = existing.map(x => (x == null ? '' : String(x))).join('|');
  const expectedJoined = expectedHeader.join('|');

  const allBlank = existing.every(c => (c == null || String(c).trim() === ''));
  if (allBlank) {
    sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]).setFontWeight('bold');
    return;
  }

  if (existingJoined !== expectedJoined) {
    throw new Error(
      `Schema mismatch in sheet: ${sheet.getName()}\n` +
      `Expected: ${expectedJoined}\nFound: ${existingJoined}\n` +
      `Refusing to continue to prevent corruption.`
    );
  }
}
/**
 * This inserts 2 columns AFTER in_after (col 7), shifting created_at to the end.
 * If sheet header is neither OLD nor NEW, STOP to prevent corruption.
 */
function invMigrateAutoTopupLogSchema_(sheet) {
  const OLD = ['run_id','report_date','item_name','in_before','out','auto_topup_added','in_after','created_at'];
  const NEW = INV.headers.autoTopup;

  if (!sheet) throw new Error('AutoTopup log sheet missing');

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return; // will be created by ensureHeaderRow_

  // Check old header (first 8 cols)
  const oldExisting = sheet.getRange(1, 1, 1, OLD.length).getValues()[0].map(x => (x == null ? '' : String(x).trim()));
  const oldJoined = oldExisting.join('|');
  const oldExpected = OLD.join('|');

  // Check new header (first NEW.length cols)
  const newExisting = sheet.getRange(1, 1, 1, NEW.length).getValues()[0].map(x => (x == null ? '' : String(x).trim()));
  const newJoined = newExisting.join('|');
  const newExpected = NEW.join('|');

  // If already NEW, done
  if (newJoined === newExpected) return;

  // If exactly OLD, migrate
  if (oldJoined === oldExpected) {
    // Insert rate,total after column 7 (in_after)
    sheet.insertColumnsAfter(7, 2);
    sheet.getRange(1, 1, 1, NEW.length).setValues([NEW]).setFontWeight('bold');
    return;
  }

  // Unknown schema => STOP
  throw new Error(
    `Schema mismatch in sheet: ${sheet.getName()}\n` +
    `Expected OLD: ${oldExpected}\nFound: ${oldJoined}\n` +
    `Expected NEW: ${newExpected}\nFound: ${newJoined}\n` +
    `Refusing to continue to prevent corruption.`
  );
}

function invSetTextFormat_(sheet, a1Range) {
  sheet.getRange(a1Range).setNumberFormat('@');
}

/** ============= PRODUCT LIST (Mapping) ============= */

function invNormName_(s) {
  return (s == null ? '' : String(s))
    .replace(/\u00A0/g, ' ')
    .replace(/[\t\r\n]+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function invNormHeader_(s) {
  return invNormName_(s);
}

/** ============= PRODUCT LIST (Mapping - CORRECTED INDEX A-G) ============= */

function invGetProductMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.productList);
  if (!sh) throw new Error(`Product list sheet not found: ${INV.sheets.productList}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    return { map: {}, meta: { ok: false, reason: 'Product list is empty.' } };
  }

  /**
   * আপনার হেডার অনুযায়ী সঠিক ইনডেক্স:
   * SL(0), Seals Name(1), Purchase name(2), Code(3), Price(4), Unit(5), R-Price(6)
   */
  const idx = {
    posName: 1,      // Column B (Seals Name)
    purchaseName: 2, // Column C (Purchase name)
    unit: 5,         // Column F (Unit) - সংশোধিত
    rate: 6          // Column G (R-Price) - সংশোধিত
  };

  const data = sh.getRange(2, 1, lastRow - 1, Math.max(lastCol, 7)).getValues();
  const map = {};

  data.forEach(r => {
    const posRaw = r[idx.posName];
    const posKey = invNormName_(posRaw);
    if (!posKey) return;

    const purchase = String(r[idx.purchaseName] || '').trim();
    const unitVal = r[idx.unit]; // Column F

    let unit = NaN;
    if (typeof unitVal === 'number') unit = unitVal;
    else {
      const s = String(unitVal || '').trim();
      unit = s ? Number(s) : NaN;
    }
    
    // Unit 0 বা ভুল হলে ক্যালকুলেশন নষ্ট হবে না, ১ ধরে নেওয়া হবে
    if (!Number.isFinite(unit) || unit <= 0) unit = 1;

    let rate = NaN;
    const rv = r[idx.rate]; // Column G (R-Price)
    if (typeof rv === 'number') rate = rv;
    else {
      const s = String(rv || '').trim().replace(/,/g, '');
      rate = s ? Number(s) : NaN;
    }

    map[posKey] = {
      posName: String(posRaw || '').trim(),
      purchaseName: purchase,
      unit: unit,
      rate: Number.isFinite(rate) ? rate : 0
    };
  });

  return { map, meta: { ok: true, idx } };
}

/**
 * Product list থেকে ইনডেক্স তৈরি (CORRECTED INDEX B,C,G)
 */
function invBuildProductIndexes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.productList);
  if (!sh) throw new Error('Product list sheet not found: ' + INV.sheets.productList);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { byPurchase: {}, salesToPurchase: {} };

  /**
   * SL(0), Seals Name(1), Purchase name(2), Code(3), Price(4), Unit(5), R-Price(6)
   */
  const iSales = 1; // Column B
  const iPurch = 2; // Column C
  const iPrice = 6; // Column G

  const byPurchase = {};
  const salesToPurchase = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const purchRaw = String(row[iPurch] || '').trim();
    if (!purchRaw) continue;
    const purchKey = invNormName_(purchRaw);

    const price = Number(row[iPrice]); // Column G এর রেট
    if (Number.isFinite(price)) {
      byPurchase[purchKey] = price;
    }

    const salesRaw = String(row[iSales] || '').trim();
    if (salesRaw) {
      const salesKey = invNormName_(salesRaw);
      salesToPurchase[salesKey] = purchKey;
    }
  }

  return { byPurchase, salesToPurchase };
}

/**
 * item_name যেটাই আসুক (purchase বা sales alias),
 * deterministic ভাবে rate বের করে।
 */
function invResolveRateForItem_(idx, itemName) {
  const k = invNormName_(itemName);

  // 1) direct purchase match
  if (idx.byPurchase[k] != null) return idx.byPurchase[k];

  // 2) sales alias -> purchase -> rate
  const purchKey = idx.salesToPurchase[k];
  if (purchKey && idx.byPurchase[purchKey] != null) return idx.byPurchase[purchKey];

  return null;
}

/**
 * Build { normalizedPurchaseName -> rate } map from Product list.
 * Uses Purchase name column and Price/Rate column.
 * Throws on duplicate purchase names with conflicting rates.
 */
function invGetPurchaseRateMap_() {
  const out = {};
  const seen = {};

  const res = invGetProductMap_();
  const map = res.map || {};
  Object.keys(map).forEach(k => {
    const p = map[k];
    const rawName = (p && p.purchaseName) ? String(p.purchaseName).trim() : '';
    if (!rawName) return;

    const key = invNormName_(rawName);
    if (!key) return;

    const rate = (p && typeof p.rate === 'number') ? p.rate : null;
    if (rate == null || !Number.isFinite(rate)) return;

    if (seen[key] != null && seen[key] !== rate) {
      throw new Error(`Product list duplicate purchase name with conflicting rate: "${rawName}" (${seen[key]} vs ${rate})`);
    }
    seen[key] = rate;
    out[key] = rate;
  });

  return out;
}

function invFindHeaderIndex_(headersNorm, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const c = candidates[i];
    const idx = headersNorm.indexOf(invNormHeader_(c));
    if (idx >= 0) return idx;
  }
  return null;
}

/** ============= CHALAN ITEMS ============= */

function invGetChalanItemsForUi() {
  invEnsureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.chalanItems);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { ok: false, items: [], message: 'Chalan_Items খালি। চাইলে Product list থেকে auto-build করতে পারো।' };
  }

  const data = sh.getRange(2, 1, lastRow - 1, INV.headers.chalanItems.length).getValues();
  const items = [];
  data.forEach(r => {
    const sl = Number(r[0]) || null;
    const name = (r[1] == null ? '' : String(r[1]).trim());
    const rate = (typeof r[2] === 'number') ? r[2] : Number(String(r[2] || '').replace(/,/g, ''));
    const active = (r[3] == null ? 'yes' : String(r[3]).trim().toLowerCase());
    if (!name) return;
    if (active && active !== 'yes' && active !== 'y' && active !== 'true' && active !== '1' && active !== 'active') return;
    if (!Number.isFinite(rate) || rate < 0) return;
    items.push({ sl: sl || 999999, purchaseName: name, rate });
  });

  items.sort((a,b) => a.sl - b.sl);
  return { ok: items.length > 0, items, message: items.length ? '' : 'Chalan_Items এ active item নাই।' };
}


function invGetLastChalanQtyMapForUi() {
  // Read-only helper for UI: last submitted qty per item (append-only scan).
  invEnsureSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.chalanEntries);
  if (!sh) return { ok: true, qtyByName: {} };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, qtyByName: {} };

  // Performance guard: scan recent rows only.
  const MAX_SCAN = 5000;
  const startRow = Math.max(2, lastRow - MAX_SCAN + 1);
  const numRows = lastRow - startRow + 1;

  const data = sh.getRange(startRow, 1, numRows, INV.headers.chalanEntries.length).getValues();

  // [entry_id, entry_type, report_date, purchase_name, rate, qty_in, line_total, created_at, created_by]
  const qtyByName = {};
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const name = (r[3] == null ? '' : String(r[3]).trim());
    if (!name) continue;
    if (qtyByName[name] != null) continue;

    const q = (typeof r[5] === 'number') ? r[5] : Number(String(r[5] || '').trim());
    qtyByName[name] = (Number.isFinite(q) && q >= 0) ? Math.floor(q) : 0;
  }

  return { ok: true, qtyByName };
}

function invSeedChalanItemsFromProductList() {
  // Additive: fills Chalan_Items if empty. Will NOT overwrite if already has rows.
  invEnsureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.chalanItems);
  if (sh.getLastRow() >= 2) {
    return { ok: false, message: 'Chalan_Items এ ডাটা আছে। Auto-build করা হবে না।' };
  }

  const pm = invGetProductMap_();
  if (!pm.meta.ok) return { ok: false, message: pm.meta.reason };

  // Collect unique purchase names with a usable rate
  const seen = new Set();
  const rows = [];
  let sl = 1;
  Object.keys(pm.map).forEach(posKey => {
    const rec = pm.map[posKey];
    const purchase = (rec.purchaseName || '').trim();
    if (!purchase) return;
    if (seen.has(purchase)) return;
    // rate can be null; still include but set 0 so user must fix
    const rate = (rec.rate == null) ? 0 : rec.rate;
    rows.push([sl, purchase, rate, 'yes']);
    seen.add(purchase);
    sl += 1;
  });

  if (!rows.length) {
    return { ok: false, message: 'Product list থেকে কোনো Purchase name পাওয়া যায়নি।' };
  }

  sh.getRange(2, 1, rows.length, INV.headers.chalanItems.length).setValues(rows);
  return { ok: true, message: `Auto-build done: ${rows.length} items.` };
}

/** ============= CHALAN SUBMIT ============= */

function invGetDefaultDatesForUi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || INV.tzFallback;
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const tomorrow = Utilities.formatDate(new Date(now.getTime() + 24*3600*1000), tz, 'yyyy-MM-dd');
  return { today, tomorrow };
}

/**
 * চ্যালান সাবমিট করা এবং প্রিন্ট কপি জেনারেট করা
 */
function invSubmitChalan(payload) {
  invEnsureSheets_();

  if (!payload || typeof payload !== 'object') throw new Error('Invalid payload.');
  const entryType = String(payload.entryType || '').trim().toUpperCase();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportDate = String(payload.reportDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(reportDate)) throw new Error('Invalid reportDate. Use YYYY-MM-DD.');

  // ডুপ্লিকেট চেক
  const entriesSh = ss.getSheetByName(INV.sheets.chalanEntries);
  const lastRow = entriesSh.getLastRow();
  if (lastRow > 1) {
    const existing = entriesSh.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existing.length; i++) {
      const exType = String(existing[i][1] || '').trim().toUpperCase();
      const exDate = String(existing[i][2] || '').trim();
      if (exType === entryType && exDate === reportDate) {
        throw new Error(`${entryType} already submitted for ${reportDate}.`);
      }
    }
  }

  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) throw new Error('No items provided.');

  // আইডি এবং টাইমস্ট্যাম্প তৈরি (Very Important)
  const entryId = Utilities.getUuid().split('-')[0].toUpperCase(); // ছোট এবং সুন্দর আইডি
  const createdAt = new Date();
  const createdBy = Session.getActiveUser().getEmail() || '';

  const rows = [];
  let total = 0;

  items.forEach(it => {
    const name = (it.purchaseName == null ? '' : String(it.purchaseName).trim());
    if (!name) return;
    const qty = Number(it.qty);
    if (qty === 0) return;
    const rate = Number(it.rate);
    const lineTotal = qty * rate;
    total += lineTotal;

    rows.push([entryId, entryType, reportDate, name, rate, qty, lineTotal, createdAt, createdBy]);
  });

  if (!rows.length) throw new Error('সব Qty 0। কিছুই সাবমিট হয়নি।');

  // ডেটা সেভ করা
  const sh = ss.getSheetByName(INV.sheets.chalanEntries);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, INV.headers.chalanEntries.length).setValues(rows);
  invSetTextFormat_(sh, 'C:C');

  // সরাসরি এইচটিএমএল জেনারেট করে রিটার্ন করা
  const chalanHtml = getChalanPrintHtml({
    entryId: entryId,
    reportDate: reportDate,
    rows: rows,
    total: total
  });

  return {
    ok: true,
    entryId: entryId,
    reportDate: reportDate,
    total: total,
    chalanHtml: chalanHtml
  };
}

/**
 * চ্যালান প্রিন্ট করার জন্য HTML জেনারেট করা (30 Rows A4 Version)
 */
function getChalanPrintHtml_A4_30Rows(payload) {
  const template = HtmlService.createTemplateFromFile('chalan_template');
  
  template.date = payload.reportDate;
  template.grandTotal = Number(payload.total).toLocaleString('en-IN');
  
  let rowsHtml = '';
  const TOTAL_ROWS = 30; // আপনি ৩০টি রো চেয়েছেন

  for (let i = 0; i < TOTAL_ROWS; i++) {
    if (i < payload.rows.length) {
      // আইটেম থাকলে ডেটা বসবে
      const r = payload.rows[i];
      const name = r[3];
      const rate = Number(r[4]).toLocaleString('en-IN');
      const qty = r[5];
      const total = Number(r[6]).toLocaleString('en-IN');
      
      rowsHtml += `<tr class="item-row">
        <td class="text-center">${i + 1}</td>
        <td>${name}</td>
        <td class="text-right">${rate}</td>
        <td class="text-center">${qty}</td>
        <td class="text-right">${total}</td>
      </tr>`;
    } else {
      // ৩০টি পূর্ণ করতে বাকি রোগুলো খালি থাকবে
      rowsHtml += `<tr class="item-row empty">
        <td class="text-center">${i + 1}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
      </tr>`;
    }
  }
  
  template.itemRows = rowsHtml;
  return template.evaluate().getContent();
}

function invComputeChalanSignature_({ entryType, reportDate, createdBy, rows }) {
  // rows format: [entryId, entryType, reportDate, name, rate, qty, lineTotal, createdAt, createdBy]
  const items = rows
    .map(r => [
      String(r[3] || '').trim().toLowerCase(), // name
      Number(r[4]) || 0,                       // rate
      Number(r[5]) || 0                        // qty
    ])
    .filter(t => t[0] && t[2] > 0)
    .sort((a, b) => {
      if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
      if (a[1] !== b[1]) return a[1] - b[1];
      return a[2] - b[2];
    });

  return [
    String(entryType || '').trim().toUpperCase(),
    String(reportDate || '').trim(),
    String(createdBy || '').trim().toLowerCase(),
    items.map(t => `${t[0]}~${t[1]}~${t[2]}`).join('|')
  ].join('||');
}

function invFindDuplicateChalanEntryId_(sheet, { entryType, reportDate, createdBy, rows, now }) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const WINDOW_MS = 3 * 60 * 1000;     // 3 minutes
  const LOOKBACK_ROWS = 2000;          // bounded scan

  const sigIncoming = invComputeChalanSignature_({ entryType, reportDate, createdBy, rows });

  const startRow = Math.max(2, lastRow - LOOKBACK_ROWS + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, INV.headers.chalanEntries.length).getValues();

  const nowMs = (now instanceof Date) ? now.getTime() : Date.now();
  const targetType = String(entryType || '').trim().toUpperCase();
  const targetDate = String(reportDate || '').trim();
  const targetBy = String(createdBy || '').trim().toLowerCase();

  // Group rows by entry_id
  const grouped = new Map(); // entry_id -> rows[]
  for (const r of data) {
    const eId = String(r[0] || '').trim();
    if (!eId) continue;

    const eType = String(r[1] || '').trim().toUpperCase();
    const eDate = String(r[2] || '').trim();
    const eBy = String(r[8] || '').trim().toLowerCase();

    if (eType !== targetType) continue;
    if (eDate !== targetDate) continue;
    if (eBy !== targetBy) continue;

    const createdAt = r[7];
    if (!(createdAt instanceof Date) || isNaN(createdAt)) continue;
    if ((nowMs - createdAt.getTime()) > WINDOW_MS) continue;

    if (!grouped.has(eId)) grouped.set(eId, []);
    grouped.get(eId).push(r);
  }

  for (const [eId, eRows] of grouped.entries()) {
    const sig = invComputeChalanSignature_({ entryType: targetType, reportDate: targetDate, createdBy: targetBy, rows: eRows });
    if (sig === sigIncoming) return eId;
  }

  return null;
}

function invBuildWhatsAppChalanText_({ entryType, reportDate, rows, total, tz }) {
  // WhatsApp format: "100% straight start" + stable structure
  // - SL + Item(short) + Qty table
  // - Per-item total removed (reduces alignment drift with Bangla)
  // - Detailed summary at bottom: Lines, Total Qty, Grand Total
  // - Title line under Date: "Chalan Deshivoj"

  const SL_W   = 2;
  const NAME_W = 18;  // short item name width
  const QTY_W  = 3;

  const toStr = v => (v == null ? '' : String(v));

  const addCommas_ = (s) => s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const fmtInt_ = (n) => {
    const x = Math.round(Number(n) || 0);
    return addCommas_(String(x));
  };

  const fmtQty2_ = (n) => {
    const x = Math.round(Number(n) || 0);
    return String(x).padStart(2, '0'); // 01, 10
  };

  const padR = (v, w) => {
    const s = toStr(v);
    return s.length >= w ? s.slice(0, w) : s + ' '.repeat(w - s.length);
  };

  const padL = (v, w) => {
    const s = toStr(v);
    return s.length >= w ? s.slice(-w) : ' '.repeat(w - s.length) + s;
  };

  // Short Bangla name: first 1–2 words, then truncate
  const shortName_ = (name) => {
    const raw = toStr(name).trim().replace(/\s+/g, ' ');
    if (!raw) return '';
    const parts = raw.split(' ');
    const keep = parts.slice(0, Math.min(2, parts.length)).join(' ');
    return keep.length > NAME_W ? keep.slice(0, NAME_W) : keep;
  };

  const items = [];
  let totalQty = 0;

  (rows || []).forEach(r => {
    const name = shortName_(r[3]); // purchase_name
    const qty  = Math.round(Number(r[5]) || 0); // qty_in
    if (!name || qty <= 0) return;
    items.push({ name, qty });
    totalQty += qty;
  });

  const lines = [];
  lines.push('```');
  lines.push('Date: ' + toStr(reportDate).trim());
  lines.push('Chalan Deshivoj');
  lines.push('');

  // Header
  lines.push(
    padR('SL', SL_W) + ' | ' +
    padR('Item', NAME_W) + ' | ' +
    padL('Qty', QTY_W)
  );

  // Separator
  lines.push('-'.repeat(SL_W) + '-+-' + '-'.repeat(NAME_W) + '-+-' + '-'.repeat(QTY_W));

  // Rows
  items.forEach((it, idx) => {
    const sl = String(idx + 1).padStart(2, '0');
    const qty2 = fmtQty2_(it.qty);
    lines.push(
      padR(sl, SL_W) + ' | ' +
      padR(it.name, NAME_W) + ' | ' +
      padL(qty2, QTY_W)
    );
  });

  lines.push('-'.repeat(SL_W) + '-+-' + '-'.repeat(NAME_W) + '-+-' + '-'.repeat(QTY_W));

  // Detailed totals
  lines.push('Lines: ' + items.length);
  lines.push('Total Qty: ' + fmtInt_(totalQty));
  lines.push('Grand Total: ' + fmtInt_(total));

  lines.push('```');

  return lines.join('\n');
}


/** ============= SLIP GENERATION ============= */

function invGetSlipForUi(reportDate) {
  invEnsureSheets_();
  const dateKey = String(reportDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Invalid date. Use YYYY-MM-DD.');

  const result = invGenerateDailySlip_(dateKey);
  return result;
}


/**
 * Inventory Range Report (read-only). No sheet writes, no auto-topup logs.
 * Returns per-item totals across dates: IN (after safe auto-topup), OUT, REMAIN (return).
 */
function invGetRangeReportForUi(startDate, endDate) {
  invEnsureSheets_();
  const startKey = String(startDate || '').trim();
  const endKey = String(endDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startKey) || !/^\d{4}-\d{2}-\d{2}$/.test(endKey)) {
    throw new Error('Invalid date format. Use YYYY-MM-DD.');
  }
  if (startKey > endKey) throw new Error('Start date must be <= end date.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || INV.tzFallback;

  // Build a rate map from Chalan_Items (purchase_name -> rate)
  const rateMap = {};
  try {
    const items = invGetChalanItemsForUi();
    if (items && items.ok && Array.isArray(items.items)) {
      items.items.forEach(it => {
        const k = invNormName_(it.purchaseName);
        if (!k) return;
        const r = Number(it.rate);
        if (Number.isFinite(r)) rateMap[k] = r;
      });
    }
  } catch (e) {
    // Ignore, report can still run without rates.
  }

  // Date iteration (inclusive)
  const startObj = parseYMD_(startKey);
  const endObj = parseYMD_(endKey);
  if (!startObj || !endObj) throw new Error('Invalid date values.');

  const dayMs = 24 * 3600 * 1000;
  const dates = [];
  for (let t = startObj.getTime(); t <= endObj.getTime(); t += dayMs) {
    const k = Utilities.formatDate(new Date(t), tz, 'yyyy-MM-dd');
    // Defensive: avoid drifting across DST (BD has no DST, but keep stable)
    if (k < startKey || k > endKey) continue;
    dates.push(k);
  }

  // Accumulators by item key
  const acc = {}; // key -> {name, inQty, outQty, remainQty}
  const names = {}; // best display name

  dates.forEach(dateKey => {
    const inMap = invLoadInByDate_(dateKey);
    const outMap = invLoadOutByDate_(dateKey);

    const keys = new Set();
    Object.keys(inMap).forEach(k => { if (!k.startsWith('__')) keys.add(k); });
    Object.keys(outMap).forEach(k => { if (!k.startsWith('__')) keys.add(k); });

    keys.forEach(k => {
      const inQty = Number(inMap[k] || 0);
      const outQty = Number(outMap[k] || 0);

      // Safe auto-topup: ensure remain never negative (does NOT write logs here)
      let adjIn = inQty;
      let remain = adjIn - outQty;
      if (remain < 0) {
        adjIn = adjIn + Math.abs(remain);
        remain = 0;
      }

      if (!acc[k]) {
        acc[k] = { inQty: 0, outQty: 0, remainQty: 0 };
      }

      acc[k].inQty += adjIn;
      acc[k].outQty += outQty;
      acc[k].remainQty += remain;

      // Best display name
      if (!names[k]) {
        names[k] = (inMap.__names && inMap.__names[k]) ? inMap.__names[k]
                : (outMap.__names && outMap.__names[k]) ? outMap.__names[k]
                : k;
      }
    });
  });

  // Ordering: Chalan_Items order then extras
  const order = invGetItemOrder_();
  const ordered = [];
  const seen = new Set();
  order.forEach(n => {
    const k = invNormName_(n);
    if (k && acc[k] && !seen.has(k)) {
      ordered.push(k);
      seen.add(k);
    }
  });
  Object.keys(acc).sort().forEach(k => {
    if (!seen.has(k)) ordered.push(k);
  });

  const rows = ordered.map(k => {
    const r = acc[k];
    const rate = (rateMap[k] == null) ? null : Number(rateMap[k]);
    const inAmount = (rate == null) ? null : r.inQty * rate;
    const outAmount = (rate == null) ? null : r.outQty * rate;
    const remainAmount = (rate == null) ? null : r.remainQty * rate;

    return {
      name: String(names[k] || k),
      inQty: invFmtQty_(r.inQty),
      outQty: invFmtQty_(r.outQty),
      remainQty: invFmtQty_(r.remainQty),
      rate: rate,
      inAmount: inAmount,
      outAmount: outAmount,
      remainAmount: remainAmount
    };
  });

  // Totals
  let tIn=0, tOut=0, tRem=0, tInAmt=0, tOutAmt=0, tRemAmt=0;
  rows.forEach(r => {
    const inQ = Number(r.inQty) || 0;
    const outQ = Number(r.outQty) || 0;
    const remQ = Number(r.remainQty) || 0;
    tIn += inQ; tOut += outQ; tRem += remQ;
    if (typeof r.inAmount === 'number') tInAmt += r.inAmount;
    if (typeof r.outAmount === 'number') tOutAmt += r.outAmount;
    if (typeof r.remainAmount === 'number') tRemAmt += r.remainAmount;
  });

  return {
    ok: true,
    startDate: startKey,
    endDate: endKey,
    totals: {
      inQty: invFmtQty_(tIn),
      outQty: invFmtQty_(tOut),
      remainQty: invFmtQty_(tRem),
      inAmount: tInAmt,
      outAmount: tOutAmt,
      remainAmount: tRemAmt
    },
    rows
  };
}

function invGenerateDailySlip_(dateKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || INV.tzFallback;

  const inMap = invLoadInByDate_(dateKey);
  const outMap = invLoadOutByDate_(dateKey);

  const allKeys = new Set();
  Object.keys(inMap).forEach(k => allKeys.add(k));
  Object.keys(outMap).forEach(k => allKeys.add(k));

  // Ordering: Chalan_Items order first, then extras sorted
  const order = invGetItemOrder_();
  const orderedKeys = [];
  const seen = new Set();

  order.forEach(name => {
    const k = invNormName_(name);
    if (!k) return;
    if (allKeys.has(k) && !seen.has(k)) {
      orderedKeys.push(k);
      seen.add(k);
    }
  });

  const extras = Array.from(allKeys).filter(k => !seen.has(k)).sort();
  extras.forEach(k => orderedKeys.push(k));

  const runId = Utilities.getUuid();
  const autoLogs = [];

  const rows = orderedKeys.map(k => {
    const inQty = Number(inMap[k] || 0);
    const outQty = Number(outMap[k] || 0);

    let adjIn = inQty;
    let remain = adjIn - outQty;
    let autoTopup = 0;

    if (remain < 0) {
      autoTopup = Math.abs(remain);
      adjIn = adjIn + autoTopup;
      remain = 0;

      autoLogs.push({
        runId,
        reportDate: dateKey,
        itemName: invKeyToDisplayName_(k, inMap, outMap),
        inBefore: inQty,
        out: outQty,
        autoTopup,
        inAfter: adjIn,
        createdAt: new Date()
      });
    }

    return {
      key: k,
      name: invKeyToDisplayName_(k, inMap, outMap),
      inQty: adjIn,
      outQty,
      remain
    };
  });

  invAppendAutoTopupLogs_(autoLogs);
  invWriteSlipSheet_(dateKey, rows);

  return {
    ok: true,
    reportDate: dateKey,
    runId,
    rows: rows.map(r => ({
      name: r.name,
      IN: invFmtQty_(r.inQty),
      OUT: invFmtQty_(r.outQty),
      REMAIN: invFmtQty_(r.remain)
    }))
  };
}

function invKeyToDisplayName_(key, inMap, outMap) {
  // key is normalized name. Prefer any original name stored in maps.
  const src = (inMap.__names && inMap.__names[key]) ? inMap.__names[key]
            : (outMap.__names && outMap.__names[key]) ? outMap.__names[key]
            : key;
  return String(src || key);
}

function invFmtQty_(n) {
  const x = Number(n || 0);
  if (!Number.isFinite(x)) return '0';
  const s = x.toFixed(3);
  // trim trailing zeros
  return s.replace(/\.?(0+)$/,'').replace(/\.$/,'');
}
/**
 * চ্যালান প্রিন্ট করার জন্য HTML জেনারেট করা
 */
function getChalanPrintHtml(payload) {
  const template = HtmlService.createTemplateFromFile('chalan_template');
  template.date = payload.reportDate;
  template.grandTotal = payload.total.toLocaleString();
  
  let rowsHtml = '';
  payload.rows.forEach((r, i) => {
    // r[3] = item name, r[4] = rate, r[5] = qty, r[6] = total
    rowsHtml += `<tr style="height: 22px">
      <td class="s8">${i + 1}</td>
      <td class="s10">${r[3]}</td>
      <td class="s9">${r[4]}</td>
      <td class="s9">${r[5]}</td>
      <td class="s9">${r[6].toLocaleString()}</td>
    </tr>`;
  });
  
  template.itemRows = rowsHtml;
  return template.evaluate().getContent();
}
function invGetItemOrder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.chalanItems);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, INV.headers.chalanItems.length).getValues();
  const items = [];
  data.forEach(r => {
    const sl = Number(r[0]) || 999999;
    const name = (r[1] == null ? '' : String(r[1]).trim());
    const active = (r[3] == null ? 'yes' : String(r[3]).trim().toLowerCase());
    if (!name) return;
    if (active && active !== 'yes' && active !== 'y' && active !== 'true' && active !== '1' && active !== 'active') return;
    items.push({ sl, name });
  });
  items.sort((a,b) => a.sl - b.sl);
  return items.map(x => x.name);
}

function invLoadInByDate_(dateKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.chalanEntries);
  if (!sh) throw new Error(`Sheet not found: ${INV.sheets.chalanEntries}`);

  const lastRow = sh.getLastRow();
  const map = { __names: {} };
  if (lastRow < 2) return map;

  const data = sh.getRange(2, 1, lastRow - 1, INV.headers.chalanEntries.length).getValues();
  // columns: entry_id, entry_type, report_date, purchase_name, rate, qty_in, line_total, ...
  data.forEach(r => {
    const d = (r[2] == null ? '' : String(r[2]).trim());
    if (d !== dateKey) return;

    const name = (r[3] == null ? '' : String(r[3]).trim());
    if (!name) return;
    const k = invNormName_(name);
    const qty = Number(r[5] || 0);
    if (!Number.isFinite(qty) || qty === 0) return;

    map[k] = (map[k] || 0) + qty;
    if (!map.__names[k]) map.__names[k] = name;
  });

  return map;
}

function invLoadOutByDate_(dateKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.posProduct);
  if (!sh) throw new Error(`POS product sheet not found: ${INV.sheets.posProduct}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) {
    return { __names: {} };
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(invNormHeader_);

  const idxDate = headers.indexOf('report_date');
  const idxName = headers.indexOf('name');
  const idxQty = headers.indexOf('total_sale');

  if (idxDate < 0 || idxName < 0 || idxQty < 0) {
    throw new Error('Q-Sales Data header missing required columns: report_date, name, total_sale');
  }

  const pm = invGetProductMap_();
  if (!pm.meta.ok) throw new Error(pm.meta.reason);

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = { __names: {} };

  data.forEach(r => {
    const d = (r[idxDate] == null ? '' : String(r[idxDate]).trim());
    if (d !== dateKey) return;

    const posName = (r[idxName] == null ? '' : String(r[idxName]).trim());
    if (!posName) return;

    const qtyRaw = r[idxQty];
    const posQty = (typeof qtyRaw === 'number') ? qtyRaw : Number(String(qtyRaw || '').trim());
    if (!Number.isFinite(posQty) || posQty === 0) return;

    const posKey = invNormName_(posName);
    const rec = pm.map[posKey];

    const purchaseName = rec ? (rec.purchaseName || '') : '';
    const unit = rec ? rec.unit : 1;

    const outQty = posQty * unit;
    const keyName = purchaseName ? purchaseName : posName; // if no chalan name, keep POS name
    const k = invNormName_(keyName);

    out[k] = (out[k] || 0) + outQty;
    if (!out.__names[k]) out.__names[k] = keyName;
  });

  return out;
}
function invAppendAutoTopupLogs_(logs) {
  if (!logs || !logs.length) return;
  invEnsureSheets_();

  const idx = invBuildProductIndexes_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.autoTopupLog);

  const rateMap = invGetPurchaseRateMap_();
  const missingRates = [];

  // Dedup exact signatures to avoid repeat spam when user clicks print multiple times
  const existingSig = invLoadRecentAutoSig_(sh, 3000);

  const rows = [];
  logs.forEach(l => {
    const sig = invAutoSig_(l);
    if (existingSig.has(sig)) return;
    existingSig.add(sig);

    const itemKey = invNormName_(l.itemName);
    const rate = invResolveRateForItem_(idx, l.itemName);

    if (rate == null) {
      missingRates.push(l.itemName);
      return;
    }

    const total = Number(l.inAfter || 0) * Number(rate || 0);

    rows.push([
      l.runId,
      l.reportDate,
      l.itemName,
      l.inBefore,
      l.out,
      l.autoTopup,
      l.inAfter,
      rate,
      total,
      l.createdAt
    ]);
  });

  // Option B: rate missing হলে BLOCK (append হবে না)
  if (missingRates.length) {
    const uniq = Array.from(new Set(missingRates.map(x => String(x || '').trim()).filter(Boolean)));
    throw new Error('Inventory_AutoTopup_Log: Product list এ rate পাওয়া যায়নি: ' + uniq.slice(0, 50).join(', '));
  }

  if (!rows.length) return;
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, INV.headers.autoTopup.length).setValues(rows);
  invSetTextFormat_(sh, 'B:B');
}

function invAutoSig_(l) {
  return [l.reportDate, invNormName_(l.itemName), Number(l.inBefore||0), Number(l.out||0), Number(l.autoTopup||0)].join('|');
}

function invLoadRecentAutoSig_(sheet, limitRows) {
  const sigs = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return sigs;

  const start = Math.max(2, lastRow - limitRows + 1);
  const num = lastRow - start + 1;
  const data = sheet.getRange(start, 1, num, INV.headers.autoTopup.length).getValues();
  // columns: run_id, report_date, item_name, in_before, out, auto_topup_added
  data.forEach(r => {
    const reportDate = (r[1] == null ? '' : String(r[1]).trim());
    const itemName = (r[2] == null ? '' : String(r[2]).trim());
    if (!reportDate || !itemName) return;
    const inBefore = Number(r[3] || 0);
    const out = Number(r[4] || 0);
    const autoTopup = Number(r[5] || 0);
    sigs.add([reportDate, invNormName_(itemName), inBefore, out, autoTopup].join('|'));
  });
  return sigs;
}

function invWriteSlipSheet_(dateKey, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INV.sheets.dailySlip);
  if (!sh) throw new Error(`Sheet not found: ${INV.sheets.dailySlip}`);

  // Clear output region but keep sheet
  sh.clear();

  sh.getRange(1, 1).setValue('DATE');
  sh.getRange(1, 2).setValue(dateKey);
  sh.getRange(3, 1, 1, INV.headers.slipTable.length).setValues([INV.headers.slipTable]).setFontWeight('bold');

  if (!rows.length) {
    sh.getRange(5, 1).setValue('No data for this date.');
    return;
  }

  const outRows = rows.map(r => [r.name, r.inQty, r.outQty, r.remain]);
  sh.getRange(4, 1, outRows.length, INV.headers.slipTable.length).setValues(outRows);

  // number formats for qty columns
  sh.getRange(4, 2, outRows.length, 3).setNumberFormat('0.###');
  sh.autoResizeColumns(1, 4);
}