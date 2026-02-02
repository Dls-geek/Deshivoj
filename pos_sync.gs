/**
 * POS Sync (Apps Script)
 * - Password prompt each run (NOT saved)
 * - Session-based login (Laravel CSRF)
 * - Fetches DataTables JSON via GET with retry (min -> full params)
 * - Append-only write to: Sales Data, Q-Sales Data, Sales run.log
 * - Dedup via runs sheet (status=success only)
 *
 * Script Properties required:
 *   POS_EMAIL          = your POS login email
 *   POS_BUSSINESS_ID   = 257
 */

const POS_SYNC = {
  tabs: {
    summary: 'Sales Data',
    product: 'Q-Sales Data',
    runs: 'Sales run.log'
  },
  runsHeader: ['report_date','report_type','sha256','file_name','row_count','run_ts_utc','status','error'],
};

function posSyncOnOpen_() {
  SpreadsheetApp.getUi()
    .createMenu('POS Sync')
    .addItem('Import EOD', 'posSyncImportEOD')
    .addToUi();
}

function posSyncImportEOD() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert('POS Sync is already running. Try again in 30 seconds.');
    return;
  }

  let reportDate = '';
  try {
    const props = PropertiesService.getScriptProperties();
    const baseUrl = (props.getProperty('POS_BASE_URL') || '').trim();
    const email = (props.getProperty('POS_EMAIL') || '').trim();
    const businessId = (props.getProperty('POS_BUSSINESS_ID') || '').trim();

    if (!baseUrl || !email || !businessId) {
      ui.alert('Missing Script Properties: POS_BASE_URL, POS_EMAIL, POS_BUSSINESS_ID');
      return;
    }

    // Date prompt
    const now = new Date();
    // FIXED: Default is NOW (Today) instead of Yesterday
    const defaultDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    const dateResp = ui.prompt(
      'POS Sync — Import EOD',
      `Enter report date (YYYY-MM-DD)\nDefault: ${defaultDate}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (dateResp.getSelectedButton() !== ui.Button.OK) return;

    reportDate = (dateResp.getResponseText() || '').trim() || defaultDate;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(reportDate)) {
      ui.alert('Invalid date format. Use YYYY-MM-DD.');
      return;
    }

    // Password prompt (NOT saved)
    const passResp = ui.prompt(
      'POS Sync — Password',
      'Enter POS password (will NOT be saved):',
      ui.ButtonSet.OK_CANCEL
    );
    if (passResp.getSelectedButton() !== ui.Button.OK) return;

    const password = (passResp.getResponseText() || '').trim();
    if (!password) {
      ui.alert('Password is required.');
      return;
    }

    // Ensure tabs + runs schema + prevent date auto convert
    posSyncEnsureTabs_();

    // Dedup map (success only)
    const dedup = posSyncGetDedupMap_();
    const summaryKey = `${reportDate}|summary`;
    const productKey = `${reportDate}|product`;

    if (dedup[summaryKey] && dedup[productKey]) {
      posSyncAppendRun_({
        report_date: reportDate,
        report_type: 'ALL',
        sha256: '',
        file_name: '',
        row_count: 0,
        run_ts_utc: new Date().toISOString(),
        status: 'skipped',
        error: 'Both report types already imported successfully.'
      });
      ss.toast(`SKIPPED: ${reportDate} already imported`, 'POS Sync', 8);
      return;
    }

    ss.toast(`POS Sync started for ${reportDate}…`, 'POS Sync', 6);

    // Login once
    const jar = posSyncLogin_(baseUrl, email, password);

    // SUMMARY
    if (!dedup[summaryKey]) {
      try {
        const summaryResult = posSyncFetchSummary_(baseUrl, jar, reportDate);
        posSyncAppendRaw_({
          tabName: POS_SYNC.tabs.summary,
          reportDate,
          headers: summaryResult.headers,
          rows: summaryResult.rows
        });
        posSyncAppendRun_({
          report_date: reportDate,
          report_type: 'summary',
          sha256: summaryResult.sha256,
          file_name: summaryResult.source,
          row_count: summaryResult.rows.length,
          run_ts_utc: summaryResult.run_ts_utc,
          status: 'success',
          error: ''
        });
      } catch (e) {
        posSyncAppendRun_({
          report_date: reportDate,
          report_type: 'summary',
          sha256: '',
          file_name: 'xhr:/sale',
          row_count: 0,
          run_ts_utc: new Date().toISOString(),
          status: 'error',
          error: String(e)
        });
        throw e;
      }
    } else {
      posSyncAppendRun_({
        report_date: reportDate,
        report_type: 'summary',
        sha256: '',
        file_name: 'xhr:/sale',
        row_count: 0,
        run_ts_utc: new Date().toISOString(),
        status: 'skipped',
        error: 'Already imported successfully earlier.'
      });
    }

    // PRODUCT
    if (!dedup[productKey]) {
      try {
        const productResult = posSyncFetchProduct_(baseUrl, jar, reportDate, String(businessId));
        posSyncAppendRaw_({
          tabName: POS_SYNC.tabs.product,
          reportDate,
          headers: productResult.headers,
          rows: productResult.rows
        });
        posSyncAppendRun_({
          report_date: reportDate,
          report_type: 'product',
          sha256: productResult.sha256,
          file_name: productResult.source,
          row_count: productResult.rows.length,
          run_ts_utc: productResult.run_ts_utc,
          status: 'success',
          error: ''
        });
      } catch (e) {
        posSyncAppendRun_({
          report_date: reportDate,
          report_type: 'product',
          sha256: '',
          file_name: 'xhr:/report/top-selling-product',
          row_count: 0,
          run_ts_utc: new Date().toISOString(),
          status: 'error',
          error: String(e)
        });
        throw e;
      }
    } else {
      posSyncAppendRun_({
        report_date: reportDate,
        report_type: 'product',
        sha256: '',
        file_name: 'xhr:/report/top-selling-product',
        row_count: 0,
        run_ts_utc: new Date().toISOString(),
        status: 'skipped',
        error: 'Already imported successfully earlier.'
      });
    }

    ss.toast(`POS Sync OK: ${reportDate}`, 'POS Sync', 10);

  } finally {
    lock.releaseLock();
  }
}

/** =========================
 *   Schema / Dedup Helpers
 *  ========================= */

function posSyncNormalizeDateKey_(v, tz) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return (v == null ? '' : String(v)).trim();
}

function posSyncSetTextFormat_(sheet, a1Range) {
  sheet.getRange(a1Range).setNumberFormat('@'); // Plain text
}

function posSyncEnsureTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create tabs if missing
  [POS_SYNC.tabs.summary, POS_SYNC.tabs.product, POS_SYNC.tabs.runs].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // runs schema
  const runsSh = ss.getSheetByName(POS_SYNC.tabs.runs);
  const lastRow = runsSh.getLastRow();

  if (lastRow === 0) {
    runsSh.getRange(1, 1, 1, POS_SYNC.runsHeader.length).setValues([POS_SYNC.runsHeader]);
  } else {
    const header = runsSh.getRange(1, 1, 1, POS_SYNC.runsHeader.length).getValues()[0];
    if (header.join('|') !== POS_SYNC.runsHeader.join('|')) {
      throw new Error("runs tab schema mismatch. Refusing to continue to prevent corruption.");
    }
  }

  // Prevent report_date auto conversion
  posSyncSetTextFormat_(ss.getSheetByName(POS_SYNC.tabs.runs), 'A:A');
  posSyncSetTextFormat_(ss.getSheetByName(POS_SYNC.tabs.summary), 'A:A');
  posSyncSetTextFormat_(ss.getSheetByName(POS_SYNC.tabs.product), 'A:A');
}

function posSyncGetDedupMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(POS_SYNC.tabs.runs);
  const lastRow = sh.getLastRow();
  const map = {};
  if (lastRow < 2) return map;

  const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  data.forEach(r => {
    const reportDate = posSyncNormalizeDateKey_(r[0], tz);
    const reportType = (r[1] || '').toString().trim();
    const status = (r[6] || '').toString().trim().toLowerCase();
    if (reportDate && reportType && status === 'success') {
      map[`${reportDate}|${reportType}`] = true;
    }
  });
  return map;
}

function posSyncAppendRun_(runObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(POS_SYNC.tabs.runs);
  sh.appendRow([
    runObj.report_date,
    runObj.report_type,
    runObj.sha256 || '',
    runObj.file_name || '',
    Number(runObj.row_count || 0),
    runObj.run_ts_utc || new Date().toISOString(),
    runObj.status || '',
    runObj.error || ''
  ]);
}

function posSyncAppendRaw_({tabName, reportDate, headers, rows}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(tabName);

  const runTs = new Date().toISOString();
  const expectedHeader = ['report_date'].concat(headers).concat(['run_ts_utc']);

  const lastRow = sh.getLastRow();
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]);
  } else {
    const existing = sh.getRange(1, 1, 1, expectedHeader.length).getValues()[0];
    if (existing.join('|') !== expectedHeader.join('|')) {
      throw new Error(
        `Header mismatch in ${tabName}. Import STOPPED.\n` +
        `Expected: ${expectedHeader.join(', ')}\nFound: ${existing.join(', ')}`
      );
    }
  }

  if (!rows.length) return;

  const out = rows.map(r => [reportDate].concat(r).concat([runTs]));
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, out.length, expectedHeader.length).setValues(out);
}

/** =========================
 *    POS Login + Cookies
 *  ========================= */

function posSyncLogin_(baseUrl, email, password) {
  const jar = {};
  const cleanBase = baseUrl.replace(/\/+$/, '');
  const loginUrl = cleanBase + '/login';

  const res1 = posSyncFetch_(loginUrl, {
    method: 'get',
    jar,
    headers: { 'Accept': 'text/html' }
  });

  const token = posSyncExtractCsrfToken_(res1.bodyText);
  if (!token) throw new Error('Cannot find CSRF token on login page.');

  const payload = posSyncFormEncode_({ email, password, _token: token });

  const res2 = posSyncFetch_(loginUrl, {
    method: 'post',
    jar,
    headers: {
      'Accept': 'text/html',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Referer': loginUrl,
      'Origin': cleanBase
    },
    payload,
    followRedirects: false
  });

  if (res2.code !== 302 && res2.code !== 200) {
    throw new Error(`Login failed. HTTP ${res2.code}`);
  }

  const loc = res2.headers.Location || res2.headers.location;
  if (loc) {
    const nextUrl = loc.startsWith('http') ? loc : (cleanBase + loc);
    posSyncFetch_(nextUrl, { method: 'get', jar, headers: { 'Accept': 'text/html' } });
  }

  return jar;
}

function posSyncExtractCsrfToken_(html) {
  const m1 = html.match(/name="_token"\s+value="([^"]+)"/);
  if (m1) return m1[1];
  const m2 = html.match(/name="csrf-token"\s+content="([^"]+)"/);
  if (m2) return m2[1];
  return null;
}

function posSyncFormEncode_(obj) {
  return Object.keys(obj).map(k =>
    encodeURIComponent(k) + '=' + encodeURIComponent(obj[k] == null ? '' : String(obj[k]))
  ).join('&');
}

function posSyncCookieHeader_(jar) {
  return Object.keys(jar).map(k => `${k}=${jar[k]}`).join('; ');
}

function posSyncUpdateJarFromSetCookie_(jar, headers) {
  const sc = headers['Set-Cookie'] || headers['set-cookie'];
  if (!sc) return;

  let parts = [];
  if (Array.isArray(sc)) {
    parts = sc.slice();
  } else {
    parts = String(sc).split(/,(?=\s*[A-Za-z0-9_\-]+=)/);
  }

  parts.forEach(line => {
    const first = String(line).trim().split(';')[0];
    const idx = first.indexOf('=');
    if (idx > 0) {
      const name = first.substring(0, idx).trim();
      const val = first.substring(idx + 1).trim();
      if (name) jar[name] = val;
    }
  });
}

function posSyncFetch_(url, opt) {
  const method = (opt.method || 'get').toUpperCase();
  const jar = opt.jar || {};
  const headers = opt.headers || {};
  const followRedirects = (opt.followRedirects === undefined) ? true : !!opt.followRedirects;

  const cookie = posSyncCookieHeader_(jar);
  if (cookie) headers['Cookie'] = cookie;

  const params = {
    method,
    muteHttpExceptions: true,
    followRedirects,
    headers
  };
  if (opt.payload) params.payload = opt.payload;

  const resp = UrlFetchApp.fetch(url, params);
  const code = resp.getResponseCode();
  const allHeaders = resp.getAllHeaders() || {};
  posSyncUpdateJarFromSetCookie_(jar, allHeaders);

  return { code, headers: allHeaders, bodyText: resp.getContentText() };
}

/** =========================
 *   DataTables Fetch (GET)
 *  ========================= */

function posSyncFetchJsonWithRetry_(urlBuilderFn, fetchFn, contextLabel) {
  let res = fetchFn(urlBuilderFn('min'));
  if (res.code === 500) {
    res = fetchFn(urlBuilderFn('full'));
  }
  if (res.code !== 200) {
    const snippet = (res.bodyText || '').toString().slice(0, 200);
    throw new Error(`${contextLabel} XHR failed. HTTP ${res.code}. Body: ${snippet}`);
  }
  if ((res.bodyText || '').trim().startsWith('<')) {
    throw new Error(`${contextLabel} XHR returned HTML (login/session issue).`);
  }
  return JSON.parse(res.bodyText);
}

/** ---------- Summary ---------- */
function posSyncFetchSummary_(baseUrl, jar, reportDate) {
  const runTs = new Date().toISOString();
  const cleanBase = baseUrl.replace(/\/+$/, '');
  const endpoint = cleanBase + '/sale';

  const cols = [
    {data:'DT_RowIndex', name:'', searchable:'false', orderable:'false'},
    {data:'sale_date', name:'sale_date', searchable:'true', orderable:'false'},
    {data:'invoice_no', name:'invoice_no', searchable:'true', orderable:'false'},
    {data:'customer.name', name:'customer.name', searchable:'true', orderable:'false'},
    {data:'user.name', name:'user.name', searchable:'true', orderable:'false'},
    {data:'sub_total', name:'sub_total', searchable:'true', orderable:'false'},
    {data:'vat', name:'vat', searchable:'true', orderable:'false'},
    {data:'discount_amount', name:'discount_amount', searchable:'true', orderable:'false'},
    {data:'service_charge', name:'service_charge', searchable:'true', orderable:'false'},
    {data:'total', name:'total', searchable:'true', orderable:'false'},
    {data:'final_paying_amount', name:'final_paying_amount', searchable:'true', orderable:'false'},
    {data:'payment_method', name:'payment_method', searchable:'true', orderable:'false'},
    {data:'final_due_amount', name:'final_due_amount', searchable:'true', orderable:'false'},
    {data:'payment_status', name:'payment_status', searchable:'true', orderable:'false'},
    {data:'action', name:'action', searchable:'true', orderable:'false'}
  ];
  const headers = cols.map(c => c.data);

  const allRows = [];
  let start = 0;
  const length = 200;
  let draw = 1;

  while (true) {
    const builder = (mode) => posSyncBuildSummaryUrl_(endpoint, reportDate, cols, draw, start, length, mode);

    const obj = posSyncFetchJsonWithRetry_(
      builder,
      (url) => posSyncFetch_(url, {
        method: 'get',
        jar,
        headers: {
          'Accept': 'application/json, text/javascript, */*; q=0.01',
          'X-Requested-With': 'XMLHttpRequest',
          'Referer': endpoint
        }
      }),
      'Summary'
    );

    const data = obj.data || obj.aaData || [];
    if (!Array.isArray(data)) throw new Error('Summary JSON: data is not an array.');

    data.forEach(rowObj => {
      allRows.push(headers.map(h => posSyncGetByPath_(rowObj, h)));
    });

    if (data.length < length) break;
    start += length;
    draw += 1;
    if (start > 20000) throw new Error('Summary pagination exceeded safety limit.');
    Utilities.sleep(150);
  }

  const sha256 = posSyncSha256Hex_(JSON.stringify({headers, rows: allRows}));
  return { headers, rows: allRows, sha256, run_ts_utc: runTs, source: 'xhr:GET /sale' };
}

function posSyncBuildSummaryUrl_(endpoint, reportDate, cols, draw, start, length, mode) {
  const pairs = [];
  pairs.push(['sale_date', reportDate]);
  pairs.push(['term', '']);
  pairs.push(['user', '']);
  pairs.push(['customer', '']);
  pairs.push(['delivery_company', '']);
  pairs.push(['type', 'Regular']);
  pairs.push(['online_status', '']);

  pairs.push(['draw', String(draw)]);
  pairs.push(['start', String(start)]);
  pairs.push(['length', String(length)]);
  pairs.push(['search[value]', '']);
  pairs.push(['search[regex]', 'false']);
  pairs.push(['order[0][column]', '0']);
  pairs.push(['order[0][dir]', 'asc']);

  cols.forEach((c, i) => {
    pairs.push([`columns[${i}][data]`, c.data]);
    pairs.push([`columns[${i}][name]`, c.name || '']);
    pairs.push([`columns[${i}][searchable]`, c.searchable]);
    pairs.push([`columns[${i}][orderable]`, c.orderable]);

    if (mode === 'full') {
      pairs.push([`columns[${i}][search][value]`, '']);
      pairs.push([`columns[${i}][search][regex]`, 'false']);
    }
  });

  pairs.push(['_', String(Date.now())]);
  const qs = pairs.map(([k, v]) => k + '=' + encodeURIComponent(v == null ? '' : String(v))).join('&');
  return endpoint + '?' + qs;
}

/** ---------- Product ---------- */
function posSyncFetchProduct_(baseUrl, jar, reportDate, businessId) {
  const runTs = new Date().toISOString();
  const cleanBase = baseUrl.replace(/\/+$/, '');
  const endpoint = cleanBase + '/report/top-selling-product';

  const cols = [
    {data:'DT_RowIndex', name:'', searchable:'false', orderable:'false'},
    {data:'name', name:'name', searchable:'true', orderable:'false'},
    {data:'total_sale', name:'total_sale', searchable:'true', orderable:'false'},
    {data:'total_sale_amount', name:'total_sale_amount', searchable:'true', orderable:'false'}
  ];
  const headers = cols.map(c => c.data);

  const allRows = [];
  let start = 0;
  const length = 500;
  let draw = 1;

  while (true) {
    const builder = (mode) => posSyncBuildProductUrl_(endpoint, reportDate, businessId, cols, draw, start, length, mode);

    const obj = posSyncFetchJsonWithRetry_(
      builder,
      (url) => posSyncFetch_(url, {
        method: 'get',
        jar,
        headers: {
          'Accept': 'application/json, text/javascript, */*; q=0.01',
          'X-Requested-With': 'XMLHttpRequest',
          'Referer': endpoint
        }
      }),
      'Product'
    );

    const data = obj.data || obj.aaData || [];
    if (!Array.isArray(data)) throw new Error('Product JSON: data is not an array.');

    data.forEach(rowObj => {
      allRows.push(headers.map(h => posSyncGetByPath_(rowObj, h)));
    });

    if (data.length < length) break;
    start += length;
    draw += 1;
    if (start > 20000) throw new Error('Product pagination exceeded safety limit.');
    Utilities.sleep(150);
  }

  const sha256 = posSyncSha256Hex_(JSON.stringify({headers, rows: allRows}));
  return { headers, rows: allRows, sha256, run_ts_utc: runTs, source: 'xhr:GET /report/top-selling-product' };
}

function posSyncBuildProductUrl_(endpoint, reportDate, businessId, cols, draw, start, length, mode) {
  const pairs = [];
  pairs.push(['bussiness_id', String(businessId)]);
  pairs.push(['starting_date', reportDate]);
  pairs.push(['ending_date', reportDate]);

  pairs.push(['draw', String(draw)]);
  pairs.push(['start', String(start)]);
  pairs.push(['length', String(length)]);
  pairs.push(['search[value]', '']);
  pairs.push(['order[0][column]', '0']);
  pairs.push(['order[0][dir]', 'asc']);

  cols.forEach((c, i) => {
    pairs.push([`columns[${i}][data]`, c.data]);
    pairs.push([`columns[${i}][name]`, c.name || '']);
    pairs.push([`columns[${i}][searchable]`, c.searchable]);
    pairs.push([`columns[${i}][orderable]`, c.orderable]);

    if (mode === 'full') {
      pairs.push([`columns[${i}][search][value]`, '']);
      pairs.push([`columns[${i}][search][regex]`, 'false']);
    }
  });

  pairs.push(['_', String(Date.now())]);
  const qs = pairs.map(([k, v]) => k + '=' + encodeURIComponent(v == null ? '' : String(v))).join('&');
  return endpoint + '?' + qs;
}

/** ---------- JSON helpers ---------- */

function posSyncGetByPath_(obj, path) {
  if (obj == null) return '';
  if (Object.prototype.hasOwnProperty.call(obj, path)) return posSyncStringify_(obj[path]);

  const parts = path.split('.');
  let cur = obj;
  for (let i = 0; i < parts.length; i++) {
    if (cur == null) return '';
    const key = parts[i];
    if (Object.prototype.hasOwnProperty.call(cur, key)) cur = cur[key];
    else return '';
  }
  return posSyncStringify_(cur);
}

function posSyncStringify_(v) {
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  return String(v);
}

function posSyncSha256Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const x = (b < 0) ? b + 256 : b;
    return ('0' + x.toString(16)).slice(-2);
  }).join('');
}