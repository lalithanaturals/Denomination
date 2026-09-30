/**
 * Lalitha Naturals — Shift Register — Google Sheets Cloud Sync
 * =============================================================
 *
 * WHAT THIS IS
 * A tiny free "Google Apps Script Web App" that receives shift records
 * posted by the Shift Register HTML app and appends/updates them as rows
 * in a Google Sheet named "ShiftRegister" (created automatically the first
 * time it runs). This is what makes your shift data reachable from any
 * device — the Sheet lives in your own Google account.
 *
 * ONE-TIME SETUP (do this once; takes about 3 minutes):
 *   1. Go to https://sheets.google.com and create a new, blank Google Sheet
 *      (or open an existing one you want to use for this).
 *   2. In the menu, click Extensions > Apps Script. A code editor opens in
 *      a new tab, with a default empty "Code.gs" file.
 *   3. Select all the placeholder code in that editor and delete it. Paste
 *      in the ENTIRE contents of this file instead.
 *   4. Click the disk/Save icon (or Ctrl/Cmd+S).
 *   5. Click "Deploy" (top-right) > "New deployment".
 *   6. Next to "Select type", click the gear icon and choose "Web app".
 *   7. Fill in:
 *        Description:      Shift Register Sync (or anything you like)
 *        Execute as:        Me (your own Google account)
 *        Who has access:    Anyone with the link
 *      (These exact choices matter — "Execute as: Me" lets the script write
 *      to your Sheet on the app's behalf even though the app itself never
 *      logs into Google, and "Anyone with the link" lets the offline app
 *      call the URL without you having to log in from inside it.)
 *   8. Click "Deploy". The first time, Google will ask you to authorize the
 *      script — click through "Review permissions" > pick your account >
 *      "Advanced" > "Go to (project name) (unsafe)" > "Allow". This warning
 *      is expected and normal for a script you wrote/pasted yourself.
 *   9. Copy the "Web app URL" it gives you (ends in /exec). That is the URL
 *      you paste into the Shift Register app's Settings > Cloud Sync panel.
 *
 * If you ever edit this script again, you must create a NEW deployment (or
 * use "Manage deployments" > edit > "New version") for the changes to take
 * effect on the same URL.
 *
 * WHAT IT STORES
 * Every time a shift is saved in the app (with Cloud Sync turned on), it
 * POSTs the shift's data as JSON here. This script appends one row per
 * shift into a sheet tab named "ShiftRegister", creating that tab and its
 * header row automatically if they don't exist yet. If a row already
 * exists for the same Date + Branch + Shift, it's updated in place instead
 * of duplicated (an "upsert").
 *
 * The display columns cover the most useful fields at a glance. A final
 * "Full Record (JSON)" column also stores the ENTIRE record exactly as
 * received, so nothing is ever lost even if a future version of the app
 * adds fields this script doesn't explicitly list yet.
 *
 * This script also creates two more tabs automatically, the first time each
 * kind of item is synced:
 *   - "ItemCatalog": the Cash Bill's barcode/item catalog (name, MRP, price).
 *   - "DamageExpiryLedger": the Damage/Expiry write-off log (date, item,
 *     qty, reason, loss amount, branch) — a simple record, push-only from
 *     the app, upserted by its own id so re-syncs never duplicate a row.
 *
 * OWNER PIN CHANGE — EMAIL OTP VERIFICATION
 * Changing the app's Owner PIN (Settings > Owner Access) now requires a
 * one-time 6-digit code emailed to the "Owner Email" configured in that
 * same Settings panel, so that simply knowing the current PIN is no longer
 * enough to change it to something else. This script is what sends that
 * email — via MailApp.sendEmail(), which sends as your own Google account
 * for free within Google's normal daily quota — since a static HTML app has
 * no other way to send mail. Two new request types on doPost:
 *   - type: "send_pin_otp"   generates a random 6-digit code, stores it in
 *     Script Properties (PIN_OTP_CODE / PIN_OTP_EXPIRES / PIN_OTP_EMAIL /
 *     PIN_OTP_ATTEMPTS) with a 10-minute expiry, and emails it to the
 *     address the app sends in the request. The code itself is never
 *     echoed back in the HTTP response — only the email carries it.
 *   - type: "verify_pin_otp" checks a submitted code against the stored
 *     one (must match, and not be expired). A correct code is deleted
 *     immediately after use (one-time use). Up to 5 wrong attempts are
 *     allowed before the stored code is invalidated, as a simple brute-
 *     force guard (Apps Script has no built-in rate limiting).
 * Script Properties (not the Sheet) are used because they persist across
 * every future request to this deployment, executing "as Me", which is
 * exactly the shared server-side state a short-lived OTP needs.
 */

var SHEET_NAME = 'ShiftRegister';
var CATALOG_SHEET_NAME = 'ItemCatalog';
var CATALOG_HEADERS = ['barcode', 'name', 'mrp', 'sellingPrice', 'discount', 'discountMode', 'updatedAt'];

function getOrCreateCatalogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CATALOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CATALOG_SHEET_NAME);
    sheet.appendRow(CATALOG_HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(CATALOG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Upserts one catalog item row, keyed on barcode.
 */
function upsertCatalogItem_(item) {
  var sheet = getOrCreateCatalogSheet_();
  var data = sheet.getDataRange().getValues();
  var rowValues = [
    item.barcode || '',
    item.name || '',
    item.mrp != null ? item.mrp : '',
    item.sellingPrice != null ? item.sellingPrice : '',
    item.discount != null ? item.discount : '',
    item.discountMode || '',
    item.updatedAt != null ? item.updatedAt : ''
  ];
  var matchRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(item.barcode)) {
      matchRow = i + 1;
      break;
    }
  }
  if (matchRow > -1) {
    sheet.getRange(matchRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

/**
 * Returns all ItemCatalog rows as an array of plain objects.
 */
function getCatalogRecords_() {
  var sheet = getOrCreateCatalogSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
    if (obj.barcode) rows.push(obj);
  }
  return rows;
}

var DAMAGE_SHEET_NAME = 'DamageExpiryLedger';
var DAMAGE_HEADERS = ['id', 'date', 'item', 'qty', 'reason', 'note', 'lossAmount', 'branch', 'updatedAt'];

function getOrCreateDamageSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DAMAGE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DAMAGE_SHEET_NAME);
    sheet.appendRow(DAMAGE_HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(DAMAGE_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Upserts one damage/expiry ledger row, keyed on its unique id
 * (last-write-wins by updatedAt, same pattern as the item catalog).
 */
function upsertDamageEntry_(entry) {
  var sheet = getOrCreateDamageSheet_();
  var data = sheet.getDataRange().getValues();
  var rowValues = [
    entry.id || '',
    entry.date || '',
    entry.item || '',
    entry.qty != null ? entry.qty : '',
    entry.reason || '',
    entry.note || '',
    entry.lossAmount != null ? entry.lossAmount : '',
    entry.branch || '',
    entry.updatedAt != null ? entry.updatedAt : ''
  ];
  var matchRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entry.id)) {
      matchRow = i + 1;
      break;
    }
  }
  if (matchRow > -1) {
    sheet.getRange(matchRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

/**
 * Returns all DamageExpiryLedger rows as plain objects — supports the
 * optional ?type=damage pull-merge path on doGet for symmetry with the
 * catalog sync, even though the app itself currently only pushes entries
 * (a logged write-off rarely needs to be pulled back and re-merged the way
 * a shared product catalog does, but the endpoint supports it either way).
 */
function getDamageRecords_() {
  var sheet = getOrCreateDamageSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
    if (obj.id) rows.push(obj);
  }
  return rows;
}

// ---- OWNER PIN CHANGE: EMAIL OTP VERIFICATION -----------------------------
var PIN_OTP_TTL_MS = 10 * 60 * 1000; // 10 minutes
var PIN_OTP_MAX_ATTEMPTS = 5;

function generateOtpCode_() {
  var n = Math.floor(Math.random() * 1000000);
  var s = String(n);
  while (s.length < 6) s = '0' + s;
  return s;
}

/**
 * Generates a fresh 6-digit code, stores it (with expiry + target email +
 * a reset attempt counter) in Script Properties, and emails it to toEmail.
 * Responds { ok: true } on success, or { ok: false, error: ... } if the
 * email address is missing/invalid or MailApp fails (e.g. daily quota).
 */
function sendPinOtp_(body) {
  var toEmail = body && body.email;
  if (!toEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(toEmail)) {
    return { ok: false, error: 'A valid Owner Email is required to send a verification code.' };
  }
  try {
    var code = generateOtpCode_();
    var props = PropertiesService.getScriptProperties();
    props.setProperties({
      PIN_OTP_CODE: code,
      PIN_OTP_EXPIRES: String(Date.now() + PIN_OTP_TTL_MS),
      PIN_OTP_EMAIL: toEmail,
      PIN_OTP_ATTEMPTS: '0'
    });
    var subject = 'Your Owner PIN change code — Lalitha Naturals Shift Register';
    var body_ = 'Your one-time verification code to change the Owner PIN is:\n\n' +
      code + '\n\n' +
      'This code expires in 10 minutes and can only be used once. ' +
      'If you did not request this, you can ignore this email — your PIN will not change without this code.';
    MailApp.sendEmail(toEmail, subject, body_);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * Verifies a submitted OTP code against the stored one. On a correct,
 * unexpired code: deletes the stored properties (one-time use) and returns
 * { ok: true }. On a wrong or expired code: increments the attempt counter
 * (invalidating the code entirely after PIN_OTP_MAX_ATTEMPTS) and returns
 * { ok: false, error: 'invalid_or_expired' } WITHOUT deleting a still-valid
 * stored code, so a mistyped attempt doesn't burn the real code.
 */
function verifyPinOtp_(body) {
  var submitted = body && String(body.code || '').trim();
  var props = PropertiesService.getScriptProperties();
  var storedCode = props.getProperty('PIN_OTP_CODE');
  var storedExpires = Number(props.getProperty('PIN_OTP_EXPIRES') || 0);
  var attempts = Number(props.getProperty('PIN_OTP_ATTEMPTS') || 0);

  if (!storedCode || !submitted) {
    return { ok: false, error: 'invalid_or_expired' };
  }
  if (Date.now() > storedExpires || attempts >= PIN_OTP_MAX_ATTEMPTS) {
    props.deleteProperty('PIN_OTP_CODE');
    props.deleteProperty('PIN_OTP_EXPIRES');
    props.deleteProperty('PIN_OTP_EMAIL');
    props.deleteProperty('PIN_OTP_ATTEMPTS');
    return { ok: false, error: 'invalid_or_expired' };
  }
  if (submitted !== storedCode) {
    props.setProperty('PIN_OTP_ATTEMPTS', String(attempts + 1));
    return { ok: false, error: 'invalid_or_expired' };
  }

  // Correct code — one-time use, delete immediately so it can't be replayed.
  props.deleteProperty('PIN_OTP_CODE');
  props.deleteProperty('PIN_OTP_EXPIRES');
  props.deleteProperty('PIN_OTP_EMAIL');
  props.deleteProperty('PIN_OTP_ATTEMPTS');
  return { ok: true };
}

var HEADERS = [
  'Date', 'Branch', 'Shift', 'Cashier',
  'Grand Total', 'Tray Balance (Closing Bal)', 'Total Outflows',
  'Software Sale', 'Variance', 'Opening Bal', 'Closing Bal',
  'Extra Change Given', 'Change Left by Customer', 'Staff Bills Total',
  'Notes', 'Synced At', 'Full Record (JSON)'
];

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  // The "Date" column (column 1) holds strings like "2026-09-29". Google
  // Sheets' default "Automatic" number format silently reinterprets any
  // string that LOOKS like a date as an actual Date value the moment it's
  // written — even when written via the API, not just typed in the UI. Once
  // that happens, getValues() hands back a real JS Date object for that
  // cell instead of the original string, which breaks the exact-string
  // match this script's own upsert logic (recordToRow_/doPost below) relies
  // on to find an existing row for the same Date+Branch+Shift, so instead of
  // updating that row it appends a duplicate — and a pulling device's merge
  // (which also matches by that same Date string) can then quietly disagree
  // with what's actually stored. Forcing this column to Plain Text ("@")
  // stops Sheets from ever auto-converting it going forward. This is
  // idempotent and cheap, so it's safe to call on every request rather than
  // only right after creating the sheet (an already-existing sheet from
  // before this fix still has "Automatic" format on this column otherwise).
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), 1).setNumberFormat('@');
  return sheet;
}

// Normalizes a "Date" cell's value back into a plain "yyyy-MM-dd" string,
// whether it's already a string (the intended, post-fix case) or a real JS
// Date object (rows written before the Plain Text fix above, or any other
// cell Sheets auto-converted regardless). Used on every read/compare of
// that column so already-corrupted existing rows keep matching/merging
// correctly instead of silently diverging forever.
function normalizeDateValue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v);
}

function sumField_(list, key) {
  var total = 0;
  (list || []).forEach(function (item) {
    total += Number(item && item[key]) || 0;
  });
  return total;
}

function recordToRow_(record) {
  // Staff Bills (items a staff member takes now and pays for later, by cash
  // or salary deduction) are folded into Total Outflows the same way Owner
  // Bills and Vendor Payments already are, matching how the app's own
  // Grand Total / "Total" formula treats them (see the app's Settings >
  // Formulas & Advanced panel).
  var staffBillsTotal = sumField_(record.staffBills, 'amt');
  var totalOutflows =
    sumField_(record.expenses, 'amt') +
    sumField_(record.owners, 'amt') +
    sumField_(record.vendors, 'amt') +
    staffBillsTotal;

  return [
    record.date || '',
    record.branch || '',
    record.shift || 'Full Day',
    record.cashier || '',
    record.grandTotal != null ? record.grandTotal : '',
    record.closingBal != null ? record.closingBal : '', // "Tray Balance" note field
    totalOutflows,
    record.vasySale != null ? record.vasySale : '',
    record.variance != null ? record.variance : '',
    record.openingBal != null ? record.openingBal : '',
    record.closingBal != null ? record.closingBal : '',
    record.extraChangeGiven != null ? record.extraChangeGiven : '',
    record.changeLeftByCustomer != null ? record.changeLeftByCustomer : '',
    staffBillsTotal,
    record.notes || '',
    new Date().toISOString(),
    JSON.stringify(record)
  ];
}

/**
 * Receives a POST from the Shift Register app. Body is JSON for a single
 * shift record. Upserts a row keyed on Date + Branch + Shift.
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // A catalog-item post is additive, new behavior — the existing shift-
    // record behavior below (no "type", or type: "shift") is unchanged.
    if (body && body.type === 'catalog_item') {
      upsertCatalogItem_(body);
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, type: 'catalog_item' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // A damage/expiry ledger entry — also additive, separate sheet tab.
    if (body && body.type === 'damage_entry') {
      upsertDamageEntry_(body);
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, type: 'damage_entry' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Owner PIN change — step 1: send a 6-digit email OTP.
    if (body && body.type === 'send_pin_otp') {
      return ContentService.createTextOutput(
        JSON.stringify(sendPinOtp_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Owner PIN change — step 2: verify the submitted OTP.
    if (body && body.type === 'verify_pin_otp') {
      return ContentService.createTextOutput(
        JSON.stringify(verifyPinOtp_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getOrCreateSheet_();

    // Accept either a single record, or a { records: [...] } batch.
    var records = Array.isArray(body) ? body
      : (body.records ? body.records : [body]);

    var data = sheet.getDataRange().getValues();
    records.forEach(function (record) {
      var rowValues = recordToRow_(record);
      var matchRow = -1;
      for (var i = 1; i < data.length; i++) {
        if (normalizeDateValue_(data[i][0]) === record.date && data[i][1] === record.branch &&
            (data[i][2] || 'Full Day') === (record.shift || 'Full Day')) {
          matchRow = i + 1; // 1-indexed sheet row
          break;
        }
      }
      if (matchRow > -1) {
        sheet.getRange(matchRow, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
        data.push(rowValues); // keep local copy in sync for subsequent records in this batch
      }
    });

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, count: records.length })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Health check / fetch-existing-records. Called with a plain GET (e.g. the
 * app's "Test Connection" button, or a future "Restore from Cloud" feature).
 * Optional query param ?since=YYYY-MM-DD filters to rows on/after that date.
 */
function doGet(e) {
  try {
    var type = e && e.parameter && e.parameter.type;

    // New, additive behavior: ?type=catalog returns the ItemCatalog sheet
    // instead of shift records. Existing behavior (no type, or type=shifts)
    // is unchanged below.
    if (type === 'catalog') {
      var catalogRecords = getCatalogRecords_();
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, count: catalogRecords.length, records: catalogRecords })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ?type=damage — symmetry with the catalog pull, for a future/manual
    // pull-merge of the damage/expiry ledger. The app itself only pushes
    // today (see notes above getDamageRecords_).
    if (type === 'damage') {
      var damageRecords = getDamageRecords_();
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, count: damageRecords.length, records: damageRecords })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getOrCreateSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var since = e && e.parameter && e.parameter.since;

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowDate = normalizeDateValue_(data[i][0]);
      if (since && rowDate < since) continue;
      var obj = {};
      headers.forEach(function (h, idx) {
        obj[h] = (h === 'Date') ? rowDate : data[i][idx];
      });
      rows.push(obj);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, count: rows.length, records: rows })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
