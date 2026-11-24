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

// ---- CASH LOCKER STATE CLOUD SYNC -----------------------------------------
// The Cash Locker (branch cash/denomination stock, soiled notes, transfers,
// vendor payments) used to live only in each device's own browser storage,
// so two devices/people never saw the same numbers. This stores the ENTIRE
// locker state as one JSON blob in a single row of its own sheet tab, keyed
// by a fixed id ("main") since there is only one shared locker state for
// the whole app (it already covers both branches internally). Whichever
// device last pushed a newer "updatedAt" wins — the app does the
// last-write-wins comparison on its side before overwriting its local copy,
// this script just stores/returns whatever it's given/asked for.
var LOCKER_STATE_SHEET_NAME = 'CashLockerState';
var LOCKER_STATE_HEADERS = ['id', 'updatedAt', 'state'];

function getOrCreateLockerStateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOCKER_STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOCKER_STATE_SHEET_NAME);
    sheet.appendRow(LOCKER_STATE_HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(LOCKER_STATE_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Upserts the single shared locker-state row, keyed on id (always "main").
 * The full state object is stored as a JSON string in the "state" column.
 */
function upsertLockerState_(payload) {
  var sheet = getOrCreateLockerStateSheet_();
  var data = sheet.getDataRange().getValues();
  var id = payload.id || 'main';
  var rowValues = [
    id,
    payload.updatedAt != null ? payload.updatedAt : Date.now(),
    JSON.stringify(payload.state || {})
  ];
  var matchRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
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
 * Returns the shared locker state row (id "main") as { id, updatedAt, state },
 * with "state" parsed back into an object — or null if nothing has been
 * pushed yet.
 */
function getLockerState_() {
  var sheet = getOrCreateLockerStateSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === 'main') {
      var parsedState = {};
      try { parsedState = JSON.parse(data[i][2]); } catch (e) {}
      return { id: data[i][0], updatedAt: Number(data[i][1]) || 0, state: parsedState };
    }
  }
  return null;
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
 * CASH ADDED — FINAL CONFIRM EMAIL
 * The app's "Final Confirm & Email" button (Cash Added section) posts
 * type: "cash_added_email" with the store, branch, date, shift, cashier,
 * confirmed denominations, total and the locker balance left. The
 * recipient is fixed HERE (not taken from the request) so this Web App
 * can never be used to send mail to anyone else. Each email is also
 * logged to a "CashAddedLog" sheet tab.
 */
var CASH_ADDED_NOTIFY_EMAIL = 'kanthirava.b@gmail.com';

function esc_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function inr_(n) {
  return '₹' + Number(n || 0).toLocaleString('en-IN');
}

function sendCashAddedEmail_(b) {
  try {
    var store = b.store || {};
    var rows = Array.isArray(b.denominations) ? b.denominations : [];
    var left = Array.isArray(b.lockerLeft) ? b.lockerLeft : [];
    var outside = Array.isArray(b.outsideLocker) ? b.outsideLocker : [];
    if (!rows.length && !(Array.isArray(b.changesSinceLast) && b.changesSinceLast.length)) return { ok: false, error: 'No denominations to send.' };

    var changes = Array.isArray(b.changesSinceLast) ? b.changesSinceLast : [];
    var subject = 'Cash Added confirmed' + (b.isResend ? ' (UPDATED)' : '') + ' - ' + (b.branch || '') + ' - ' + (b.date || '') + ' - ' + inr_(b.total);
    var changeLine_ = function (c) {
      return (c.at ? c.at + ' ' : '') + (c.action || '') + (c.otp ? ' (OTP)' : '') + ': ' + c.label + ' ' + c.fromQty + ' -> ' + c.toQty +
        ' (' + (c.amount >= 0 ? '+' : '-') + inr_(Math.abs(c.amount)) + ')';
    };

    var text = 'CASH ADDED - FINAL CONFIRMATION' + (b.isResend ? ' (UPDATED)' : '') + '\n\n' +
      'Store: ' + (store.name || '') + '\n' +
      (store.address ? 'Address: ' + store.address + '\n' : '') +
      (store.phone ? 'Phone: ' + store.phone + '\n' : '') +
      'Branch: ' + (b.branch || '') + '\n' +
      'Date: ' + (b.date || '') + (b.shift ? ' (' + b.shift + ')' : '') + '\n' +
      'Cashier: ' + (b.cashier || '') + '\n' +
      'Confirmed at: ' + (b.confirmedAt || '') + '\n\n' +
      'DENOMINATIONS\n' +
      rows.map(function (r) { return r.label + ' x ' + r.qty + ' = ' + inr_(r.amount); }).join('\n') +
      '\n\nTOTAL CASH ADDED: ' + inr_(b.total) + '\n';
    if (outside.length) {
      text += '\nNOT FROM LOCKER\n' + outside.map(function (o) {
        return o.denomSummary + ' = ' + inr_(o.amount) + ' | Reason: ' + o.reason + ' | From: ' + o.source;
      }).join('\n') + '\n';
    }
    if (changes.length) {
      text += '\nCHANGES SINCE PREVIOUS FINAL CONFIRM\n' + changes.map(changeLine_).join('\n') + '\n';
    }
    text += '\nLocker balance left (' + (b.branch || '') + '): ' + inr_(b.lockerLeftTotal) + '\n' +
      left.map(function (r) { return r.label + ' x ' + r.qty; }).join(', ') + '\n';

    var td = 'style="padding:6px 10px;border:1px solid #e2e8f0;"';
    var tdr = 'style="padding:6px 10px;border:1px solid #e2e8f0;text-align:right;"';
    var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;">' +
      '<h2 style="margin:0 0 4px;color:#065f46;">Cash Added - Final Confirmation' + (b.isResend ? ' (UPDATED)' : '') + '</h2>' +
      '<div style="color:#475569;margin-bottom:12px;">' + esc_(store.name) +
        (store.address ? ' &middot; ' + esc_(store.address) : '') +
        (store.phone ? ' &middot; ' + esc_(store.phone) : '') + '</div>' +
      '<table style="border-collapse:collapse;margin-bottom:14px;">' +
        '<tr><td ' + td + '><b>Branch</b></td><td ' + td + '>' + esc_(b.branch) + '</td></tr>' +
        '<tr><td ' + td + '><b>Date</b></td><td ' + td + '>' + esc_(b.date) + (b.shift ? ' (' + esc_(b.shift) + ')' : '') + '</td></tr>' +
        '<tr><td ' + td + '><b>Cashier</b></td><td ' + td + '>' + esc_(b.cashier) + '</td></tr>' +
        '<tr><td ' + td + '><b>Confirmed at</b></td><td ' + td + '>' + esc_(b.confirmedAt) + '</td></tr>' +
      '</table>' +
      '<table style="border-collapse:collapse;min-width:280px;">' +
        '<tr style="background:#0f172a;color:#fff;"><th ' + td + '>Denomination</th><th ' + tdr + '>Qty</th><th ' + tdr + '>Amount</th></tr>' +
        rows.map(function (r) {
          return '<tr><td ' + td + '>' + esc_(r.label) + '</td><td ' + tdr + '>' + esc_(r.qty) + '</td><td ' + tdr + '>' + inr_(r.amount) + '</td></tr>';
        }).join('') +
        '<tr style="background:#ecfdf5;font-weight:bold;"><td ' + td + ' colspan="2">TOTAL CASH ADDED</td><td ' + tdr + '>' + inr_(b.total) + '</td></tr>' +
      '</table>';
    if (outside.length) {
      html += '<p style="margin:14px 0 4px;color:#b45309;"><b>Not from locker:</b></p><ul style="margin:0;">' +
        outside.map(function (o) {
          return '<li>' + esc_(o.denomSummary) + ' = ' + inr_(o.amount) + ' &mdash; Reason: ' + esc_(o.reason) + ' &mdash; From: ' + esc_(o.source) + '</li>';
        }).join('') + '</ul>';
    }
    if (changes.length) {
      html += '<p style="margin:14px 0 4px;color:#b91c1c;"><b>Changes since previous Final Confirm:</b></p><ul style="margin:0;">' +
        changes.map(function (c) { return '<li>' + esc_(changeLine_(c)) + '</li>'; }).join('') + '</ul>';
    }
    html += '<p style="margin:14px 0 0;color:#475569;">Locker balance left (' + esc_(b.branch) + '): <b>' + inr_(b.lockerLeftTotal) + '</b><br>' +
      esc_(left.map(function (r) { return r.label + ' x ' + r.qty; }).join(', ')) + '</p></div>';

    MailApp.sendEmail({ to: CASH_ADDED_NOTIFY_EMAIL, subject: subject, body: text, htmlBody: html });
    var emailedTo = CASH_ADDED_NOTIFY_EMAIL;

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName('CashAddedLog') || ss.insertSheet('CashAddedLog');
      if (sh.getLastRow() === 0) sh.appendRow(['Sent At', 'Date', 'Branch', 'Shift', 'Cashier', 'Denominations', 'Total']);
      sh.appendRow([new Date(), b.date || '', b.branch || '', b.shift || '', b.cashier || '',
        rows.map(function (r) { return r.label + 'x' + r.qty; }).join(', '), Number(b.total) || 0]);
    } catch (logErr) {}

    return { ok: true, type: 'cash_added_email', emailed: true, to: emailedTo };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * CASH ADDED — PER-CHANGE EMAIL
 * Every confirmed Cash Added change (✓ add, lower + ✓, ↺ revert, Clear)
 * is emailed to CASH_ADDED_NOTIFY_EMAIL with what changed, the running
 * Cash Added breakdown and the locker balance left. Also logged to a
 * "CashAddedChanges" sheet tab.
 */
function sendCashAddedChangeEmail_(b) {
  try {
    var changes = Array.isArray(b.changes) ? b.changes : [];
    if (!changes.length) return { ok: false, error: 'No changes to send.' };
    var store = b.store || {};
    var rows = Array.isArray(b.breakdown) ? b.breakdown : [];
    var left = Array.isArray(b.lockerLeft) ? b.lockerLeft : [];
    var action = String(b.action || 'CHANGED').toUpperCase();
    var colour = action === 'ADDED' ? '#047857' : '#b91c1c';
    var net = changes.reduce(function (a, c) { return a + (Number(c.amount) || 0); }, 0);
    var changeTxt = changes.map(function (c) {
      return c.label + ': ' + c.fromQty + ' -> ' + c.toQty + ' (' + (c.deltaQty > 0 ? '+' : '') + c.deltaQty + ' = ' + (c.amount >= 0 ? '+' : '-') + inr_(Math.abs(c.amount)) + ')';
    });
    var subject = 'Cash Added ' + action + (b.otpVerified ? ' (OTP)' : '') + ' - ' + (b.branch || '') + ' - ' +
      (changes.length === 1 ? changes[0].label + ' x ' + Math.abs(changes[0].deltaQty) : changes.length + ' denominations') +
      ' (' + (net >= 0 ? '+' : '-') + inr_(Math.abs(net)) + ')';

    var text = 'CASH ADDED - ' + action + (b.otpVerified ? ' (approved with OTP)' : '') + '\n\n' +
      'Store: ' + (store.name || '') + '\nBranch: ' + (b.branch || '') + '\n' +
      'Date: ' + (b.date || '') + (b.shift ? ' (' + b.shift + ')' : '') + '\n' +
      'Cashier: ' + (b.cashier || '') + '\nTime: ' + (b.at || '') + '\n\n' +
      'CHANGE\n' + changeTxt.join('\n') + '\n\n' +
      'CASH ADDED NOW: ' + inr_(b.cashAddedTotal) + '\n' +
      (rows.length ? rows.map(function (r) { return r.label + ' x ' + r.qty + ' = ' + inr_(r.amount); }).join('\n') : '(none)') + '\n\n' +
      'Locker balance left (' + (b.branch || '') + '): ' + inr_(b.lockerLeftTotal) + '\n' +
      left.map(function (r) { return r.label + ' x ' + r.qty; }).join(', ') + '\n' +
      (b.finalPending ? '\nNOTE: Final Confirm has NOT been pressed for these figures yet.\n' : '');

    var td = 'style="padding:6px 10px;border:1px solid #e2e8f0;"';
    var tdr = 'style="padding:6px 10px;border:1px solid #e2e8f0;text-align:right;"';
    var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;">' +
      '<h2 style="margin:0 0 4px;color:' + colour + ';">Cash Added - ' + esc_(action) + (b.otpVerified ? ' <span style="font-size:13px;color:#475569;">(approved with OTP)</span>' : '') + '</h2>' +
      '<div style="color:#475569;margin-bottom:12px;">' + esc_(store.name) + ' &middot; <b>' + esc_(b.branch) + '</b> &middot; ' + esc_(b.date) + (b.shift ? ' (' + esc_(b.shift) + ')' : '') +
        ' &middot; Cashier: ' + esc_(b.cashier) + ' &middot; ' + esc_(b.at) + '</div>' +
      '<table style="border-collapse:collapse;margin-bottom:14px;">' +
        '<tr style="background:#0f172a;color:#fff;"><th ' + td + '>Denomination</th><th ' + tdr + '>Before</th><th ' + tdr + '>After</th><th ' + tdr + '>Change</th><th ' + tdr + '>Amount</th></tr>' +
        changes.map(function (c) {
          return '<tr><td ' + td + '>' + esc_(c.label) + '</td><td ' + tdr + '>' + esc_(c.fromQty) + '</td><td ' + tdr + '>' + esc_(c.toQty) + '</td><td ' + tdr + ' style="color:' + (c.deltaQty > 0 ? '#047857' : '#b91c1c') + ';text-align:right;padding:6px 10px;border:1px solid #e2e8f0;"><b>' + (c.deltaQty > 0 ? '+' : '') + esc_(c.deltaQty) + '</b></td><td ' + tdr + '>' + (c.amount >= 0 ? '+' : '-') + inr_(Math.abs(c.amount)) + '</td></tr>';
        }).join('') +
      '</table>' +
      '<p style="margin:0 0 4px;"><b>Cash Added now: ' + inr_(b.cashAddedTotal) + '</b></p>' +
      '<p style="margin:0 0 12px;color:#475569;">' + (rows.length ? esc_(rows.map(function (r) { return r.label + ' x ' + r.qty; }).join(', ')) : '(none)') + '</p>' +
      '<p style="margin:0;color:#475569;">Locker balance left (' + esc_(b.branch) + '): <b>' + inr_(b.lockerLeftTotal) + '</b><br>' +
        esc_(left.map(function (r) { return r.label + ' x ' + r.qty; }).join(', ')) + '</p>' +
      (b.finalPending ? '<p style="margin:12px 0 0;color:#b45309;"><b>Final Confirm has not been pressed for these figures yet.</b></p>' : '') +
      '</div>';

    MailApp.sendEmail({ to: CASH_ADDED_NOTIFY_EMAIL, subject: subject, body: text, htmlBody: html });

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName('CashAddedChanges') || ss.insertSheet('CashAddedChanges');
      if (sh.getLastRow() === 0) sh.appendRow(['Sent At', 'Date', 'Branch', 'Shift', 'Cashier', 'Action', 'OTP', 'Changes', 'Net Amount', 'Cash Added Total']);
      sh.appendRow([new Date(), b.date || '', b.branch || '', b.shift || '', b.cashier || '', action, b.otpVerified ? 'Yes' : '',
        changeTxt.join('; '), net, Number(b.cashAddedTotal) || 0]);
    } catch (logErr) {}

    return { ok: true, type: 'cash_added_change_email', to: CASH_ADDED_NOTIFY_EMAIL };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * CASH DENOMINATIONS — BALANCE REPORT (PDF)
 * The "Email PDF Report" button on the Cash Denominations tab posts the
 * balance pieces/values per store. This builds an A4 PDF from it and
 * emails it (as an attachment) to CASH_ADDED_NOTIFY_EMAIL.
 */
function sendDenominationReport_(b) {
  try {
    var branches = Array.isArray(b.branches) ? b.branches : [];
    var totals = b.totals || { rows: [], totalQty: 0, totalAmount: 0, damaged: 0 };
    if (!branches.length) return { ok: false, error: 'No store data to report.' };
    var store = (b.store && b.store.name) || 'Lalitha Naturals';
    var th = 'style="padding:6px 8px;border:1px solid #cbd5e1;background:#0f172a;color:#fff;font-size:11px;"';
    var td = 'style="padding:5px 8px;border:1px solid #cbd5e1;font-size:11px;"';
    var tdr = 'style="padding:5px 8px;border:1px solid #cbd5e1;font-size:11px;text-align:right;"';
    var tf = 'style="padding:6px 8px;border:1px solid #cbd5e1;font-size:11px;text-align:right;font-weight:bold;background:#ecfdf5;"';

    var head = '<tr><th ' + th + '>Denomination</th>' + branches.map(function (br) {
      return '<th ' + th + '>' + esc_(br.name) + ' (Qty)</th><th ' + th + '>' + esc_(br.name) + ' (Value)</th>';
    }).join('') + '<th ' + th + '>Total Qty</th><th ' + th + '>Total Value</th></tr>';
    var body = (totals.rows || []).map(function (tr, i) {
      return '<tr><td ' + td + '><b>' + esc_(tr.label) + '</b></td>' + branches.map(function (br) {
        var r = (br.rows || [])[i] || { qty: 0, amount: 0 };
        return '<td ' + tdr + '>' + esc_(r.qty) + ' pcs</td><td ' + tdr + '>' + inr_(r.amount) + '</td>';
      }).join('') + '<td ' + tdr + '><b>' + esc_(tr.qty) + ' pcs</b></td><td ' + tdr + '><b>' + inr_(tr.amount) + '</b></td></tr>';
    }).join('');
    var tfl = 'style="padding:6px 8px;border:1px solid #cbd5e1;font-size:11px;text-align:left;font-weight:bold;background:#ecfdf5;"';
    var foot = '<tr><td ' + tfl + '>TOTAL CASH</td>' + branches.map(function (br) {
      return '<td ' + tf + '>' + esc_(br.totalQty) + ' pcs</td><td ' + tf + '>' + inr_(br.totalAmount) + '</td>';
    }).join('') + '<td ' + tf + '>' + esc_(totals.totalQty) + ' pcs</td><td ' + tf + '>' + inr_(totals.totalAmount) + '</td></tr>';

    var summary = branches.map(function (br) {
      return '<td style="padding:8px 14px;border:1px solid #cbd5e1;"><div style="font-size:11px;color:#475569;">' + esc_(br.name) + '</div>' +
        '<div style="font-size:18px;font-weight:bold;">' + inr_(br.totalAmount) + '</div>' +
        '<div style="font-size:10px;color:#b45309;">Damaged held: ' + inr_(br.damaged) + '</div></td>';
    }).join('') + '<td style="padding:8px 14px;border:1px solid #cbd5e1;background:#ecfdf5;"><div style="font-size:11px;color:#475569;">All stores</div>' +
      '<div style="font-size:18px;font-weight:bold;color:#065f46;">' + inr_(totals.totalAmount) + '</div>' +
      '<div style="font-size:10px;color:#b45309;">Damaged held: ' + inr_(totals.damaged) + '</div></td>';

    var html = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;color:#0f172a;margin:24px;">' +
      '<div style="font-size:20px;font-weight:bold;color:#065f46;">' + esc_(store) + '</div>' +
      '<div style="font-size:14px;font-weight:bold;margin:2px 0 4px;">Cash Denominations - Balance Report</div>' +
      '<div style="font-size:11px;color:#475569;margin-bottom:14px;">Generated: ' + esc_(b.generatedAt) + (b.generatedBy ? ' &middot; By: ' + esc_(b.generatedBy) : '') + '</div>' +
      '<table style="border-collapse:collapse;margin-bottom:16px;"><tr>' + summary + '</tr></table>' +
      '<table style="border-collapse:collapse;width:100%;">' + head + body + foot + '</table>' +
      '<div style="font-size:10px;color:#64748b;margin-top:12px;">Usable cash only. Damaged notes held in the vault are shown separately above.</div>' +
      '</body></html>';

    var datePart = String(b.generatedAt || '').split(' ')[0].replace(/\//g, '-');
    var fileName = 'Lalitha_Naturals_Denomination_Balance_' + (datePart || 'report') + '.pdf';
    var pdf = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf').setName(fileName);

    var text = 'Cash Denominations - Balance Report (' + (b.generatedAt || '') + ')\n\n' +
      branches.map(function (br) { return br.name + ': ' + inr_(br.totalAmount) + ' (damaged held ' + inr_(br.damaged) + ')'; }).join('\n') +
      '\nTotal: ' + inr_(totals.totalAmount) + '\n\nThe full denomination breakdown is in the attached PDF.';
    var htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;"><p><b>Cash Denominations - Balance Report</b><br>' + esc_(b.generatedAt) + '</p>' +
      '<p>' + branches.map(function (br) { return esc_(br.name) + ': <b>' + inr_(br.totalAmount) + '</b>'; }).join('<br>') +
      '<br>Total: <b>' + inr_(totals.totalAmount) + '</b></p><p>The full denomination breakdown is in the attached PDF.</p></div>';

    MailApp.sendEmail({
      to: CASH_ADDED_NOTIFY_EMAIL,
      subject: 'Cash Denominations Balance Report - ' + (b.generatedAt || '') + ' - ' + inr_(totals.totalAmount),
      body: text, htmlBody: htmlBody, attachments: [pdf]
    });
    return { ok: true, type: 'denomination_report_email', to: CASH_ADDED_NOTIFY_EMAIL };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * CASH ADDED — REVERT OTP
 * After Cash Added has been Final Confirmed (emailed), reverting or
 * lowering any denomination needs a one-time 6-digit code emailed to
 * CASH_ADDED_NOTIFY_EMAIL (fixed here, never taken from the request).
 * Stored separately from the Owner PIN OTP (REVERT_OTP_* properties),
 * 10-minute expiry, one-time use, max 5 wrong attempts.
 */
function sendRevertOtp_(b) {
  try {
    var code = generateOtpCode_();
    PropertiesService.getScriptProperties().setProperties({
      REVERT_OTP_CODE: code,
      REVERT_OTP_EXPIRES: String(Date.now() + PIN_OTP_TTL_MS),
      REVERT_OTP_ATTEMPTS: '0'
    });
    var subject = 'OTP ' + code + ' - Cash Added revert - ' + (b.branch || '');
    var text = 'Your one-time code to approve a Cash Added revert is:\n\n' + code + '\n\n' +
      'Branch: ' + (b.branch || '') + '\n' +
      'Date: ' + (b.date || '') + (b.shift ? ' (' + b.shift + ')' : '') + '\n' +
      'Requested by: ' + (b.cashier || '') + '\n' +
      'Change requested: ' + (b.action || '') + '\n\n' +
      'The code expires in 10 minutes and works only once. If you did not expect this, do NOT share the code.';
    var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;">' +
      '<p>Your one-time code to approve a <b>Cash Added revert</b>:</p>' +
      '<div style="font-size:30px;font-weight:bold;letter-spacing:6px;color:#b91c1c;margin:8px 0 14px;">' + code + '</div>' +
      '<table style="border-collapse:collapse;">' +
      '<tr><td style="padding:4px 10px 4px 0;color:#475569;">Branch</td><td><b>' + esc_(b.branch) + '</b></td></tr>' +
      '<tr><td style="padding:4px 10px 4px 0;color:#475569;">Date</td><td>' + esc_(b.date) + (b.shift ? ' (' + esc_(b.shift) + ')' : '') + '</td></tr>' +
      '<tr><td style="padding:4px 10px 4px 0;color:#475569;">Requested by</td><td>' + esc_(b.cashier) + '</td></tr>' +
      '<tr><td style="padding:4px 10px 4px 0;color:#475569;">Change requested</td><td><b>' + esc_(b.action) + '</b></td></tr>' +
      '</table><p style="color:#64748b;">Expires in 10 minutes, works only once. If you did not expect this, do not share the code.</p></div>';
    MailApp.sendEmail({ to: CASH_ADDED_NOTIFY_EMAIL, subject: subject, body: text, htmlBody: html });
    return { ok: true, type: 'send_revert_otp', to: CASH_ADDED_NOTIFY_EMAIL };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function verifyRevertOtp_(b) {
  var submitted = b && String(b.code || '').trim();
  var props = PropertiesService.getScriptProperties();
  var storedCode = props.getProperty('REVERT_OTP_CODE');
  var storedExpires = Number(props.getProperty('REVERT_OTP_EXPIRES') || 0);
  var attempts = Number(props.getProperty('REVERT_OTP_ATTEMPTS') || 0);
  function clear_() {
    props.deleteProperty('REVERT_OTP_CODE');
    props.deleteProperty('REVERT_OTP_EXPIRES');
    props.deleteProperty('REVERT_OTP_ATTEMPTS');
  }
  if (!storedCode || !submitted) return { ok: false, type: 'verify_revert_otp', error: 'invalid_or_expired' };
  if (Date.now() > storedExpires || attempts >= PIN_OTP_MAX_ATTEMPTS) {
    clear_();
    return { ok: false, type: 'verify_revert_otp', error: 'invalid_or_expired' };
  }
  if (submitted !== storedCode) {
    props.setProperty('REVERT_OTP_ATTEMPTS', String(attempts + 1));
    return { ok: false, type: 'verify_revert_otp', error: 'wrong_code', attemptsLeft: PIN_OTP_MAX_ATTEMPTS - attempts - 1 };
  }
  clear_();
  return { ok: true, type: 'verify_revert_otp' };
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

    // Cash Locker's entire shared state (branch stock, soiled notes,
    // transactions, transfers, vendor payments) — pushed as one JSON blob
    // any time it changes, so another device/person sees the same numbers.
    if (body && body.type === 'cash_locker') {
      upsertLockerState_(body);
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, type: 'cash_locker' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Health check — lets the app see which features this deployment has.
    if (body && body.type === 'ping') {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, type: 'ping', features: ['cash_added_email', 'revert_otp', 'cash_added_change_email', 'denomination_report_email'] })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Cash Added — revert after Final Confirm: send / verify an email OTP.
    if (body && body.type === 'send_revert_otp') {
      return ContentService.createTextOutput(
        JSON.stringify(sendRevertOtp_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }
    if (body && body.type === 'verify_revert_otp') {
      return ContentService.createTextOutput(
        JSON.stringify(verifyRevertOtp_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Cash Added — one email per confirmed change (added / reduced /
    // reverted / cleared), sent the moment the change is confirmed.
    if (body && body.type === 'cash_added_change_email') {
      return ContentService.createTextOutput(
        JSON.stringify(sendCashAddedChangeEmail_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Cash Denominations — email the balance report as a PDF.
    if (body && body.type === 'denomination_report_email') {
      return ContentService.createTextOutput(
        JSON.stringify(sendDenominationReport_(body))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Cash Added — Final Confirm: email the confirmed Cash Added summary.
    if (body && body.type === 'cash_added_email') {
      return ContentService.createTextOutput(
        JSON.stringify(sendCashAddedEmail_(body))
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

    if (body && body.type && body.type !== 'shift') {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, error: 'Unknown request type: ' + body.type })
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

    // ?type=cash_locker — returns the single shared Cash Locker state blob
    // (branch stock, soiled notes, transactions, transfers, vendor
    // payments) so every device/person works off the same numbers.
    if (type === 'cash_locker') {
      var lockerRow = getLockerState_();
      return ContentService.createTextOutput(
        JSON.stringify(lockerRow ? { ok: true, found: true, updatedAt: lockerRow.updatedAt, state: lockerRow.state }
                                  : { ok: true, found: false })
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


/**
 * ONE-TIME SETUP: select "testEmailSetup" in the function dropdown above and
 * click Run. Google asks you to allow "Send email as you" — click Allow.
 * Sends a test email to CASH_ADDED_NOTIFY_EMAIL so you can see it works.
 */
function testEmailSetup() {
  MailApp.sendEmail(CASH_ADDED_NOTIFY_EMAIL,
    'Lalitha Naturals - email test',
    'Email sending from the Shift Register is working. Cash Added confirmations and revert OTP codes will arrive at this address.');
  Logger.log('Test email sent to ' + CASH_ADDED_NOTIFY_EMAIL + '. Emails left today: ' + MailApp.getRemainingDailyQuota());
}
