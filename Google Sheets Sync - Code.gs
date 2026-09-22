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

var HEADERS = [
  'Date', 'Branch', 'Shift', 'Cashier',
  'Grand Total', 'Tray Balance (Closing Bal)', 'Total Outflows',
  'Software Sale', 'Variance', 'Opening Bal', 'Closing Bal',
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
  return sheet;
}

function sumField_(list, key) {
  var total = 0;
  (list || []).forEach(function (item) {
    total += Number(item && item[key]) || 0;
  });
  return total;
}

function recordToRow_(record) {
  var totalOutflows =
    sumField_(record.expenses, 'amt') +
    sumField_(record.owners, 'amt') +
    sumField_(record.vendors, 'amt');

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

    var sheet = getOrCreateSheet_();

    // Accept either a single record, or a { records: [...] } batch.
    var records = Array.isArray(body) ? body
      : (body.records ? body.records : [body]);

    var data = sheet.getDataRange().getValues();
    records.forEach(function (record) {
      var rowValues = recordToRow_(record);
      var matchRow = -1;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === record.date && data[i][1] === record.branch &&
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

    var sheet = getOrCreateSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var since = e && e.parameter && e.parameter.since;

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (since && String(rowDate) < since) continue;
      var obj = {};
      headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
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
