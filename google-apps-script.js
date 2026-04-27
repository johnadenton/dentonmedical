// dentonmedical - Google Apps Script v3
// Conflict-free multi-device sync
//
// SETUP:
// 1. sheets.google.com - new blank spreadsheet - name it: dentonmedical
// 2. Extensions > Apps Script > delete all code > paste this file
// 3. Save (floppy disk icon)
// 4. Deploy > New deployment > gear icon > Web app
// 5. Execute as: Me  |  Who has access: Anyone
// 6. Deploy > Authorize > copy the Web App URL
// 7. Farm app > Settings > Google Sheets Sync > paste URL > Save & Connect
// 8. Tap Sync - sheets populate automatically
//
// HOW IT WORKS (no conflicts):
// - Each record has a unique ID. Records are NEVER bulk-overwritten.
// - Every save = upsert by ID (insert if new, update if exists).
// - Every delete = remove that one row by ID only.
// - ScriptLock prevents two devices writing at the same instant.
// - GET returns all records so both devices always read the full truth.

var SHEET_EXPENSES = 'Expenses';
var SHEET_SALES    = 'Sales';
var SHEET_MILEAGE  = 'Mileage';
var SHEET_VENDORS  = 'Vendors';
var SHEET_SETTINGS = 'AppSettings';
var DRIVE_FOLDER   = 'dentonmedical';
var DRIVE_SUBFOLDER = 'receipts';

var EXP_HEADERS  = ['Record ID','Date','Category','Description','Vendor','Miles',
                    'Rate Per Mile','Mileage Amount','Amount','Receipt URL','Notes','Source','Device','Synced At'];
var SALE_HEADERS = ['Record ID','Date','Buyer','Item','Unit','Qty',
                    'Price Per Unit','Subtotal','Sale Total','Payment','Notes','Device','Synced At'];
var MILE_HEADERS = ['Record ID','Date','Description','Miles','Rate Per Mile','Amount','Notes','Synced At'];
var VEND_HEADERS = ['Record ID','Type','Name','Phone','Email','Address','Website','Products','Notes','Created At'];

// -------------------------------------------------------
// respond
// -------------------------------------------------------
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------
// GET - read all data and send to app
// -------------------------------------------------------
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'read';
    if (action === 'ping') {
      return respond({ success: true, status: 'dentonmedical v3', ts: new Date().getTime() });
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return respond({
      success:  true,
      expenses: readExpenses(ss),
      sales:    readSales(ss),
      vendors:  readVendors(ss),
      settings: readSettings(ss),
      ts:       new Date().getTime()
    });
  } catch(err) {
    return respond({ success: false, error: err.message });
  }
}

// -------------------------------------------------------
// POST - write one record at a time (upsert or delete)
// -------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch(err) {
    return respond({ success: false, error: 'Server busy, please retry.' });
  }

  try {
    var data   = JSON.parse(e.postData.contents);
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var action = data.action || '';

    if (action === 'saveExpense') {
      upsertExpense(ss, data.record);
      if (data.record && data.record.mileage) {
        upsertMileage(ss, data.record);
      }
      return respond({ success: true, action: action, id: data.record.id });
    }

    if (action === 'saveSale') {
      upsertSale(ss, data.record);
      return respond({ success: true, action: action, id: data.record.id });
    }

    if (action === 'deleteExpense') {
      deleteById(ss, SHEET_EXPENSES, data.id);
      deleteById(ss, SHEET_MILEAGE,  data.id);
      return respond({ success: true, action: action });
    }

    if (action === 'deleteSale') {
      deleteAllRowsById(ss, SHEET_SALES, data.id);
      return respond({ success: true, action: action });
    }

    if (action === 'saveVendor') {
      upsertVendor(ss, data.record);
      return respond({ success: true, action: action });
    }

    if (action === 'deleteVendor') {
      deleteById(ss, SHEET_VENDORS, data.id);
      return respond({ success: true, action: action });
    }

    if (action === 'uploadReceipt') {
      var url = saveReceiptToDrive(data.imageBase64, data.filename, data.expenseId);
      return respond({ success: true, action: action, url: url || '' });
    }

    if (action === 'saveSettings') {
      writeSettings(ss, data.settings || {});
      return respond({ success: true, action: action });
    }

    if (action === 'fullSync') {
      var expenses = data.expenses || [];
      var sales    = data.sales    || [];
      var vendors  = data.vendors  || [];
      for (var i = 0; i < expenses.length; i++) {
        upsertExpense(ss, expenses[i]);
        if (expenses[i].mileage) { upsertMileage(ss, expenses[i]); }
      }
      for (var j = 0; j < sales.length; j++) {
        upsertSale(ss, sales[j]);
      }
      for (var k = 0; k < vendors.length; k++) {
        upsertVendor(ss, vendors[k]);
      }
      if (data.settings) { writeSettings(ss, data.settings); }
      return respond({ success: true, action: 'fullSync',
        expenseCount: expenses.length, saleCount: sales.length });
    }

    return respond({ success: false, error: 'Unknown action: ' + action });

  } catch(err) {
    return respond({ success: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------
// Sheet setup helpers
// -------------------------------------------------------
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    writeHeaders(sheet, headers);
  } else if (sheet.getLastRow() === 0) {
    writeHeaders(sheet, headers);
  }
  return sheet;
}

function writeHeaders(sheet, headers) {
  var r = sheet.getRange(1, 1, 1, headers.length);
  r.setValues([headers]);
  r.setFontWeight('bold');
  r.setBackground('#2D5016');
  r.setFontColor('#ffffff');
  r.setFontSize(11);
  sheet.setFrozenRows(1);
  // Force column A (Record ID) to plain text so Google never converts to scientific notation
  sheet.getRange('A:A').setNumberFormat('@');
}

function setIdColumnFormat(sheet) {
  sheet.getRange('A:A').setNumberFormat('@');
}

// -------------------------------------------------------
// Find / delete row by ID (ID is always column 1)
// -------------------------------------------------------
function findRowById(sheet, id) {
  if (!id) { return -1; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return -1; }
  var searchId = String(id).trim();
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === searchId) { return i + 2; }
  }
  return -1;
}

function deleteById(ss, sheetName, id) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !id) { return; }
  var row = findRowById(sheet, id);
  if (row > 0) { sheet.deleteRow(row); }
}

function deleteAllRowsById(ss, sheetName, id) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !id) { return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return; }
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) { sheet.deleteRow(i + 2); }
  }
}

// -------------------------------------------------------
// Expenses
// -------------------------------------------------------
function upsertExpense(ss, e) {
  if (!e || !e.id) { return; }
  var sheet    = getOrCreateSheet(ss, SHEET_EXPENSES, EXP_HEADERS);
  setIdColumnFormat(sheet);
  var mileAmt  = e.mileage ? Math.round((e.mileage * (e.mileageRate || 0.67)) * 100) / 100 : '';
  var row      = [
    String(e.id),
    e.date        || '',
    catLabel(e.category),
    e.description || '',
    e.vendor      || '',
    e.mileage     || '',
    e.mileageRate || '',
    mileAmt,
    e.amount      || 0,
    e.receiptUrl  || '',
    e.notes       || '',
    e.source      || 'manual',
    e.deviceName  || '',
    new Date().toLocaleString()
  ];
  var existing = findRowById(sheet, e.id);
  if (existing > 0) {
    sheet.getRange(existing, 1, 1, EXP_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function readExpenses(ss) {
  var sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var rows   = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r  = rows[i];
    var id = String(r['Record ID'] || '').trim();
    if (!id) { continue; }
    var rawDate = r['Date'];
    var dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = String(rawDate || '').substring(0, 10);
    }
    result.push({
      id:          id,
      date:        dateStr,
      category:    catIdFromLabel(String(r['Category'] || '')),
      description: String(r['Description'] || ''),
      vendor:      String(r['Vendor']       || ''),
      mileage:     r['Miles']          ? parseFloat(r['Miles'])          : null,
      mileageRate: r['Rate Per Mile']  ? parseFloat(r['Rate Per Mile'])  : null,
      amount:      parseFloat(r['Amount']) || 0,
      receiptUrl:  String(r['Receipt URL']  || ''),
      notes:       String(r['Notes']        || ''),
      source:      String(r['Source']       || 'manual')
    });
  }
  return result;
}

// -------------------------------------------------------
// Sales
// -------------------------------------------------------
function upsertSale(ss, s) {
  if (!s || !s.id) { return; }
  deleteAllRowsById(ss, SHEET_SALES, s.id);
  var sheet = getOrCreateSheet(ss, SHEET_SALES, SALE_HEADERS);
  setIdColumnFormat(sheet);
  var now   = new Date().toLocaleString();
  var items = s.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    sheet.appendRow([
      String(s.id),
      s.date    || '',
      s.buyer   || '',
      item.name || '',
      item.unit || '',
      item.qty  || 0,
      item.price    || 0,
      item.subtotal || 0,
      s.total   || 0,
      s.payment || '',
      s.notes   || '',
      s.deviceName  || '',
      now
    ]);
  }
}

function readSales(ss) {
  var sheet = ss.getSheetByName(SHEET_SALES);
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var rows = sheetToObjects(sheet);
  var map  = {};
  for (var i = 0; i < rows.length; i++) {
    var r  = rows[i];
    var id = String(r['Record ID'] || '').trim();
    if (!id) { continue; }
    var rawDate = r['Date'];
    var dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = String(rawDate || '').substring(0, 10);
    }
    if (!map[id]) {
      map[id] = {
        id:      id,
        date:    dateStr,
        buyer:   String(r['Buyer']   || ''),
        total:   parseFloat(r['Sale Total']) || 0,
        payment: String(r['Payment'] || ''),
        notes:   String(r['Notes']   || ''),
        items:   []
      };
    }
    map[id].items.push({
      name:     String(r['Item']         || ''),
      unit:     String(r['Unit']         || ''),
      qty:      parseFloat(r['Qty'])          || 0,
      price:    parseFloat(r['Price Per Unit']) || 0,
      subtotal: parseFloat(r['Subtotal'])     || 0
    });
  }
  var result = [];
  for (var key in map) { result.push(map[key]); }
  return result;
}

// -------------------------------------------------------
// Mileage
// -------------------------------------------------------
function upsertMileage(ss, e) {
  if (!e || !e.id || !e.mileage) { return; }
  var sheet = getOrCreateSheet(ss, SHEET_MILEAGE, MILE_HEADERS);
  setIdColumnFormat(sheet);
  var amt   = Math.round((e.mileage * (e.mileageRate || 0.67)) * 100) / 100;
  var row   = [
    String(e.id),
    e.date        || '',
    e.description || '',
    e.mileage,
    e.mileageRate || 0.67,
    amt,
    e.notes       || '',
    new Date().toLocaleString()
  ];
  var existing = findRowById(sheet, e.id);
  if (existing > 0) {
    sheet.getRange(existing, 1, 1, MILE_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

// -------------------------------------------------------
// Vendors
// -------------------------------------------------------
function upsertVendor(ss, v) {
  if (!v || !v.id) { return; }
  var sheet = getOrCreateSheet(ss, SHEET_VENDORS, VEND_HEADERS);
  setIdColumnFormat(sheet);
  var row   = [
    String(v.id),
    v.type     || '',
    v.name     || '',
    v.phone    || '',
    v.email    || '',
    v.address  || '',
    v.website  || '',
    v.products || '',
    v.notes    || '',
    v.createdAt|| ''
  ];
  var existing = findRowById(sheet, v.id);
  if (existing > 0) {
    sheet.getRange(existing, 1, 1, VEND_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function readVendors(ss) {
  var sheet = ss.getSheetByName(SHEET_VENDORS);
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var rows   = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r  = rows[i];
    var id = String(r['Record ID'] || '');
    if (!id) { continue; }
    result.push({
      id:        id,
      type:      String(r['Type']      || 'seller'),
      name:      String(r['Name']      || ''),
      phone:     String(r['Phone']     || ''),
      email:     String(r['Email']     || ''),
      address:   String(r['Address']   || ''),
      website:   String(r['Website']   || ''),
      products:  String(r['Products']  || ''),
      notes:     String(r['Notes']     || ''),
      createdAt: String(r['Created At']|| '')
    });
  }
  return result;
}

// -------------------------------------------------------
// App Settings
// -------------------------------------------------------
function readSettings(ss) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) { return {}; }
  var vals = sheet.getDataRange().getValues();
  var obj  = {};
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0]) { obj[String(vals[i][0])] = String(vals[i][1]); }
  }
  return obj;
}

function writeSettings(ss, settings) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    var hr = sheet.getRange(1, 1, 1, 2);
    hr.setValues([['Key', 'Value']]);
    hr.setBackground('#2D5016');
    hr.setFontColor('#ffffff');
    hr.setFontWeight('bold');
  }
  var keys = Object.keys(settings);
  for (var i = 0; i < keys.length; i++) {
    var key     = keys[i];
    var val     = String(settings[key]);
    var lastRow = sheet.getLastRow();
    var found   = -1;
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var j = 0; j < existing.length; j++) {
        if (String(existing[j][0]) === key) { found = j + 2; break; }
      }
    }
    if (found > 0) {
      sheet.getRange(found, 2).setValue(val);
    } else {
      sheet.appendRow([key, val]);
    }
  }
}

// -------------------------------------------------------
// Google Drive - receipt photos
// -------------------------------------------------------
function saveReceiptToDrive(base64Data, filename, expenseId) {
  try {
    var rootIt = DriveApp.getFoldersByName(DRIVE_FOLDER);
    var root   = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder(DRIVE_FOLDER);
    var subIt  = root.getFoldersByName(DRIVE_SUBFOLDER);
    var folder = subIt.hasNext() ? subIt.next() : root.createFolder(DRIVE_SUBFOLDER);

    var clean    = base64Data.replace(/^data:[^;]+;base64,/, '');
    var bytes    = Utilities.base64Decode(clean);
    var fname    = filename || ('receipt_' + expenseId + '.jpg');
    var blob     = Utilities.newBlob(bytes, 'image/jpeg', fname);
    var file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(err) {
    Logger.log('Drive error: ' + err.message);
    return null;
  }
}

// -------------------------------------------------------
// Helpers
// -------------------------------------------------------
function sheetToObjects(sheet) {
  var all     = sheet.getDataRange().getValues();
  if (all.length < 2) { return []; }
  var headers = all[0];
  var result  = [];
  for (var i = 1; i < all.length; i++) {
    var row = all[i];
    if (row[0] === '' || row[0] === null) { continue; }
    var obj = {};
    for (var j = 0; j < headers.length; j++) { obj[headers[j]] = row[j]; }
    result.push(obj);
  }
  return result;
}

function catLabel(id) {
  var m = {
    lodging: 'Lodging', meals: 'Meals', mileage: 'Mileage',
    deductible: 'Deductibles', copay: 'Co-Pay', prescription: 'Prescriptions',
    procedure: 'Procedures', lab: 'Lab/Tests', equipment: 'Medical Equipment',
    parking: 'Parking/Tolls', insurance: 'Insurance Premium', dental: 'Dental',
    vision: 'Vision', therapy: 'Therapy', other: 'Other'
  };
  return m[id] || (id ? id.charAt(0).toUpperCase() + id.slice(1) : 'Other');
}

function catIdFromLabel(label) {
  var m = {
    'Lodging': 'lodging', 'Meals': 'meals', 'Mileage': 'mileage',
    'Deductibles': 'deductible', 'Co-Pay': 'copay', 'Prescriptions': 'prescription',
    'Procedures': 'procedure', 'Lab/Tests': 'lab', 'Medical Equipment': 'equipment',
    'Parking/Tolls': 'parking', 'Insurance Premium': 'insurance', 'Dental': 'dental',
    'Vision': 'vision', 'Therapy': 'therapy', 'Other': 'other'
  };
  return m[label] || (label ? label.toLowerCase().replace(/\s+/g, '_') : 'other');
}
