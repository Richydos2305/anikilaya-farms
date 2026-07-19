/**
 * Anikilaya Farms — backend for the static frontend.
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 * Bound to the Google Sheet containing the "Settings", "AF Pen 1", "AF Pen 2",
 * "Sales", and "Expenses" tabs.
 */

var SUBMISSIONS_SHEET_BY_FORM = {
  main: 'AF Pen 1',
  afPen2: 'AF Pen 2',
  sales: 'Sales',
  expenses: 'Expenses',
  inventory: 'Inventory'
};

var FORM_LABELS = {
  main: 'AF Pen 1',
  afPen2: 'AF Pen 2',
  sales: 'Sales',
  expenses: 'Expenses',
  receipts: 'Receipt',
  inventory: 'Inventory',
  broilerRearing: 'Broiler Rearing',
  broilerExpenses: 'Broiler Expense'
};

var RECEIPTS_HEADER = [
  'Timestamp', 'Number', 'Date', 'CustomerName', 'CustomerPhone', 'Items',
  'Subtotal', 'Discount', 'Total', 'AmountPaid', 'PaymentMethod', 'Notes'
];

var BROILER_REARING_HEADER = [
  'Timestamp', 'Date', 'Age', 'OpeningStock', 'Mortality', 'ClosingStock',
  'FeedConsumed', 'Medication', 'AvgBodyWeight', 'Comment'
];

var BROILER_EXPENSES_HEADER = [
  'Timestamp', 'Date', 'Particulars', 'Qty', 'Price', 'Amount'
];

function normalizeForm_(form) {
  return (form === 'sales' || form === 'expenses' || form === 'inventory' || form === 'afPen2') ? form : 'main';
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSettingsSheet_() {
  return getSpreadsheet_().getSheetByName('Settings');
}

function getSubmissionsSheet_(form) {
  return getSpreadsheet_().getSheetByName(SUBMISSIONS_SHEET_BY_FORM[normalizeForm_(form)]);
}

function getReceiptsSheet_() {
  return getSpreadsheet_().getSheetByName('Receipts');
}

function getBroilerRearingSheet_() {
  return getSpreadsheet_().getSheetByName('BroilerRearing');
}

function getBroilerExpensesSheet_() {
  return getSpreadsheet_().getSheetByName('BroilerExpenses');
}

function getOwnerEmail_() {
  return PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
}

/**
 * Run this once manually from the Apps Script editor (select "setup" in the
 * function dropdown, click Run) to create/seed the Settings, AF Pen 1,
 * AF Pen 2, Sales, and Expenses tabs. Safe to re-run — it only creates tabs
 * and seeds field rows that don't already exist, never touches what's
 * already there (and renames a pre-existing "Submissions" tab to "AF Pen 1"
 * in place if found).
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var settings = ss.getSheetByName('Settings');
  if (!settings) {
    settings = ss.insertSheet('Settings');
    settings.getRange(1, 1, 1, 6).setValues([['RowType', 'Key', 'Value', 'Extra', 'Form', 'Options']]);
    settings.getRange(2, 1, 6, 6).setValues([
      ['passcode', 'staff', 'changeme-staff', '', '', ''],
      ['passcode', 'admin', 'changeme-admin', '', '', ''],
      ['field', 'itemName', 'Item Name', 'text', 'main', ''],
      ['field', 'quantity', 'Quantity', 'number', 'main', ''],
      ['field', 'unitCost', 'Unit Cost', 'number', 'main', ''],
      ['field', 'unitPrice', 'Unit Price', 'number', 'main', '']
    ]);
    settings.setFrozenRows(1);
    settings.autoResizeColumns(1, 6);
  }

  // AF Pen 1 used to be called "Submissions" — rename the existing tab in
  // place on re-setup so the entrepreneur's data isn't touched or duplicated.
  var legacySubmissions = ss.getSheetByName('Submissions');
  if (legacySubmissions && !ss.getSheetByName('AF Pen 1')) {
    legacySubmissions.setName('AF Pen 1');
  }
  ensureSubmissionsTab_(ss, 'AF Pen 1');

  // AF Pen 2 is a carbon copy of AF Pen 1 (same fields, same staff-submit
  // behavior) for a second, independent pen — copy whatever fields AF Pen 1
  // currently has (including any the admin has already customized) rather
  // than reseeding the original defaults.
  seedAfPen2FieldsFromMain_(settings);
  ensureSubmissionsTab_(ss, 'AF Pen 2');

  seedFormFields_(settings, 'sales', [
    ['date', 'Date', 'date', ''],
    ['description', 'Description', 'select', 'Crates of Egg,Pig,Broiler,Fish,Others'],
    ['quantity', 'Quantity', 'number', ''],
    ['price', 'Price (₦)', 'number', ''],
    ['total', 'Total (₦)', 'number', '']
  ]);
  ensureSubmissionsTab_(ss, 'Sales');

  seedFormFields_(settings, 'expenses', [
    ['date', 'Date', 'date', ''],
    ['description', 'Description', 'text', ''],
    ['amount', 'Amount (₦)', 'number', '']
  ]);
  ensureSubmissionsTab_(ss, 'Expenses');

  ensureReceiptsTab_(ss);
  seedReceiptSettings_(settings);

  seedFormFields_(settings, 'inventory', [
    ['date', 'Date', 'date', ''],
    ['dept', 'Department', 'select', 'Layers,Broiler,Piggery,Fish,Cattle'],
    ['itemName', 'Item Name', 'text', ''],
    ['quantity', 'Quantity', 'number', '']
  ]);
  // Unlike AF Pen 1/AF Pen 2/Sales/Expenses, Inventory's columns are seeded here
  // right away instead of growing lazily on first submit — otherwise the
  // sheet looks broken/empty (just a bare "Timestamp" column) the moment it's
  // created, before anyone has entered data.
  ensureSubmissionsTab_(ss, 'Inventory', ['date', 'dept', 'itemName', 'quantity']);

  ensureBroilerRearingTab_(ss);
  ensureBroilerExpensesTab_(ss);

  Logger.log('Setup complete. Settings, AF Pen 1, AF Pen 2, Sales, Expenses, Receipts, Inventory, BroilerRearing, and BroilerExpenses tabs are ready.');
}

// Copies AF Pen 1's current field rows (whatever the admin has customized
// them to, including legacy rows with a blank Form column) into a fresh
// 'afPen2' block, so AF Pen 2 starts as a true carbon copy. No-ops once
// afPen2 has its own field rows, so a later AF Pen 1 customization won't
// retroactively overwrite AF Pen 2.
function seedAfPen2FieldsFromMain_(settingsSheet) {
  var lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return;

  var existing = settingsSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var alreadySeeded = existing.some(function (row) { return row[0] === 'field' && row[4] === 'afPen2'; });
  if (alreadySeeded) return;

  var mainFieldRows = existing.filter(function (row) {
    return row[0] === 'field' && (row[4] === 'main' || !row[4]);
  });
  if (mainFieldRows.length === 0) return;

  var newRows = mainFieldRows.map(function (row) {
    return ['field', row[1], row[2], row[3], 'afPen2', row[5]];
  });
  var startRow = settingsSheet.getLastRow() + 1;
  settingsSheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
}

function seedFormFields_(settingsSheet, form, fieldDefs) {
  var lastRow = settingsSheet.getLastRow();
  var alreadySeeded = false;
  if (lastRow >= 2) {
    var existing = settingsSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    alreadySeeded = existing.some(function (row) { return row[0] === 'field' && row[4] === form; });
  }
  if (alreadySeeded) return;

  var newRows = fieldDefs.map(function (f) {
    return ['field', f[0], f[1], f[2], form, f[3]];
  });
  var startRow = settingsSheet.getLastRow() + 1;
  settingsSheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
}

function ensureSubmissionsTab_(ss, name, extraColumns) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var header = ['Timestamp'].concat(extraColumns || []);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, header.length);
    return sheet;
  }

  // Repairs a sheet that was already created before its expected columns
  // existed (e.g. Inventory, first seeded with just "Timestamp") — appends
  // any missing header columns in place without touching existing data rows.
  if (extraColumns && extraColumns.length) {
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    extraColumns.forEach(function (col) {
      if (headerRow.indexOf(col) === -1) {
        headerRow.push(col);
        sheet.getRange(1, headerRow.length).setValue(col);
      }
    });
  }
  return sheet;
}

function ensureReceiptsTab_(ss) {
  var sheet = ss.getSheetByName('Receipts');
  if (!sheet) {
    sheet = ss.insertSheet('Receipts');
    sheet.getRange(1, 1, 1, RECEIPTS_HEADER.length).setValues([RECEIPTS_HEADER]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, RECEIPTS_HEADER.length);
  }
  return sheet;
}

function ensureBroilerRearingTab_(ss) {
  var sheet = ss.getSheetByName('BroilerRearing');
  if (!sheet) {
    sheet = ss.insertSheet('BroilerRearing');
    sheet.getRange(1, 1, 1, BROILER_REARING_HEADER.length).setValues([BROILER_REARING_HEADER]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, BROILER_REARING_HEADER.length);
  }
  return sheet;
}

function ensureBroilerExpensesTab_(ss) {
  var sheet = ss.getSheetByName('BroilerExpenses');
  if (!sheet) {
    sheet = ss.insertSheet('BroilerExpenses');
    sheet.getRange(1, 1, 1, BROILER_EXPENSES_HEADER.length).setValues([BROILER_EXPENSES_HEADER]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, BROILER_EXPENSES_HEADER.length);
  }
  return sheet;
}

// Seeds a default footer message once, the first time setup() runs after
// this feature ships. Leaves phone/whatsapp/address blank for the admin to
// fill in from the Receipts view — only seeds if no 'setting' rows exist yet,
// so re-running setup() never clobbers values the admin has already saved.
function seedReceiptSettings_(settingsSheet) {
  var lastRow = settingsSheet.getLastRow();
  var alreadySeeded = false;
  if (lastRow >= 2) {
    var existing = settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    alreadySeeded = existing.some(function (row) { return row[0] === 'setting'; });
  }
  if (alreadySeeded) return;

  var startRow = settingsSheet.getLastRow() + 1;
  settingsSheet.getRange(startRow, 1, 1, 6).setValues([
    ['setting', 'footer', 'Thank you for your patronage!', '', '', '']
  ]);
}

function padSeq_(n) {
  var s = String(n);
  while (s.length < 4) s = '0' + s;
  return s;
}

// Sheets silently converts a 'yyyy-MM-dd' string written via setValues() into
// a real Date-typed cell, so reading the Date column back needs to reformat
// it to plain 'yyyy-MM-dd' or the frontend's date parsing breaks.
function formatDateCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return jsonResponse_({ ok: false, error: 'use_post' });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'server_error' });
  }

  try {
    switch (body.action) {
      case 'login':
        return handleLogin_(body);
      case 'getFields':
        return handleGetFields_(body);
      case 'saveFields':
        return handleSaveFields_(body);
      case 'submit':
        return handleSubmit_(body);
      case 'getReceiptSettings':
        return handleGetReceiptSettings_(body);
      case 'saveReceiptSettings':
        return handleSaveReceiptSettings_(body);
      case 'listReceipts':
        return handleListReceipts_(body);
      case 'saveReceipt':
        return handleSaveReceipt_(body);
      case 'deleteReceipt':
        return handleDeleteReceipt_(body);
      case 'saveBroilerRearing':
        return handleSaveBroilerRearing_(body);
      case 'listBroilerRearing':
        return handleListBroilerRearing_(body);
      case 'saveBroilerExpense':
        return handleSaveBroilerExpense_(body);
      case 'listBroilerExpenses':
        return handleListBroilerExpenses_(body);
      default:
        return jsonResponse_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    Logger.log(err);
    return jsonResponse_({ ok: false, error: 'server_error' });
  }
}

// ─── Settings helpers ───

function readSettingsAllRows_() {
  var sheet = getSettingsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), 6);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      return {
        rowType: row[0],
        key: row[1],
        value: row[2],
        extra: row[3],
        form: row[4] || '',
        options: row[5] || ''
      };
    });
}

function readPasscodeMap_() {
  var map = {};
  readSettingsAllRows_().forEach(function (r) {
    if (r.rowType === 'passcode') map[r.key] = String(r.value);
  });
  return map;
}

// Receipts is an admin-only tool — unlike Sales/Expenses there's no staff
// half of the flow, so every action (including just reading settings) needs
// the admin passcode specifically, not just any valid login.
function isAdminAuthorized_(passcode, passcodes) {
  return String(passcode || '').trim() === passcodes.admin;
}

function readReceiptSettings_() {
  var settings = { phone: '', whatsapp: '', address: '', footer: 'Thank you for your patronage!' };
  readSettingsAllRows_().forEach(function (r) {
    if (r.rowType === 'setting' && settings.hasOwnProperty(r.key)) {
      settings[r.key] = r.value;
    }
  });
  return settings;
}

function readFieldRows_(form) {
  form = normalizeForm_(form);
  return readSettingsAllRows_()
    .filter(function (r) {
      if (r.rowType !== 'field') return false;
      var rowForm = r.form || 'main';
      return rowForm === form;
    })
    .map(function (r) {
      var options =
        r.extra === 'select' && r.options
          ? String(r.options).split(',').map(function (o) { return o.trim(); }).filter(Boolean)
          : [];
      return { id: r.key, label: r.value, type: r.extra, options: options };
    });
}

// ─── Actions ───

function handleLogin_(body) {
  var passcodes = readPasscodeMap_();
  var entered = String(body.passcode || '').trim();

  var role = null;
  if (entered && entered === passcodes.staff) role = 'staff';
  else if (entered && entered === passcodes.admin) role = 'admin';

  if (!role) {
    return jsonResponse_({ ok: false, error: 'invalid_passcode' });
  }
  return jsonResponse_({ ok: true, role: role });
}

function handleGetFields_(body) {
  var form = normalizeForm_(body.form);
  return jsonResponse_({ ok: true, fields: readFieldRows_(form) });
}

function handleSaveFields_(body) {
  var form = normalizeForm_(body.form);
  var passcodes = readPasscodeMap_();

  if (String(body.passcode || '').trim() !== passcodes.admin) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var incoming = body.fields;
  var isValid =
    Array.isArray(incoming) &&
    incoming.length > 0 &&
    incoming.every(function (f) {
      if (
        !f ||
        typeof f.id !== 'string' ||
        f.id.trim() === '' ||
        typeof f.label !== 'string' ||
        f.label.trim() === ''
      ) {
        return false;
      }
      if (f.type !== 'text' && f.type !== 'number' && f.type !== 'date' && f.type !== 'select') {
        return false;
      }
      if (f.type === 'select') {
        return Array.isArray(f.options) && f.options.some(function (o) { return String(o).trim() !== ''; });
      }
      return true;
    });

  if (!isValid) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSettingsSheet_();
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 6);
    var allValues = lastRow >= 1 ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : [];

    // Remember where this form's block starts before touching anything, so
    // the rewritten rows can go back to the same spot instead of jumping to
    // the bottom of the sheet — that's what lets a blank divider row stay
    // put as a visual separator between forms' field blocks.
    var firstMatchIndex = -1;
    for (var i = 1; i < allValues.length; i++) {
      var rowForm = allValues[i][4] || 'main';
      if (allValues[i][0] === 'field' && rowForm === form) {
        firstMatchIndex = i;
        break;
      }
    }

    // Delete this form's existing field rows (bottom-to-top so row indices
    // stay valid as we go). Safer than clearing and rewriting a fixed-size
    // block in place: this form's row count can grow or shrink from one save
    // to the next, and its rows may sit right next to another form's block,
    // so an in-place overwrite could clobber a neighboring form's rows.
    for (var j = allValues.length - 1; j >= 1; j--) {
      var delRowForm = allValues[j][4] || 'main';
      if (allValues[j][0] === 'field' && delRowForm === form) {
        sheet.deleteRow(j + 1); // allValues is 0-indexed from row 1; sheet rows are 1-based
      }
    }

    var newRows = incoming.map(function (f) {
      var optionsStr =
        f.type === 'select' && Array.isArray(f.options)
          ? f.options.map(function (o) { return String(o).trim(); }).filter(Boolean).join(',')
          : '';
      return ['field', f.id, f.label, f.type, form, optionsStr];
    });

    // firstMatchIndex is only valid post-deletion because every row deleted
    // above sat at or after it (it was the first match) — nothing before it
    // shifted, so it's still the right sheet row (1-based) to insert at.
    if (firstMatchIndex !== -1) {
      sheet.insertRowsBefore(firstMatchIndex + 1, newRows.length);
      sheet.getRange(firstMatchIndex + 1, 1, newRows.length, 6).setValues(newRows);
    } else {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
    }

    var savedFields = incoming.map(function (f) {
      return {
        id: f.id,
        label: f.label,
        type: f.type,
        options:
          f.type === 'select'
            ? (f.options || []).map(function (o) { return String(o).trim(); }).filter(Boolean)
            : []
      };
    });
    return jsonResponse_({ ok: true, fields: savedFields });
  } finally {
    lock.releaseLock();
  }
}

function handleSubmit_(body) {
  var form = normalizeForm_(body.form);
  var passcodes = readPasscodeMap_();
  var entered = String(body.passcode || '').trim();

  // Inventory is the one generic form both staff and admin can submit to —
  // everything else (Main/Sales/Expenses) stays staff-only.
  var authorized = form === 'inventory'
    ? (entered === passcodes.staff || entered === passcodes.admin)
    : entered === passcodes.staff;

  if (!authorized) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var fields = readFieldRows_(form);
  var values = body.values || {};

  var missing = fields.some(function (f) {
    var v = values[f.id];
    return v === undefined || v === null || String(v).trim() === '';
  });
  if (missing) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSubmissionsSheet_(form);
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setValue(new Date());

    fields.forEach(function (field) {
      var colIndex = headerRow.indexOf(field.id);
      if (colIndex === -1) {
        colIndex = headerRow.length;
        sheet.getRange(1, colIndex + 1).setValue(field.id);
        headerRow.push(field.id);
      }
      var raw = values[field.id];
      var cellValue =
        field.type === 'number' ? Number(raw) :
        field.type === 'date' ? new Date(raw) :
        String(raw);
      sheet.getRange(newRow, colIndex + 1).setValue(cellValue);
    });

    if (form === 'main' || form === 'afPen2') {
      var eggRateColIndex = headerRow.indexOf('Egg Production Rate');
      if (eggRateColIndex === -1) {
        eggRateColIndex = headerRow.length;
        sheet.getRange(1, eggRateColIndex + 1).setValue('Egg Production Rate');
        headerRow.push('Egg Production Rate');
      }
      var eggRateCell = sheet.getRange(newRow, eggRateColIndex + 1);
      var eggRate = computeEggProductionRate_(values);
      eggRateCell.setValue(eggRate);
      if (eggRate !== '') {
        eggRateCell.setNumberFormat('0.00%');
      }
    }

    sendOwnerAlertEmail_(FORM_LABELS[form]);

    return jsonResponse_({ ok: true });
  } catch (err) {
    Logger.log(err);
    return jsonResponse_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// ─── Receipts ───

function handleGetReceiptSettings_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }
  return jsonResponse_({ ok: true, settings: readReceiptSettings_() });
}

function handleSaveReceiptSettings_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var incoming = {
    phone: String(body.phone || '').trim(),
    whatsapp: String(body.whatsapp || '').trim(),
    address: String(body.address || '').trim(),
    footer: String(body.footer || '').trim()
  };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSettingsSheet_();
    var lastRow = sheet.getLastRow();
    var values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 6).getValues() : [];

    Object.keys(incoming).forEach(function (key) {
      var rowIndex = -1;
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === 'setting' && values[i][1] === key) {
          rowIndex = i;
          break;
        }
      }
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 2, 3).setValue(incoming[key]);
      } else {
        var newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1, 1, 6).setValues([['setting', key, incoming[key], '', '', '']]);
        values.push(['setting', key, incoming[key], '', '', '']);
      }
    });

    return jsonResponse_({ ok: true, settings: incoming });
  } finally {
    lock.releaseLock();
  }
}

function handleListReceipts_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var sheet = getReceiptsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ ok: true, receipts: [] });
  }

  var values = sheet.getRange(2, 1, lastRow - 1, RECEIPTS_HEADER.length).getValues();
  var receipts = values.map(function (row) {
    var items = [];
    try {
      items = JSON.parse(row[5] || '[]');
    } catch (e) {
      items = [];
    }
    return {
      id: row[1],
      number: row[1],
      date: formatDateCell_(row[2]),
      customerName: row[3],
      customerPhone: row[4],
      items: items,
      subtotal: Number(row[6]) || 0,
      discount: Number(row[7]) || 0,
      total: Number(row[8]) || 0,
      amountPaid: Number(row[9]) || 0,
      paymentMethod: row[10],
      notes: row[11],
      createdAt: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || '')
    };
  });

  receipts.reverse(); // sheet rows grow downward on append; newest-first for the log
  return jsonResponse_({ ok: true, receipts: receipts });
}

function handleSaveReceipt_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var customerName = String(body.customerName || '').trim();
  var items = Array.isArray(body.items) ? body.items : [];
  var hasItem = items.some(function (i) {
    return i && (String(i.desc || '').trim() || Number(i.qty) || Number(i.price));
  });

  if (!customerName || !hasItem) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getReceiptsSheet_();

    // Number assigned here, under the same lock as the row insert, so two
    // people saving a receipt at nearly the same moment can't land on the
    // same sequence number the way the original client-side counter could.
    var year = new Date().getFullYear();
    var lastRow = sheet.getLastRow();
    var seq = 1;
    if (lastRow >= 2) {
      var numbers = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      var countThisYear = numbers.filter(function (r) {
        return String(r[0] || '').indexOf('-' + year + '-') !== -1;
      }).length;
      seq = countThisYear + 1;
    }
    var number = 'AF-' + year + '-' + padSeq_(seq);

    var timestamp = new Date();
    var newRow = sheet.getLastRow() + 1;
    var receipt = {
      id: number,
      number: number,
      date: String(body.date || ''),
      customerName: customerName,
      customerPhone: String(body.customerPhone || '').trim(),
      items: items,
      subtotal: Number(body.subtotal) || 0,
      discount: Number(body.discount) || 0,
      total: Number(body.total) || 0,
      amountPaid: Number(body.amountPaid) || 0,
      paymentMethod: String(body.paymentMethod || ''),
      notes: String(body.notes || '').trim(),
      createdAt: timestamp.toISOString()
    };

    sheet.getRange(newRow, 1, 1, RECEIPTS_HEADER.length).setValues([[
      timestamp, receipt.number, receipt.date, receipt.customerName, receipt.customerPhone,
      JSON.stringify(items), receipt.subtotal, receipt.discount, receipt.total,
      receipt.amountPaid, receipt.paymentMethod, receipt.notes
    ]]);

    sendOwnerAlertEmail_(FORM_LABELS.receipts);

    return jsonResponse_({ ok: true, receipt: receipt });
  } finally {
    lock.releaseLock();
  }
}

function handleDeleteReceipt_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var number = String(body.number || '').trim();
  if (!number) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getReceiptsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var numbers = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = numbers.length - 1; i >= 0; i--) {
        if (numbers[i][0] === number) {
          sheet.deleteRow(i + 2);
          break;
        }
      }
    }

    return jsonResponse_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// ─── Broiler Rearing & Broiler Expenses ───

function handleSaveBroilerRearing_(body) {
  var passcodes = readPasscodeMap_();
  if (String(body.passcode || '').trim() !== passcodes.staff) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var date = String(body.date || '').trim();
  var age = String(body.age || '').trim();
  var openingStock = Number(body.openingStock);
  var mortality = Number(body.mortality);

  if (!date || !age || body.openingStock === undefined || body.openingStock === '' ||
      body.mortality === undefined || body.mortality === '' || isNaN(openingStock) || isNaN(mortality)) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getBroilerRearingSheet_();
    var timestamp = new Date();
    var closingStock = openingStock - mortality;
    var entry = {
      date: date,
      age: age,
      openingStock: openingStock,
      mortality: mortality,
      closingStock: closingStock,
      feedConsumed: String(body.feedConsumed || '').trim(),
      medication: String(body.medication || '').trim(),
      avgBodyWeight: String(body.avgBodyWeight || '').trim(),
      comment: String(body.comment || '').trim(),
      createdAt: timestamp.toISOString()
    };

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, BROILER_REARING_HEADER.length).setValues([[
      timestamp, entry.date, entry.age, entry.openingStock, entry.mortality, entry.closingStock,
      entry.feedConsumed, entry.medication, entry.avgBodyWeight, entry.comment
    ]]);

    sendOwnerAlertEmail_(FORM_LABELS.broilerRearing);

    return jsonResponse_({ ok: true, entry: entry });
  } finally {
    lock.releaseLock();
  }
}

function handleListBroilerRearing_(body) {
  var passcodes = readPasscodeMap_();
  var entered = String(body.passcode || '').trim();
  if (!entered || (entered !== passcodes.staff && entered !== passcodes.admin)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var sheet = getBroilerRearingSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ ok: true, entries: [] });
  }

  var values = sheet.getRange(2, 1, lastRow - 1, BROILER_REARING_HEADER.length).getValues();
  var entries = values.map(function (row) {
    return {
      date: formatDateCell_(row[1]),
      age: String(row[2] || ''),
      openingStock: Number(row[3]) || 0,
      mortality: Number(row[4]) || 0,
      closingStock: Number(row[5]) || 0,
      feedConsumed: String(row[6] || ''),
      medication: String(row[7] || ''),
      avgBodyWeight: String(row[8] || ''),
      comment: String(row[9] || ''),
      createdAt: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || '')
    };
  });

  entries.reverse(); // sheet rows grow downward on append; newest-first for the log
  return jsonResponse_({ ok: true, entries: entries });
}

function handleSaveBroilerExpense_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var date = String(body.date || '').trim();
  var particulars = String(body.particulars || '').trim();
  var amount = Number(body.amount);

  if (!date || !particulars || body.amount === undefined || body.amount === '' || isNaN(amount)) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getBroilerExpensesSheet_();
    var timestamp = new Date();
    var entry = {
      date: date,
      particulars: particulars,
      qty: body.qty === undefined || body.qty === '' ? '' : Number(body.qty),
      price: body.price === undefined || body.price === '' ? '' : Number(body.price),
      amount: amount,
      createdAt: timestamp.toISOString()
    };

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, BROILER_EXPENSES_HEADER.length).setValues([[
      timestamp, entry.date, entry.particulars, entry.qty, entry.price, entry.amount
    ]]);

    sendOwnerAlertEmail_(FORM_LABELS.broilerExpenses);

    return jsonResponse_({ ok: true, entry: entry });
  } finally {
    lock.releaseLock();
  }
}

function handleListBroilerExpenses_(body) {
  var passcodes = readPasscodeMap_();
  if (!isAdminAuthorized_(body.passcode, passcodes)) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var sheet = getBroilerExpensesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ ok: true, entries: [] });
  }

  var values = sheet.getRange(2, 1, lastRow - 1, BROILER_EXPENSES_HEADER.length).getValues();
  var runningBalance = 0;
  var entries = values.map(function (row) {
    var amount = Number(row[5]) || 0;
    runningBalance += amount;
    return {
      date: formatDateCell_(row[1]),
      particulars: String(row[2] || ''),
      qty: row[3] === '' || row[3] === null ? '' : Number(row[3]),
      price: row[4] === '' || row[4] === null ? '' : Number(row[4]),
      amount: amount,
      cumulativeBalance: runningBalance,
      createdAt: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || '')
    };
  });

  entries.reverse(); // reverse only after the running sum is computed, so the math stays chronological
  return jsonResponse_({ ok: true, entries: entries });
}

// Egg Production Rate = (Crate Collected × 30) / Closing Stock, stored as a
// fraction and cell-formatted as a percentage (e.g. 0.8531 displays as 85.31%).
// Returns '' (blank cell) instead of erroring when either input is missing,
// non-numeric, or Closing Stock is 0 — avoids writing NaN/Infinity to the sheet.
function computeEggProductionRate_(values) {
  var crateCollected = Number(values.crateCollected);
  var closingStock = Number(values.closingStock);
  if (isNaN(crateCollected) || isNaN(closingStock) || closingStock === 0) {
    return '';
  }
  return (crateCollected * 30) / closingStock;
}

function sendOwnerAlertEmail_(formLabel) {
  try {
    var ownerEmail = getOwnerEmail_();
    if (!ownerEmail) return;
    var sheetUrl = getSpreadsheet_().getUrl();
    var timestamp = new Date().toString();
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'New ' + formLabel + ' Submission Received from Anikilaya Farms',
      body:
        'A new ' + formLabel + ' submission was received at ' + timestamp +
        '. Open the Sheet to review: ' + sheetUrl,
      htmlBody:
        'A new ' + formLabel + ' submission was received at ' + timestamp + '. Open the ' +
        '<a href="' + sheetUrl + '">Sheet</a> to review.'
    });
  } catch (err) {
    Logger.log(err); // never let a mail failure fail the submission
  }
}
