/**
 * Anikilaya Farms — backend for the static frontend.
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 * Bound to the Google Sheet containing the "Settings", "Submissions", "Sales",
 * and "Expenses" tabs.
 */

var SUBMISSIONS_SHEET_BY_FORM = {
  main: 'Submissions',
  sales: 'Sales',
  expenses: 'Expenses'
};

var FORM_LABELS = {
  main: 'Daily Farm Record',
  sales: 'Sales',
  expenses: 'Expenses',
  receipts: 'Receipt'
};

var RECEIPTS_HEADER = [
  'Timestamp', 'Number', 'Date', 'CustomerName', 'CustomerPhone', 'Items',
  'Subtotal', 'Discount', 'Total', 'AmountPaid', 'PaymentMethod', 'Notes'
];

function normalizeForm_(form) {
  return (form === 'sales' || form === 'expenses') ? form : 'main';
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

function getOwnerEmail_() {
  return PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
}

/**
 * Run this once manually from the Apps Script editor (select "setup" in the
 * function dropdown, click Run) to create/seed the Settings, Submissions,
 * Sales, and Expenses tabs. Safe to re-run — it only creates tabs and seeds
 * field rows that don't already exist, never touches what's already there.
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

  ensureSubmissionsTab_(ss, 'Submissions');

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

  Logger.log('Setup complete. Settings, Submissions, Sales, Expenses, and Receipts tabs are ready.');
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

function ensureSubmissionsTab_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1).setValue('Timestamp');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumn(1);
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

  if (String(body.passcode || '').trim() !== passcodes.staff) {
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

    if (form === 'main') {
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
      date: String(row[2] || ''),
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
