/**
 * Anikilaya Farms — backend for the static frontend.
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 * Bound to the Google Sheet containing the "Settings" and "Submissions" tabs.
 */

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSettingsSheet_() {
  return getSpreadsheet_().getSheetByName('Settings');
}

function getSubmissionsSheet_() {
  return getSpreadsheet_().getSheetByName('Submissions');
}

function getOwnerEmail_() {
  return PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
}

/**
 * Run this once manually from the Apps Script editor (select "setup" in the
 * function dropdown, click Run) to create and seed the Settings and
 * Submissions tabs. Safe to re-run — it only creates tabs that don't already
 * exist and never touches one that's already there.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var settings = ss.getSheetByName('Settings');
  if (!settings) {
    settings = ss.insertSheet('Settings');
    settings.getRange(1, 1, 1, 4).setValues([['RowType', 'Key', 'Value', 'Extra']]);
    settings.getRange(2, 1, 6, 4).setValues([
      ['passcode', 'staff', 'changeme-staff', ''],
      ['passcode', 'admin', 'changeme-admin', ''],
      ['field', 'itemName', 'Item Name', 'text'],
      ['field', 'quantity', 'Quantity', 'number'],
      ['field', 'unitCost', 'Unit Cost', 'number'],
      ['field', 'unitPrice', 'Unit Price', 'number']
    ]);
    settings.setFrozenRows(1);
    settings.autoResizeColumns(1, 4);
  }

  var submissions = ss.getSheetByName('Submissions');
  if (!submissions) {
    submissions = ss.insertSheet('Submissions');
    submissions.getRange(1, 1).setValue('Timestamp');
    submissions.setFrozenRows(1);
    submissions.autoResizeColumn(1);
  }

  Logger.log('Setup complete. Settings and Submissions tabs are ready.');
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
      default:
        return jsonResponse_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    Logger.log(err);
    return jsonResponse_({ ok: false, error: 'server_error' });
  }
}

// ─── Settings helpers ───

function readSettingsRows_() {
  var sheet = getSettingsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return values
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      return { rowType: row[0], key: row[1], value: row[2], extra: row[3] };
    });
}

function getPasscodeMap_(rows) {
  var map = {};
  rows.forEach(function (r) {
    if (r.rowType === 'passcode') map[r.key] = String(r.value);
  });
  return map;
}

function getFieldRows_(rows) {
  return rows
    .filter(function (r) { return r.rowType === 'field'; })
    .map(function (r) { return { id: r.key, label: r.value, type: r.extra }; });
}

// ─── Actions ───

function handleLogin_(body) {
  var rows = readSettingsRows_();
  var passcodes = getPasscodeMap_(rows);
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
  var rows = readSettingsRows_();
  return jsonResponse_({ ok: true, fields: getFieldRows_(rows) });
}

function handleSaveFields_(body) {
  var rows = readSettingsRows_();
  var passcodes = getPasscodeMap_(rows);

  if (String(body.passcode || '').trim() !== passcodes.admin) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var incoming = body.fields;
  var isValid =
    Array.isArray(incoming) &&
    incoming.length > 0 &&
    incoming.every(function (f) {
      return (
        f &&
        typeof f.id === 'string' &&
        f.id.trim() !== '' &&
        typeof f.label === 'string' &&
        f.label.trim() !== '' &&
        (f.type === 'text' || f.type === 'number' || f.type === 'date')
      );
    });

  if (!isValid) {
    return jsonResponse_({ ok: false, error: 'validation' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSettingsSheet_();
    var allValues = sheet.getDataRange().getValues();
    var fieldRowIndices = [];
    allValues.forEach(function (row, i) {
      if (row[0] === 'field') fieldRowIndices.push(i);
    });

    var firstFieldRow;
    if (fieldRowIndices.length > 0) {
      firstFieldRow = fieldRowIndices[0] + 1; // 1-based
      var lastFieldRow = fieldRowIndices[fieldRowIndices.length - 1] + 1;
      var numExisting = lastFieldRow - firstFieldRow + 1;
      sheet.getRange(firstFieldRow, 1, numExisting, 4).clearContent();
    } else {
      firstFieldRow = sheet.getLastRow() + 1;
    }

    var newRows = incoming.map(function (f) {
      return ['field', f.id, f.label, f.type];
    });
    sheet.getRange(firstFieldRow, 1, newRows.length, 4).setValues(newRows);

    return jsonResponse_({ ok: true, fields: incoming });
  } finally {
    lock.releaseLock();
  }
}

function handleSubmit_(body) {
  var rows = readSettingsRows_();
  var passcodes = getPasscodeMap_(rows);

  if (String(body.passcode || '').trim() !== passcodes.staff) {
    return jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  var fields = getFieldRows_(rows);
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

    var sheet = getSubmissionsSheet_();
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

    sendOwnerAlertEmail_();

    return jsonResponse_({ ok: true });
  } catch (err) {
    Logger.log(err);
    return jsonResponse_({ ok: false, error: 'server_error' });
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

function sendOwnerAlertEmail_() {
  try {
    var ownerEmail = getOwnerEmail_();
    if (!ownerEmail) return;
    var sheetUrl = getSpreadsheet_().getUrl();
    var timestamp = new Date().toString();
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'New Submission Received from Anikilaya Farms',
      body: 'A new submission was received at ' + timestamp + '. Open the Sheet to review: ' + sheetUrl,
      htmlBody:
        'A new submission was received at ' + timestamp + '. Open the ' +
        '<a href="' + sheetUrl + '">Sheet</a> to review.'
    });
  } catch (err) {
    Logger.log(err); // never let a mail failure fail the submission
  }
}
