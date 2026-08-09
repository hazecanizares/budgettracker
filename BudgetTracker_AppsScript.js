// ============================================================
// BUDGET TRACKER — Google Apps Script (verified sync logic —
// locking + merge-by-ID + deletion tombstones, tested against
// add/edit/delete/stale-device/concurrent-save scenarios)
// Paste this entire file into script.google.com, then
// Deploy > Manage deployments > Edit > New version > Deploy
// ============================================================

var SHEET_ID = '1Q47S8vigKMZobBj1ZqiBdklBsoY0PXNIEbqJD8P8QxY'; // from the sheet's URL, between /d/ and /edit
var SHEET_NAME = 'BudgetData'; // Do not change


// ============================================================
// doGet — called when the HTML app loads data (manual "Load from Sheet")
// ============================================================
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'load') {
    return handleLoad();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Budget Tracker script is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// doPost — called when the HTML app saves data (manual "Save to Sheet")
// Locked so two simultaneous saves (two tabs/devices) can't
// interleave and corrupt each other.
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // wait up to 15s for any other save to finish
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Server busy — please try again in a moment.' });
  }

  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save') {
      return handleSave(body.data);
    }
    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// LOAD — read the JSON blob from the sheet and return it as-is.
// ============================================================
function handleLoad() {
  try {
    var sheet = getOrCreateSheet();
    var raw = sheet.getRange(1, 1).getValue();

    if (!raw || raw === '') {
      return jsonResponse({ status: 'ok', data: defaultData() });
    }

    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (parseErr) {
      // A1 has non-JSON content in it (e.g. something got pasted in by
      // mistake). Don't crash the load or silently discard it — report
      // it clearly so it can be fixed by hand in the sheet.
      return jsonResponse({
        status: 'error',
        message: 'Cell A1 does not contain valid data (found: "' +
          String(raw).substring(0, 40) + '..."). Check A1 in the BudgetData sheet directly.'
      });
    }

    return jsonResponse({ status: 'ok', data: parsed });

  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Load failed: ' + err.toString() });
  }
}


// ============================================================
// SAVE — merge incoming data into what's already on the sheet
// (instead of blindly overwriting it) and write the result.
// This is what protects you if two devices save close together.
// ============================================================
function handleSave(incoming) {
  try {
    var sheet = getOrCreateSheet();
    var raw = sheet.getRange(1, 1).getValue();
    var current;
    if (!raw || raw === '') {
      current = defaultData();
    } else {
      try {
        current = JSON.parse(raw);
      } catch (e) {
        current = defaultData(); // corrupted cell — don't block saving forever
      }
    }

    // Union of every id ever marked deleted, from both sides — once
    // deleted, always excluded, so a stale device that hasn't heard
    // about the deletion yet can't bring it back.
    var deletedSet = {};
    (current.deletedIds || []).forEach(function (id) { deletedSet[id] = true; });
    (incoming.deletedIds || []).forEach(function (id) { deletedSet[id] = true; });

    var merged = {
      expenses: mergeById(current.expenses, incoming.expenses, deletedSet),
      banks: mergeById(current.banks, incoming.banks, deletedSet),
      finEntries: {
        allan: mergeById((current.finEntries || {}).allan, (incoming.finEntries || {}).allan, deletedSet),
        hazel: mergeById((current.finEntries || {}).hazel, (incoming.finEntries || {}).hazel, deletedSet)
      },
      nextId: Math.max(current.nextId || 1, incoming.nextId || 1),
      nextFinId: Math.max(current.nextFinId || 1, incoming.nextFinId || 1),
      deletedIds: Object.keys(deletedSet)
    };

    // Carry over any other fields (noteTables, etc.) that aren't
    // array-based — these take whichever save was last.
    for (var k in incoming) {
      if (!(k in merged)) merged[k] = incoming[k];
    }
    for (var k2 in current) {
      if (!(k2 in merged)) merged[k2] = current[k2];
    }

    sheet.getRange(1, 1).setValue(JSON.stringify(merged));
    sheet.getRange(1, 2).setValue('Last saved: ' + new Date().toLocaleString());
    writeReadableExpenses(sheet, merged);

    return jsonResponse({ status: 'ok', message: 'Saved successfully', data: merged });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Save failed: ' + err.toString() });
  }
}

// Merge two arrays of {id, ...} objects. Incoming wins on matching
// ids (edits apply); ids only present in "current" are kept (so a
// stale save from another device can't erase items it didn't know
// about); ids only in "incoming" are added. Anything in deletedSet
// is excluded from either side — this is how an intentional delete
// survives being merged against a device that doesn't know about it yet.
function mergeById(currentArr, incomingArr, deletedSet) {
  currentArr = currentArr || [];
  incomingArr = incomingArr || [];
  deletedSet = deletedSet || {};
  var map = {};
  currentArr.forEach(function (item) { if (item && item.id && !deletedSet[item.id]) map[item.id] = item; });
  incomingArr.forEach(function (item) { if (item && item.id && !deletedSet[item.id]) map[item.id] = item; });
  return Object.keys(map).map(function (id) { return map[id]; });
}


// ============================================================
// READABLE ROWS — writes expenses as a human-readable table
// starting at row 3, for reference only (not read back into the app)
// ============================================================
function writeReadableExpenses(sheet, data) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, 10).clearContent();
    }

    var headers = ['Month', 'Year', 'Category', 'Description', 'Bank', 'Amount', 'Paid By', 'Paid?', 'Installment', 'Months Left'];
    sheet.getRange(3, 1, 1, headers.length).setValues([headers]);

    var expenses = data.expenses || [];
    if (expenses.length === 0) return;

    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var rows = expenses.map(function(e) {
      return [
        MONTHS[e.month] || e.month,
        e.year,
        e.category || '',
        e.desc || '',
        e.bank || '',
        e.amount || 0,
        e.paidBy || '',
        e.paid ? 'Yes' : 'No',
        e.installment === 'yes' ? 'Yes' : 'No',
        e.installment === 'yes' ? (e.installmentMonths || 0) : ''
      ];
    });

    sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);

  } catch (err) {
    Logger.log('writeReadableExpenses error: ' + err.toString());
  }
}


// ============================================================
// HELPERS
// ============================================================
function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function defaultData() {
  return {
    expenses: [],
    banks: [
      { id: 'b1', name: 'MB Platinum' },
      { id: 'b2', name: 'BPI' },
      { id: 'b3', name: 'MB MFREE' }
    ],
    finEntries: { allan: [], hazel: [] },
    nextId: 1,
    nextFinId: 1,
    deletedIds: []
  };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
