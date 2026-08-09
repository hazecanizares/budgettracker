// ============================================================
// BUDGET TRACKER — Google Apps Script
// Storage design: ONE ROW PER RECORD (not one JSON blob in one
// cell). Google Sheets has a hard 50,000-character-per-cell limit —
// with 100+ expenses, the old single-blob design silently exceeded
// it and Sheets rejected the write without throwing an error. This
// version stores each expense/bank/entry as its own small row, so
// no single cell ever comes close to that limit no matter how much
// data accumulates over time.
//
// Paste this entire file into script.google.com, then
// Deploy > Manage deployments > Edit > New version > Deploy
// ============================================================

var SHEET_ID = '1Q47S8vigKMZobBj1ZqiBdklBsoY0PXNIEbqJD8P8QxY';
var SHEET_NAME = 'BudgetData'; // Do not change

var HEADERS = ['ID', 'Type', 'Data'];


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
    lock.waitLock(15000);
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
// LOAD — read every row and reconstruct the data object the app expects
// ============================================================
function handleLoad() {
  try {
    var sheet = getOrCreateSheet();
    var data = readAllRows(sheet);
    return jsonResponse({ status: 'ok', data: data });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Load failed: ' + err.toString() });
  }
}


// ============================================================
// SAVE — merge incoming data into what's already on the sheet
// (by ID, respecting deletions), then rewrite every row fresh.
// ============================================================
function handleSave(incoming) {
  try {
    var sheet = getOrCreateSheet();
    var current = readAllRows(sheet);

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
      deletedIds: Object.keys(deletedSet),
      noteTables: incoming.noteTables || current.noteTables || {}
    };

    writeAllRows(sheet, merged);

    // Recompute nextId/nextFinId the same way readAllRows does, for
    // the response the app uses locally between saves.
    var withCounters = readAllRows(sheet);

    return jsonResponse({
      status: 'ok',
      message: 'Saved successfully',
      data: withCounters,
      debugRowsWritten: (merged.expenses.length + merged.banks.length +
        merged.finEntries.allan.length + merged.finEntries.hazel.length + merged.deletedIds.length + 1),
      debugExpenseCount: merged.expenses.length
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Save failed: ' + err.toString() });
  }
}

// Merge two arrays of {id, ...} objects. Incoming wins on matching
// ids (edits apply); ids only present in "current" are kept (so a
// stale save from another device can't erase items it didn't know
// about); ids only in "incoming" are added. Anything in deletedSet
// is excluded from either side.
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
// ROW STORAGE — one record per row: [ID, Type, Data(json)]
// Type is one of: expense, bank, fin_allan, fin_hazel, deleted, meta
// ============================================================
function readAllRows(sheet) {
  var data = { expenses: [], banks: [], finEntries: { allan: [], hazel: [] }, deletedIds: [], noteTables: {}, nextId: 1, nextFinId: 1 };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // Empty sheet — seed with default banks so the app has something to work with
    data.banks = [
      { id: 'b1', name: 'BDO' },
      { id: 'b2', name: 'BPI' },
      { id: 'b3', name: 'GCash' }
    ];
    return data;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var maxExpId = 0, maxFinId = 0;

  values.forEach(function (row) {
    var id = row[0], type = row[1], json = row[2];
    if (!type) return; // skip blank rows

    try {
      if (type === 'expense') {
        var e = JSON.parse(json);
        data.expenses.push(e);
        var n = parseInt(String(e.id).replace('e', ''), 10);
        if (!isNaN(n) && n > maxExpId) maxExpId = n;
      } else if (type === 'bank') {
        data.banks.push(JSON.parse(json));
      } else if (type === 'fin_allan') {
        var fa = JSON.parse(json);
        data.finEntries.allan.push(fa);
        var na = parseInt(String(fa.id).replace('f', ''), 10);
        if (!isNaN(na) && na > maxFinId) maxFinId = na;
      } else if (type === 'fin_hazel') {
        var fh = JSON.parse(json);
        data.finEntries.hazel.push(fh);
        var nh = parseInt(String(fh.id).replace('f', ''), 10);
        if (!isNaN(nh) && nh > maxFinId) maxFinId = nh;
      } else if (type === 'deleted') {
        data.deletedIds.push(String(id));
      } else if (type === 'meta') {
        var m = JSON.parse(json);
        data.noteTables = m.noteTables || {};
      }
    } catch (parseErr) {
      // one bad row shouldn't take down the whole load — skip it
      Logger.log('Skipped unreadable row: ' + JSON.stringify(row) + ' — ' + parseErr);
    }
  });

  data.finEntries.allan.sort(function (a, b) { return (b.date || '').localeCompare(a.date || ''); });
  data.finEntries.hazel.sort(function (a, b) { return (b.date || '').localeCompare(a.date || ''); });
  data.nextId = maxExpId + 1;
  data.nextFinId = maxFinId + 1;

  return data;
}

function writeAllRows(sheet, merged) {
  var rows = [];
  merged.expenses.forEach(function (e) { rows.push([e.id, 'expense', JSON.stringify(e)]); });
  merged.banks.forEach(function (b) { rows.push([b.id, 'bank', JSON.stringify(b)]); });
  (merged.finEntries.allan || []).forEach(function (f) { rows.push([f.id, 'fin_allan', JSON.stringify(f)]); });
  (merged.finEntries.hazel || []).forEach(function (f) { rows.push([f.id, 'fin_hazel', JSON.stringify(f)]); });
  merged.deletedIds.forEach(function (id) { rows.push([id, 'deleted', '']); });
  rows.push(['app', 'meta', JSON.stringify({ noteTables: merged.noteTables || {} })]);

  // Clear everything below the header, then write the full fresh set.
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  sheet.getRange(1, 5).setValue('Last saved: ' + new Date().toLocaleString());
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
  // Make sure the header row exists
  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== 'ID') {
    sheet.getRange(1, 1, 1, 3).setValues([HEADERS]);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
