// ============================================================
// Ø£ÙƒØ§Ø¯ÙŠÙ…ÙŠØ© Ø§Ù„Ø´Ù…Ø§Ù„ÙŠ - Google Apps Script API
// Ø§Ø³ØªØ¨Ø¯Ù„ Ù‡Ø°Ø§ Ø§Ù„ÙƒÙˆØ¯ ÙÙŠ Ù…Ø­Ø±Ø± Apps Script Ø«Ù… Ø§Ù†Ø´Ø±Ù‡ Ù…Ù† Ø¬Ø¯ÙŠØ¯
// ============================================================

var SHEET_ID = '16Qc1DUjhpbT-kIra74YeFQxzVxvBY7tSquUrHdacZKI';

// ==================== GET (Ù‚Ø±Ø§Ø¡Ø© Ø§Ù„Ø¨ÙŠØ§Ù†Ø§Øª) ====================
function doGet(e) {
  try {
    var action = (e.parameter.action || 'get');

    if (action === 'get') {
      var sheetName = e.parameter.sheet || 'Ø§Ù„ÙˆØ±Ù‚Ø©1';
      var limit = parseInt(e.parameter.limit) || 500;
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return ok([]);

      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) return ok([]);

      var headers = data[0].map(function(h){ return h.toString().trim(); });
      var rows = [];
      for (var i = 1; i < data.length && rows.length < limit; i++) {
        var row = {};
        var empty = true;
        headers.forEach(function(h, j) {
          var val = data[i][j] !== undefined && data[i][j] !== null ? data[i][j].toString() : '';
          row[h] = val;
          if (val !== '') empty = false;
        });
        if (!empty) rows.push(row);
      }
      return ok(rows);
    }

    return ok({ error: 'unknown action' });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ==================== POST (ÙƒØªØ§Ø¨Ø© ÙˆØ­Ø°Ù) ====================
function doPost(e) {
  try {
    // Ø±ÙØ¹ Ù…Ù„ÙØ§Øª (Ø§Ù„ÙˆØ¸ÙŠÙØ© Ø§Ù„Ù‚Ø¯ÙŠÙ…Ø©)
    if (e.parameter.action === 'upload' || e.parameter.type === 'upload') {
      return handleUpload(e);
    }

    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'append';
    var sheetName = body.sheet || e.parameter.sheet || 'Ø§Ù„ÙˆØ±Ù‚Ø©1';
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // â”€â”€ Ø¥Ø¶Ø§ÙØ© ØµÙÙˆÙ â”€â”€
    if (action === 'append') {
      var sheet = getOrCreateSheet(ss, sheetName);
      var rows = Array.isArray(body.data) ? body.data : [body.data];
      var headers = getHeaders(sheet);

      // Ø¥Ù†Ø´Ø§Ø¡ Ø§Ù„Ø£Ø¹Ù…Ø¯Ø© ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹ Ø¥Ø°Ø§ Ù„Ù… ØªÙˆØ¬Ø¯
      if (headers.length === 0) {
        headers = Object.keys(rows[0]);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }

      rows.forEach(function(row) {
        var values = headers.map(function(h) {
          var v = row[h];
          return v !== undefined && v !== null ? v.toString() : '';
        });
        sheet.appendRow(values);
      });
      return ok({ status: 'ok', added: rows.length });
    }

    // â”€â”€ Ø­Ø°Ù ØµÙÙˆÙ â”€â”€
    if (action === 'delete') {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return ok({ status: 'ok', deleted: 0 });
      var col = body.column;
      var val = body.value ? body.value.toString() : '';
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h){ return h.toString().trim(); });
      var colIdx = headers.indexOf(col);
      if (colIdx < 0) return ok({ status: 'ok', deleted: 0 });
      var deleted = 0;
      for (var i = data.length - 1; i >= 1; i--) {
        if (data[i][colIdx].toString().trim() === val) {
          sheet.deleteRow(i + 1);
          deleted++;
        }
      }
      return ok({ status: 'ok', deleted: deleted });
    }

    // â”€â”€ Ø±ÙØ¹ Ù…Ù„Ù (Ø¹Ø¨Ø± POST Ø£ÙŠØ¶Ø§Ù‹) â”€â”€
    if (action === 'upload') {
      return handleUpload(e);
    }

    return ok({ error: 'unknown action: ' + action });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ==================== Ù…Ø³Ø§Ø¹Ø¯Ø§Øª ====================
function ok(data) {
  var out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getHeaders(sheet) {
  var last = sheet.getLastColumn();
  if (last === 0) return [];
  return sheet.getRange(1, 1, 1, last).getValues()[0].map(function(h){ return h.toString().trim(); });
}

// ==================== Ø±ÙØ¹ Ø§Ù„Ù…Ù„ÙØ§Øª (Ø§Ù„ÙƒÙˆØ¯ Ø§Ù„Ù‚Ø¯ÙŠÙ…) ====================
function handleUpload(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var folder = DriveApp.getFolderById('1oZ1DCyEDg-oXSBF9RH8bkOKvmOiqL7_y');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.file),
      data.mimeType,
      data.filename
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return ok({ id: file.getId(), url: file.getUrl() });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ============================================================
// ðŸŽ¨ Ø¯Ø§Ù„Ø© Ø§Ù„ØªÙ†Ø¸ÙŠÙ… Ø§Ù„Ø´Ø§Ù…Ù„ â€” Ø´ØºÙ‘Ù„Ù‡Ø§ Ù…Ø±Ø© ÙˆØ§Ø­Ø¯Ø© ÙÙ‚Ø·
// Ù…Ù† Ù…Ø­Ø±Ø± Apps Script: Ø´ØºÙ‘Ù„ Ø§Ù„Ø¯Ø§Ù„Ø© organizeAllSheets
// ============================================================
function organizeAllSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();

  // Ø£Ù„ÙˆØ§Ù† Ù…Ø®ØµØµØ© Ù„ÙƒÙ„ ÙˆØ±Ù‚Ø©
  var colors = {
    'Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„Ù„Ø§Ø¹Ø¨ÙŠÙ†':      { bg: '#7a1015', fg: '#ffffff' },
    'ØªØ¬Ø¯ÙŠØ¯Ø§Øª Ø§Ù„Ø§Ø´ØªØ±Ø§ÙƒØ§Øª':   { bg: '#1e3a5f', fg: '#ffffff' },
    'Ø³Ø¬Ù„ Ø§Ù„Ø­Ø¶ÙˆØ± Ø§Ù„ÙŠÙˆÙ…ÙŠ':   { bg: '#14532d', fg: '#ffffff' },
    'Ù†ØªØ§Ø¦Ø¬ Ø§Ù„Ù…Ø¨Ø§Ø±ÙŠØ§Øª':      { bg: '#713f12', fg: '#ffffff' },
    'Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„Ù…Ø¯Ø±Ø¨ÙŠÙ†':      { bg: '#312e81', fg: '#ffffff' },
    'Ø£Ø®Ø¨Ø§Ø± Ø§Ù„Ø£ÙƒØ§Ø¯ÙŠÙ…ÙŠØ©':    { bg: '#134e4a', fg: '#ffffff' },
    'Ø§Ù„ÙØ¹Ø§Ù„ÙŠØ§Øª Ø§Ù„Ù‚Ø§Ø¯Ù…Ø©':    { bg: '#4a1d96', fg: '#ffffff' },
    'Ø£Ø±Ø´ÙŠÙ Ø§Ù„Ø­Ø¶ÙˆØ±':        { bg: '#374151', fg: '#ffffff' },
  };
  var defaultColor = { bg: '#3b0a0a', fg: '#ffffff' };

  var log = [];

  sheets.forEach(function(sheet) {
    try {
      var name = sheet.getName();
      if (name.indexOf('Ø£Ø±Ø´ÙŠÙ') >= 0 && name !== 'Ø£Ø±Ø´ÙŠÙ Ø§Ù„Ø­Ø¶ÙˆØ±') return; // ØªØ¬Ø§Ù‡Ù„ Ø£Ø±Ø´ÙŠÙØ§Øª Ø£Ø®Ø±Ù‰

      var lastCol = sheet.getLastColumn();
      var lastRow = sheet.getLastRow();
      if (lastCol === 0) return;

      // 1. ØªØ¬Ù…ÙŠØ¯ Ø§Ù„ØµÙ Ø§Ù„Ø£ÙˆÙ„
      sheet.setFrozenRows(1);

      // 2. ØªÙ„ÙˆÙŠÙ† Ø±Ø£Ø³ Ø§Ù„Ø¬Ø¯ÙˆÙ„
      var c = colors[name] || defaultColor;
      sheet.getRange(1, 1, 1, lastCol)
        .setBackground(c.bg)
        .setFontColor(c.fg)
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center');

      // 3. ÙÙ„ØªØ± ØªÙ„Ù‚Ø§Ø¦ÙŠ
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      if (lastRow > 1) {
        sheet.getRange(1, 1, lastRow, lastCol).createFilter();
      }

      // 4. Ø¶Ø¨Ø· Ø¹Ø±Ø¶ Ø§Ù„Ø£Ø¹Ù…Ø¯Ø© ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹
      for (var c2 = 1; c2 <= lastCol; c2++) {
        sheet.autoResizeColumn(c2);
        // Ø­Ø¯ Ø£Ù‚ØµÙ‰ 250 Ù„Ø¹Ø±Ø¶ Ø§Ù„Ø£Ø¹Ù…Ø¯Ø©
        if (sheet.getColumnWidth(c2) > 250) sheet.setColumnWidth(c2, 250);
        if (sheet.getColumnWidth(c2) < 80)  sheet.setColumnWidth(c2, 80);
      }

      // 5. ØªØ¨Ø¯ÙŠÙ„ Ø£Ù„ÙˆØ§Ù† Ø§Ù„ØµÙÙˆÙ (zebra stripes) Ù„Ù„Ù‚Ø±Ø§Ø¡Ø© Ø§Ù„Ø³Ù‡Ù„Ø©
      if (lastRow > 2) {
        for (var r = 2; r <= lastRow; r++) {
          var rowColor = (r % 2 === 0) ? '#ffffff' : '#f8fafc';
          sheet.getRange(r, 1, 1, lastCol).setBackground(rowColor);
        }
      }

      log.push('âœ… ' + name);
    } catch(err) {
      log.push('âš ï¸ ' + sheet.getName() + ': ' + err.message);
    }
  });

  // Ø¥Ù†Ø´Ø§Ø¡ ÙˆØ±Ù‚Ø© ÙÙ‡Ø±Ø³ ÙÙŠ Ø§Ù„Ø¨Ø¯Ø§ÙŠØ©
  createIndexSheet(ss);
  log.push('âœ… ÙˆØ±Ù‚Ø© Ø§Ù„ÙÙ‡Ø±Ø³');

  Logger.log('ØªÙ… Ø§Ù„ØªÙ†Ø¸ÙŠÙ…:\n' + log.join('\n'));
  Logger.log('âœ… ØªÙ… Ø§Ù„ØªÙ†Ø¸ÙŠÙ… Ø¨Ù†Ø¬Ø§Ø­!\n\n' + log.join('\n'));
}

// ============================================================
// ðŸ“‘ ÙˆØ±Ù‚Ø© ÙÙ‡Ø±Ø³ Ø¨Ø±ÙˆØ§Ø¨Ø· Ù„ÙƒÙ„ Ø§Ù„Ø£ÙˆØ±Ø§Ù‚
// ============================================================
function createIndexSheet(ss) {
  var existing = ss.getSheetByName('ðŸ“‹ Ø§Ù„ÙÙ‡Ø±Ø³');
  if (existing) ss.deleteSheet(existing);

  var index = ss.insertSheet('ðŸ“‹ Ø§Ù„ÙÙ‡Ø±Ø³', 0);
  index.setTabColor('#d4af37');

  var sheets = ss.getSheets().filter(function(s){ return s.getName() !== 'ðŸ“‹ Ø§Ù„ÙÙ‡Ø±Ø³'; });

  // Ù‡ÙŠØ¯Ø±
  index.getRange('A1:C1').merge()
    .setValue('ðŸ“‹ ÙÙ‡Ø±Ø³ Ø£ÙˆØ±Ø§Ù‚ Ø£ÙƒØ§Ø¯ÙŠÙ…ÙŠØ© Ø§Ù„Ø´Ù…Ø§Ù„ÙŠ')
    .setBackground('#7a1015')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');

  index.getRange('A2:C2').setValues([['Ø§Ù„ÙˆØ±Ù‚Ø©', 'Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ', 'Ø¢Ø®Ø± ØªØ­Ø¯ÙŠØ«']])
    .setBackground('#f1f5f9')
    .setFontWeight('bold');

  sheets.forEach(function(sheet, i) {
    var row = i + 3;
    var lastRow = Math.max(0, sheet.getLastRow() - 1);
    var name = sheet.getName();
    var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID +
              '/edit#gid=' + sheet.getSheetId();

    index.getRange(row, 1).setFormula('=HYPERLINK("' + url + '","ðŸ“„ ' + name + '")')
      .setFontColor('#1e3a5f').setFontWeight('bold');
    index.getRange(row, 2).setValue(lastRow + ' ØµÙ')
      .setHorizontalAlignment('center');
    index.getRange(row, 3).setValue(new Date())
      .setNumberFormat('dd/MM/yyyy HH:mm')
      .setHorizontalAlignment('center');

    if (i % 2 === 0) index.getRange(row, 1, 1, 3).setBackground('#f8fafc');
  });

  index.setFrozenRows(2);
  index.setColumnWidth(1, 220);
  index.setColumnWidth(2, 100);
  index.setColumnWidth(3, 150);
}

// ============================================================
// ðŸ—„ï¸ Ø£Ø±Ø´ÙØ© Ø§Ù„Ø­Ø¶ÙˆØ± Ø§Ù„Ù‚Ø¯ÙŠÙ… â€” Ø´ØºÙ‘Ù„Ù‡Ø§ ÙŠØ¯ÙˆÙŠØ§Ù‹ Ø£Ùˆ ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹
// ØªÙ†Ù‚Ù„ Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„Ø­Ø¶ÙˆØ± Ø§Ù„Ø£Ù‚Ø¯Ù… Ù…Ù† 3 Ø£Ø´Ù‡Ø± Ø¥Ù„Ù‰ ÙˆØ±Ù‚Ø© Ø§Ù„Ø£Ø±Ø´ÙŠÙ
// ============================================================
function archiveOldAttendance() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var source = ss.getSheetByName('Ø³Ø¬Ù„ Ø§Ù„Ø­Ø¶ÙˆØ± Ø§Ù„ÙŠÙˆÙ…ÙŠ');
  if (!source) { Logger.log('Ø§Ù„ÙˆØ±Ù‚Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©'); return; }

  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 3); // 3 Ø£Ø´Ù‡Ø± Ù„Ù„Ø®Ù„Ù

  var data = source.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0];
  var dateCol = headers.indexOf('Ø§Ù„ØªØ§Ø±ÙŠØ®');
  if (dateCol < 0) dateCol = 0; // Ø§Ù„Ø¹Ù…ÙˆØ¯ Ø§Ù„Ø£ÙˆÙ„ Ø§ÙØªØ±Ø§Ø¶ÙŠØ§Ù‹

  var keep = [headers];
  var archive = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[dateCol];
    var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (!isNaN(d.getTime()) && d < cutoff) {
      archive.push(row);
    } else {
      keep.push(row);
    }
  }

  if (archive.length === 0) {
    Logger.log('Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨ÙŠØ§Ù†Ø§Øª Ù‚Ø¯ÙŠÙ…Ø© Ù„Ù„Ø£Ø±Ø´ÙØ©');
    Logger.log('Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨ÙŠØ§Ù†Ø§Øª Ø£Ù‚Ø¯Ù… Ù…Ù† 3 Ø£Ø´Ù‡Ø± Ù„Ù„Ø£Ø±Ø´ÙØ©');
    return;
  }

  // ÙˆØ±Ù‚Ø© Ø§Ù„Ø£Ø±Ø´ÙŠÙ
  var archiveSheet = ss.getSheetByName('Ø£Ø±Ø´ÙŠÙ Ø§Ù„Ø­Ø¶ÙˆØ±');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Ø£Ø±Ø´ÙŠÙ Ø§Ù„Ø­Ø¶ÙˆØ±');
    archiveSheet.setTabColor('#374151');
    archiveSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#374151')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  // Ø£Ø¶Ù Ù„Ù„Ø£Ø±Ø´ÙŠÙ
  var archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, archive.length, headers.length)
    .setValues(archive);

  // Ø§Ø­ØªÙØ¸ Ø¨Ø§Ù„Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„Ø­Ø¯ÙŠØ«Ø© ÙÙ‚Ø·
  source.clearContents();
  source.getRange(1, 1, keep.length, headers.length).setValues(keep);

  // Ø£Ø¹Ø¯ ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ØªÙ†Ø³ÙŠÙ‚
  source.getRange(1, 1, 1, headers.length)
    .setBackground('#14532d').setFontColor('#ffffff').setFontWeight('bold');
  source.setFrozenRows(1);

  var msg = 'âœ… ØªÙ… Ø£Ø±Ø´ÙØ© ' + archive.length + ' ØµÙ\n' +
            'ðŸ“‹ ØªØ¨Ù‚Ù‘Ù‰ ' + (keep.length - 1) + ' ØµÙ ÙÙŠ Ø§Ù„ÙˆØ±Ù‚Ø© Ø§Ù„Ø±Ø¦ÙŠØ³ÙŠØ©\n' +
            'ðŸ—„ï¸ Ø§Ù„Ø£Ø±Ø´ÙŠÙ: ÙˆØ±Ù‚Ø© "Ø£Ø±Ø´ÙŠÙ Ø§Ù„Ø­Ø¶ÙˆØ±"';
  Logger.log(msg);
  Logger.log(msg);
}

// ============================================================
// â° ØªØ´ØºÙŠÙ„ Ø§Ù„Ø£Ø±Ø´ÙØ© ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹ ÙƒÙ„ Ø£ÙˆÙ„ ÙŠÙˆÙ… ÙÙŠ Ø§Ù„Ø´Ù‡Ø±
// Ø´ØºÙ‘Ù„ Ù‡Ø°Ù‡ Ø§Ù„Ø¯Ø§Ù„Ø© Ù…Ø±Ø© ÙˆØ§Ø­Ø¯Ø© Ù„Ø¥Ø¹Ø¯Ø§Ø¯ Ø§Ù„ØªØ´ØºÙŠÙ„ Ø§Ù„ØªÙ„Ù‚Ø§Ø¦ÙŠ
// ============================================================
function setupMonthlyArchiveTrigger() {
  // Ø§Ø­Ø°Ù Ø£ÙŠ trigger Ù‚Ø¯ÙŠÙ…
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'archiveOldAttendance') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Ø£Ù†Ø´Ø¦ trigger Ø¬Ø¯ÙŠØ¯: Ø£ÙˆÙ„ ÙŠÙˆÙ… ÙƒÙ„ Ø´Ù‡Ø± Ø§Ù„Ø³Ø§Ø¹Ø© 3 ØµØ¨Ø§Ø­Ø§Ù‹
  ScriptApp.newTrigger('archiveOldAttendance')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
  Logger.log('âœ… ØªÙ… Ø¥Ø¹Ø¯Ø§Ø¯ Ø§Ù„Ø£Ø±Ø´ÙØ© Ø§Ù„ØªÙ„Ù‚Ø§Ø¦ÙŠØ© ÙƒÙ„ Ø£ÙˆÙ„ ÙŠÙˆÙ… ÙÙŠ Ø§Ù„Ø´Ù‡Ø±');
}
