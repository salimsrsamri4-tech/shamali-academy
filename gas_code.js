// ============================================================
// أكاديمية الشمالي - Google Apps Script API
// استبدل هذا الكود في محرر Apps Script ثم انشره من جديد
// ============================================================

var SHEET_ID = '16Qc1DUjhpbT-kIra74YeFQxzVxvBY7tSquUrHdacZKI';

// ==================== GET (قراءة البيانات) ====================
function doGet(e) {
  try {
    var action = (e.parameter.action || 'get');

    if (action === 'get') {
      var sheetName = e.parameter.sheet || 'الورقة1';
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

// ==================== POST (كتابة وحذف) ====================
function doPost(e) {
  try {
    // رفع ملفات (الوظيفة القديمة)
    if (e.parameter.action === 'upload' || e.parameter.type === 'upload') {
      return handleUpload(e);
    }

    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'append';
    var sheetName = body.sheet || e.parameter.sheet || 'الورقة1';
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // ── إضافة صفوف ──
    if (action === 'append') {
      var sheet = getOrCreateSheet(ss, sheetName);
      var rows = Array.isArray(body.data) ? body.data : [body.data];
      var headers = getHeaders(sheet);

      // إنشاء الأعمدة تلقائياً إذا لم توجد
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

    // ── حذف صفوف ──
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

    // ── رفع ملف (عبر POST أيضاً) ──
    if (action === 'upload') {
      return handleUpload(e);
    }

    return ok({ error: 'unknown action: ' + action });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ==================== مساعدات ====================
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

// ==================== رفع الملفات (الكود القديم) ====================
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
