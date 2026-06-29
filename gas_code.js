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

// ============================================================
// 🎨 دالة التنظيم الشامل — شغّلها مرة واحدة فقط
// من محرر Apps Script: شغّل الدالة organizeAllSheets
// ============================================================
function organizeAllSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();

  // ألوان مخصصة لكل ورقة
  var colors = {
    'بيانات اللاعبين':      { bg: '#7a1015', fg: '#ffffff' },
    'تجديدات الاشتراكات':   { bg: '#1e3a5f', fg: '#ffffff' },
    'سجل الحضور اليومي':   { bg: '#14532d', fg: '#ffffff' },
    'نتائج المباريات':      { bg: '#713f12', fg: '#ffffff' },
    'بيانات المدربين':      { bg: '#312e81', fg: '#ffffff' },
    'أخبار الأكاديمية':    { bg: '#134e4a', fg: '#ffffff' },
    'الفعاليات القادمة':    { bg: '#4a1d96', fg: '#ffffff' },
    'أرشيف الحضور':        { bg: '#374151', fg: '#ffffff' },
  };
  var defaultColor = { bg: '#3b0a0a', fg: '#ffffff' };

  var log = [];

  sheets.forEach(function(sheet) {
    try {
      var name = sheet.getName();
      if (name.indexOf('أرشيف') >= 0 && name !== 'أرشيف الحضور') return; // تجاهل أرشيفات أخرى

      var lastCol = sheet.getLastColumn();
      var lastRow = sheet.getLastRow();
      if (lastCol === 0) return;

      // 1. تجميد الصف الأول
      sheet.setFrozenRows(1);

      // 2. تلوين رأس الجدول
      var c = colors[name] || defaultColor;
      sheet.getRange(1, 1, 1, lastCol)
        .setBackground(c.bg)
        .setFontColor(c.fg)
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center');

      // 3. فلتر تلقائي
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      if (lastRow > 1) {
        sheet.getRange(1, 1, lastRow, lastCol).createFilter();
      }

      // 4. ضبط عرض الأعمدة تلقائياً
      for (var c2 = 1; c2 <= lastCol; c2++) {
        sheet.autoResizeColumn(c2);
        // حد أقصى 250 لعرض الأعمدة
        if (sheet.getColumnWidth(c2) > 250) sheet.setColumnWidth(c2, 250);
        if (sheet.getColumnWidth(c2) < 80)  sheet.setColumnWidth(c2, 80);
      }

      // 5. تبديل ألوان الصفوف (zebra stripes) للقراءة السهلة
      if (lastRow > 2) {
        for (var r = 2; r <= lastRow; r++) {
          var rowColor = (r % 2 === 0) ? '#ffffff' : '#f8fafc';
          sheet.getRange(r, 1, 1, lastCol).setBackground(rowColor);
        }
      }

      log.push('✅ ' + name);
    } catch(err) {
      log.push('⚠️ ' + sheet.getName() + ': ' + err.message);
    }
  });

  // إنشاء ورقة فهرس في البداية
  createIndexSheet(ss);
  log.push('✅ ورقة الفهرس');

  Logger.log('تم التنظيم:\n' + log.join('\n'));
  SpreadsheetApp.getUi().alert('✅ تم التنظيم بنجاح!\n\n' + log.join('\n'));
}

// ============================================================
// 📑 ورقة فهرس بروابط لكل الأوراق
// ============================================================
function createIndexSheet(ss) {
  var existing = ss.getSheetByName('📋 الفهرس');
  if (existing) ss.deleteSheet(existing);

  var index = ss.insertSheet('📋 الفهرس', 0);
  index.setTabColor('#d4af37');

  var sheets = ss.getSheets().filter(function(s){ return s.getName() !== '📋 الفهرس'; });

  // هيدر
  index.getRange('A1:C1').merge()
    .setValue('📋 فهرس أوراق أكاديمية الشمالي')
    .setBackground('#7a1015')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');

  index.getRange('A2:C2').setValues([['الورقة', 'عدد الصفوف', 'آخر تحديث']])
    .setBackground('#f1f5f9')
    .setFontWeight('bold');

  sheets.forEach(function(sheet, i) {
    var row = i + 3;
    var lastRow = Math.max(0, sheet.getLastRow() - 1);
    var name = sheet.getName();
    var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID +
              '/edit#gid=' + sheet.getSheetId();

    index.getRange(row, 1).setFormula('=HYPERLINK("' + url + '","📄 ' + name + '")')
      .setFontColor('#1e3a5f').setFontWeight('bold');
    index.getRange(row, 2).setValue(lastRow + ' صف')
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
// 🗄️ أرشفة الحضور القديم — شغّلها يدوياً أو تلقائياً
// تنقل بيانات الحضور الأقدم من 3 أشهر إلى ورقة الأرشيف
// ============================================================
function archiveOldAttendance() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var source = ss.getSheetByName('سجل الحضور اليومي');
  if (!source) { Logger.log('الورقة غير موجودة'); return; }

  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 3); // 3 أشهر للخلف

  var data = source.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0];
  var dateCol = headers.indexOf('التاريخ');
  if (dateCol < 0) dateCol = 0; // العمود الأول افتراضياً

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
    Logger.log('لا توجد بيانات قديمة للأرشفة');
    SpreadsheetApp.getUi().alert('لا توجد بيانات أقدم من 3 أشهر للأرشفة');
    return;
  }

  // ورقة الأرشيف
  var archiveSheet = ss.getSheetByName('أرشيف الحضور');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('أرشيف الحضور');
    archiveSheet.setTabColor('#374151');
    archiveSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#374151')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  // أضف للأرشيف
  var archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, archive.length, headers.length)
    .setValues(archive);

  // احتفظ بالبيانات الحديثة فقط
  source.clearContents();
  source.getRange(1, 1, keep.length, headers.length).setValues(keep);

  // أعد تطبيق التنسيق
  source.getRange(1, 1, 1, headers.length)
    .setBackground('#14532d').setFontColor('#ffffff').setFontWeight('bold');
  source.setFrozenRows(1);

  var msg = '✅ تم أرشفة ' + archive.length + ' صف\n' +
            '📋 تبقّى ' + (keep.length - 1) + ' صف في الورقة الرئيسية\n' +
            '🗄️ الأرشيف: ورقة "أرشيف الحضور"';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// ⏰ تشغيل الأرشفة تلقائياً كل أول يوم في الشهر
// شغّل هذه الدالة مرة واحدة لإعداد التشغيل التلقائي
// ============================================================
function setupMonthlyArchiveTrigger() {
  // احذف أي trigger قديم
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'archiveOldAttendance') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // أنشئ trigger جديد: أول يوم كل شهر الساعة 3 صباحاً
  ScriptApp.newTrigger('archiveOldAttendance')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
  SpreadsheetApp.getUi().alert('✅ تم إعداد الأرشفة التلقائية كل أول يوم في الشهر');
}
