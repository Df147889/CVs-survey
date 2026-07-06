/**
 * ==========================================================================
 * Impetus - CV Data Form  |  Google Apps Script Backend
 * استمارة بيانات السيرة الذاتية - كود الربط مع Google Sheets + Drive
 * ==========================================================================
 * ماذا يفعل هذا الكود؟
 *  1. يستقبل بيانات النموذج (JSON) من صفحة index.html
 *  2. يحفظ ملف السيرة الذاتية في مجلد داخل Google Drive ويجلب رابطه
 *  3. يكتب صفاً واحداً لكل موظف في الشيت، وكل مشاريعه في خانة واحدة
 *
 * طريقة التركيب موجودة في نهاية هذا الملف (خطوات النشر / Deploy).
 * ==========================================================================
 */

// اسم مجلد الدرايف الذي ستُحفظ فيه ملفات السير الذاتية
var DRIVE_FOLDER_NAME = 'CV_Attachments';

// اسم مجلد الدرايف الذي ستُحفظ فيه الصور الشخصية
var PHOTO_FOLDER_NAME = 'Personal_Photos';

// اسم الورقة (التاب) داخل الشيت
var SHEET_NAME = 'Responses';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000); // يمنع تضارب الكتابة عند إرسال عدة نماذج بنفس اللحظة

  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = getSheet_();

    var safeName = (data.fullName || 'Employee').replace(/[\\/:*?"<>|]/g, ' ');

    // ---- 1) حفظ الصورة الشخصية في Drive (إن وُجدت) ----
    var photoLink = '';
    if (data.photoFile && data.photoFile.base64) {
      var pFolder = getOrCreateFolder_(PHOTO_FOLDER_NAME);
      var pBytes = Utilities.base64Decode(data.photoFile.base64);
      var pBlob = Utilities.newBlob(pBytes, data.photoFile.mimeType, data.photoFile.fileName);
      var pFile = pFolder.createFile(pBlob);
      pFile.setName(safeName + ' - ' + data.photoFile.fileName);
      photoLink = pFile.getUrl();
    }

    // ---- 2) حفظ ملف السيرة الذاتية في Drive (إن وُجد) ----
    var cvLink = '';
    if (data.cvFile && data.cvFile.base64) {
      var folder = getOrCreateFolder_(DRIVE_FOLDER_NAME);
      var bytes = Utilities.base64Decode(data.cvFile.base64);
      var blob = Utilities.newBlob(bytes, data.cvFile.mimeType, data.cvFile.fileName);
      var file = folder.createFile(blob);
      file.setName(safeName + ' - ' + data.cvFile.fileName);
      cvLink = file.getUrl();
    }

    // ---- 3) تجهيز المشاريع في خانة واحدة منسّقة ----
    var projectsText = '';
    if (data.projects && data.projects.length) {
      projectsText = data.projects.map(function (p, i) {
        return 'مشروع ' + (i + 1) + ':\n' +
               '• اسم المشروع: ' + (p.projectName || '') + '\n' +
               '• الدور: ' + (p.role || '') + '\n' +
               '• وصف الدور: ' + (p.roleDescription || '') + '\n' +
               '• الجهة/العميل: ' + (p.client || '');
      }).join('\n———————————————\n');
    }

    // ---- 4) كتابة الصف في الشيت ----
    sheet.appendRow([
      new Date(),                               // Timestamp
      data.fullName || '',                      // الاسم واللقب
      data.jobTitle || '',                      // المسمى الوظيفي
      data.yearsExperience || '',               // سنوات الخبرة
      data.qualification || '',                 // المؤهل
      data.university || '',                     // الجامعة
      data.specialization || '',                // التخصص
      data.certifications || '',                // الشهادات المهنية
      (data.languages || []).join(', '),        // اللغات
      data.bio || '',                           // النبذة التعريفية
      projectsText,                             // الخبرات العملية (المشاريع)
      data.previousExperience || '',            // الخبرات السابقة
      photoLink,                                // رابط الصورة الشخصية
      cvLink                                    // رابط ملف السيرة الذاتية
    ]);

    return json_({ result: 'success' });

  } catch (err) {
    return json_({ result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** يفتح ورقة الردود وينشئها مع صف العناوين إن لم تكن موجودة */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'التاريخ', 'الاسم واللقب', 'المسمى الوظيفي', 'سنوات الخبرة',
      'المؤهل', 'الجامعة', 'التخصص', 'الشهادات المهنية', 'اللغات',
      'النبذة التعريفية', 'الخبرات العملية (المشاريع)', 'الخبرات السابقة',
      'رابط الصورة الشخصية', 'رابط السيرة الذاتية'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** يجلب مجلد الدرايف أو ينشئه إن لم يكن موجوداً */
function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/** رد JSON قياسي */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================================
   خطوات التركيب والنشر:
   1. افتح Google Sheet جديد (sheets.new)
   2. من القائمة: Extensions ▸ Apps Script
   3. احذف الكود الموجود والصق هذا الملف بالكامل
   4. اضغط Save (أيقونة القرص)
   5. Deploy ▸ New deployment ▸ اختر النوع: Web app
        - Description: CV Form
        - Execute as: Me (بريدك)
        - Who has access: Anyone
   6. اضغط Deploy ثم وافق على الصلاحيات (Authorize access)
   7. انسخ رابط "Web app URL" (ينتهي بـ /exec)
   8. الصقه في ملف index.html مكان: PASTE_YOUR_SCRIPT_URL_HERE
   ملاحظة: كل ما عدّلت الكود لاحقاً، اعمل Deploy ▸ Manage deployments ▸ Edit ▸
           Version: New version حتى تُحدَّث النسخة المنشورة.
   ========================================================================== */
