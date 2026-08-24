/**
 * منصة حقول التشغيلية - خادم Google Apps Script المحدث
 * يتيح عرض الواجهة مباشرة كـ Web App واستقبال البيانات وحفظها في الأوراق.
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('منصة حقول التشغيلية - لواء حيفا')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Save / Update Schools Dashboard Sheet
    if (data.schools && Array.isArray(data.schools)) {
      var sheetSchools = getOrCreateSheet(ss, 'المدارس (الداشبورد)');
      if (sheetSchools.getLastRow() > 1) {
        sheetSchools.getRange(2, 1, sheetSchools.getLastRow() - 1, sheetSchools.getLastColumn()).clearContent();
      }
      if (sheetSchools.getLastRow() === 0) {
        sheetSchools.appendRow(['اسم المدرسة', 'اسم المرشد/ة', 'حالة الخطة', 'وحدات مطبقة', 'مشاهدات صفية', 'تعلم مهني', 'تقييم بديل', 'تقدم الطلاب', 'التحدي', 'الأولوية', 'تاريخ التحديث']);
      }
      data.schools.forEach(function(s) {
        sheetSchools.appendRow([
          s.name,
          s.mentor,
          s.status,
          s.units,
          s.observations,
          s.pd ? 'نعم' : 'لا',
          s.altEval ? 'نعم' : 'لا',
          s.studentProgress,
          s.challenge,
          s.priority,
          new Date()
        ]);
      });
    }

    // 2. Save Visit Protocol Log
    if (data.visitForm) {
      var sheetVisits = getOrCreateSheet(ss, 'سجلات الزيارات');
      if (sheetVisits.getLastRow() === 0) {
        sheetVisits.appendRow(['اسم المدرسة', 'المرشد والتاريخ', 'مهمة للأسبوع', 'المسؤول', 'الفئة', 'موعد التنفيذ', 'الدليل', 'سؤال انعكاسي', 'مؤشر النجاح', 'تاريخ التسجيل']);
      }
      var vf = data.visitForm;
      var mand = vf.mandatory || {};
      sheetVisits.appendRow([
        vf.schoolName,
        vf.mentorDate,
        mand.task,
        mand.responsible,
        mand.targetAudience,
        mand.date,
        mand.evidence,
        mand.reflectiveQ,
        mand.successIndicator,
        new Date()
      ]);
    }

    // 3. Save Absence Protocol Log
    if (data.absenceForm) {
      var sheetAbsence = getOrCreateSheet(ss, 'أسبوع غياب المرشد');
      if (sheetAbsence.getLastRow() === 0) {
        sheetAbsence.appendRow(['المهمة السابقة', 'المسؤول', 'الفئة', 'الموعد', 'نوع الدليل', 'تفاصيل الدليل', 'الانعكاس', 'سؤال للقاء القادم', 'تاريخ التسجيل']);
      }
      var af = data.absenceForm;
      sheetAbsence.appendRow([
        af.prevTask,
        af.responsible,
        af.target,
        af.date,
        af.evidenceType,
        af.evidenceDetails,
        af.reflection,
        af.nextQuestion,
        new Date()
      ]);
    }

    // 4. Save Student Participation Map
    if (data.studentsMap && Array.isArray(data.studentsMap)) {
      var sheetStudents = getOrCreateSheet(ss, 'خريطة مشاركة الطلاب');
      if (sheetStudents.getLastRow() > 1) {
        sheetStudents.getRange(2, 1, sheetStudents.getLastRow() - 1, sheetStudents.getLastColumn()).clearContent();
      }
      if (sheetStudents.getLastRow() === 0) {
        sheetStudents.appendRow(['اسم/رمز الطالب', 'خط الأساس', 'الفرصة المخططة', 'الدور المتاح', 'ما حدث فعليًا', 'المستوى التالي', 'تاريخ التحديث']);
      }
      data.studentsMap.forEach(function(st) {
        sheetStudents.appendRow([
          st.name,
          st.baseline,
          st.plannedOpportunity,
          st.role,
          st.actual,
          st.nextLevel,
          new Date()
        ]);
      });
    }

    // 5. Save Monthly Report
    if (data.monthlyReport) {
      var sheetReports = getOrCreateSheet(ss, 'التقارير الشهرية');
      if (sheetReports.getLastRow() === 0) {
        sheetReports.appendRow(['اسم المدرسة', 'ما نُفذ', 'الوحدة', 'الدليل المختار', 'وضع الطلاب الأقل ظهورًا', 'التحدي', 'ممارسة ناجحة', 'أولوية الشهر القادم', 'تاريخ الإرسال']);
      }
      var mr = data.monthlyReport;
      sheetReports.appendRow([
        mr.school,
        mr.executed,
        mr.unit,
        mr.evidence,
        mr.studentsStatus,
        mr.challenge,
        mr.successPractice,
        mr.nextPriority,
        new Date()
      ]);
    }

    // 6. Save Checklist Status
    if (data.checklist && Array.isArray(data.checklist)) {
      var sheetChecklist = getOrCreateSheet(ss, 'قائمة الفحص');
      if (sheetChecklist.getLastRow() > 1) {
        sheetChecklist.getRange(2, 1, sheetChecklist.getLastRow() - 1, sheetChecklist.getLastColumn()).clearContent();
      }
      if (sheetChecklist.getLastRow() === 0) {
        sheetChecklist.appendRow(['البند', 'الحالة (منجز / غير منجز)', 'تاريخ التحديث']);
      }
      data.checklist.forEach(function(item) {
        sheetChecklist.appendRow([
          item.text,
          item.checked ? 'منجز' : 'قيد العمل',
          new Date()
        ]);
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'تم حفظ البيانات بنجاح في Google Sheets' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper function to retrieve an existing sheet by name or create a new one.
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}
