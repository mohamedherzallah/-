// ==========================================
// أكواد Apps Script - مدرسة المجد النموذجية
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('مدرسة المجد النموذجية - نظام الحضور والغياب')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SYSTEM_DATA = {
  teachers: {
    "101": { name: "ميساء الجمل", grade: "الصف الأول", students: ["محمد سعيد", "أحمد سعيد", "منار سعيد"] },
    "102": { name: "نهاد عزيز", grade: "الصف الثاني", students: ["جابر عبدالله", "ميساء عبدالله", "أحمد الجبري"] },
    "103": { name: "منال حرزالله", grade: "الصف الثالث", students: ["راغب أحمد", "إيناس حرزالله", "حامد أحمد"] }
  }
};

function getFormattedDate(d) {
  if (!d) return "";
  const dateObj = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dateObj.getTime())) return String(d).trim();
  
  const year = dateObj.getFullYear();
  const month = dateObj.getMonth() + 1;
  const day = dateObj.getDate();
  return `${year}-${month}-${day}`;
}

function loginTeacher(code) {
  const teacher = SYSTEM_DATA.teachers[code];
  if (teacher) {
    const todayData = getTodayAttendance(teacher.name);
    return { success: true, teacher: teacher, todayData: todayData };
  } else {
    return { success: false, message: "كود الدخول غير صحيح، يرجى التأكد والمحاولة مجدداً." };
  }
}

function getTodayAttendance(teacherName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("تسجيل الحضور");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  const todayStr = getFormattedDate(new Date());
  let attendanceMap = {};

  for (let i = 1; i < data.length; i++) {
    const rowDate = getFormattedDate(data[i][0]);
    const rowTeacher = String(data[i][2]).trim();
    const studentName = String(data[i][4]).trim();
    const status = String(data[i][5]).trim();

    if (rowDate === todayStr && rowTeacher === teacherName.trim()) {
      attendanceMap[studentName] = status;
    }
  }
  return attendanceMap;
}

function saveAttendance(data) {
  try {
    // 🔒 1. التحقق الأمامي والأمني من الكود واسم المعلمة والصف
    const teacherAccount = SYSTEM_DATA.teachers[data.authCode];
    
    if (!teacherAccount) {
      return { success: false, message: "رمز المرور غير صحيح أو غير مصرح لك!" };
    }

    if (teacherAccount.name !== data.teacherName || teacherAccount.grade !== data.grade) {
      return { success: false, message: "غير مصرح لك بتسجيل حضور لصف أو معلمة أخرى!" };
    }

    // 2. إذا نجح التحقق، يتم كتابة البيانات في Google Sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("تسجيل الحضور");
    
    if (!sheet) {
      sheet = ss.insertSheet("تسجيل الحضور");
      sheet.appendRow(["التاريخ", "الوقت", "المعلمة", "الصف", "اسم الطالب", "الحالة"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#008A45").setFontColor("#ffffff");
    }

    const now = new Date();
    const dateStr = getFormattedDate(now);
    const timeStr = now.toLocaleTimeString('ar-EG');

    const gradeColors = {
      "الصف الأول": "#E8F5E9",
      "الصف الثاني": "#E1F5FE",
      "الصف الثالث": "#FFF8E1"
    };
    const rowColor = gradeColors[data.grade] || "#FFFFFF";

    const allData = sheet.getDataRange().getValues();

    data.records.forEach(record => {
      let isUpdated = false;
      const targetStudent = String(record.studentName).trim();
      const targetTeacher = String(data.teacherName).trim();

      for (let i = 1; i < allData.length; i++) {
        const rowDate = getFormattedDate(allData[i][0]);
        const rowTeacher = String(allData[i][2]).trim();
        const rowStudent = String(allData[i][4]).trim();

        if (rowDate === dateStr && rowTeacher === targetTeacher && rowStudent === targetStudent) {
          sheet.getRange(i + 1, 2).setValue(timeStr);
          sheet.getRange(i + 1, 6).setValue(record.status);
          sheet.getRange(i + 1, 1, 1, 6).setBackground(rowColor);
          isUpdated = true;
          break;
        }
      }

      if (!isUpdated) {
        sheet.appendRow([now, timeStr, data.teacherName, data.grade, record.studentName, record.status]);
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow, 1, 1, 6).setBackground(rowColor);
      }
    });

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
