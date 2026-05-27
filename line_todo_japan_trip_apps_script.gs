const SPREADSHEET_ID = "1IMfv-j08A_SF2DuDSX3j3Zp9dcyTtb0wtFOsgCEv9Wc";
const STATUS_SHEET_NAME = "todo_status";
const TASKS_SHEET_NAME = "trip_tasks";

const TASKS = [
  ["t1", "Wed", "พุธ 27 พ.ค.", "ทั้งวัน", "", "move", "เดินทาง กรุงเทพฯ → สนามบินฮาเนดะ", "เช็กอินที่ Prince Hotel"],
  ["t2", "Thu", "พฤหัส 28 พ.ค.", "06:45", "09:30", "move", "เดินทางไปโตเกียว / ชินโยโกฮามะ", "พบที่ล็อบบี้ 4 ชั้น 1 โรงแรม Shin-Yokohama Prince Hotel"],
  ["t3", "Thu", "พฤหัส 28 พ.ค.", "07:06", "07:24", "move", "JR Shinkansen Kodama No. 902", "Shin-Yokohama → Tokyo"],
  ["t4", "Thu", "พฤหัส 28 พ.ค.", "07:49", "08:28", "move", "JR Shinkansen Yamabiko No. 125", "Tokyo → Utsunomiya"],
  ["t5", "Thu", "พฤหัส 28 พ.ค.", "08:54", "09:00", "move", "JR Shonan-Shinjuku Line", "Utsunomiya → Suzumenomiya"],
  ["t6", "Thu", "พฤหัส 28 พ.ค.", "10:00", "", "move", "พบที่ UTPC / Transfer", "เดินทางจากโรงแรมไป UTPC"],
  ["t7", "Thu", "พฤหัส 28 พ.ค.", "10:30", "11:30", "meet", "ทัวร์กระบวนการผลิต DS-5000 และ FT-2500", "Group OSTH และรอบร่วมกับ Nichizu"],
  ["t8", "Thu", "พฤหัส 28 พ.ค.", "12:00", "13:00", "food", "อาหารกลางวัน", "Lunch"],
  ["t9", "Thu", "พฤหัส 28 พ.ค.", "13:00", "14:30", "meet", "ทัวร์พื้นที่", "Area Tour"],
  ["t10", "Thu", "พฤหัส 28 พ.ค.", "14:30", "16:00", "meet", "ประชุม Kumakuro กับ Nichizu", "หัวข้อ: วัสดุ และทัวร์พื้นที่จัดจำหน่าย"],
  ["t11", "Thu", "พฤหัส 28 พ.ค.", "16:30", "18:30", "move", "ออกเดินทาง / Transfer", "ไปทัวร์สำนักงานใหญ่ หรือโรงแรม"],
  ["t12", "Thu", "พฤหัส 28 พ.ค.", "19:00", "", "move", "กลับโรงแรม Shin Yokohama Prince Hotel", "Return to hotel"],
  ["t13", "Fri", "ศุกร์ 29 พ.ค.", "08:30", "", "move", "ถึงห้องบอลรูม InterContinental Hotel", "Minato-Mirai"],
  ["t14", "Fri", "ศุกร์ 29 พ.ค.", "10:30", "11:00", "meet", "ประชุมกับ Mitsui", "Meeting with Mitsui"],
  ["t15", "Fri", "ศุกร์ 29 พ.ค.", "12:00", "13:00", "food", "อาหารกลางวันกับ Sato san", "Lunch with Sato san"],
  ["t16", "Fri", "ศุกร์ 29 พ.ค.", "13:30", "14:00", "meet", "ทัวร์สำนักงานใหญ่", "Headquarters Tour"],
  ["t17", "Fri", "ศุกร์ 29 พ.ค.", "14:00", "14:30", "meet", "ประชุมกับ Kubo san", "Isuzu LC-8300A solution"],
  ["t18", "Fri", "ศุกร์ 29 พ.ค.", "14:30", "", "free", "เวลาว่าง / Automotive Engineering Exposition", "กรุณาเยี่ยมชมงาน Automotive Engineering Exposition"]
];

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || "callback";
  const action = params.action || "get";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets_(ss);

  if (action === "setup") {
    return jsonp_(callback, { ok: true, message: "setup complete" });
  }

  if (action === "save") {
    saveStatus_(
      ss.getSheetByName(STATUS_SHEET_NAME),
      params.userId || "unknown",
      params.displayName || "",
      params.done || "{}",
      params.updatedAt || new Date().toISOString()
    );
    return jsonp_(callback, { ok: true });
  }

  if (action === "ping") {
    return jsonp_(callback, { ok: true, message: "pong" });
  }

  const status = findStatus_(ss.getSheetByName(STATUS_SHEET_NAME), params.userId || "unknown");
  return jsonp_(callback, { ok: true, done: status.done, updatedAt: status.updatedAt });
}

function ensureSheets_(ss) {
  let statusSheet = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!statusSheet) {
    statusSheet = ss.insertSheet(STATUS_SHEET_NAME);
  }
  if (statusSheet.getLastRow() === 0) {
    statusSheet.appendRow(["userId", "displayName", "doneJson", "updatedAt"]);
  }

  let tasksSheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet(TASKS_SHEET_NAME);
  }
  if (tasksSheet.getLastRow() === 0) {
    tasksSheet.appendRow(["id", "day", "date", "time", "end", "type", "title", "detail"]);
    tasksSheet.getRange(2, 1, TASKS.length, TASKS[0].length).setValues(TASKS);
  }
}

function findStatus_(sheet, userId) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId) {
      let done = {};
      try {
        done = JSON.parse(values[i][2] || "{}");
      } catch (err) {}
      return { done, updatedAt: values[i][3] || "" };
    }
  }
  return { done: {}, updatedAt: "" };
}

function saveStatus_(sheet, userId, displayName, doneJson, updatedAt) {
  const values = sheet.getDataRange().getValues();
  const rowValues = [userId, displayName, doneJson, updatedAt];

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId) {
      sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  sheet.appendRow(rowValues);
}

function jsonp_(callback, data) {
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(data) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
