const SHEET_ID = '1uoptlfMVeePnwQBlVmUawYZPIetSHW6ih3nwNJk6NXM'; 
const SHEET_NAME = 'Students';

// ฟังก์ชันช่วยตอบกลับเป็น JSON
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 1. รับ GET Request (สำหรับดึงข้อมูล)
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getAllStudents') {
    const result = getAllStudents();
    return responseJSON(result);
  }
  
  return responseJSON({ success: false, message: 'Invalid Action (GET)' });
}

// 2. รับ POST Request (สำหรับ สร้าง/แก้ไข/ลบ)
function doPost(e) {
  try {
    // รับข้อมูลที่ส่งมาจาก Vercel
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data = params.data; // ข้อมูลที่แนบมา (อาจเป็น object หรือ array)

    let result;

    if (action === 'createStudent') {
      result = createStudent(data); // ส่ง Object ไปตรงๆ ไม่ต้อง stringify ซ้ำ
    } else if (action === 'createStudentsBulk') {
      result = createStudentsBulk(data);
    } else if (action === 'updateStudent') {
      result = updateStudent(data);
    } else if (action === 'updateScoresBulk') {
      result = updateScoresBulk(data);
    } else if (action === 'deleteStudent') {
      result = deleteStudent(data.id);
    } else if (action === 'deleteStudentsBulk') {
      result = deleteStudentsBulk(data.ids);
    } else {
      result = { success: false, message: 'Invalid Action (POST)' };
    }

    return responseJSON(result);

  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

// --- Logic เดิม (ปรับปรุงนิดหน่อยให้รับ Object แทน JSON String) ---

function getAllStudents() {
  const sheet = getSheet();
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 13).getDisplayValues();
  
  if (data.length <= 1) return { success: true, students: [] };
  
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    students.push({
      id: row[0],
      exam_id: String(row[1]).replace(/^'/, ''), // ลบ ' ออกตอนส่งกลับ
      full_name: row[2],
      previous_school: row[3],
      grade_level: row[4],
      thai_score: parseFloat(row[5]) || 0,
      math_score: parseFloat(row[6]) || 0,
      science_score: parseFloat(row[7]) || 0,
      english_score: parseFloat(row[8]) || 0,
      aptitude_score: parseFloat(row[9]) || 0,
      total_score: parseFloat(row[10]) || 0,
      rank: parseInt(row[11]) || 0,
      national_id: String(row[12] || '').replace(/^'/, '')
    });
  }
  return { success: true, students: students };
}

function createStudent(data) {
  // รับ data เป็น object เลย ไม่ต้อง JSON.parse
  try {
    const sheet = getSheet();
    const id = Utilities.getUuid();
    const newRow = [
      id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
      data.thai_score || 0, data.math_score || 0, data.science_score || 0,
      data.english_score || 0, data.aptitude_score || 0, data.total_score || 0,
      data.rank || 0, "'" + (data.national_id || '')
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'บันทึกข้อมูลสำเร็จ', id: id };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function createStudentsBulk(students) {
  try {
    const sheet = getSheet();
    if (!students || students.length === 0) return { success: false, message: 'ไม่พบข้อมูล' };
    const newRows = students.map(data => [
      Utilities.getUuid(), "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
      data.thai_score || 0, data.math_score || 0, data.science_score || 0,
      data.english_score || 0, data.aptitude_score || 0, data.total_score || 0,
      data.rank || 0, "'" + (data.national_id || '')
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 13).setValues(newRows);
    return { success: true, message: `นำเข้าข้อมูลสำเร็จ ${newRows.length} รายการ` };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateStudent(data) {
  try {
    const sheet = getSheet();
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.id) {
        sheet.getRange(i + 1, 1, 1, 13).setValues([[
          data.id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
          data.thai_score || 0, data.math_score || 0, data.science_score || 0,
          data.english_score || 0, data.aptitude_score || 0, data.total_score || 0,
          data.rank || 0, "'" + (data.national_id || '')
        ]]);
        return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateScoresBulk(scoresData) {
  try {
    const sheet = getSheet();
    const range = sheet.getDataRange();
    const allValues = range.getValues(); 
    let studentMap = {};
    for (let i = 1; i < allValues.length; i++) {
      let eid = String(allValues[i][1]).replace(/^'/, '').trim();
      studentMap[eid] = i;
    }
    let updatedCount = 0;
    scoresData.forEach(item => {
      let targetId = String(item.exam_id).trim();
      if (studentMap.hasOwnProperty(targetId)) {
        let r = studentMap[targetId];
        let sc = parseFloat(item.science_score) || 0;
        let ma = parseFloat(item.math_score) || 0;
        let en = parseFloat(item.english_score) || 0;
        let total = sc + ma + en;
        allValues[r][6] = ma; allValues[r][7] = sc; allValues[r][8] = en; allValues[r][10] = total;
        updatedCount++;
      }
    });
    if (updatedCount > 0) {
      range.setValues(allValues);
      return { success: true, message: `อัปเดตคะแนนสำเร็จ ${updatedCount} รายการ` };
    } else { return { success: false, message: 'ไม่พบรหัสผู้สอบที่ตรงกัน' }; }
  } catch (e) { return { success: false, message: e.toString() }; }
}

function deleteStudent(id) {
  try {
    const sheet = getSheet();
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function deleteStudentsBulk(idsToDelete) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (idsToDelete.includes(data[i][0])) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    if (deletedCount > 0) return { success: true, message: `ลบข้อมูลสำเร็จ ${deletedCount} รายการ` };
    else return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID','Exam ID','Full Name','Previous School','Grade Level','Thai','Math','Science','English','Aptitude','Total','Rank','National ID']);
  }
  return sheet;
}
