const CONFIG = {
  SPREADSHEET_ID: '', // ถ้า Apps Script เปิดมาจาก Google Sheets ไฟล์เดียวกัน ให้เว้นว่าง
  TIMEZONE: 'Asia/Bangkok',
  APP_NAME: 'ระบบจัดการกีฬาสี'
};

const SHEETS = {
  students: '01_Students',
  rooms: '02_Rooms',
  schedule: '03_DailySchedule',
  activities: '04_Activities',
  assignments: '05_Assignments',
  attendance: '06_Attendance',
  settings: '07_Settings'
};

const HEADERS = {
  students: ['student_id', 'room_id', 'number', 'first_name', 'last_name', 'color_id', 'active'],
  rooms: ['room_id', 'room_name', 'color_id', 'color_name', 'teacher', 'active'],
  schedule: ['room_id', 'room_name', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'note'],
  activities: ['activity_id', 'activity_name', 'date', 'start_time', 'end_time', 'venue', 'status'],
  assignments: ['assignment_id', 'student_id', 'activity_id', 'note'],
  attendance: ['attendance_id', 'date', 'student_id', 'room_id', 'first_name', 'last_name', 'status', 'check_time', 'checked_by', 'note'],
  settings: ['key', 'value']
};

function getSS_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ไม่พบ Google Sheets กรุณาเปิด Apps Script จากไฟล์ Google Sheets หรือใส่ SPREADSHEET_ID ใน CONFIG');
  return ss;
}

function getSheet_(key) {
  const sh = getSS_().getSheetByName(SHEETS[key]);
  if (!sh) throw new Error('ไม่พบชีต ' + SHEETS[key] + ' กรุณารัน setupSheets ก่อน');
  return sh;
}

function setupSheets() {
  const ss = getSS_();
  Object.keys(SHEETS).forEach(key => {
    let sh = ss.getSheetByName(SHEETS[key]);
    if (!sh) sh = ss.insertSheet(SHEETS[key]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[key].length)
        .setFontWeight('bold').setBackground('#fce8e6');
      sh.autoResizeColumns(1, HEADERS[key].length);
    }
  });
  migrateAttendanceHeaders_();
  const settings = getSheet_('settings');
  const current = rows_('settings');
  const defaults = [
    ['APP_NAME', CONFIG.APP_NAME],
    ['TIMEZONE', CONFIG.TIMEZONE]
  ];
  defaults.forEach(([key, value]) => {
    if (!current.some(r => clean_(r.key) === key)) {
      settings.appendRow([key, value]);
    }
  });
  return 'สร้าง/ตรวจสอบชีต 7 ชีตเรียบร้อยแล้ว (ไม่มี 08_Users)';
}

function migrateAttendanceHeaders_() {
  const sh = getSheet_('attendance');
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const oldValues = lastRow ? sh.getRange(1, 1, lastRow, lastCol).getValues() : [];
  const oldHeaders = oldValues.length ? oldValues[0].map(v => clean_(v)) : [];
  const targetHeaders = HEADERS.attendance;
  if (oldHeaders.join('|') === targetHeaders.join('|')) return;

  const studentsById = {};
  rows_('students').forEach(st => { studentsById[clean_(st.student_id)] = st; });
  const migrated = [targetHeaders];
  oldValues.slice(1).forEach(oldRow => {
    const old = {};
    oldHeaders.forEach((h, i) => { if (h) old[h] = oldRow[i]; });
    if (!Object.keys(old).some(k => old[k] !== '' && old[k] !== null)) return;
    const st = studentsById[clean_(old.student_id)] || {};
    migrated.push([
      old.attendance_id || Utilities.getUuid(),
      old.date || '',
      clean_(old.student_id),
      clean_(st.room_id || old.room_id),
      clean_(st.first_name || old.first_name),
      clean_(st.last_name || old.last_name),
      clean_(old.status),
      old.check_time || '',
      clean_(old.checked_by),
      clean_(old.note)
    ]);
  });
  sh.clearContents();
  sh.getRange(1, 1, migrated.length, targetHeaders.length).setValues(migrated);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, targetHeaders.length)
    .setFontWeight('bold').setBackground('#fce8e6');
  sh.autoResizeColumns(1, targetHeaders.length);
}

function rows_(key) {
  const sh = getSheet_(key);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(v => String(v).trim());
  return values.slice(1).map((row, i) => {
    const obj = { __row: i + 2 };
    headers.forEach((h, j) => obj[h] = row[j]);
    return obj;
  }).filter(obj => Object.keys(obj).some(k => k !== '__row' && obj[k] !== '' && obj[k] !== null));
}

function clean_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function isActive_(v) {
  const s = clean_(v).toLowerCase();
  return v === true || s === '' || ['true', '1', 'yes', 'y', 'ใช่', 'active'].includes(s);
}

function dateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const s = clean_(value);
  if (!s) return '';
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  const d = new Date(value);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function today_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function timeText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'HH:mm');
  }
  return clean_(value);
}

function publicStudent_(s) {
  return {
    student_id: clean_(s.student_id),
    room_id: clean_(s.room_id),
    number: clean_(s.number),
    first_name: clean_(s.first_name),
    last_name: clean_(s.last_name),
    color_id: clean_(s.color_id)
  };
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบจัดการกีฬาสี')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialData() {
  const students = rows_('students').filter(s => isActive_(s.active));
  const rooms = rows_('rooms').filter(r => isActive_(r.active));
  const activities = rows_('activities');
  const attendance = rows_('attendance').filter(a => dateKey_(a.date) === today_());
  const presentStatuses = ['present', 'late', 'absent', 'leave'];
  const todayMap = {};
  attendance.forEach(a => todayMap[clean_(a.student_id)] = clean_(a.status));
  return {
    appName: getSetting_('APP_NAME') || CONFIG.APP_NAME,
    today: today_(),
    totalStudents: students.length,
    totalRooms: rooms.length,
    totalActivities: activities.length,
    checkedToday: students.filter(s => presentStatuses.includes(todayMap[clean_(s.student_id)])).length,
    rooms: rooms.map(r => ({
      room_id: clean_(r.room_id),
      room_name: clean_(r.room_name),
      color_id: clean_(r.color_id),
      color_name: clean_(r.color_name)
    }))
  };
}

function getSetting_(key) {
  const found = rows_('settings').find(r => clean_(r.key) === key);
  return found ? clean_(found.value) : '';
}

function getRoomSchedules() {
  return rows_('schedule').map(r => ({
    room_id: clean_(r.room_id),
    room_name: clean_(r.room_name),
    monday: timeText_(r.monday),
    tuesday: timeText_(r.tuesday),
    wednesday: timeText_(r.wednesday),
    thursday: timeText_(r.thursday),
    friday: timeText_(r.friday),
    note: clean_(r.note)
  }));
}

function getStudents(roomId) {
  const rid = clean_(roomId);
  return rows_('students')
    .filter(s => isActive_(s.active) && (!rid || clean_(s.room_id) === rid))
    .map(publicStudent_)
    .sort((a, b) => a.room_id.localeCompare(b.room_id, 'th', {numeric:true}) ||
      Number(a.number || 0) - Number(b.number || 0) ||
      a.first_name.localeCompare(b.first_name, 'th'));
}

function searchStudents(query) {
  const q = clean_(query).toLowerCase();
  if (!q) return [];
  return getStudents('').filter(s =>
    [s.student_id, s.number, s.first_name, s.last_name, s.room_id]
      .some(v => clean_(v).toLowerCase().includes(q))
  ).slice(0, 50);
}

function getActivities() {
  const names = {};
  rows_('activities').forEach(a => {
    names[clean_(a.activity_id)] = {
      activity_id: clean_(a.activity_id),
      activity_name: clean_(a.activity_name),
      date: dateKey_(a.date),
      start_time: timeText_(a.start_time),
      end_time: timeText_(a.end_time),
      venue: clean_(a.venue),
      status: clean_(a.status)
    };
  });
  return Object.values(names);
}

function getAssignments(studentId) {
  const sid = clean_(studentId);
  const activityMap = {};
  getActivities().forEach(a => activityMap[a.activity_id] = a);
  return rows_('assignments')
    .filter(a => !sid || clean_(a.student_id) === sid)
    .map(a => ({
      assignment_id: clean_(a.assignment_id),
      student_id: clean_(a.student_id),
      activity: activityMap[clean_(a.activity_id)] || {
        activity_id: clean_(a.activity_id),
        activity_name: 'ไม่พบรายการกีฬา',
        date: '', start_time: '', end_time: '', venue: '', status: ''
      },
      note: clean_(a.note)
    }));
}

function getAttendance(date, roomId) {
  const day = dateKey_(date || today_());
  const rid = clean_(roomId);
  const students = getStudents(rid);
  const attendanceRows = rows_('attendance').filter(a => dateKey_(a.date) === day);
  const attendanceMap = {};
  attendanceRows.forEach(a => {
    attendanceMap[clean_(a.student_id)] = {
      status: clean_(a.status),
      check_time: timeText_(a.check_time),
      checked_by: clean_(a.checked_by),
      note: clean_(a.note)
    };
  });
  return students.map(s => ({
    ...s,
    attendance: attendanceMap[s.student_id] || {
      status: 'pending', check_time: '', checked_by: '', note: ''
    }
  }));
}

function saveAttendance(date, studentId, status, checkerName, note) {
  const day = dateKey_(date);
  const sid = clean_(studentId);
  const statusValue = clean_(status);
  const checkedBy = clean_(checkerName).slice(0, 80);
  const noteValue = clean_(note).slice(0, 300);

  // ตรวจสอบข้อมูลก่อนบันทึก
  if (!day || !/^\d{4}-\d{2}-\d{2}$/.test(day)) {
    throw new Error('วันที่ไม่ถูกต้อง');
  }

  if (!sid) throw new Error('ไม่พบรหัสนักเรียน');

  const allowed = ['present', 'late', 'absent', 'leave'];
  if (!allowed.includes(statusValue)) {
    throw new Error('สถานะเช็กชื่อไม่ถูกต้อง');
  }

  if (!checkedBy) {
    throw new Error('กรุณากรอกชื่อผู้เช็กชื่อ');
  }

  const student = rows_('students').find(
    s => clean_(s.student_id) === sid && isActive_(s.active)
  );

  if (!student) {
    throw new Error('ไม่พบนักเรียน หรือรายชื่อนี้ถูกปิดใช้งาน');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = getSheet_('attendance');
    const all = rows_('attendance');

    // ค้นหาเฉพาะนักเรียนคนนี้ในวันที่เลือก
    // ไม่แก้ไขแถวของวันอื่น
    const existing = all.find(a =>
      dateKey_(a.date) === day &&
      clean_(a.student_id) === sid
    );

    const now = new Date();

    const record = [
      existing
        ? clean_(existing.attendance_id)
        : Utilities.getUuid(),
      day,
      sid,
      clean_(student.room_id),
      clean_(student.first_name),
      clean_(student.last_name),
      statusValue,
      now,
      checkedBy,
      noteValue
    ];

    if (existing) {
      // มีข้อมูลวันเดียวกันแล้ว: แก้ไขแถวนั้นเท่านั้น
      sh.getRange(
        existing.__row,
        1,
        1,
        record.length
      ).setValues([record]);
    } else {
      // ยังไม่มีข้อมูลวันนั้น: เพิ่มแถวใหม่
      sh.appendRow(record);
    }

    return {
      ok: true,
      date: day,
      student_id: sid,
      room_id: clean_(student.room_id),
      first_name: clean_(student.first_name),
      last_name: clean_(student.last_name),
      status: statusValue,
      check_time: Utilities.formatDate(
        now,
        CONFIG.TIMEZONE,
        'HH:mm:ss'
      ),
      checked_by: checkedBy
    };

  } finally {
    lock.releaseLock();
  }
}




