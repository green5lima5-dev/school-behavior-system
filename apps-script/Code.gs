const CONFIG = {
  SPREADSHEET_ID: '1BKKplEs9ZG82ZKHOx6hJljHDg8a6Mn8n9m91bBlwmt0',
  SOURCE_GID: 1659414265,
  TX_SHEET: 'BehaviorTransactions',
  USER_SHEET: 'BehaviorUsers',
  BASE_SCORE: 100,
  SESSION_SECONDS: 21600,
  INITIAL_ADMIN: {
    username: 'admin',
    password: '', // กำหนดค่า INITIAL_ADMIN_PASSWORD ใน Script Properties ก่อน setupSystem
    fullName: 'ผู้ดูแลระบบ',
    role: 'admin'
  }
};

const TX_HEADERS = [
  'TransactionId','Timestamp','Date','StudentKey','StudentId','No','ClassName','StudentName',
  'Type','ReasonId','Reason','Score','Hours','Note','RecorderUsername','RecorderName','RecorderRole'
];
const USER_HEADERS = [
  'Username',
  'PasswordHash',
  'Salt',
  'FullName',
  'Role',
  'Active',
  'CreatedAt',
  'UpdatedAt',
  'StudentKey',
  'StudentId',
  'ClassName',
  'LastLoginAt'
];


/** ห้องที่บัญชีครูทั่วไปได้รับอนุญาตให้ดู (ไม่มีการส่งข้อมูลห้องอื่นออกจากเซิร์ฟเวอร์) */
const TEACHER_CLASS_MAP = Object.freeze({
  'yuwariyah': 'ม.2/1',
  'nuriyah': 'ม.3/2',
  'nasilah': 'ม.6/2',
  'noreesan': 'ม.2/1',
  'hafis': 'ม.5/1',
  'fatimah': 'ม.1/1',
  'naseebah': 'ม.4/1',
  'deyoh': 'ม.3/2',
  'natchaya': 'ม.4/2',
  'sarilla': 'ม.4/1',
  'ilham': 'ม.6/1',
  'nadiyah': 'ม.5/2',
  'suhaisan': 'ม.5/1',
  'nurizzah': 'ม.4/1',
  'bassam': 'ม.2/2',
  'asiyah': 'ม.3/1',
  'afandi': 'ม.1/2',
  'phanuphong': 'ม.3/1',
  'nazrin': 'ม.3/3',
  'safwanee': 'ม.2/2',
  'faizah': 'ม.4/2',
  'anisah': 'ม.6/1',
  'ketsara': 'ม.6/2',
  'haisam': 'ม.1/2',
  'sitirokiyoh': 'ม.5/2',
  'sitinurulhuda': 'ม.1/1',
  'sapiing': 'ม.3/3',
  'halimoh': 'ม.6/2',
  'husna': 'ม.2/2'
});

function normalizeTeacherClass_(value) {
  const s = String(value || '').trim().replace(/\s+/g, '');
  const m = s.match(/^(?:ม\.?)?(\d+)\/(\d+)$/);
  return m ? 'ม.' + Number(m[1]) + '/' + Number(m[2]) : '';
}
function authorizedTeacherClass_(user) {
  // รายชื่อครู 29 คน: ยึดการจับคู่ห้องจากข้อมูลที่ได้รับ ไม่เปิดให้ client เลือกห้องเอง
  const key = String(user.username || '').trim().toLowerCase();
  const mapped = TEACHER_CLASS_MAP[key];
  return mapped || normalizeTeacherClass_(user.className);
}

/** ซิงค์ห้องประจำของครู 29 คนลง BehaviorUsers คอลัมน์ K (ไม่แก้รหัสผ่าน) */
function syncTeacherClassAssignments() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
    const values = sheet.getDataRange().getValues();
    let updated = 0;
    let found = 0;
    const missing = Object.keys(TEACHER_CLASS_MAP);
    const existing = new Set();
    for (let i = 1; i < values.length; i++) {
      const username = String(values[i][0] || '').trim().toLowerCase();
      if (!Object.prototype.hasOwnProperty.call(TEACHER_CLASS_MAP, username)) continue;
      found++;
      existing.add(username);
      const room = TEACHER_CLASS_MAP[username];
      if (String(values[i][4]).trim() !== 'teacher') {
        Logger.log('ข้ามบัญชีที่ไม่ใช่ teacher: ' + username);
        continue;
      }
      if (normalizeTeacherClass_(values[i][10]) !== room) {
        sheet.getRange(i + 1, 11).setValue(room);
        updated++;
      }
    }
    SpreadsheetApp.flush();
    const result = {
      ok: true, found: found, updated: updated,
      missing: missing.filter(name => !existing.has(name)),
      spreadsheetUrl: getSS_().getUrl(), sheetName: sheet.getName()
    };
    Logger.log(JSON.stringify(result));
    return result;
  } finally { lock.releaseLock(); }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('บันทึกคะแนนความประพฤติ')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

function setupSystem() {
  const ss = getSS_();
  const tx = getOrCreateSheet_(ss, CONFIG.TX_SHEET, TX_HEADERS);
  const users = getOrCreateSheet_(ss, CONFIG.USER_SHEET, USER_HEADERS);
  formatSheet_(tx, TX_HEADERS.length);
  formatSheet_(users, USER_HEADERS.length);

  const existing = readUsers_().find(u => u.username.toLowerCase() === CONFIG.INITIAL_ADMIN.username.toLowerCase());
  if (!existing) {
    const initialPassword = PropertiesService.getScriptProperties().getProperty('INITIAL_ADMIN_PASSWORD');
    if (!initialPassword || initialPassword.length < 12) throw new Error('กรุณาตั้ง Script Property INITIAL_ADMIN_PASSWORD อย่างน้อย 12 ตัวอักษรก่อน setupSystem');
    const salt = Utilities.getUuid();
    users.appendRow([
      CONFIG.INITIAL_ADMIN.username,
      hashPassword_(initialPassword, salt),
      salt,
      CONFIG.INITIAL_ADMIN.fullName,
      CONFIG.INITIAL_ADMIN.role,
      true,
      new Date(),
      new Date()
    ]);
  }
  return 'ตั้งค่าระบบสำเร็จ | หากเพิ่งสร้างแอดมิน กรุณาเปลี่ยนรหัสผ่านทันที';
}

function testConnection(token) {
  requireSession_(token);
  const ss = getSS_();
  const source = getSourceSheet_(ss);
  return {ok:true, spreadsheetName:ss.getName(), sourceSheet:source.getName(), timestamp:new Date().toISOString()};
}

function login(username, password) {

  username = String(username || '').trim();
  password = String(password || '');

  if (!username || !password) {
    return {
      ok: false,
      message: 'กรุณากรอก Username และ Password'
    };
  }

  const users = readUsers_();

 const normalizeUsername = value => {

  const text =
    String(value || '')
      .trim()
      .toLowerCase();

  // ถ้า Username เป็นตัวเลขล้วน
  // จะตัดเลข 0 ที่อยู่ด้านหน้าออก
  if (/^\d+$/.test(text)) {

    const noZero =
      text.replace(/^0+/, '');

    return noZero || '0';
  }

  // ถ้าเป็น Username ของครู / Admin
  // จะไม่เปลี่ยนแปลง
  return text;
};


const normalizedInput =
  normalizeUsername(username);


const user = users.find(u => {

  return normalizeUsername(u.username) ===
         normalizedInput;

});
  if (!user) {
    return {
      ok: false,
      message: 'ไม่พบบัญชีผู้ใช้'
    };
  }

  if (!user.active) {
    return {
      ok: false,
      message: 'บัญชีนี้ถูกระงับการใช้งาน'
    };
  }

  const inputHash =
    hashPassword_(password, user.salt);

  if (inputHash !== user.passwordHash) {

    return {
      ok: false,
      message: 'รหัสผ่านไม่ถูกต้อง'
    };

  }

  // นักเรียนต้องเชื่อมกับข้อมูลนักเรียน
  if (
    user.role === 'student' &&
    !user.studentKey
  ) {

    return {
      ok: false,
      message: 'บัญชีนักเรียนยังไม่ได้เชื่อมโยงกับข้อมูลนักเรียน'
    };

  }


  // =========================
  // บันทึกเวลา Login ล่าสุด
  // =========================

  try {

    const sh =
      getOrCreateSheet_(
        getSS_(),
        CONFIG.USER_SHEET,
        USER_HEADERS
      );

    const data =
      sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0] || '')
          .trim()
          .toLowerCase() ===
        username.toLowerCase()
      ) {

        // คอลัมน์ L = LastLoginAt
        sh.getRange(i + 1, 12)
          .setValue(new Date());

        break;
      }
    }

  } catch (e) {

    console.log(
      'บันทึก LastLogin ไม่สำเร็จ: ' +
      e.message
    );

  }


  const token =
    Utilities.getUuid() +
    Utilities.getUuid();


  const session = {

    username: user.username,

    fullName: user.fullName,

    role: user.role,

    studentKey:
      user.studentKey || '',

    studentId:
      user.studentId || '',

    className:
      user.role === 'teacher' ? authorizedTeacherClass_(user) : (user.className || '')

  };


  CacheService
    .getScriptCache()
    .put(
      'SESSION_' + token,
      JSON.stringify(session),
      CONFIG.SESSION_SECONDS
    );


  return {

    ok: true,

    token: token,

    user: session

  };
}


/** ใช้ตรวจเลขประจำตัวนักเรียนให้ตรงกัน แม้พิมพ์เลขศูนย์นำหน้า */
function normalizeStudentIdForRegistration_(value) {
  const id = String(value == null ? '' : value).trim();
  if (!/^\d+$/.test(id)) return id.toLowerCase();
  return id.replace(/^0+/, '') || '0';
}

/** หนึ่งนักเรียนสมัครได้หนึ่งบัญชี โดยตรวจและเพิ่มข้อมูลภายใต้ Script Lock เดียวกัน */
function registerStudent(payload) {
  payload = payload || {};
  const studentId = String(payload.studentId || '').trim();
  const className = normalizeTeacherClass_(payload.className);
  const fullName = String(payload.fullName || '').trim();
  const password = String(payload.password || '');
  if (!studentId || !className || !fullName || !password) {
    throw new Error('กรุณากรอกข้อมูลให้ครบ');
  }
  if (!/^\d+$/.test(studentId)) {
    throw new Error('รหัสนักเรียนต้องเป็นตัวเลขเท่านั้น');
  }
  if (password.length < 8) {
    throw new Error('Password ต้องมีอย่างน้อย 8 ตัวอักษร');
  }

  const normalizeName = text => String(text || '')
    .replace(/\s+/g, '').trim().toLowerCase();
  const normalizeId = normalizeStudentIdForRegistration_;

  // ต้องล็อกก่อนค้นหาบัญชี แล้วค่อยเขียนข้อมูล เพื่อกันคำขอที่เข้ามาพร้อมกัน
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const student = readStudents_().find(s =>
      normalizeId(s.studentId) === normalizeId(studentId) &&
      normalizeTeacherClass_(s.className) === className &&
      normalizeName(s.name) === normalizeName(fullName)
    );
    if (!student) {
      throw new Error('ข้อมูลไม่ตรงกับฐานข้อมูลนักเรียน กรุณาตรวจรหัสนักเรียน ชั้น และชื่อ-นามสกุล');
    }

    const username = String(student.studentId).trim();
    const studentKey = String(student.key || '').trim();
    const users = readUsers_(); // อ่านข้อมูลล่าสุดขณะถือ lock
    const duplicate = users.some(u => {
      const oldUsername = String(u.username || '').trim();
      const oldId = String(u.studentId || '').trim();
      const oldKey = String(u.studentKey || '').trim();
      // รวมบัญชีเก่าที่อาจไม่ได้บันทึก StudentId/StudentKey แต่ใช้รหัสนักเรียนเป็น Username
      return normalizeId(oldUsername) === normalizeId(username) ||
        (oldId !== '' && normalizeId(oldId) === normalizeId(username)) ||
        (studentKey !== '' && oldKey !== '' && oldKey === studentKey);
    });
    if (duplicate) {
      throw new Error('รหัสนักเรียนนี้มีบัญชีอยู่แล้ว ไม่สามารถสมัครซ้ำหรือสร้าง 2 Username ได้ กรุณาเข้าสู่ระบบหรือติดต่อผู้ดูแลระบบหากลืมรหัสผ่าน');
    }

    const salt = Utilities.getUuid();
    const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
    sh.appendRow([
      username, hashPassword_(password, salt), salt, student.name,
      'student', true, new Date(), new Date(),
      studentKey, student.studentId, student.className, ''
    ]);
    SpreadsheetApp.flush();
    return {
      ok: true, username: username, fullName: student.name,
      className: student.className, message: 'สมัครใช้งานสำเร็จ'
    };
  } finally {
    lock.releaseLock();
  }
}

function getSession(token) {
  try { return {ok:true, user:requireSession_(token)}; }
  catch (e) { return {ok:false, message:e.message}; }
}

function getDashboardData(token) {
  const session = requireSession_(token);
  if (session.role === 'student') {
    throw new Error('นักเรียนไม่มีสิทธิ์เรียกข้อมูลภาพรวม');
  }
  const allStudents = readStudents_();
  const allTransactions = readTransactions_();
  let students = allStudents;
  let transactions = allTransactions;
  let allowedClass = '';
  if (session.role === 'teacher') {
    // ครูทั่วไปต้องถูกผูกกับห้องเรียน มิฉะนั้นปฏิเสธการดึงข้อมูลทั้งหมด
    allowedClass = normalizeTeacherClass_(session.className);
    if (!allowedClass) throw new Error('ยังไม่ได้กำหนดห้องประจำของบัญชีครู กรุณาติดต่อผู้ดูแลระบบ');
    students = allStudents.filter(s => normalizeTeacherClass_(s.className) === allowedClass);
    const studentKeys = new Set(students.map(s => String(s.key)));
    const totals = new Map();
    allTransactions.forEach(t => {
      const key = String(t.studentKey);
      if (!studentKeys.has(key)) return;
      if (!totals.has(key)) totals.set(key, {positive:0, negative:0});
      const item = totals.get(key);
      const score = Number(t.score || 0);
      if (score > 0) item.positive += score;
      else if (score < 0) item.negative += score;
    });
    // ส่งเฉพาะยอดรวมบวกและลบเพื่อให้รายงาน Excel ถูกต้อง ไม่เผยประวัติรายรายการ
    transactions = students.flatMap(s => {
      const total = totals.get(String(s.key)) || {positive:0,negative:0};
      return [{studentKey:s.key,score:total.positive},{studentKey:s.key,score:total.negative}];
    });
  }
  return {
    ok:true,
    students:students,
    transactions:transactions,
    users: session.role === 'admin' ? publicUsers_(readUsers_()) : [],
    baseScore:CONFIG.BASE_SCORE,
    allowedClass:allowedClass
  };
}
function getStudentPortalData(token) {

  const session =
    requireSession_(token);

  if (session.role !== 'student') {

    throw new Error(
      'เมนูนี้สำหรับนักเรียนเท่านั้น'
    );

  }

  if (!session.studentKey) {

    throw new Error(
      'ไม่พบข้อมูลนักเรียนที่เชื่อมกับบัญชีนี้'
    );

  }

  const student =
    readStudents_().find(
      s =>
        String(s.key) ===
        String(session.studentKey)
    );

  if (!student) {

    throw new Error(
      'ไม่พบข้อมูลนักเรียนในฐานข้อมูล'
    );

  }

  // *** สำคัญ ***
  // ดึงเฉพาะรายการของนักเรียนที่ Login
  const transactions =
    readTransactions_()
      .filter(
        t =>
          String(t.studentKey) ===
          String(session.studentKey)
      );

  const positive =
    transactions
      .filter(
        t =>
          Number(t.score) > 0
      )
      .reduce(
        (sum, t) =>
          sum + Number(t.score),
        0
      );

  const negative =
    transactions
      .filter(
        t =>
          Number(t.score) < 0
      )
      .reduce(
        (sum, t) =>
          sum +
          Math.abs(Number(t.score)),
        0
      );

  const currentScore =
    CONFIG.BASE_SCORE +
    transactions.reduce(
      (sum, t) =>
        sum + Number(t.score || 0),
      0
    );

  return {

    ok: true,

    student: student,

    transactions:
      transactions,

    baseScore:
      CONFIG.BASE_SCORE,

    positive:
      positive,

    negative:
      negative,

    currentScore:
      currentScore

  };
}
function saveTransaction(token, payload) {
  const session = requireSession_(token);
  if (!['admin', 'discipline'].includes(session.role)) {
    throw new Error('อนุญาตเฉพาะครูฝ่ายปกครองและแอดมินให้บันทึกคะแนน');
  }
  payload = payload || {};
  const allowedTypes = ['positive','minor','moderate','serious'];
  if (!allowedTypes.includes(payload.type)) throw new Error('ประเภทรายการไม่ถูกต้อง');
  if (payload.type === 'serious' && !['admin','discipline'].includes(session.role)) {
    throw new Error('เฉพาะครูฝ่ายปกครองหรือแอดมินเท่านั้นที่บันทึกความผิดหมวด 3 ได้');
  }

  const students = readStudents_();
  const student = students.find(s => String(s.key) === String(payload.studentKey));
  if (!student) throw new Error('ไม่พบนักเรียน กรุณาโหลดข้อมูลใหม่');
  if (!payload.date || !payload.reason || payload.score === undefined || payload.score === null) throw new Error('ข้อมูลรายการไม่ครบถ้วน');

  const score = Number(payload.score);
  if (!Number.isFinite(score) || score === 0) throw new Error('คะแนนไม่ถูกต้อง');
  if (payload.type === 'positive' && score <= 0) throw new Error('รายการเพิ่มคะแนนต้องเป็นคะแนนบวก');
  if (payload.type !== 'positive' && score >= 0) throw new Error('รายการตัดคะแนนต้องเป็นคะแนนลบ');

  // บังคับคะแนนตามระเบียบจากฝั่ง Server ป้องกันแก้ค่าใน Browser
  const verifiedScore = verifyRuleScore_(payload.type, payload.reasonId, payload.hours);
  if (score !== verifiedScore) throw new Error('คะแนนไม่ตรงตามมาตรการของโรงเรียน');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // หมวด 3 เป็นความผิดร้ายแรง: ห้ามบันทึกคะแนนบวกเพิ่มภายหลัง
    // ตรวจจากประวัติจริงบน Server เพื่อไม่ให้ข้ามกติกาด้วยการแก้หน้าเว็บ
    if (payload.type === 'positive') {
      const hasSerious = readTransactions_().some(t =>
        String(t.studentKey) === String(student.key) &&
        (t.type === 'serious' || /^M3-/.test(String(t.reasonId || '')))
      );
      if (hasSerious) {
        throw new Error('นักเรียนมีประวัติถูกหักคะแนนหมวด 3 (ความผิดร้ายแรง) จึงไม่สามารถเพิ่มคะแนนความดีหรือใช้คะแนนหักล้างความผิดหมวด 3 ได้');
      }
    }

    const sh = getOrCreateSheet_(getSS_(), CONFIG.TX_SHEET, TX_HEADERS);
    sh.appendRow([
      Utilities.getUuid(), new Date(), String(payload.date), student.key, student.studentId || '',
      student.no || '', student.className || '', student.name || '', payload.type, payload.reasonId || '',
      payload.reason || '', verifiedScore, payload.hours || '', payload.note || '',
      session.username, session.fullName, session.role
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return {ok:true};
}

function listUsers(token) {
  requireRole_(token, ['admin']);
  return {ok:true, users:publicUsers_(readUsers_())};
}

function createUser(token, payload) {
  requireRole_(token, ['admin']);
  payload = payload || {};
  const username = String(payload.username || '').trim();
  const fullName = String(payload.fullName || '').trim();
  const password = String(payload.password || '');
  const role = String(payload.role || 'teacher');
  if (!username || !fullName || password.length < 8) throw new Error('กรุณากรอกข้อมูลให้ครบ และรหัสผ่านอย่างน้อย 8 ตัวอักษร');
  if (!/^[A-Za-z0-9._-]{3,40}$/.test(username)) throw new Error('Username ใช้ได้เฉพาะ A-Z, a-z, 0-9, จุด ขีดกลาง และขีดล่าง');
  if (!['admin','teacher','discipline'].includes(role)) throw new Error('สิทธิ์ผู้ใช้ไม่ถูกต้อง');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (readUsers_().some(u => u.username.toLowerCase() === username.toLowerCase())) throw new Error('Username นี้มีอยู่แล้ว');
    const salt = Utilities.getUuid();
    const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
    sh.appendRow([username, hashPassword_(password,salt), salt, fullName, role, true, new Date(), new Date()]);
  } finally { lock.releaseLock(); }
  return {ok:true, users:publicUsers_(readUsers_())};
}

function resetUserPassword(token, username, newPassword) {
  requireRole_(token, ['admin']);
  username = String(username || '').trim();
  newPassword = String(newPassword || '');
  if (newPassword.length < 8) throw new Error('รหัสผ่านต้องอย่างน้อย 8 ตัวอักษร');
  const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
  const data = sh.getDataRange().getValues();
  const row = data.findIndex((r,i) => i>0 && String(r[0]).toLowerCase() === username.toLowerCase());
  if (row < 1) throw new Error('ไม่พบผู้ใช้');
  const salt = Utilities.getUuid();
  sh.getRange(row+1,2).setValue(hashPassword_(newPassword,salt));
  sh.getRange(row+1,3).setValue(salt);
  sh.getRange(row+1,8).setValue(new Date());
  return {ok:true};
}
function deleteUser(token, username) {

  const session =
    requireRole_(token, ['admin']);

  username =
    String(username || '').trim();

  if (!username) {
    throw new Error('ไม่พบ Username ที่ต้องการลบ');
  }

  // ห้ามลบบัญชีตัวเอง
  if (
    session.username.toLowerCase() ===
    username.toLowerCase()
  ) {
    throw new Error(
      'ไม่สามารถลบบัญชี Admin ที่กำลังใช้งานอยู่ได้'
    );
  }

  // ห้ามลบ admin หลัก
  if (
    username.toLowerCase() === 'admin'
  ) {
    throw new Error(
      'ไม่อนุญาตให้ลบบัญชี Admin หลัก'
    );
  }

  const sh =
    getOrCreateSheet_(
      getSS_(),
      CONFIG.USER_SHEET,
      USER_HEADERS
    );

  const data =
    sh.getDataRange().getValues();

  const row =
    data.findIndex(
      (r, i) =>
        i > 0 &&
        String(r[0] || '')
          .trim()
          .toLowerCase() ===
        username.toLowerCase()
    );

  if (row < 1) {
    throw new Error('ไม่พบบัญชีผู้ใช้');
  }

  sh.deleteRow(row + 1);

  SpreadsheetApp.flush();

  return {
    ok: true,
    users: publicUsers_(readUsers_())
  };
}
function setUserActive(token, username, active) {
  const session = requireRole_(token, ['admin']);
  username = String(username || '').trim();
  if (session.username.toLowerCase() === username.toLowerCase() && !active) throw new Error('ไม่สามารถระงับบัญชีที่กำลังใช้งานอยู่ได้');
  const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
  const data = sh.getDataRange().getValues();
  const row = data.findIndex((r,i) => i>0 && String(r[0]).toLowerCase() === username.toLowerCase());
  if (row < 1) throw new Error('ไม่พบผู้ใช้');
  sh.getRange(row+1,6).setValue(Boolean(active));
  sh.getRange(row+1,8).setValue(new Date());
  return {ok:true, users:publicUsers_(readUsers_())};
}

function verifyRuleScore_(type, reasonId, hours) {
  const positive = {P1:5,P2:10,P3:10,P4:25};
  if (type === 'positive') {
    if (reasonId === 'P5') {
      const h = Math.max(1, parseInt(hours,10) || 1);
      return h * 5;
    }
    if (positive[reasonId] !== undefined) return positive[reasonId];
  }
  if (type === 'minor' && /^M1-(10|[1-9])$/.test(reasonId)) return -5;
  if (type === 'moderate' && /^M2-(10|[1-9])$/.test(reasonId)) return -10;
  if (type === 'serious' && /^M3-[1-9]$/.test(reasonId)) return -50;
  throw new Error('รายการคะแนนไม่อยู่ในมาตรการที่กำหนด');
}
function readStudents_() {

  const ss = getSS_();

  const ignoreSheets = [
    'SYS_Users',
    'SYS_BehaviorTransactions',
    'BehaviorUsers',
    'BehaviorTransactions',
    'บันทึกคะแนนความประพฤติ'
  ];

  const students = [];
  const seenKeys = new Set();

  const sheets = ss.getSheets();

  sheets.forEach(function(sh) {

    const sheetName = String(sh.getName() || '').trim();

    // ข้ามชีตระบบ
    if (ignoreSheets.includes(sheetName)) {
      return;
    }

    // รับเฉพาะชื่อชีตห้อง เช่น
    // ม.1/1, ม.1/2, 1/1, 2/1 ฯลฯ
    const normalizedSheetName =
      sheetName.replace(/\s+/g, '');

    const classMatch =
      normalizedSheetName.match(
        /^(?:ม\.?)?(\d+)\/(\d+)$/
      );

    if (!classMatch) {
      return;
    }

    const className =
      'ม.' +
      classMatch[1] +
      '/' +
      classMatch[2];

    const values =
      sh.getDataRange()
        .getDisplayValues();

    if (!values || values.length === 0) {
      return;
    }

    // ==============================
    // หาแถวหัวตาราง
    // ==============================

    let headerRow = -1;

    let idxNo = -1;
    let idxId = -1;
    let idxName = -1;
    let idxClass = -1;

    const maxCheckRows =
      Math.min(values.length, 15);

    for (
      let r = 0;
      r < maxCheckRows;
      r++
    ) {

      const headers =
        values[r].map(function(v) {
          return String(v || '')
            .trim()
            .toLowerCase()
            .replace(/[\s._-]+/g, '');
        });

      idxNo = findHeader_(headers, [
        'เลขที่',
        'ลำดับ',
        'no',
        'number'
      ]);

      idxId = findHeader_(headers, [
        'รหัสนักเรียน',
        'รหัส',
        'studentid'
      ]);

      idxName = findHeader_(headers, [
        'รายชื่อนักเรียน',
        'ชื่อนักเรียน',
        'ชื่อ-นามสกุล',
        'ชื่อสกุล',
        'ชื่อ',
        'name'
      ]);

      idxClass = findHeader_(headers, [
        'ระดับชั้น',
        'ชั้น',
        'ห้อง',
        'class'
      ]);

      // ขอเพียงเจอคอลัมน์ชื่อก็ถือว่าเจอหัวตาราง
      if (idxName >= 0) {

        headerRow = r;

        break;
      }
    }

    if (headerRow < 0) {

      Logger.log(
        'ข้ามชีต ' +
        sheetName +
        ' : ไม่พบหัวคอลัมน์รายชื่อนักเรียน'
      );

      return;
    }

    // ==============================
    // อ่านนักเรียน
    // ==============================

    for (
      let r = headerRow + 1;
      r < values.length;
      r++
    ) {

      const row =
        values[r];

      const name =
        idxName >= 0
          ? String(
              row[idxName] || ''
            ).trim()
          : '';

      if (!name) {
        continue;
      }

      // ป้องกันแถวรวม/หมายเหตุ
      if (
        name === 'รายชื่อนักเรียน' ||
        name === 'ชื่อ-นามสกุล' ||
        name === 'ชื่อนักเรียน'
      ) {
        continue;
      }

      const no =
        idxNo >= 0
          ? String(
              row[idxNo] || ''
            ).trim()
          : '';

      const studentId =
        idxId >= 0
          ? String(
              row[idxId] || ''
            ).trim()
          : '';

      let studentClass =
        className;

      // ถ้าในชีตมีคอลัมน์ระดับชั้น
      // และมีค่า ให้ใช้ค่านั้น
      if (
        idxClass >= 0 &&
        String(
          row[idxClass] || ''
        ).trim()
      ) {

        studentClass =
          normalizeClassNameNew_(
            row[idxClass]
          );

      }

      if (!studentClass) {
        studentClass =
          className;
      }

      // สร้าง key นักเรียน
      const key =
        studentId
          ? 'ID:' + studentId
          : 'K:' +
            [
              studentClass,
              no,
              name
            ].join('|');

      if (seenKeys.has(key)) {
        continue;
      }

      seenKeys.add(key);

      students.push({

        key: key,

        studentId:
          studentId,

        no:
          no,

        className:
          studentClass,

        name:
          name,

        sourceSheet:
          sheetName,

        sourceRow:
          r + 1

      });

    }

  });

  // ==============================
  // เรียงตามระดับชั้นและเลขที่
  // ==============================

  students.sort(
    function(a, b) {

      const ca =
        classSortValueNew_(
          a.className
        );

      const cb =
        classSortValueNew_(
          b.className
        );

      if (ca !== cb) {
        return ca - cb;
      }

      const na =
        parseInt(
          a.no,
          10
        );

      const nb =
        parseInt(
          b.no,
          10
        );

      if (
        !isNaN(na) &&
        !isNaN(nb)
      ) {
        return na - nb;
      }

      return String(a.name)
        .localeCompare(
          String(b.name),
          'th'
        );

    }
  );

  Logger.log(
    'พบนักเรียนทั้งหมด: ' +
    students.length
  );

  Logger.log(
    'ระดับชั้น: ' +
    [
      ...new Set(
        students.map(
          s => s.className
        )
      )
    ].join(', ')
  );

  return students;
}


/* =========================
   แปลงชื่อระดับชั้น
========================= */

function normalizeClassNameNew_(value) {

  let text =
    String(value || '')
      .trim()
      .replace(/\s+/g, '');

  if (!text) {
    return '';
  }

  let match =
    text.match(
      /^(?:ม\.?)?(\d+)\/(\d+)$/
    );

  if (match) {

    return (
      'ม.' +
      match[1] +
      '/' +
      match[2]
    );

  }

  return text;
}


/* =========================
   เรียง ม.1/1 → ม.6/2
========================= */

function classSortValueNew_(className) {

  const match =
    String(className || '')
      .match(
        /(\d+)\/(\d+)/
      );

  if (!match) {
    return 99999;
  }

  return (
    Number(match[1]) * 100 +
    Number(match[2])
  );
}

 function classSortValueNew_(className) {

  const match =
    String(className || '')
      .match(/(\d+)\/(\d+)/);

  if (!match) {
    return 99999;
  }

  return (
    Number(match[1]) * 100 +
    Number(match[2])
  );
}
function deleteTransaction(token, transactionId) {

  const session = requireSession_(token);

  // อนุญาตเฉพาะ Admin และครูฝ่ายปกครอง
  if (!['admin', 'discipline'].includes(session.role)) {
    throw new Error(
      'เฉพาะผู้ดูแลระบบหรือครูฝ่ายปกครองเท่านั้นที่สามารถลบรายการได้'
    );
  }

  transactionId = String(transactionId || '').trim();

  if (!transactionId) {
    throw new Error('ไม่พบรหัสรายการที่ต้องการลบ');
  }

  const sh =
    getOrCreateSheet_(
      getSS_(),
      CONFIG.TX_SHEET,
      TX_HEADERS
    );

  const data =
    sh.getDataRange().getValues();

  let rowToDelete = -1;

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0] || '').trim() ===
      transactionId
    ) {

      rowToDelete = i + 1;

      break;
    }
  }

  if (rowToDelete < 2) {
    throw new Error('ไม่พบรายการที่ต้องการลบ');
  }

  sh.deleteRow(rowToDelete);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'ลบรายการสำเร็จ'
  };
}

function readTransactions_() {
  const sh = getOrCreateSheet_(getSS_(), CONFIG.TX_SHEET, TX_HEADERS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({
    transactionId:String(r[0]||''), timestamp:toIso_(r[1]), date:dateOnly_(r[2]), studentKey:String(r[3]||''),
    studentId:String(r[4]||''), no:String(r[5]||''), className:String(r[6]||''), studentName:String(r[7]||''),
    type:String(r[8]||''), reasonId:String(r[9]||''), reason:String(r[10]||''), score:Number(r[11]||0), hours:r[12]||'',
    note:String(r[13]||''), username:String(r[14]||''), recorderName:String(r[15]||''), recorderRole:String(r[16]||'')
  }));
}

function readUsers_() {

  const sh = getOrCreateSheet_(
    getSS_(),
    CONFIG.USER_SHEET,
    USER_HEADERS
  );

  const data = sh.getDataRange().getValues();

  if (data.length < 2) return [];

  return data
    .slice(1)
    .filter(r => r[0])
    .map(r => ({

      username: String(r[0] || '').trim(),
      passwordHash: String(r[1] || '').trim(),
      salt: String(r[2] || '').trim(),
      fullName: String(r[3] || '').trim(),
      role: String(r[4] || 'teacher').trim(),
      active: toBool_(r[5]),

      createdAt: toIso_(r[6]),
      updatedAt: toIso_(r[7]),

      studentKey: String(r[8] || '').trim(),
      studentId: String(r[9] || '').trim(),
      className: String(r[10] || '').trim(),

      lastLoginAt: toIso_(r[11])

    }));
}

function publicUsers_(users) {

  return users.map(u => ({

    username:
      u.username,

    fullName:
      u.fullName,

    role:
      u.role,

    active:
      u.active,

    studentId:
      u.studentId || '',

    className:
      u.className || '',

    createdAt:
      u.createdAt || '',

    updatedAt:
      u.updatedAt || '',

    lastLoginAt:
      u.lastLoginAt || ''

  }));

}

function requireSession_(token) {
  token = String(token || '');
  if (!token) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const raw = CacheService.getScriptCache().get('SESSION_' + token);
  if (!raw) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const session = JSON.parse(raw);
  const live = readUsers_().find(u => u.username.toLowerCase() === session.username.toLowerCase());
  if (!live || !live.active) throw new Error('บัญชีถูกระงับหรือไม่พบผู้ใช้');
  // อ่านสิทธิ์และห้องเรียนปัจจุบันจากฐานข้อมูลทุกครั้งเพื่อไม่ใช้ข้อมูล cache เก่า
  session.role = live.role;
  session.fullName = live.fullName;
  session.className = live.role === 'teacher' ? authorizedTeacherClass_(live) : live.className;
  session.studentKey = live.studentKey || '';
  session.studentId = live.studentId || '';
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), CONFIG.SESSION_SECONDS);
  return session;
}

function requireRole_(token, roles) {
  const session = requireSession_(token);
  if (!roles.includes(session.role)) throw new Error('คุณไม่มีสิทธิ์ดำเนินการนี้');
  return session;
}

function getSS_() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getSourceSheet_(ss) {
  const sh = ss.getSheets().find(s => s.getSheetId() === Number(CONFIG.SOURCE_GID));
  if (!sh) throw new Error('ไม่พบชีตรายชื่อนักเรียน GID: ' + CONFIG.SOURCE_GID);
  return sh;
}
function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const actual = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getDisplayValues()[0];
    headers.forEach((h,i)=>{ if (!actual[i]) sh.getRange(1,i+1).setValue(h); });
  }
  return sh;
}
function formatSheet_(sh, cols) {
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#EEF2FF');
  sh.autoResizeColumns(1,cols);
}
function normalizeHeader_(s) { return String(s||'').trim().toLowerCase().replace(/[\s_-]+/g,''); }
function findHeader_(headers, names) {
  const opts = names.map(normalizeHeader_);
  return headers.findIndex(h => opts.includes(h));
}
function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '|' + String(salt), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function toBool_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v) === '1'; }
function toIso_(v) { if (!v) return ''; const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? String(v) : d.toISOString(); }
function dateOnly_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  const s=String(v); return /^\d{4}-\d{2}-\d{2}$/.test(s)?s:s;
}
function resetAdmin() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let sh = ss.getSheetByName(CONFIG.USER_SHEET);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.USER_SHEET);

    sh.getRange(1, 1, 1, USER_HEADERS.length)
      .setValues([USER_HEADERS]);

    sh.setFrozenRows(1);
  }

  const username = 'admin';
  const password = PropertiesService.getScriptProperties().getProperty('INITIAL_ADMIN_PASSWORD');
  if (!password || password.length < 12) throw new Error('กรุณาตั้ง Script Property INITIAL_ADMIN_PASSWORD อย่างน้อย 12 ตัวอักษรก่อนเรียกฟังก์ชันนี้');
  const fullName = 'ผู้ดูแลระบบ';
  const role = 'admin';

  const salt = Utilities.getUuid();
  const passwordHash = hashPassword_(password, salt);

  const data = sh.getDataRange().getValues();

  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]).trim().toLowerCase() === 'admin'
    ) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {

    sh.getRange(foundRow, 1, 1, 8).setValues([[
      username,
      passwordHash,
      salt,
      fullName,
      role,
      true,
      new Date(),
      new Date()
    ]]);

  } else {

    sh.appendRow([
      username,
      passwordHash,
      salt,
      fullName,
      role,
      true,
      new Date(),
      new Date()
    ]);

  }

  SpreadsheetApp.flush();

  Logger.log('รีเซ็ต Admin สำเร็จ');
  Logger.log('Username: admin');
  Logger.log('รหัสผ่านแอดมินตั้งจาก Script Properties');

  return 'รีเซ็ต Admin สำเร็จ | Username: admin';
}// ฟังก์ชัน CREATE_DATABASE_NOW เดิมถูกถอดออก เพราะล้างชีตและอาจทำข้อมูลสูญหาย

/** รีเซ็ตคะแนนทุกคนเฉพาะ Admin พร้อมสำรองประวัติเดิมแบบ snapshot */
function resetAllStudentScores(token, confirmation) {
  requireRole_(token, ['admin']);
  if (confirmation !== 'RESET ALL') throw new Error('ข้อความยืนยันไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSS_();
    const tx = getOrCreateSheet_(ss, CONFIG.TX_SHEET, TX_HEADERS);
    const last = tx.getLastRow();
    const count = Math.max(0, last - 1);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    let backupName = 'Backup_Behavior_' + stamp;
    let suffix = 1;
    while (ss.getSheetByName(backupName)) backupName = 'Backup_Behavior_' + stamp + '_' + (suffix++);
    // สำรองครบทั้งหัวตารางและรายการก่อนดำเนินการล้าง
    const backup = tx.copyTo(ss);
    backup.setName(backupName);
    SpreadsheetApp.flush();
    if (backup.getLastRow() !== last) {
      throw new Error('สำรองข้อมูลไม่ครบ ระบบยกเลิกการรีเซ็ต');
    }
    if (count > 0) tx.deleteRows(2, count);
    SpreadsheetApp.flush();
    return {ok:true, backupSheet:backupName, archivedCount:count, baseScore:CONFIG.BASE_SCORE};
  } finally {
    lock.releaseLock();
  }
}

/** เพิ่มครู 29 คนครั้งเดียว โดยไม่แก้ไขรหัสผ่านหรือสิทธิ์ของบัญชีเดิม */
function setup29TeacherAccounts() {
  const entries = [
    ['Yuwariyah','2/1'], ['Nuriyah','3/2'], ['Nasilah','6/2'],
    ['Noreesan','2/1'], ['Hafis','5/1'], ['Fatimah','1/1'],
    ['Naseebah','4/1'], ['Deyoh','3/2'], ['Natchaya','4/2'],
    ['Sarilla','4/1'], ['Ilham','6/1'], ['Nadiyah','5/2'],
    ['Suhaisan','5/1'], ['Nurizzah','4/1'], ['Bassam','2/2'],
    ['Asiyah','3/1'], ['Afandi','1/2'], ['Phanuphong','3/1'],
    ['Nazrin','3/3'], ['Safwanee','2/2'], ['Faizah','4/2'],
    ['Anisah','6/1'], ['Ketsara','6/2'], ['Haisam','1/2'],
    ['Siti Rokiyoh','5/2'], ['Siti Nurul Huda','1/1'],
    ['Sapiing','3/3'], ['Halimoh','6/2'], ['Husna','2/2']
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
    const existing = new Set(readUsers_().map(u => u.username.toLowerCase()));
    const rows = [];
    const created = [];
    const skipped = [];
    const now = new Date();
    entries.forEach(function(entry) {
      const name = entry[0];
      const classCode = entry[1];
      const username = name.replace(/\s+/g, '');
      if (existing.has(username.toLowerCase())) {
        skipped.push(username);
        return;
      }
      const classParts = classCode.split('/');
      const roomCode = classParts[0] + classParts[1].padStart(2, '0');
      const password = username + roomCode;
      const salt = Utilities.getUuid();
      rows.push([username, hashPassword_(password, salt), salt, name, 'teacher', true,
        now, now, '', '', 'ม.' + classCode, '']);
      created.push(username);
      existing.add(username.toLowerCase());
    });
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, USER_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    const summary = {spreadsheetUrl: getSS_().getUrl(), sheetName: sheet.getName(), createdCount:created.length, skippedCount:skipped.length, created:created, skipped:skipped};
    Logger.log(JSON.stringify(summary));
    return summary;
  } finally {
    lock.releaseLock();
  }
}


/** ตรวจสอบบัญชีในชีตจริง โดยไม่แก้ไขข้อมูล */
function check29TeacherAccounts() {
  const ss = getSS_();
  const sh = getOrCreateSheet_(ss, CONFIG.USER_SHEET, USER_HEADERS);
  const expected = ['Yuwariyah','Nuriyah','Nasilah','Noreesan','Hafis','Fatimah','Naseebah','Deyoh','Natchaya','Sarilla','Ilham','Nadiyah','Suhaisan','Nurizzah','Bassam','Asiyah','Afandi','Phanuphong','Nazrin','Safwanee','Faizah','Anisah','Ketsara','Haisam','SitiRokiyoh','SitiNurulHuda','Sapiing','Halimoh','Husna'];
  const actual = new Set(readUsers_().map(u => u.username.toLowerCase()));
  const found = expected.filter(u => actual.has(u.toLowerCase()));
  const missing = expected.filter(u => !actual.has(u.toLowerCase()));
  const result = {spreadsheetUrl:ss.getUrl(), sheet:sh.getName(), foundCount:found.length, missingCount:missing.length, missing};
  Logger.log(JSON.stringify(result));
  return result;
}

function checkTeacherRows() {
  const ss = getSS_();
  const sheet = ss.getSheetByName('BehaviorUsers');

  if (!sheet) {
    throw new Error('ไม่พบชีต BehaviorUsers');
  }

  const data = sheet.getDataRange().getValues();

  const names = [
    'Yuwariyah', 'Nuriyah', 'Nasilah',
    'Noreesan', 'Hafis', 'Fatimah',
    'Naseebah', 'Deyoh', 'Natchaya',
    'Sarilla', 'Ilham', 'Nadiyah',
    'Suhaisan', 'Nurizzah', 'Bassam',
    'Asiyah', 'Afandi', 'Phanuphong',
    'Nazrin', 'Safwanee', 'Faizah',
    'Anisah', 'Ketsara', 'Haisam',
    'SitiRokiyoh', 'SitiNurulHuda',
    'Sapiing', 'Halimoh', 'Husna'
  ];

  const found = [];
  const missing = [];

  names.forEach(name => {
    const index = data.findIndex((row, i) =>
      i > 0 &&
      String(row[0]).trim().toLowerCase() ===
      name.toLowerCase()
    );

    if (index >= 0) {
      found.push({
        username: name,
        row: index + 1,
        role: String(data[index][4]),
        active: data[index][5]
      });
    } else {
      missing.push(name);
    }
  });

  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('Sheet URL: ' +
    ss.getUrl() + '#gid=' + sheet.getSheetId());
  Logger.log('จำนวนแถวทั้งหมด: ' + data.length);
  Logger.log('พบบัญชี: ' + found.length);
  Logger.log(JSON.stringify(found, null, 2));
  Logger.log('ไม่พบ: ' + JSON.stringify(missing));

  return { found, missing };
}




/* ===== โปรไฟล์ส่วนตัว: ใช้สิทธิ์บัญชีที่ล็อกอินเท่านั้น ===== */
function getMyProfile(token) {
  const session = requireSession_(token);
  const user = readUsers_().find(u => u.username.toLowerCase() === session.username.toLowerCase());
  if (!user) throw new Error('ไม่พบบัญชีผู้ใช้งาน');
  const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r,i) => i > 0 && String(r[0]).trim().toLowerCase() === session.username.toLowerCase());
  const fileId = idx > 0 ? String(rows[idx][12] || '').trim() : '';
  let photoData = '';
  if (fileId) {
    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const mime = blob.getContentType();
      if (/^image\/(jpeg|png|webp)$/.test(mime) && blob.getBytes().length <= 600000) {
        photoData = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    } catch (e) { console.log('โหลดรูปโปรไฟล์ไม่ได้: ' + e.message); }
  }
  return {ok:true, username:user.username, fullName:user.fullName, role:user.role,
    className:session.className || '', studentId:user.studentId || '', photoData:photoData};
}
function updateMyProfile(token, payload) {
  const session = requireSession_(token);
  payload = payload || {};
  const name = String(payload.fullName || '').trim();
  const newPassword = String(payload.newPassword || '');
  const oldPassword = String(payload.oldPassword || '');
  if (name.length < 2 || name.length > 100 || /^[=+@\-]/.test(name)) throw new Error('ชื่อที่แสดงไม่ถูกต้อง (2–100 ตัวอักษร)');
  if (newPassword && (newPassword.length < 8 || newPassword.length > 128)) throw new Error('รหัสผ่านใหม่ต้องมี 8–128 ตัวอักษร');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getOrCreateSheet_(getSS_(), CONFIG.USER_SHEET, USER_HEADERS);
    const values = sh.getDataRange().getValues();
    const idx = values.findIndex((r,i) => i > 0 && String(r[0]).trim().toLowerCase() === session.username.toLowerCase());
    if (idx < 1) throw new Error('ไม่พบบัญชี');
    const row = values[idx];
    if (newPassword && (!oldPassword || hashPassword_(oldPassword, String(row[2])) !== String(row[1]))) {
      throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    }
    let newPhotoId = '';
    if (payload.photoData) {
      const match = String(payload.photoData).match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
      if (!match) throw new Error('รูปภาพต้องเป็น JPEG, PNG หรือ WebP');
      if (match[2].length > 650000) throw new Error('รูปภาพใหญ่เกินกำหนด');
      const bytes = Utilities.base64Decode(match[2]);
      if (bytes.length > 450000) throw new Error('รูปภาพใหญ่เกิน 450 KB กรุณาเลือกรูปที่เล็กลง');
      const ext = match[1].split('/')[1].replace('jpeg','jpg');
      const blob = Utilities.newBlob(bytes,match[1],'profile-' + Utilities.getUuid() + '.' + ext);
      const file = DriveApp.createFile(blob);
      file.setDescription('ภาพโปรไฟล์ระบบคะแนนความประพฤติของ ' + session.username);
      newPhotoId = file.getId();
    }
    sh.getRange(idx + 1, 4).setValue(name);
    if (newPassword) {
      const salt = Utilities.getUuid();
      sh.getRange(idx + 1, 2, 1, 2).setValues([[hashPassword_(newPassword,salt), salt]]);
    }
    if (newPhotoId) {
      const oldId = String(row[12] || '').trim();
      sh.getRange(idx + 1, 13).setValue(newPhotoId);
      if (oldId) try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) { console.log(e.message); }
    }
    sh.getRange(idx + 1, 8).setValue(new Date());
    SpreadsheetApp.flush();
    return {ok:true, fullName:name, photoUpdated:!!newPhotoId};
  } finally { lock.releaseLock(); }
}


/**
 * Cloudflare Pages -> Google Apps Script RPC gateway.
 * Do not expose the shared secret or allow arbitrary function evaluation.
 * Configure Apps Script Script Property: RPC_SHARED_SECRET.
 * WARNING: keep original permission checks inside every server function.
 */
const RPC_ALLOWED_METHODS = Object.freeze(["login", "registerStudent", "getSession", "getDashboardData", "getStudentPortalData", "testConnection", "saveTransaction", "deleteTransaction", "listUsers", "createUser", "resetUserPassword", "deleteUser", "setUserActive", "resetAllStudentScores", "getMyProfile", "updateMyProfile"]);

function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const secret = PropertiesService.getScriptProperties().getProperty('RPC_SHARED_SECRET');
    if (!secret || secret.length < 32 || body.secret !== secret) {
      return rpcJson_({ok:false,error:'Unauthorized',code:'UNAUTHORIZED'});
    }
    const method = String(body.method || '');
    if (RPC_ALLOWED_METHODS.indexOf(method) < 0) {
      return rpcJson_({ok:false,error:'Unsupported method',code:'FORBIDDEN'});
    }
    if (!Array.isArray(body.args) || body.args.length > 4) {
      return rpcJson_({ok:false,error:'Invalid arguments',code:'BAD_REQUEST'});
    }
    const handlers = {
      login: login, registerStudent: registerStudent,
      getSession: getSession, getDashboardData: getDashboardData,
      getStudentPortalData: getStudentPortalData, testConnection: testConnection,
      saveTransaction: saveTransaction, deleteTransaction: deleteTransaction,
      listUsers: listUsers, createUser: createUser,
      resetUserPassword: resetUserPassword, deleteUser: deleteUser,
      setUserActive: setUserActive, resetAllStudentScores: resetAllStudentScores,
      getMyProfile: getMyProfile, updateMyProfile: updateMyProfile
    };
    const result = handlers[method].apply(null, body.args);
    return rpcJson_({ok:true,result:result});
  } catch (err) {
    // Do not disclose stack traces or internal IDs to visitors
    return rpcJson_({ok:false,error:String(err && err.message || 'Server error')});
  }
}
function rpcJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
