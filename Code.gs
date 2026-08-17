const APP = {
  NAME: 'ABOUT US STAFF',
  TZ: 'Asia/Bangkok',
  SHEETS: {
    SETTINGS: 'Settings',
    EMPLOYEES: 'Employees',
    DEVICES: 'PendingDevices',
    ATTENDANCE: 'Attendance',
    SCHEDULE: 'Schedule',
    AUDIT: 'AuditLog'
  },
  RADIUS_M: 15,
  MAX_ACCURACY_M: 20,
  SHIFT_START_HOUR: 18,
  STAFF_TOKEN_DAYS: 30,
  ADMIN_TOKEN_HOURS: 12,
  POLL_TTL_SEC: 120
};

const HEADERS = {
  Settings: ['Key','Value','UpdatedAt'],
  Employees: ['EmployeeId','Name','PINHash','Active','DeviceHash','DeviceLabel','DeviceApprovedAt','DefaultStart','DefaultEnd','CreatedAt','UpdatedAt'],
  PendingDevices: ['RequestId','EmployeeId','Name','DeviceHash','DeviceLabel','RequestedAt','Status','ReviewedAt'],
  Attendance: ['AttendanceId','ShiftKey','EmployeeId','Name','ClockIn','ClockOut','WorkMinutes','InLat','InLng','InDistanceM','InAccuracyM','OutLat','OutLng','OutDistanceM','OutAccuracyM','DeviceHash','Status','Note','UpdatedAt'],
  Schedule: ['ShiftKey','EmployeeId','Name','WorkStatus','StartTime','EndTime','Note','UpdatedAt'],
  AuditLog: ['Timestamp','ActorType','ActorId','Action','Detail']
};

/**
 * RUN THIS ONCE from a Google Sheet-bound Apps Script project.
 * IMPORTANT: change FIRST_ADMIN_PIN before running for the first time.
 */
function setupSystem() {
  const FIRST_ADMIN_PIN = '2580'; // <-- CHANGE THIS BEFORE FIRST RUN
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('กรุณาสร้าง Apps Script จาก Google Sheet แล้วรัน setupSystem() อีกครั้ง');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('APP_SECRET')) props.setProperty('APP_SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('ADMIN_PIN_HASH')) props.setProperty('ADMIN_PIN_HASH', pinHash_('ADMIN', FIRST_ADMIN_PIN));

  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));

  upsertSetting_('STORE_NAME', 'ABOUT US');
  upsertSetting_('STORE_LAT', '');
  upsertSetting_('STORE_LNG', '');
  upsertSetting_('RADIUS_M', String(APP.RADIUS_M));
  upsertSetting_('MAX_ACCURACY_M', String(APP.MAX_ACCURACY_M));
  upsertSetting_('SHIFT_START_HOUR', String(APP.SHIFT_START_HOUR));

  ss.setSpreadsheetTimeZone(APP.TZ);
  logAudit_('SYSTEM', 'SETUP', 'SETUP_SYSTEM', 'สร้างโครงสร้าง ABOUT US STAFF');
  return 'SETUP สำเร็จ — เปลี่ยน Admin PIN ในหน้า Admin หลัง Deploy';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'health';
  if (action === 'poll') return jsonp_(p.callback, pollResponse_(p.requestId));
  if (action === 'health') return jsonp_(p.callback, {ok:true, app:APP.NAME, version:'MVP-1.0.0', serverTime:new Date().toISOString()});
  return jsonp_(p.callback, {ok:false, error:'GET route ไม่ถูกต้อง'});
}

function doPost(e) {
  let requestId = '';
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    requestId = String(body.requestId || '');
    if (!requestId) throw new Error('ไม่มี requestId');
    const result = apiRoute_(String(body.action || ''), body.payload || {});
    cacheResponse_(requestId, {ok:true, data:result});
  } catch (err) {
    if (requestId) cacheResponse_(requestId, {ok:false, error:String(err && err.message ? err.message : err)});
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function apiRoute_(action, payload) {
  switch (action) {
    case 'staffLogin': return staffLogin_(payload);
    case 'staffStatus': return staffStatus_(payload);
    case 'staffLocationCheck': return staffLocationCheck_(payload);
    case 'clockIn': return clockAction_(payload, 'IN');
    case 'clockOut': return clockAction_(payload, 'OUT');
    case 'adminLogin': return adminLogin_(payload);
    case 'adminDashboard': return adminDashboard_(payload);
    case 'adminEmployees': return adminEmployees_(payload);
    case 'adminSaveEmployee': return adminSaveEmployee_(payload);
    case 'adminPendingDevices': return adminPendingDevices_(payload);
    case 'adminApproveDevice': return adminApproveDevice_(payload);
    case 'adminRejectDevice': return adminRejectDevice_(payload);
    case 'adminScheduleWeek': return adminScheduleWeek_(payload);
    case 'adminSaveSchedule': return adminSaveSchedule_(payload);
    case 'adminSettings': return adminSettings_(payload);
    case 'adminSaveStoreLocation': return adminSaveStoreLocation_(payload);
    case 'adminChangePin': return adminChangePin_(payload);
    default: throw new Error('API route ไม่ถูกต้อง');
  }
}

// ---------- STAFF AUTH ----------
function staffLogin_(p) {
  const employeeId = cleanId_(p.employeeId);
  const pin = String(p.pin || '');
  const deviceId = String(p.deviceId || '');
  const deviceLabel = cleanText_(p.deviceLabel || '', 120);
  if (!employeeId || !pin || !deviceId) throw new Error('กรอกข้อมูล Login ให้ครบ');

  const emp = findEmployee_(employeeId);
  if (!emp || !isTrue_(emp.Active)) throw new Error('ไม่พบพนักงาน หรือบัญชีถูกปิด');
  if (emp.PINHash !== pinHash_(employeeId, pin)) throw new Error('รหัสพนักงานหรือ PIN ไม่ถูกต้อง');

  const deviceHash = hash_('device:' + deviceId);
  if (!emp.DeviceHash) {
    createOrRefreshDeviceRequest_(emp, deviceHash, deviceLabel);
    return {status:'PENDING_APPROVAL', message:'ส่งคำขอผูกเครื่องแล้ว กรุณาให้ Admin อนุมัติเครื่องนี้'};
  }
  if (emp.DeviceHash !== deviceHash) {
    createOrRefreshDeviceRequest_(emp, deviceHash, deviceLabel);
    return {status:'NEW_DEVICE', message:'เครื่องนี้ยังไม่ได้รับอนุมัติ ระบบส่งคำขอไปที่ Admin แล้ว'};
  }

  const token = createToken_({role:'staff', sub:employeeId, deviceHash}, APP.STAFF_TOKEN_DAYS * 24 * 60 * 60 * 1000);
  logAudit_('STAFF', employeeId, 'LOGIN', deviceLabel);
  return {status:'OK', token, employee:publicEmployee_(emp)};
}

function staffStatus_(p) {
  const auth = requireStaff_(p.token);
  const emp = auth.employee;
  const shiftKey = currentShiftKey_();
  const attendance = findAttendance_(shiftKey, emp.EmployeeId);
  const schedule = findSchedule_(shiftKey, emp.EmployeeId);
  return {
    employee: publicEmployee_(emp),
    shiftKey,
    schedule: schedule ? publicSchedule_(schedule) : defaultSchedule_(emp),
    attendance: attendance ? publicAttendance_(attendance) : null,
    serverTime: formatDateTime_(new Date()),
    radiusM: APP.RADIUS_M,
    maxAccuracyM: APP.MAX_ACCURACY_M,
    storeConfigured: !!(getSetting_('STORE_LAT') && getSetting_('STORE_LNG'))
  };
}

function staffLocationCheck_(p) {
  requireStaff_(p.token);
  return validateLocation_(p.location, false);
}

function clockAction_(p, mode) {
  const auth = requireStaff_(p.token);
  const emp = auth.employee;
  const loc = validateLocation_(p.location, true);
  const now = new Date();
  const shiftKey = currentShiftKey_(now);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(APP.SHEETS.ATTENDANCE);
    const values = sh.getDataRange().getValues();
    const headers = values[0] || HEADERS.Attendance;
    let rowIndex = -1;
    let rec = null;
    for (let i = 1; i < values.length; i++) {
      const obj = rowToObj_(headers, values[i]);
      if (String(obj.ShiftKey) === shiftKey && cleanId_(obj.EmployeeId) === emp.EmployeeId) {
        rowIndex = i + 1;
        rec = obj;
        break;
      }
    }

    if (mode === 'IN') {
      if (rec && rec.ClockIn) throw new Error('วันนี้ลงเวลาเข้างานแล้ว');
      const attendanceId = Utilities.getUuid();
      const row = [
        attendanceId, shiftKey, emp.EmployeeId, emp.Name, now, '', '',
        loc.lat, loc.lng, round1_(loc.distanceM), round1_(loc.accuracyM),
        '', '', '', '', auth.payload.deviceHash, 'WORKING', '', now
      ];
      sh.appendRow(row);
      logAudit_('STAFF', emp.EmployeeId, 'CLOCK_IN', JSON.stringify({shiftKey,distanceM:loc.distanceM,accuracyM:loc.accuracyM}));
      return {mode:'IN', time:formatTime_(now), shiftKey, distanceM:round1_(loc.distanceM), accuracyM:round1_(loc.accuracyM), attendance:publicAttendance_(rowToObj_(headers,row))};
    }

    if (!rec || !rec.ClockIn) throw new Error('ยังไม่มีเวลาเข้างานของรอบนี้');
    if (rec.ClockOut) throw new Error('วันนี้ลงเวลาออกงานแล้ว');
    const clockIn = new Date(rec.ClockIn);
    const workMinutes = Math.max(0, Math.round((now.getTime() - clockIn.getTime()) / 60000));
    const updates = {
      ClockOut: now,
      WorkMinutes: workMinutes,
      OutLat: loc.lat,
      OutLng: loc.lng,
      OutDistanceM: round1_(loc.distanceM),
      OutAccuracyM: round1_(loc.accuracyM),
      Status: 'DONE',
      UpdatedAt: now
    };
    setRowFields_(sh, rowIndex, headers, updates);
    Object.assign(rec, updates);
    logAudit_('STAFF', emp.EmployeeId, 'CLOCK_OUT', JSON.stringify({shiftKey,workMinutes,distanceM:loc.distanceM,accuracyM:loc.accuracyM}));
    return {mode:'OUT', time:formatTime_(now), shiftKey, workMinutes, distanceM:round1_(loc.distanceM), accuracyM:round1_(loc.accuracyM), attendance:publicAttendance_(rec)};
  } finally {
    lock.releaseLock();
  }
}

// ---------- ADMIN AUTH ----------
function adminLogin_(p) {
  const pin = String(p.pin || '');
  if (!pin) throw new Error('กรอก Admin PIN');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_PIN_HASH') !== pinHash_('ADMIN', pin)) throw new Error('Admin PIN ไม่ถูกต้อง');
  const token = createToken_({role:'admin', sub:'ADMIN'}, APP.ADMIN_TOKEN_HOURS * 60 * 60 * 1000);
  logAudit_('ADMIN', 'ADMIN', 'LOGIN', 'Admin login');
  return {token, expiresHours:APP.ADMIN_TOKEN_HOURS};
}

function adminChangePin_(p) {
  requireAdmin_(p.token);
  const currentPin = String(p.currentPin || '');
  const newPin = String(p.newPin || '');
  if (!/^\d{4,8}$/.test(newPin)) throw new Error('PIN ใหม่ต้องเป็นตัวเลข 4-8 หลัก');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_PIN_HASH') !== pinHash_('ADMIN', currentPin)) throw new Error('PIN ปัจจุบันไม่ถูกต้อง');
  props.setProperty('ADMIN_PIN_HASH', pinHash_('ADMIN', newPin));
  logAudit_('ADMIN', 'ADMIN', 'CHANGE_PIN', 'เปลี่ยน Admin PIN');
  return {message:'เปลี่ยน Admin PIN สำเร็จ'};
}

// ---------- ADMIN: EMPLOYEES ----------
function adminEmployees_(p) {
  requireAdmin_(p.token);
  return getRows_(APP.SHEETS.EMPLOYEES).map(publicEmployeeAdmin_);
}

function adminSaveEmployee_(p) {
  requireAdmin_(p.token);
  const employeeId = cleanId_(p.employeeId);
  const name = cleanText_(p.name, 80);
  const pin = String(p.pin || '');
  const active = p.active !== false;
  const defaultStart = validTime_(p.defaultStart || '19:00');
  const defaultEnd = validTime_(p.defaultEnd || '02:00');
  if (!employeeId || !name) throw new Error('กรอกรหัสและชื่อพนักงาน');
  if (pin && !/^\d{4,8}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4-8 หลัก');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(APP.SHEETS.EMPLOYEES);
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    let rowIndex = -1, old = null;
    for (let i=1;i<values.length;i++) {
      const obj = rowToObj_(headers, values[i]);
      if (cleanId_(obj.EmployeeId) === employeeId) { rowIndex=i+1; old=obj; break; }
    }
    const now = new Date();
    if (rowIndex > 0) {
      const updates = {Name:name, Active:active, DefaultStart:defaultStart, DefaultEnd:defaultEnd, UpdatedAt:now};
      if (pin) updates.PINHash = pinHash_(employeeId, pin);
      setRowFields_(sh, rowIndex, headers, updates);
    } else {
      if (!pin) throw new Error('พนักงานใหม่ต้องกำหนด PIN');
      sh.appendRow([employeeId,name,pinHash_(employeeId,pin),active,'','','',defaultStart,defaultEnd,now,now]);
    }
    logAudit_('ADMIN','ADMIN','SAVE_EMPLOYEE',employeeId + ' ' + name);
    return {message:'บันทึกพนักงานสำเร็จ'};
  } finally { lock.releaseLock(); }
}

// ---------- ADMIN: DEVICE APPROVAL ----------
function adminPendingDevices_(p) {
  requireAdmin_(p.token);
  return getRows_(APP.SHEETS.DEVICES).filter(r => String(r.Status) === 'PENDING').map(r => ({
    requestId:String(r.RequestId), employeeId:String(r.EmployeeId), name:String(r.Name), deviceLabel:String(r.DeviceLabel), requestedAt:formatDateTime_(r.RequestedAt)
  }));
}

function adminApproveDevice_(p) {
  requireAdmin_(p.token);
  const requestId = String(p.requestId || '');
  if (!requestId) throw new Error('ไม่มี requestId');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const dsh = getSheet_(APP.SHEETS.DEVICES);
    const dvals = dsh.getDataRange().getValues(); const dh = dvals[0];
    let req=null, drow=-1;
    for(let i=1;i<dvals.length;i++) { const o=rowToObj_(dh,dvals[i]); if(String(o.RequestId)===requestId && String(o.Status)==='PENDING'){req=o;drow=i+1;break;} }
    if(!req) throw new Error('ไม่พบคำขอ หรือคำขอนี้ถูกดำเนินการแล้ว');

    const esh = getSheet_(APP.SHEETS.EMPLOYEES);
    const evals = esh.getDataRange().getValues(); const eh=evals[0]; let erow=-1;
    for(let i=1;i<evals.length;i++){const o=rowToObj_(eh,evals[i]);if(cleanId_(o.EmployeeId)===cleanId_(req.EmployeeId)){erow=i+1;break;}}
    if(erow<0) throw new Error('ไม่พบพนักงาน');
    const now=new Date();
    setRowFields_(esh,erow,eh,{DeviceHash:req.DeviceHash,DeviceLabel:req.DeviceLabel,DeviceApprovedAt:now,UpdatedAt:now});
    setRowFields_(dsh,drow,dh,{Status:'APPROVED',ReviewedAt:now});
    // Reject other pending requests for same employee.
    for(let i=1;i<dvals.length;i++){
      const o=rowToObj_(dh,dvals[i]);
      if(i+1!==drow && cleanId_(o.EmployeeId)===cleanId_(req.EmployeeId) && String(o.Status)==='PENDING') setRowFields_(dsh,i+1,dh,{Status:'REJECTED',ReviewedAt:now});
    }
    logAudit_('ADMIN','ADMIN','APPROVE_DEVICE',req.EmployeeId+' '+req.DeviceLabel);
    return {message:'อนุมัติเครื่องของ '+req.Name+' สำเร็จ'};
  } finally { lock.releaseLock(); }
}

function adminRejectDevice_(p) {
  requireAdmin_(p.token);
  const requestId=String(p.requestId||'');
  const sh=getSheet_(APP.SHEETS.DEVICES); const vals=sh.getDataRange().getValues(); const h=vals[0];
  for(let i=1;i<vals.length;i++){const o=rowToObj_(h,vals[i]);if(String(o.RequestId)===requestId && String(o.Status)==='PENDING'){setRowFields_(sh,i+1,h,{Status:'REJECTED',ReviewedAt:new Date()});return {message:'ปฏิเสธคำขอแล้ว'};}}
  throw new Error('ไม่พบคำขอ');
}

// ---------- ADMIN: SETTINGS ----------
function adminSettings_(p) {
  requireAdmin_(p.token);
  return {
    storeName:getSetting_('STORE_NAME')||'ABOUT US',
    lat:numOrNull_(getSetting_('STORE_LAT')),
    lng:numOrNull_(getSetting_('STORE_LNG')),
    radiusM:APP.RADIUS_M,
    maxAccuracyM:APP.MAX_ACCURACY_M,
    shiftStartHour:APP.SHIFT_START_HOUR
  };
}

function adminSaveStoreLocation_(p) {
  requireAdmin_(p.token);
  const lat=Number(p.lat), lng=Number(p.lng), accuracy=Number(p.accuracy||999);
  if(!Number.isFinite(lat)||!Number.isFinite(lng)) throw new Error('พิกัดไม่ถูกต้อง');
  if(accuracy>APP.MAX_ACCURACY_M) throw new Error('GPS ยังไม่แม่นพอ (±'+Math.round(accuracy)+' ม.) กรุณาลองใหม่');
  upsertSetting_('STORE_LAT',String(lat));
  upsertSetting_('STORE_LNG',String(lng));
  logAudit_('ADMIN','ADMIN','SET_STORE_LOCATION',JSON.stringify({lat,lng,accuracy}));
  return {message:'บันทึกพิกัดร้านแล้ว',lat,lng,accuracy,radiusM:APP.RADIUS_M};
}

// ---------- ADMIN: SCHEDULE ----------
function adminScheduleWeek_(p) {
  requireAdmin_(p.token);
  const start = normalizeMonday_(String(p.weekStart || currentShiftKey_()));
  const dates = Array.from({length:7},(_,i)=>addDaysKey_(start,i));
  const employees = getRows_(APP.SHEETS.EMPLOYEES).filter(r=>isTrue_(r.Active)).map(publicEmployeeAdmin_);
  const schedules = getRows_(APP.SHEETS.SCHEDULE).filter(r=>dates.includes(String(r.ShiftKey))).map(publicSchedule_);
  return {weekStart:start,dates,employees,schedules};
}

function adminSaveSchedule_(p) {
  requireAdmin_(p.token);
  const items = Array.isArray(p.items) ? p.items : [];
  if(!items.length) throw new Error('ไม่มีตารางงานให้บันทึก');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const sh=getSheet_(APP.SHEETS.SCHEDULE);const vals=sh.getDataRange().getValues();const h=vals[0];
    const map={};for(let i=1;i<vals.length;i++){const o=rowToObj_(h,vals[i]);map[String(o.ShiftKey)+'|'+cleanId_(o.EmployeeId)]={row:i+1,obj:o};}
    const emps={};getRows_(APP.SHEETS.EMPLOYEES).forEach(e=>emps[cleanId_(e.EmployeeId)]=e);
    const now=new Date();
    items.forEach(it=>{
      const date=validDateKey_(it.shiftKey);const id=cleanId_(it.employeeId);const emp=emps[id];if(!emp)return;
      const status=['WORK','OFF','LEAVE','UNSET'].includes(String(it.workStatus))?String(it.workStatus):'UNSET';
      const start=validTime_(it.startTime||emp.DefaultStart||'19:00');const end=validTime_(it.endTime||emp.DefaultEnd||'02:00');
      const key=date+'|'+id;
      if(map[key]) setRowFields_(sh,map[key].row,h,{Name:emp.Name,WorkStatus:status,StartTime:start,EndTime:end,Note:cleanText_(it.note||'',120),UpdatedAt:now});
      else sh.appendRow([date,id,emp.Name,status,start,end,cleanText_(it.note||'',120),now]);
    });
    logAudit_('ADMIN','ADMIN','SAVE_SCHEDULE','จำนวน '+items.length+' รายการ');
    return {message:'บันทึกตารางงานสำเร็จ'};
  }finally{lock.releaseLock();}
}

// ---------- ADMIN: DASHBOARD ----------
function adminDashboard_(p) {
  requireAdmin_(p.token);
  const range=String(p.range||'today');
  let start,end;
  if(range==='today'){start=end=currentShiftKey_();}
  else if(range==='week'){start=normalizeMonday_(currentShiftKey_());end=addDaysKey_(start,6);}
  else {
    start=validDateKey_(p.startDate);end=validDateKey_(p.endDate||p.startDate);
    if(dateKeyToDate_(end)<dateKeyToDate_(start)) throw new Error('ช่วงวันที่ไม่ถูกต้อง');
    const days=Math.round((dateKeyToDate_(end)-dateKeyToDate_(start))/86400000)+1;
    if(days>62) throw new Error('เลือกช่วงได้สูงสุด 62 วันต่อครั้ง');
  }
  const dates=dateRangeKeys_(start,end);
  const employees=getRows_(APP.SHEETS.EMPLOYEES).filter(r=>isTrue_(r.Active));
  const attendance=getRows_(APP.SHEETS.ATTENDANCE).filter(r=>dates.includes(String(r.ShiftKey)));
  const schedules=getRows_(APP.SHEETS.SCHEDULE).filter(r=>dates.includes(String(r.ShiftKey)));
  const aMap={};attendance.forEach(a=>aMap[String(a.ShiftKey)+'|'+cleanId_(a.EmployeeId)]=a);
  const sMap={};schedules.forEach(s=>sMap[String(s.ShiftKey)+'|'+cleanId_(s.EmployeeId)]=s);
  const details=[];
  employees.forEach(emp=>dates.forEach(date=>{
    const key=date+'|'+cleanId_(emp.EmployeeId);const a=aMap[key];const s=sMap[key];
    const sched=s?publicSchedule_(s):defaultSchedule_(emp);
    let status='UNSET';
    if(sched.workStatus==='OFF') status='OFF';
    else if(sched.workStatus==='LEAVE') status='LEAVE';
    else if(a && a.ClockOut) status='DONE';
    else if(a && a.ClockIn) status='WORKING';
    else if(sched.workStatus==='WORK') status='NOT_IN';
    const mins=a&&a.WorkMinutes!==''?Number(a.WorkMinutes||0):(a&&a.ClockIn&&!a.ClockOut&&date===currentShiftKey_()?Math.max(0,Math.round((Date.now()-new Date(a.ClockIn).getTime())/60000)):0);
    details.push({
      date,employeeId:cleanId_(emp.EmployeeId),name:String(emp.Name),workStatus:sched.workStatus,startTime:sched.startTime,endTime:sched.endTime,
      clockIn:a&&a.ClockIn?formatTime_(a.ClockIn):'',clockOut:a&&a.ClockOut?formatTime_(a.ClockOut):'',workMinutes:mins,
      inDistanceM:a&&a.InDistanceM!==''?Number(a.InDistanceM):null,inAccuracyM:a&&a.InAccuracyM!==''?Number(a.InAccuracyM):null,
      outDistanceM:a&&a.OutDistanceM!==''?Number(a.OutDistanceM):null,outAccuracyM:a&&a.OutAccuracyM!==''?Number(a.OutAccuracyM):null,status
    });
  }));
  const todayDetails=details.filter(d=>d.date===currentShiftKey_());
  const summary={
    employees:employees.length,
    working:todayDetails.filter(d=>d.status==='WORKING').length,
    notIn:todayDetails.filter(d=>d.status==='NOT_IN').length,
    off:todayDetails.filter(d=>d.status==='OFF'||d.status==='LEAVE').length,
    totalMinutes:details.reduce((s,d)=>s+Number(d.workMinutes||0),0)
  };
  return {range,startDate:start,endDate:end,dates,summary,details,serverTime:formatDateTime_(new Date())};
}

// ---------- LOCATION ----------
function validateLocation_(loc, strict) {
  loc=loc||{};const lat=Number(loc.lat),lng=Number(loc.lng),accuracy=Number(loc.accuracy);
  if(!Number.isFinite(lat)||!Number.isFinite(lng)||!Number.isFinite(accuracy)) throw new Error('อ่าน GPS ไม่สำเร็จ');
  const rawStoreLat=getSetting_('STORE_LAT'), rawStoreLng=getSetting_('STORE_LNG');
  if(rawStoreLat===''||rawStoreLng==='') throw new Error('Admin ยังไม่ได้ตั้งพิกัดร้าน');
  const storeLat=Number(rawStoreLat),storeLng=Number(rawStoreLng);
  if(!Number.isFinite(storeLat)||!Number.isFinite(storeLng)) throw new Error('พิกัดร้านไม่ถูกต้อง');
  const distance=haversineM_(storeLat,storeLng,lat,lng);
  const accuracyOk=accuracy<=APP.MAX_ACCURACY_M;const inside=distance<=APP.RADIUS_M;
  if(strict && !accuracyOk) throw new Error('GPS ยังไม่แม่นพอ (±'+Math.round(accuracy)+' ม.) กรุณารอแล้วลองใหม่');
  if(strict && !inside) throw new Error('อยู่นอกรัศมีร้าน '+Math.round(distance)+' ม. (อนุญาตไม่เกิน '+APP.RADIUS_M+' ม.)');
  return {lat,lng,accuracyM:accuracy,distanceM:distance,inside,accuracyOk,radiusM:APP.RADIUS_M,maxAccuracyM:APP.MAX_ACCURACY_M};
}

function haversineM_(lat1,lon1,lat2,lon2){const R=6371000,toRad=x=>x*Math.PI/180;const dLat=toRad(lat2-lat1),dLon=toRad(lon2-lon1);const a=Math.sin(dLat/2)**2+Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)**2;return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));}

// ---------- TOKEN / SECURITY ----------
function createToken_(payload, ttlMs){const obj=Object.assign({},payload,{iat:Date.now(),exp:Date.now()+ttlMs,nonce:Utilities.getUuid()});const b64=Utilities.base64EncodeWebSafe(JSON.stringify(obj),Utilities.Charset.UTF_8);const sig=hash_('token:'+b64);return b64+'.'+sig;}
function verifyToken_(token,role){const parts=String(token||'').split('.');if(parts.length!==2)throw new Error('Session ไม่ถูกต้อง กรุณา Login ใหม่');const b64=parts[0],sig=parts[1];if(hash_('token:'+b64)!==sig)throw new Error('Session ไม่ถูกต้อง');let payload;try{payload=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(b64)).getDataAsString());}catch(e){throw new Error('Session ไม่ถูกต้อง');}if(payload.role!==role||Number(payload.exp)<Date.now())throw new Error('Session หมดอายุ กรุณา Login ใหม่');return payload;}
function requireAdmin_(token){return verifyToken_(token,'admin');}
function requireStaff_(token){const payload=verifyToken_(token,'staff');const emp=findEmployee_(payload.sub);if(!emp||!isTrue_(emp.Active))throw new Error('บัญชีพนักงานถูกปิด');if(!emp.DeviceHash||emp.DeviceHash!==payload.deviceHash)throw new Error('อุปกรณ์นี้ไม่ได้รับอนุญาต กรุณา Login ใหม่');return {payload,employee:emp};}
function pinHash_(id,pin){return hash_('pin:'+cleanId_(id)+':'+String(pin));}
function hash_(value){const secret=PropertiesService.getScriptProperties().getProperty('APP_SECRET')||'CHANGE_ME';const bytes=Utilities.computeHmacSha256Signature(String(value),secret);return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}

// ---------- DEVICE REQUEST ----------
function createOrRefreshDeviceRequest_(emp,deviceHash,deviceLabel){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const sh=getSheet_(APP.SHEETS.DEVICES);const vals=sh.getDataRange().getValues();const h=vals[0];const now=new Date();
    for(let i=1;i<vals.length;i++){const o=rowToObj_(h,vals[i]);if(cleanId_(o.EmployeeId)===cleanId_(emp.EmployeeId)&&String(o.DeviceHash)===deviceHash&&String(o.Status)==='PENDING'){setRowFields_(sh,i+1,h,{DeviceLabel:deviceLabel,RequestedAt:now});return;}}
    sh.appendRow([Utilities.getUuid(),emp.EmployeeId,emp.Name,deviceHash,deviceLabel,now,'PENDING','']);
    logAudit_('STAFF',emp.EmployeeId,'REQUEST_DEVICE',deviceLabel);
  }finally{lock.releaseLock();}
}

// ---------- SHEET HELPERS ----------
function getSs_(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw new Error('ยังไม่ได้รัน setupSystem()');return SpreadsheetApp.openById(id);}
function getSheet_(name){const sh=getSs_().getSheetByName(name);if(!sh)throw new Error('ไม่พบชีต '+name);return sh;}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);else sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold');return sh;}
function getRows_(name){const vals=getSheet_(name).getDataRange().getValues();if(vals.length<2)return[];const h=vals[0];return vals.slice(1).filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>rowToObj_(h,r));}
function rowToObj_(headers,row){const o={};headers.forEach((h,i)=>o[String(h)]=row[i]);return o;}
function setRowFields_(sh,rowIndex,headers,updates){const row=sh.getRange(rowIndex,1,1,headers.length).getValues()[0];headers.forEach((h,i)=>{if(Object.prototype.hasOwnProperty.call(updates,h))row[i]=updates[h];});sh.getRange(rowIndex,1,1,headers.length).setValues([row]);}
function findEmployee_(id){const cid=cleanId_(id);return getRows_(APP.SHEETS.EMPLOYEES).find(r=>cleanId_(r.EmployeeId)===cid)||null;}
function findAttendance_(date,id){const cid=cleanId_(id);return getRows_(APP.SHEETS.ATTENDANCE).find(r=>String(r.ShiftKey)===date&&cleanId_(r.EmployeeId)===cid)||null;}
function findSchedule_(date,id){const cid=cleanId_(id);return getRows_(APP.SHEETS.SCHEDULE).find(r=>String(r.ShiftKey)===date&&cleanId_(r.EmployeeId)===cid)||null;}

// ---------- SETTINGS ----------
function getSetting_(key){const row=getRows_(APP.SHEETS.SETTINGS).find(r=>String(r.Key)===key);return row?String(row.Value):'';}
function upsertSetting_(key,value){const sh=getSheet_(APP.SHEETS.SETTINGS);const vals=sh.getDataRange().getValues();const h=vals[0]||HEADERS.Settings;for(let i=1;i<vals.length;i++){if(String(vals[i][0])===key){sh.getRange(i+1,1,1,3).setValues([[key,value,new Date()]]);return;}}sh.appendRow([key,value,new Date()]);}

// ---------- DATE / PUBLIC ----------
function currentShiftKey_(date){date=date||new Date();const localDate=Utilities.formatDate(date,APP.TZ,'yyyy-MM-dd');const hour=Number(Utilities.formatDate(date,APP.TZ,'H'));return hour>=APP.SHIFT_START_HOUR?localDate:addDaysKey_(localDate,-1);}
function dateKeyToDate_(key){validDateKey_(key);return new Date(key+'T12:00:00+07:00');}
function addDaysKey_(key,days){const d=dateKeyToDate_(key);d.setUTCDate(d.getUTCDate()+Number(days));return Utilities.formatDate(d,APP.TZ,'yyyy-MM-dd');}
function normalizeMonday_(key){const d=dateKeyToDate_(validDateKey_(key));const dow=d.getUTCDay()||7;return addDaysKey_(key,1-dow);}
function dateRangeKeys_(start,end){const out=[];let k=start;while(k<=end){out.push(k);k=addDaysKey_(k,1);if(out.length>70)break;}return out;}
function validDateKey_(v){const s=String(v||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(s))throw new Error('วันที่ไม่ถูกต้อง');return s;}
function validTime_(v){const s=String(v||'');if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(s))throw new Error('เวลาไม่ถูกต้อง: '+s);return s;}
function formatTime_(d){return d?Utilities.formatDate(new Date(d),APP.TZ,'HH:mm'):'';}
function formatDateTime_(d){return d?Utilities.formatDate(new Date(d),APP.TZ,'yyyy-MM-dd HH:mm:ss'):'';}
function defaultSchedule_(emp){return {shiftKey:currentShiftKey_(),employeeId:cleanId_(emp.EmployeeId),name:String(emp.Name),workStatus:'UNSET',startTime:String(emp.DefaultStart||'19:00'),endTime:String(emp.DefaultEnd||'02:00'),note:''};}
function publicSchedule_(s){return {shiftKey:String(s.ShiftKey),employeeId:cleanId_(s.EmployeeId),name:String(s.Name),workStatus:String(s.WorkStatus||'UNSET'),startTime:String(s.StartTime||''),endTime:String(s.EndTime||''),note:String(s.Note||'')};}
function publicAttendance_(a){return {shiftKey:String(a.ShiftKey||''),clockIn:a.ClockIn?formatTime_(a.ClockIn):'',clockOut:a.ClockOut?formatTime_(a.ClockOut):'',workMinutes:a.WorkMinutes===''?0:Number(a.WorkMinutes||0),status:String(a.Status||''),inDistanceM:a.InDistanceM===''?null:Number(a.InDistanceM),inAccuracyM:a.InAccuracyM===''?null:Number(a.InAccuracyM),outDistanceM:a.OutDistanceM===''?null:Number(a.OutDistanceM),outAccuracyM:a.OutAccuracyM===''?null:Number(a.OutAccuracyM)};}
function publicEmployee_(e){return {employeeId:cleanId_(e.EmployeeId),name:String(e.Name),defaultStart:String(e.DefaultStart||'19:00'),defaultEnd:String(e.DefaultEnd||'02:00'),deviceLabel:String(e.DeviceLabel||'')};}
function publicEmployeeAdmin_(e){return {employeeId:cleanId_(e.EmployeeId),name:String(e.Name),active:isTrue_(e.Active),deviceBound:!!e.DeviceHash,deviceLabel:String(e.DeviceLabel||''),defaultStart:String(e.DefaultStart||'19:00'),defaultEnd:String(e.DefaultEnd||'02:00'),updatedAt:formatDateTime_(e.UpdatedAt||e.CreatedAt)};}

// ---------- MISC ----------
function cleanId_(v){return String(v||'').trim().toUpperCase().replace(/[^A-Z0-9_-]/g,'').slice(0,24);}
function cleanText_(v,max){return String(v||'').replace(/[<>]/g,'').trim().slice(0,max||200);}
function isTrue_(v){return v===true||String(v).toUpperCase()==='TRUE'||String(v)==='1';}
function numOrNull_(v){if(v===''||v===null||v===undefined)return null;const n=Number(v);return Number.isFinite(n)?n:null;}
function round1_(n){return Math.round(Number(n)*10)/10;}
function logAudit_(type,id,action,detail){try{getSheet_(APP.SHEETS.AUDIT).appendRow([new Date(),type,id,action,cleanText_(detail,500)]);}catch(e){}}

// ---------- POST/POLL TRANSPORT ----------
function cacheResponse_(requestId,obj){CacheService.getScriptCache().put('rpc:'+requestId,JSON.stringify(obj),APP.POLL_TTL_SEC);}
function pollResponse_(requestId){if(!requestId)return{ok:false,error:'ไม่มี requestId'};const raw=CacheService.getScriptCache().get('rpc:'+requestId);return raw?JSON.parse(raw):{ok:false,pending:true};}
function jsonp_(callback,obj){const cb=/^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(String(callback||''))?String(callback):'callback';return ContentService.createTextOutput(cb+'('+JSON.stringify(obj)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);}
