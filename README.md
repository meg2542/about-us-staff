# ABOUT US STAFF — MVP 1.0

ระบบลงเวลาเข้า/ออกพนักงานผ่านมือถือ

## เงื่อนไขหลัก
- ไม่มี QR
- GPS ต้องอยู่ในรัศมี 15 เมตรจากพิกัดร้าน
- GPS Accuracy ต้องไม่เกิน 20 เมตร
- Login ด้วยรหัสพนักงาน + PIN
- ผูก Browser/มือถือ และ Admin ต้องอนุมัติครั้งแรก
- บันทึกเวลาจาก Server
- รอบวันทำงาน 18:00 ถึง 17:59 วันถัดไป
- Admin ดู วันนี้ / สัปดาห์ / เลือกช่วงวันที่
- Admin จัดตาราง ทำงาน / หยุด / ลา

## ไฟล์
- `Code.gs` — Apps Script Backend
- `appsscript.json` — ตั้ง Timezone Asia/Bangkok
- `config.js` — ใส่ Apps Script `/exec` URL
- `index.html` — หน้า Staff
- `admin.html` — หน้า Admin

## วิธีติดตั้ง

### A) Google Sheet + Apps Script Backend
1. สร้าง Google Sheet ใหม่ชื่อ `ABOUT US STAFF DATA`
2. ไป Extensions > Apps Script
3. ลบ Code.gs เดิม แล้ววาง `Code.gs`
4. Project Settings > Show appsscript.json แล้วแทนด้วยไฟล์ `appsscript.json`
5. ใน `Code.gs` หา `FIRST_ADMIN_PIN = '2580'` แล้วเปลี่ยนเป็น PIN ที่ต้องการ
6. กด Run ฟังก์ชัน `setupSystem()` หนึ่งครั้ง และอนุญาตสิทธิ์
7. Deploy > New deployment > Web app
8. Execute as: Me
9. Who has access: Anyone
10. Deploy แล้วคัดลอก URL ที่ลงท้าย `/exec`

### B) GitHub Pages Frontend
1. สร้าง repo เช่น `about-us-staff`
2. อัปโหลด `index.html`, `admin.html`, `config.js`
3. เปิด `config.js` แล้วแทน `PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE` ด้วย URL `/exec`
4. Settings > Pages > Deploy from branch > main / root
5. ลิงก์ Staff = `https://USERNAME.github.io/about-us-staff/`
6. ลิงก์ Admin = `https://USERNAME.github.io/about-us-staff/admin.html`

## ตั้งค่าครั้งแรก
1. เปิดหน้า Admin และ Login ด้วย Admin PIN
2. ไป `ตั้งค่า` แล้วกด `ใช้ตำแหน่งปัจจุบันเป็นพิกัดร้าน` ขณะยืนตรงจุดที่ต้องการเป็นศูนย์กลางร้าน
3. ไป `พนักงาน` เพิ่มรหัส/ชื่อ/PIN ของพนักงาน
4. พนักงานเปิดหน้า Staff และ Login ครั้งแรก
5. หน้า Staff จะแจ้งว่ารอ Admin อนุมัติ
6. Admin ไป `อนุมัติเครื่อง` > อนุมัติ
7. พนักงาน Login อีกครั้ง จากนั้นระบบจะจำเครื่องไว้ประมาณ 30 วัน

## หมายเหตุความปลอดภัย
การผูกเครื่องใน MVP ใช้ browser-generated device ID ที่เก็บใน localStorage ซึ่งช่วยกันการฝาก Login แบบทั่วไป แต่ไม่ใช่ Hardware ID และไม่สามารถป้องกันผู้ใช้ที่มีความรู้เทคนิคสูงจากการคัดลอกข้อมูล browser ได้ 100% หากต้องการระดับสูงขึ้น รุ่นถัดไปควรใช้ Passkey/WebAuthn หรือ Native App.
