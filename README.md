School Behavior System — Cloudflare Pages + Apps Script Backend
> **สถานะ:** โค้ดพร้อมสำหรับการทดสอบและตั้งค่า Deployment แต่ **ยังไม่ได้ทดสอบกับ Google Sheets จริง**
สถาปัตยกรรมระยะเปลี่ยนผ่าน
`public/index.html` — หน้าเดิมทั้งหมด
`public/app.js` — ตัวเชื่อม `google.script.run` เดิมไป `/api/rpc`
`public/style.css` — ที่ว่างสำหรับแยกสไตล์ภายหลัง (CSS เดิมยังอยู่ใน HTML)
`functions/api/rpc.js` — Cloudflare Pages Function ส่งต่อคำขอไป Apps Script อย่างจำกัดเมธอด
`functions/api/health.js` — API สำหรับทดสอบ
`apps-script/Code.gs` — Backend เดิมพร้อม `doPost` และการป้องกันสมัครซ้ำ
นี่ไม่ใช่การแปลง Code.gs เป็น Workers แบบ native: โค้ด Google Apps Script ยังคงรันบน Google สำหรับใช้ Sheets/LockService/CacheService โดย Cloudflare เป็นโฮสต์หน้าเว็บและ API gateway ช่วยหลีกเลี่ยงการย้ายข้อมูลขณะนี้
ติดตั้ง — Google Apps Script ก่อน
สำรองไฟล์ Google Sheets และโปรเจกต์ Apps Script เดิม (แนะนำใช้ข้อมูลสำเนาสำหรับทดสอบ)
เปิด Apps Script โครงการที่เชื่อมชีตเดิม แล้วแทน `Code.gs` ด้วยไฟล์ `apps-script/Code.gs` (หรืออัปเดตเนื้อหาตามไฟล์นี้)
ไป `Project Settings > Script Properties` เพิ่ม `RPC_SHARED_SECRET` โดยสร้างค่าลับแบบสุ่มอย่างน้อย 32 ตัวอักษร อย่าเก็บค่าลับไว้ใน GitHub
ถ้าเป็นการตั้งระบบใหม่เท่านั้น ให้เพิ่ม `INITIAL_ADMIN_PASSWORD` รหัสผ่านสุ่มอย่างน้อย 12 ตัวอักษร แล้วจึงเรียก `setupSystem()`; ถ้าระบบเดิมมีบัญชีแอดมินแล้ว ไม่ต้องเรียกฟังก์ชันตั้งค่าซ้ำ
`Deploy > New deployment > Web app` เลือก `Execute as: Me` และ `Who has access: Anyone` (หน้า doPost มีการตรวจ secret เพิ่ม; ต้องตั้ง secret ก่อนเปิด)
คัดลอก Web app URL ที่ลงท้ายด้วย `/exec`
ติดตั้ง — GitHub / Cloudflare
สร้าง GitHub Repository แบบ Private ชื่อ `school-behavior-system`
อัปโหลด เฉพาะโฟลเดอร์ `public/`, `functions/`, `.gitignore`, `README.md`  อย่าอัปโหลด `apps-script/Code.gs` ซึ่งมีค่าระบุชีตและส่วนของบัญชีผู้ใช้เดิม (เก็บไว้ใน Apps Script ส่วนตัว)
Cloudflare Dashboard > Workers & Pages > Create application > Pages > Connect to Git > เลือก Repo
Framework preset: None; Build command: เว้นว่าง; Build output directory: `public`
ที่โปรเจกต์ Cloudflare > Settings > Variables and Secrets เพิ่ม Secret สองค่า:
`GAS_WEBAPP_URL`: URL ของ Apps Script Web app (`https://script.google.com/macros/s/.../exec`)
`RPC_SHARED_SECRET`: ต้องตรงกับค่าใน Apps Script Script Properties
Deploy ใหม่ แล้วเปิด `https://ชื่อโปรเจกต์.pages.dev/api/health` ควรเห็น `ok: true`
ทดสอบ login, register, duplicate registration, teacher restrictions, scores, admin และ reset บน สำเนาฐานข้อมูล ก่อนเผยแพร่จริง
ความปลอดภัย / ข้อจำกัด
เดิมมี รหัสผ่าน admin เริ่มต้นฝังในไฟล์ จึงลบค่ารหัสผ่านเริ่มต้นแบบ hardcoded ในเวอร์ชันนี้ ต้องเปลี่ยนรหัสผ่านของบัญชีที่ใช้งานจริงด้วย
API gateway เป็น endpoint สาธารณะ ซึ่งป้องกันการเรียกเมธอดที่ไม่อนุญาต แต่ยังต้องตั้ง Rate limiting / Turnstile / Monitoring ที่ Cloudflare เพื่อป้องกัน brute force และ abuse
การตรวจสิทธิ์ดำเนินการใน Apps Script เช่นเดิม ผู้ใช้ไม่มีสิทธิ์เข้าถึง Google Sheets โดยตรง
สำหรับ Production ระยะยาว แนะนำย้ายบัญชีและคะแนนสู่ D1 โดยใช้ Unique constraints, transaction, audit logs และ backup policy; ไม่ใช่เปลี่ยนโฮสต์อย่างเดียว
Cloudflare มีขีดจำกัด request body และ CPU; การอัปโหลดรูปใหญ่ผ่าน data URL ควรทดสอบเป็นพิเศษ
โหมด Apps Script `Anyone` ขึ้นอยู่กับนโยบาย Google Workspace ของโรงเรียน
