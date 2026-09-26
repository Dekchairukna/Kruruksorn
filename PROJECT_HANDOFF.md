# สรุปโปรเจกต์สมุดคะแนน → .accdb (สำหรับให้ AI แชทใหม่อ่านต่อ)

อ่านไฟล์นี้ก่อนเริ่มงานต่อ แล้วสำรวจไฟล์ในโฟลเดอร์ประกอบ

## เป้าหมาย
ครู (พศิน พิมพ์คำไหล, รหัสครู 307, ร.ร.เปรมติณสูลานนท์ อ.น้ำพอง ขอนแก่น) ต้องการคีย์คะแนนนักเรียนบนเว็บ
(ทุกเครื่อง ไม่ต้องลงโปรแกรม) แล้วได้ไฟล์ BookMark **.accdb** กลับไปส่งระบบโรงเรียน (บังคับ .accdb เท่านั้น)

## สิ่งที่ทำเสร็จแล้ว
1. **กู้ข้อมูล + ถอดรหัส** ไฟล์ .accdb ของ BookMark ทั้ง 9 วิชา (พ21201, พ22201, พ23201, พ31201, พ32201, พ33201, ว21102, ว22102, ว23102)
2. **เว็บ Kruruksorn** (Flask, ไฟล์เดียว app.py, deploy Railway ผ่าน Dockerfile, DB=Postgres/​SQLite) — เพิ่มโมดูล "สมุดคะแนน (.accdb)" แล้ว:
   - โมเดล: `AccdbSubject`, `AccdbRow`, `AccdbToken` (แยกจากระบบเดิม)
   - เมนู "📒 สมุดคะแนน (.accdb)" ในแถบข้าง
   - หน้า `/accdb` (รายวิชา + นำเข้าตั้งต้น + ลิงก์ token), `/accdb/<id>` (คีย์คะแนน 4 แท็บ: คะแนน/หน่วยย่อย/คุณลักษณะ/อ่านคิดเขียน + ปุ่มเติมทั้งห้อง)
   - API `/api/accdb/export?token=...`
3. **Path B (ทางที่เลือก): เขียน .accdb บนเซิร์ฟเวอร์เอง** — ฝังเอนจิน Java (Jackcess + jackcess-encrypt) ผ่าน Dockerfile multi-stage
   - `accdb_engine/` (pom.xml + AccdbTool.java) → build เป็น `accdbtool.jar`
   - endpoint `POST /accdb/<id>/to-accdb` (อัปโหลด .accdb → เขียนคะแนนจาก DB ลงไฟล์ → ดาวน์โหลด)
   - endpoint `/accdb/engine-status` (เช็ก jar_exists + java version)
   - **ยืนยันจากซอร์สแล้วว่า jackcess-encrypt เขียนไฟล์ agile-encrypted กลับได้** (encodePageImpl→blockEncrypt)

## ค้างอยู่ / ขั้นต่อไป
- ยังไม่ได้ทดสอบเอนจิน Java จริง (workspace โหลด Maven ไม่ได้) → ต้อง **deploy บน Railway** (build ผ่าน Dockerfile) แล้ว:
  1. เปิด `<เว็บ>/accdb/engine-status` ต้องได้ `"jar_exists": true`
  2. อัปโหลดสำเนา .accdb 1 วิชา → "เขียน + ดาวน์โหลด" → เปิดใน BookMark เช็ก
- ถ้า Path B ใช้ไม่ได้ → fallback = **Path A** (คลาวด์คีย์ + ตัวช่วย PowerShell ในเครื่อง Windows) ที่ทำงานได้แล้ว
  (ไฟล์อยู่ที่ Windows: C:\Users\Pasin\Downloads\program\ : server.ps1, index.html, embed.json, เปิดระบบคีย์คะแนน.bat)

## ความลับที่ใช้ (สำคัญ)
- การเข้ารหัสชื่อใน .accdb = **transposition (สลับตำแหน่งตัวอักษร)** คีย์ = ผลรวมไบต์ TIS-620 mod โหมด
  - FIRSTNAME โหมด 3 → key = sum%3 ; LASTNAME โหมด 5 → key = sum%5 ; รหัสไฟล์(password) โหมด 2
  - ตาราง perm ต่อความยาว/คีย์ อยู่ใน embed.json / grades_seed.json (คีย์ `perm`)
- รหัสผ่าน .accdb 9 ไฟล์ (uid=admins) อยู่ใน grades_seed.json (คีย์ `pwmap`, key = ตัวเลขล้วนของชื่อไฟล์):
  พ21201=713322926511100 พ22201=713322926521200 พ23201=713322926531300
  พ31201=713094524611203 พ32201=723094525611203 พ33201=733094526611203
  ว21102=723312926511100 ว22102=723312926521200 ว23102=723312926531300
- รายชื่อนักเรียนจริง 249 คน มาจากไฟล์ Excel ที่ครูส่ง (ถอดรหัสจับคู่แล้ว 183 คนในฐานข้อมูล)
- ตาราง .accdb ที่เขียน: `TRANSCRIPTS` (UM01,UM02,MidtermMark,FinalMark,UnitMark,TotalMark,TotalPercent,Grade,QM1-8,QualityMark,QGrade,LM1-5,LiteratureMark,LGrade),
  `TRANSCRIPTS2` (um<NN>_1..5 = คะแนนย่อย). ช่องเกรดถ้าว่างต้องเขียน NULL (มี FK กับ TabGrade/TabGradeQuality/TabGradeLiterature; ค่า Q/L เกรดใช้ 0-3)

## ไฟล์แพตช์ที่ส่งให้ครู
- Kruruksorn-PathB-webengine.zip (app.py, Dockerfile, grades_seed.json, templates/accdb_*, accdb_engine/*)
- ตัวช่วยในเครื่อง (Path A): อยู่ในโฟลเดอร์ program บน Windows แล้ว
