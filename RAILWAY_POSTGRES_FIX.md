# Railway PostgreSQL Fix

ไฟล์นี้ปรับให้ระบบใช้ฐานข้อมูล PostgreSQL บน Railway ผ่าน `DATABASE_URL` และกันไม่ให้ Production เผลอเปิด SQLite ใหม่
จนดูเหมือนข้อมูลเช็กชื่อหาย

## ต้องตั้งค่าใน Railway

ที่ service `web` > Variables ต้องมี:

```env
DATABASE_URL=${{Postgres.DATABASE_URL}}
```

หรือ copy ค่า `DATABASE_URL` จาก service PostgreSQL มาใส่ใน web โดยตรง

## สิ่งที่แก้ในโค้ด

- อ่าน `DATABASE_URL` / `POSTGRES_URL` / `POSTGRESQL_URL`
- แปลง `postgres://` เป็น `postgresql://`
- ถ้ารันบน Railway แล้วไม่มี `DATABASE_URL` จะหยุดทันที ไม่ fallback ไป SQLite
- เพิ่ม `psycopg2-binary` ใน requirements.txt
- ปรับ `ensure_schema_columns()` ให้รองรับ PostgreSQL ไม่ใช้ `PRAGMA` เฉพาะ SQLite
- เพิ่ม `.gitignore` กันไฟล์ `.db` ไม่ให้ติด GitHub อีก

## ห้ามทำ

- ห้าม `drop_all()` กับฐานข้อมูลใช้งานจริง
- ห้าม commit ไฟล์ `.db` ขึ้น GitHub
- ก่อน import ข้อมูลจำนวนมากควร export/backup PostgreSQL ก่อน
