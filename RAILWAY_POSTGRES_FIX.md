# Railway PostgreSQL Fix

ระบบนี้ต้องต่อฐานข้อมูล PostgreSQL ผ่านตัวแปร `DATABASE_URL` ใน service `web`

## วิธีใส่ที่ถูกต้อง

ไปที่ `web -> Variables`

ช่องซ้าย:

```env
DATABASE_URL
```

ช่องขวาต้องเป็น URL จริง เช่น:

```env
postgresql://postgres:password@host:5432/railway
```

หรือใช้ปุ่ม `Add Reference` แล้วเลือก `PostgreSQL -> DATABASE_URL` เท่านั้น

ห้ามใส่ค่าเหล่านี้ในช่องขวา:

```env
Postgres
PostgreSQL
DATABASE_URL
${{Postgres}}
```

ถ้าใช้ `${{Postgres.DATABASE_URL}}` แล้ว error แปลว่า Railway ไม่ได้ resolve reference ให้ ให้ copy URL จริงจาก service PostgreSQL มาใส่แทน

## สิ่งที่แก้ในโค้ด

- รองรับ `DATABASE_URL`, `DATABASE_PRIVATE_URL`, `DATABASE_PUBLIC_URL`, `POSTGRES_URL`, `POSTGRESQL_URL`
- รองรับกรณี Railway ส่ง `postgres://` โดยแปลงเป็น `postgresql://`
- ถ้าใส่ reference ผิด ระบบจะแจ้ง error ชัดเจน
- ถ้าไม่มี PostgreSQL บน Railway ระบบจะหยุด ไม่ fallback ไป SQLite
