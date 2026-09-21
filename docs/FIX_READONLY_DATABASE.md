# แก้ปัญหา SQLite readonly database

ถ้าเจอ error:

```txt
sqlite3.OperationalError: attempt to write a readonly database
```

แปลว่าไฟล์ฐานข้อมูล local `instance/krurakson.db` หรือโฟลเดอร์ `instance` เขียนไม่ได้

## วิธีแก้บน Mac / Linux

เปิด Terminal ที่โฟลเดอร์โปรเจกต์ แล้วรัน:

```bash
mkdir -p instance
chmod u+rwx instance
chmod u+rw instance/krurakson.db 2>/dev/null || true
```

ถ้ายังไม่ได้ ให้รัน:

```bash
sudo chown -R "$USER" instance
chmod -R u+rwX instance
```

## วิธีล้างฐานข้อมูล local แล้วเริ่มใหม่

ใช้เฉพาะตอนข้อมูลทดสอบไม่สำคัญ:

```bash
rm -f instance/krurakson.db
python app.py
```

## บน Railway

ควรใช้ PostgreSQL เท่านั้น โดยตั้ง `DATABASE_URL` ให้ web service ไม่ควรใช้ SQLite ใน production
