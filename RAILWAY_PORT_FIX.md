# Railway PORT Fix

สาเหตุจาก log ล่าสุด:

```txt
Error: '$PORT' is not a valid port number.
```

แปลว่า Gunicorn ได้รับค่า `$PORT` เป็นข้อความตรง ๆ ไม่ได้ถูก expand เป็นเลข port จริงจาก Railway
จึงแก้โดยเพิ่ม `start.sh` ให้ shell อ่านค่า `PORT` ก่อน แล้วส่งเลขจริงให้ Gunicorn

ไฟล์ที่แก้:

- `start.sh`
- `Dockerfile`
- `railway.json`
- `Procfile`

คำสั่งเริ่มระบบที่ถูกต้องคือ:

```bash
sh /app/start.sh
```

ถ้าใน Railway เคยตั้ง Start Command เองเป็น `gunicorn app:app --bind 0.0.0.0:$PORT`
ให้ลบออก หรือเปลี่ยนเป็น:

```bash
sh /app/start.sh
```

ห้ามใส่ `$PORT` เป็น argument ให้ Gunicorn โดยตรงใน Railway UI เพราะบางกรณี Railway จะส่งค่าไปแบบ literal ไม่ expand ให้
