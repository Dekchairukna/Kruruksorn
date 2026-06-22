# Railway Build Fix

ไฟล์ชุดนี้เพิ่มการตั้งค่า deploy ให้ชัดเจนขึ้น เพื่อแก้อาการ Railway ขึ้น `Failed to build an image` ในขั้น Build image

## เพิ่ม/แก้ไฟล์

- `Dockerfile` บังคับใช้ Python 3.11 slim และติดตั้งจาก `requirements.txt`
- `railway.json` บังคับ Railway ใช้ Dockerfile builder
- `.python-version` ระบุ Python 3.11 สำหรับกรณี Railway/Nixpacks/Railpack ตรวจ Python version
- `.dockerignore` ตัดไฟล์ cache/database/upload ออกจาก build context

## ตัวแปรที่ต้องมีบน Railway

อย่างน้อยควรมี:

```env
SECRET_KEY=ตั้งค่าเป็นข้อความยาว ๆ
DATABASE_URL=postgresql://...
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1-mini
```

ถ้ายังไม่ตั้ง `OPENAI_API_KEY` หน้า `/ai` จะเปิดได้ แต่กดสร้าง AI ไม่ได้

## หมายเหตุ

AI Assistant ไม่แก้ข้อมูลเอง ไม่บันทึกฐานข้อมูลเอง สร้างข้อความให้ครูตรวจและคัดลอกไปใช้เท่านั้น
