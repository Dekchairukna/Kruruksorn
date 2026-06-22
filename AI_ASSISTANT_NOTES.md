# Phase AI Assistant - KRURUKSORN

เพิ่มหน้า `/ai` สำหรับครู/แอดมิน

## สิ่งที่เพิ่ม

- เมนู `✨ AI Assistant` ใน sidebar
- Route `/ai`
- ปุ่มหลัก 3 งาน
  - สรุปว่าสอนถึงไหน
  - สร้างใบความรู้
  - สร้างใบงาน
- อ่าน `OPENAI_API_KEY` จาก Environment Variable
- อ่าน `OPENAI_MODEL` จาก Environment Variable ถ้าไม่ตั้ง ใช้ `gpt-4.1-mini`
- AI สร้างเฉพาะข้อความให้ครูตรวจแก้และคัดลอกเอง ยังไม่แก้ข้อมูล/ไม่บันทึกฐานข้อมูลเอง

## ตั้งค่าบน Railway

ไปที่ Service เว็บ > Variables แล้วเพิ่ม

```env
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1-mini
```

จากนั้น Redeploy หรือ Restart Service

## ไฟล์ที่แก้

- `app.py`
- `templates/base.html`
- `templates/ai_assistant.html`
- `static/css/style.css`
