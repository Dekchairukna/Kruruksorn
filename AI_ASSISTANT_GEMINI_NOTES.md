# Phase AI Assistant - Gemini Support

เพิ่มให้หน้า `/ai` รองรับ Gemini API เพื่อลดค่าใช้จ่ายเริ่มต้น

## Railway Variables ที่แนะนำ

```env
AI_PROVIDER=gemini
GEMINI_API_KEY=AIza-your-real-key-here
GEMINI_MODEL=gemini-2.5-flash
```

หมายเหตุ: `AIza...` ในตัวอย่างเป็นแค่รูปแบบของ key ต้องใช้ key จริงจาก Google AI Studio เท่านั้น

## Provider ที่รองรับ

- `gemini` ใช้ Google Gemini API
- `openai` ใช้ OpenAI Responses API เดิม
- `offline` ใช้แม่แบบออฟไลน์ ไม่ใช่ AI จริง แต่หน้าไม่พัง

ถ้าไม่ตั้ง `AI_PROVIDER` ระบบจะเลือก Gemini ก่อนถ้ามี `GEMINI_API_KEY` จากนั้นค่อย OpenAI ถ้ามี `OPENAI_API_KEY`

## ความปลอดภัย

ระบบอ่าน key จาก Environment Variables เท่านั้น ไม่เก็บ key ลงฐานข้อมูล และไม่ควร commit key ลง GitHub

## หน้าใช้งาน

- `/ai`
- ปุ่มสรุปว่าสอนถึงไหน
- ปุ่มสร้างใบความรู้
- ปุ่มสร้างใบงาน

AI ยังไม่แก้ข้อมูลเองและไม่บันทึกฐานข้อมูล ครูต้องตรวจและคัดลอก/บันทึกเอง
