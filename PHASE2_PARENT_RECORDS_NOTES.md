# KRURUKSORN Phase 2: Parent & Records

เพิ่มต่อจาก Phase 1 โดยไม่แตะ KRURUK SPORTS

## เพิ่มแล้ว

- หน้า `/phase2` ศูนย์ Phase 2
- ปพ.6 รายงานผลการเรียนรายบุคคล `/pp6/<student_id>`
- Export ปพ.6 Excel `/pp6/<student_id>/export`
- ปพ.7 หนังสือรับรองการเป็นนักเรียน `/pp7/<student_id>`
- Parent Portal แบบลิงก์เฉพาะนักเรียน `/parent/<token>`
- หน้าแก้ข้อมูลผู้ปกครองและ LINE userId `/phase2/guardian-settings`
- หน้า LINE แจ้งขาดเรียน `/phase2/line-absence`
- เพิ่ม field `guardian_line_user_id` ในตาราง user พร้อม schema migration อัตโนมัติ

## LINE

ใช้ LINE Messaging API แบบ Push message ผ่าน LINE Official Account

ตั้งค่าใน Railway Variables:

```env
LINE_CHANNEL_ACCESS_TOKEN=ใส่ Channel access token ของ LINE OA
```

หมายเหตุ: `guardian_line_user_id` ต้องเป็น LINE userId ที่ได้จาก LINE OA/LIFF/Webhook ไม่ใช่ @LINE ID และไม่ใช่เบอร์โทร

## Parent Portal

ระบบสร้างลิงก์แบบ signed token จาก `SECRET_KEY` ให้ผู้ปกครองดูคะแนนและเวลาเรียนได้อย่างเดียว ไม่สามารถแก้ข้อมูลได้

ถ้าเปลี่ยน `SECRET_KEY` ลิงก์ผู้ปกครองเก่าจะใช้ไม่ได้ ต้องคัดลอกลิงก์ใหม่
