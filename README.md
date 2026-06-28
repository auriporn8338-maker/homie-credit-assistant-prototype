# Homie Credit Assistant Prototype

เว็บ Prototype สำหรับให้พนักงานทดลองใช้งานก่อนนำไปผูกกับ LINE OA

## เมนูหลัก

1. คำนวณสินเชื่อเบื้องต้น
2. Checklist เอกสารตามอาชีพและวัตถุประสงค์กู้
3. Self Learning
4. Q&A งานสินเชื่อ

## วิธีเปิดใช้งาน

เปิดไฟล์ `index.html` หรือเปิดผ่าน GitHub Pages หลังตั้งค่า Pages ของ repository นี้

## GitHub Pages

ไปที่ Repository Settings > Pages แล้วตั้งค่า:

- Source: Deploy from a branch
- Branch: main
- Folder: /root

หลังจากบันทึก GitHub จะสร้าง URL สำหรับทดลองใช้งาน

## ข้อจำกัด Prototype

- ใช้สำหรับทดลองกับพนักงานเท่านั้น
- Calculator ยังไม่ใช่สูตรผลิตภัณฑ์จริง
- Checklist เป็นข้อมูลเบื้องต้นตามฐานความรู้ V2
- ไม่ควรกรอกข้อมูลส่วนบุคคลของลูกค้าจริงบนเว็บไซต์สาธารณะ

## Files

- `index.html` หน้าเว็บหลัก
- `styles.css` รูปแบบหน้าจอ
- `data.js` ฐานข้อมูล Prototype
- `app.js` Logic การทำงาน
