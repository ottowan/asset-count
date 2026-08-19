# Asset Count

ระบบตรวจนับครุภัณฑ์จาก Serial Number โดยอ่านข้อมูลหลักจาก Firebase Firestore และส่งออกผลเป็น Excel ได้

## เริ่มใช้งาน

```bash
npm install
npm run dev
```

## Deploy บน Netlify

เชื่อม GitHub repository กับ Netlify แล้วระบบจะใช้ค่าจาก `netlify.toml` โดยอัตโนมัติ:

- Build command: `npm run build`
- Publish directory: `dist`

## เปิดใช้ยอดรวมหลายคนด้วย Firebase Real-time

1. สร้าง Project ที่ Firebase Console แล้วเพิ่ม Web app
2. ไปที่ Firestore Database > Create database แล้วเลือก Production mode
3. นำเนื้อหา `firestore.rules` ไปวางที่ Firestore > Rules แล้วกด Publish
4. คัดลอก `.env.example` เป็น `.env` และกรอกค่าจาก Firebase web app config
5. บน Netlify เพิ่ม Environment variables ทั้ง 6 ตัวตาม `.env.example`
6. Deploy ใหม่ ยอดรวมของทุกเครื่องจะอัปเดตทันทีโดยไม่ต้อง Refresh

## นำเข้าข้อมูล serial.xlsx เข้า Firestore

1. Firebase Console > Project settings > Service accounts > Generate new private key
2. บันทึกไฟล์ JSON ไว้นอก Git repository และห้ามนำขึ้น GitHub
3. ติดตั้ง dependency และรันคำสั่งนำเข้า:

```powershell
$env:FIREBASE_SERVICE_ACCOUNT="C:\path\to\service-account.json"
npm run import:assets
```

สคริปต์จะตรวจ ID/SN ซ้ำ อัปโหลด collection `assets` และสร้าง `system/stats` สำหรับยอดรวม การรันซ้ำจะอัปเดตข้อมูลเดิมโดยไม่สร้างรายการซ้ำ

ไฟล์ `serial.xlsx` และ Service Account ใช้เฉพาะเครื่องสำหรับคำสั่ง Import และถูก `.gitignore` เพื่อไม่เผยแพร่ข้อมูลครุภัณฑ์บน GitHub เมื่อ Deploy ต้องตั้งค่า Firebase environment variables ให้ครบ
