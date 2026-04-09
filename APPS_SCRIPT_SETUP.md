# 🔗 Google Sheets Sync — ขั้นตอนการตั้งค่า

ทำให้ข้อมูลซิงค์กันทุกเครื่อง (มือถือ/คอม/แท็บเล็ต)

## ขั้นตอน

### 1. สร้าง Google Sheet
1. ไปที่ https://sheets.google.com
2. สร้างไฟล์ใหม่ ตั้งชื่อ **SHIFT KKC Data**
3. ไม่ต้องสร้าง sheet tabs เอง — โค้ดสร้างให้อัตโนมัติ

### 2. วาง Apps Script
1. ใน Google Sheet → **Extensions → Apps Script**
2. ลบโค้ดเดิมใน `Code.gs` ทั้งหมด
3. คัดลอกโค้ดจากไฟล์ `apps_script.gs` (ในโฟลเดอร์นี้) ไปวาง
4. กด **Save** (💾)

### 3. Deploy เป็น Web App
1. มุมขวาบน คลิก **Deploy → New deployment**
2. คลิกเฟือง ⚙ ข้าง "Select type" → เลือก **Web app**
3. ตั้งค่า:
   - **Description:** SHIFT KKC Sync
   - **Execute as:** Me (your@email.com)
   - **Who has access:** Anyone
4. คลิก **Deploy**
5. ครั้งแรกจะขอ permission → Allow ให้หมด
6. **คัดลอก URL** ที่ได้ (จะขึ้นต้นด้วย `https://script.google.com/macros/s/.../exec`)

### 4. ใส่ URL ในแอพ
1. เข้าแอพ SHIFT KKC → **Manager** (PIN 2043)
2. Sidebar → **ตั้งค่าระบบ**
3. วาง URL ที่ช่อง **Apps Script URL**
4. กดบันทึก

### 5. ซิงค์ครั้งแรก
- กดปุ่ม **↻ ซิงค์** มุมขวาบน
- ทำในทุกเครื่องที่ต้องการให้ข้อมูลตรงกัน

---

## วิธีใช้งาน
- ทุกครั้งที่บันทึกสต็อก/รายรับ → auto push ไป Sheet
- เมื่อสลับเครื่อง → กด **↻ ซิงค์** 1 ครั้ง จะดึงข้อมูลล่าสุดมา
- Google Sheet คือ **source of truth** — ดูได้จาก browser
