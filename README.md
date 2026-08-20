# 🐷 Pig Farm Accounting System (Account Assistant)

> ระบบผู้ช่วยบัญชีและฟาร์มสุกร ประมวลผลและสร้างชีทสรุปยอดขายสุกร (`สรุปยอด`) และสรุปยอดขายแยกเล้า (`สรุปยอดแยกเล้า`) อัตโนมัติจากไฟล์ Excel

---

## 🌟 คุณสมบัติเด่น (Features)

* 🖥️ **2-Way Dual Interface (GUI + TUI):**
  * **Desktop GUI Application:** พัฒนาด้วย `PyQt6` ดีไซน์ Dark Theme ทันสมัย ใช้งานง่าย มีดรอปดาวน์เลือกไฟล์ ปุ่ม Browse และรายงานสถานะแบบเรียลไทม์
  * **Terminal UI (TUI):** พัฒนาด้วย `Rich` หน้าจอคอนโซลโทนสีสวยงาม รันรวดเร็วผ่านคำสั่ง CLI
* 🖼️ **Custom Signature App Icon:** ไอคอนประจำโปรแกรมลุคน่ารักและเป็นเอกลักษณ์
* 📑 **Sheet Overwrite Mode (เขียนทับชีทเดิม):** เมื่อรันประมวลผลซ้ำ ระบบจะลบและเขียนทับชีท `สรุปยอด` หรือ `สรุปยอดแยกเล้า` ในชื่อเดิมทันที โดยไม่สร้างชื่อชีทซ้ำ
* 📊 **Single Week Summary Row (รวมทุกเล้าประจำสัปดาห์):** ในชีทสรุปยอดแยกเล้า มีบรรทัดสรุปยอดรวมทุกเล้าปิดท้ายในแต่ละสัปดาห์ (ไฮไลท์สีทองผึ้ง `#FFF2CC` และเส้นขอบคู่)
* 🔗 **Merged Cell Handling in Column N:** ตรวจจับและประมวลผลเซลล์ผสาน (Merged Cells) ในคอลัมน์เล้า (คอลัมน์ N) ได้อย่างถูกต้อง ไม่ตกหล่น และไม่นับซ้ำ
* 🧩 **Flexible Barn String Regex Parser:** ถอดรหัสชื่อเล้าและจำนวนหมูได้อย่างแม่นยำ แม้ว่าในไฟล์ Excel จะพิมพ์ติดกัน หรือลืมใส่เครื่องหมายคั่น `,` (เช่น `"T8=1 , N3=1 N1=3"`)
* 🟧 **Orange Fill Duplicate Row Exclusion Rule:** ตรวจสอบและตัดบรรทัดที่มีการเติมสีส้มที่ซ้ำวันและจำนวนกับบรรทัดก่อนหน้า/ถัดไปออกจากการรวมยอดโดยอัตโนมัติ

---

## 🚀 การติดตั้งและใช้งาน (Installation & Usage)

### 1. การติดตั้ง Dependencies
```bash
pip install -r requirements.txt
```

### 2. การเปิดใช้งาน Desktop GUI (โหมดกราฟิก)
```bash
python main.py
```
*(หรือ `python main.py --gui`)*

### 3. การเปิดใช้งาน Terminal UI (โหมดคอนโซล)
```bash
python main.py --tui
```

---

## 🛠️โครงสร้างโปรเจกต์ (Project Structure)

```text
account assistant/
├── main.py              # ไฟล์หลักสำหรับรันโปรแกรม (CLI Argument Switcher GUI/TUI)
├── gui.py               # Desktop GUI Application (PyQt6)
├── excel_processor.py   # เอนจินประมวลผล Excel (openpyxl)
├── utils.py             # ฟังก์ชันจัดการข้อความและหมวดหมู่สุกร
├── signature_logo.jpg   # ไฟล์ไอคอนและลายเซ็นประจำแอปพลิเคชัน
├── requirements.txt     # รายการ Python Libraries ที่ใช้งาน
└── README.md            # คู่มือการใช้งานระบบ
```

---

## 🔒 การรักษาความปลอดภัยข้อมูล (Data Privacy)

ไฟล์ข้อมูล Excel (`*.xlsx`, `*.xls`) ทั้งหมดถูกตัดออกจากการอัปโหลดด้วย `.gitignore` เพื่อป้องกันไม่ให้ข้อมูลบัญชีฟาร์มหลุดขึ้น Git Repository 100%
