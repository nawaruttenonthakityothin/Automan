# 🔐 User Access Automation Web App

ระบบอ่านข้อมูลภาพแคปเจอร์ด้วย AI (Gemini 2.0 Flash Vision AI) ➔ บันทึก Excel ➔ ส่ง Email ตอบกลับอัตโนมัติ พัฒนาด้วยภาษา Python และ Streamlit Web Framework

---

## 📁 โครงสร้างไฟล์ในโครงการ:

```text
etravelling_auto_exe/
├── web_app.py                # 🌐 โปรแกรมระบบ Web Application (Streamlit)
├── requirements.txt          # 📦 ไลบรารีสำหรับ Web Deployment
├── app.py                    # 💻 โปรแกรมระบบ Desktop App (Tkinter)
├── backup_desktop_app/       # 📁 โฟลเดอร์สำรองข้อมูล Desktop Standalone App
│   └── app.py
└── Template Column Excel...  # 📊 ไฟล์ตารางสรุป Excel
```

---

## 💻 1. วิธีรันใช้งานบนเครื่องคอมพิวเตอร์ของคุณเอง (Local Web App):

1. เปิด Terminal / Command Prompt ในโฟลเดอร์โครงการ:
   ```bash
   cd C:\Users\Nawarutte.Non\.gemini\antigravity\scratch\etravelling_auto_exe
   ```
2. ติดตั้ง Streamlit (หากยังไม่มี):
   ```bash
   pip install streamlit
   ```
3. รันคำสั่งเปิดเว็บ:
   ```bash
   streamlit run web_app.py
   ```
4. ระบบจะเปิดหน้า Web Browser ขึ้นมาอัตโนมัติที่ `http://localhost:8501`

---

## 🌐 2. วิธีขึ้นระบบเป็นเว็บไซต์ใช้งานร่วมกัน (Deploy to Streamlit Cloud - ฟรี):

1. **เอาโค้ดขึ้น GitHub**:
   - สร้าง repository ใหม่บน GitHub (เช่น `user-access-automation`)
   - พุชไฟล์ทั้งหมด (`web_app.py`, `requirements.txt`, `README.md`) ขึ้น GitHub
2. **เชื่อมต่อกับ Streamlit Cloud**:
   - เข้าไปที่ [share.streamlit.io](https://share.streamlit.io/)
   - เลือก **New app** ➔ เลือก Repository บน GitHub ของคุณ ➔ ตั้งค่า Main file path เป็น `web_app.py` ➔ กด **Deploy**
3. **การส่งผลอัปเดตแบบ Real-Time**:
   - ทุกครั้งที่คุณแก้ไขไฟล์ `web_app.py` แล้วพุชลง GitHub หน้าเว็บไซต์จะทำการอัปเดตการเปลี่ยนแปลงให้เพื่อนร่วมงานทุกคนใช้งานทันทีอัตโนมัติ 100%!

---

## 🔙 3. หากต้องการกลับมาใช้งานเวอร์ชัน Desktop Standalone (Tkinter):

ไฟล์สำรองเวอร์ชัน Desktop ถูกจัดเก็บไว้อย่างปลอดภัยในโฟลเดอร์:
[`backup_desktop_app/app.py`](file:///C:/Users/Nawarutte.Non/.gemini/antigravity/scratch/etravelling_auto_exe/backup_desktop_app/app.py)

คุณสามารถรันใช้งานโปรแกรม Desktop เดิมได้ตลอดเวลาด้วยคำสั่ง:
```bash
python app.py
```
