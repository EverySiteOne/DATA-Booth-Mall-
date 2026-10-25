# DATA Booth Mall

Dashboard Booth Mall Insurance 2569 บน GitHub Pages — ไม่มีแถบเตือนของ Google Apps Script

- หน้าเว็บ (`index.html`) ดึงข้อมูลจาก Google Apps Script Web App (`API_URL` ในไฟล์) ซึ่งอ่านจาก Google Sheets และอัปเดตอัตโนมัติ
- Apps Script ต้องใช้ `Code.gs` เวอร์ชันที่รองรับ `?api=getDashboardData` และ Deploy แบบ Who has access: Anyone
- เปิดตรงหน้าการออกบูธ: `https://chaloemphos-byte.github.io/DATA-Booth-Mall-/?tab=ops`
- ถ้าดึงข้อมูลไม่ได้ หน้าเว็บจะแสดงลิงก์เปิด Dashboard (Apps Script) ในแท็บใหม่
