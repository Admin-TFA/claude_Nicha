/*
 * ตั้งค่าการเชื่อมต่อ Google Sheet
 * --------------------------------
 * วาง URL ของ Web app ที่ได้จากการ Deploy Google Apps Script (ลงท้ายด้วย /exec)
 * ตัวอย่าง: "https://script.google.com/macros/s/AKfycb..../exec"
 *
 * ถ้าปล่อยว่างไว้ เว็บจะทำงานแบบ "บันทึกในเครื่องนี้" (ไม่ส่งข้อมูลไป Google Sheet)
 * หมายเหตุ: อย่าใส่รหัสผู้ดูแล (ADMIN_KEY) ไว้ในไฟล์นี้ เพราะไฟล์นี้ทุกคนเปิดดูได้
 */
window.ATTENDANCE_CONFIG = {
  sheetsUrl: "",

  // ดึงข้อมูลล่าสุดจาก Google Sheet ทุกกี่วินาที (ให้ทุกเครื่องเห็นตรงกัน)
  pollSeconds: 8
};
