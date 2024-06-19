import fitz
import os

# โฟลเดอร์ที่มีไฟล์ PDF หลายๆ ไฟล์
pdf_folder = r'C:\Users\ASUS\Desktop\OCR\PDF'
# ตั้งค่าโฟลเดอร์ที่จะบันทึกไฟล์รูปภาพ (ใช้ชื่อเดิมของไฟล์ PDF แต่ต่อท้ายด้วย .jpg)
output_folder = r'C:\Users\ASUS\Desktop\OCR\test_BKI_JPG'
# ตรวจสอบและสร้างโฟลเดอร์ output ถ้ายังไม่มี
os.makedirs(output_folder, exist_ok=True)
# วนลูปเพื่อเปิดและแปลงทุกไฟล์ PDF
for filename in os.listdir(pdf_folder):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder, filename)
        doc = fitz.open(pdf_path)
        # เลือกหน้าแรกของเอกสาร (หมายเลขหน้า 0)
        page = doc.load_page(0)
        # แปลงหน้าเป็น pixmap ที่มี dpi=300
        pixmap = page.get_pixmap(dpi=300)
        # สร้างชื่อไฟล์รูปภาพโดยใช้ชื่อเดิมของไฟล์ PDF
        output_filename = os.path.splitext(filename)[0] + '.jpg'
        output_path = os.path.join(output_folder, output_filename)
        # บันทึกรูปภาพเป็นไฟล์
        pixmap.save(output_path)
        # ปิดเอกสาร
        doc.close()
print("การแปลง PDF เป็นรูปภาพเสร็จสมบูรณ์")
