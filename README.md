ใช้python version 3.12.3
วิธีการใช้
ติดตั้ง Python และ dependencies ดังนี้:

OpenCV (pip install opencv-python)
Tesseract OCR (pip install pytesseract)
Pandas (pip install pandas)
XlsxWriter (pip install XlsxWriter)
Imutils (pip install imutils)
ติดตั้ง Tesseract OCR และตั้งค่าตำแหน่งไฟล์ tesseract.exe ให้ถูกต้องในโค้ด
โมเดลภาษาเพิ่มเติม ตามลิงค์นี้ https://github.com/UB-Mannheim/tesseract/wiki
(สามารถใช่ตัวล่าสุดได้แต่ในกรณีที่มีปัญหาให้ใช่้ 2024-05-19 Update Tesseract 5.4.0-rc2.)
เมื่อลงโมเดลภาษาที่ต้องการเสร็จแล้ว
จะมีการเรียกใช้ โดยใช้คำสั่ง
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

เตรียมโฟลเดอร์ที่มีภาพสแกนประกันภัยที่ต้องการประมวลผล

แก้ไขโค้ดในส่วนของ folder_path เพื่อระบุที่ตั้งของโฟลเดอร์ที่มีภาพ

รันโปรแกรม

ผลลัพธ์จะถูกเขียนลงในไฟล์ Excel ที่โฟลเดอร์ excel_BKI

คำแนะนำเพิ่มเติม
สามารถปรับแต่งพารามิเตอร์ต่างๆ เช่น ขนาดของตัวอักษรที่อ่านได้, การปรับสี, การตัดภาพ, เป็นต้น ในส่วนของ pytesseract.image_to_string เพื่อประสิทธิภาพการอ่านที่ดีขึ้น

สามารถปรับแต่งการตรวจสอบข้อมูลจากภาพ เช่น การค้นหาข้อมูลเฉพาะที่สนใจ, การแยกแยะข้อมูล, และการประมวลผลเพิ่มเติมตามความต้องการของแต่ละกรณี

อ่านเอกสาร API และคู่มือการใช้งานของ OpenCV, Tesseract OCR, และ Pandas เพื่อเรียนรู้เพิ่มเติมเกี่ยวกับฟังก์ชันและวิธีการใช้งานที่มีอยู่

ข้อจำกัด
การตรวจสอบและประมวลผลข้อมูลอาจไม่แม่นยำ 100% ขึ้นอยู่กับคุณภาพของภาพ, การตั้งค่าพารามิเตอร์, และการประมวลผล

ควรทดลองและปรับแต่งโปรแกรมให้เหมาะสมกับข้อมูลและเงื่อนไขการใช้งานของแต่ละกรณี


การเรียกใช้งานผ่าน terminal
1) เราต้องระบุโฟลเดอร์ก่อน โดยเมื่อเปิดterminal มาแล้ว ให้ใช้คำสั่ง cd "C:\Users\ASUS\Desktop\OCR" (ตย.)
 
2) จากนั้นให้ติดตั้ง pip ทีต้องใช้  คือ 
pip install pdf2image
pip install pytesseract
pip install xlsxwriter
pip install matplotlib
pip install opencv-python
pip install pandas 


3)เรียกใช้งานไฟล์โดยใช้คำสั่ง python BKI.py BKI-001
