import fitz
import os

def read_file_based_on_choice(text_choice1, text_choice2):
    pdf_folder = f'C:\\Users\\ASUS\\Desktop\\OCR\\{text_choice1}'
    output_folder = f'C:\\Users\\ASUS\\Desktop\\OCR\\{text_choice2}'
    return pdf_folder, output_folder

def main(args):
    if len(args) >= 2:
        text_choice1 = args[0]
        text_choice2 = args[1]
        print(f"โฟลเดอร์ที่ต้องการแปลงภาพ : {text_choice1}")
        print(f"โฟลเดอร์ที่ต้องการบันทึกภาพ : {text_choice2}")
        
        pdf_folder, output_folder = read_file_based_on_choice(text_choice1, text_choice2)
        
        # Create the output directory if it does not exist
        os.makedirs(output_folder, exist_ok=True)
        
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
        
        print("PDF 2 image conversion เสร็จสิ้น")
        return True
    else:
        print("กรุณาระบุที่อยุ่และที่ต้องการบันทึก")
        return False

# Example usage
if __name__ == "__main__":
    import sys
    main(sys.argv[1:])
