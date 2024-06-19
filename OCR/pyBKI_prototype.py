
import os
import cv2
import numpy as np
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import re
import sys
import pandas as pd
import xlsxwriter
import shutil
import imutils


BKI_001_x1_TYPE, BKI_001_y1_TYPE, BKI_001_x2_TYPE, BKI_001_y2_TYPE = 1100, 385, 1490, 470

BKI_001_x1_DATE, BKI_001_y1_DATE, BKI_001_x2_DATE, BKI_001_y2_DATE = 1900, 240, 2205, 310

BKI_001_x1_NAME, BKI_001_y1_NAME, BKI_001_x2_NAME, BKI_001_y2_NAME = 150, 580, 605, 720

BKI_001_x1_Policy, BKI_001_y1_Policy, BKI_001_x2_Policy, BKI_001_y2_Policy = 320, 995, 790, 1090

BKI_001_x1_fee_duty, BKI_001_y1_fee_duty, BKI_001_x2_fee_duty, BKI_001_y2_fee_duty = 2000, 675, 2305, 850

BKI_001_x1_Balanc, BKI_001_y1_Balanc, BKI_001_x2_Balanc, BKI_001_y2_Balanc= 2000, 920, 2305, 1050

BKI_001_x1_VAT, BKI_001_y1_VAT, BKI_001_x2_VAT, BKI_001_y2_VAT = 2000, 1020, 2305, 1150

BKI_001_x1_total, BKI_001_y1_total, BKI_001_x2_total, BKI_001_y2_total = 2000, 1130, 2305, 1220

def process_image(image, min_contour_area=1000, min_line_length=400):
        # Convert image to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Apply GaussianBlur to reduce noise
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)

        # Apply Canny edge detection
        edges = cv2.Canny(blurred, 50, 150)

        # Find contours
        contours, _ = cv2.findContours(edges.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # Filter out small contours and long lines
        min_contour_area = 1000
        min_line_length = 400
        max_angle = -np.inf
        max_angle_lines = []
        for contour in contours:
            epsilon = 0.02 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)
            if len(approx) >= 2:  # Check if the contour is approximated by at least 2 points (a line)
                # Iterate over each pair of points to calculate line length
                for i in range(len(approx) - 1):
                    x1, y1 = approx[i][0]
                    x2, y2 = approx[i + 1][0]
                    line_length = np.sqrt((x2 - x1)**2 + (y2 - y1)**2)
                    angle = np.arctan2(y2 - y1, x2 - x1) * 180 / np.pi
                    if abs(angle) < 10 or abs(angle - 180) < 10:
                        if line_length > min_line_length:
                            if abs(angle) > max_angle:
                                max_angle = abs(angle)
                                max_angle_lines = [(x1, y1), (x2, y2)]

        # Draw the line with maximum angle
        if max_angle_lines:
            x1, y1 = max_angle_lines[0]
            x2, y2 = max_angle_lines[1]
            # Generate random color
            color = (np.random.randint(0, 256), np.random.randint(0, 256), np.random.randint(0, 256))
            cv2.line(image, (x1, y1), (x2, y2), color, 2)

        # Filter contours with area greater than the specified threshold
        min_contour_area = 1
        filtered_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > min_contour_area]

        # Check if no contours are found
        if not filtered_contours:
            print(f"ไม่พบเส้นขอบในไฟล์ {filename}")

        # Update the image rotation code
        for contour in filtered_contours:
            # Determine the type of contour
            epsilon = 0.02 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)
            if len(approx) == 2:  # Straight line
                # Calculate the angle of the line
                angle = np.degrees(np.arctan((y2 - y1) / (x2 - x1)))

                # Get the image dimensions
                height, width = image.shape[:2]

                # Set the center point as the midpoint of the image
                center = (width // 2, height // 2)

                # Rotate the image
                height, width = image.shape[:2]
                rotation_matrix = cv2.getRotationMatrix2D(center, angle, 1)
                rotated_image = cv2.warpAffine(image, rotation_matrix, (width, height))
                return rotated_image

def process_TYPE(rotated_image, filename, x1, y1, x2, y2):
    # กำหนดค่าความกว้างและความสูงใหม่
    new_width = 100
    new_height = 100

    # ปรับขนาดรูปภาพ
    resized_image = cv2.resize(rotated_image, (new_width, new_height))

    # ตัดรูป
    cropped_image = rotated_image[y1:y2, x1:x2]

    # เพิ่มความคมชัดให้รูป
    enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)

    # ส่งรูปที่ปรับขนาดและตัดมายังฟังก์ชัน OCR
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(cropped_image, lang='eng', config=custom_config)
    print("from " + filename)
    print("The text  from  scan is")

    # ค้นหาตำแหน่งของเครื่องหมายวงเล็บเปิดและปิด
    opening_brackets = ["(", "{", "["]
    closing_brackets = [")", "}", "]"]
    opening_indexes = [text.find(bracket) for bracket in opening_brackets if text.find(bracket) != -1]
    closing_indexes = [text.rfind(bracket) for bracket in closing_brackets if text.rfind(bracket) != -1]
    if opening_indexes:
        opening_index = min(opening_indexes)
    if closing_indexes:
        closing_index = max(closing_indexes)

    # ถ้าพบเครื่องหมายเปิดและปิด แยกข้อความระหว่างวงเล็บ
    if 'opening_index' in locals() and 'closing_index' in locals() and opening_index != -1 and closing_index != -1:
        extracted_text = text[opening_index + 1 : closing_index]
        TYPE = extracted_text
        print("type:", TYPE)
        return TYPE  # รีเทิร์นค่า TYPE ออกมา
    else:
        print("ไม่พบเครื่องหมายเปิดและปิดคำสั่งในข้อความ")

def process_DATE(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                        # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(image, d=9, sigmaColor=75, sigmaSpace=75)
                        # ขยายภาพ
                        scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
                        #cv2_imshow( enhanced_image)#
                        # ใช้ EasyOCR เพื่ออ่านข้อความ
                        # reader = easyocr.Reader(['en', 'th'], gpu= True)
                        reader = pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        result = reader.image_to_string(enhanced_image, config=custom_config)
                        # ใช้ Regular Expression เพื่อคัดแยกเฉพาะตัวเลข 0-9 และเครื่องหมาย /
                        filtered_result = re.findall(r'[\d/]+', result)
                        pattern = r'(0[1-9]|1[0-9]|2[0-9]|3[01])/(0[1-9]|1[012])/(\d{4})'
                        # พิมพ์ผลลัพธ์ที่คัดแยกแล้ว
                        #print("ผลลัพธ์การสแกน:")
                        matches = re.findall(pattern, result)
                        # ตรวจสอบว่ามีผลลัพธ์หรือไม่
                        if not matches:
                            print("Not found")
                        else:
                            for match in matches:
                                DAY = "{}/{}/{}".format(match[0], match[1], match[2])
                                print("DATE:",DAY)
                            return DAY

def process_NAME(filename, rotated_image, x1, y1, x2, y2):
    # ตัดภาพตามพิกัดที่กำหนด
    cropped_image = rotated_image[y1:y2, x1:x2]
    # ขยายขนาดภาพ
    scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
    # ปรับสีให้เห็นชัดขึ้น
    enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
    # ใช้ pytesseract เพื่ออ่านข้อความ (ชื่อ)
    custom_config = r'--oem 3 --psm 6'
    detected_text = pytesseract.image_to_string(enhanced_image, lang='tha', config=custom_config)
    # ตรวจสอบคำที่ได้จาก pytesseract
    valid_prefixes = ["นาย", "นางสาว", "นาง", "คุณ", "Mr", "Mrs", "Ms", "Miss"]
    found_prefixes = []
    # ตรวจสอบและเก็บคำที่พบ
    for prefix in valid_prefixes:
        if prefix in detected_text:
            found_prefixes.append(prefix)
            break
    # หาตำแหน่งของคำนำหน้า
    remaining_text = detected_text
    for prefix in found_prefixes:
        index = remaining_text.find(prefix)
        if index != -1:
            remaining_text = remaining_text[index + len(prefix):].strip()
    # สร้างชื่อที่เหมาะสม
    if found_prefixes:
        NAME = "{} {}".format(found_prefixes[0], remaining_text)
        print("Name :", NAME)
    else:
        NAME = None
    # ส่งคืนชื่อที่ได้
    return NAME
        

def process_Policy(filename, rotated_image, x1, y1, x2, y2):
    # Crop the image based on specified coordinates
    cropped_image = rotated_image[y1:y2, x1:x2]
    # Set custom configuration for pytesseract
    custom_config = r'--oem 3 --psm 6'
    # Perform OCR to extract text from the cropped image
    text = pytesseract.image_to_string(cropped_image, lang='tha+eng', config=custom_config)
    # Extract all sequences of digits and hyphens from the OCR result
    filtered_result = re.findall(r'[0-9-]+', text)
    # Concatenate all digits and hyphens into a single string
    combined_string = ''.join(filtered_result)
    # Remove trailing "10" or "1" from each number
    cleaned_numbers = []
    for num in combined_string.split('-'):
        if num.endswith('10') or num.endswith('1'):
            num = num[:-2]
        cleaned_numbers.append(num)
    # Format the policy number as XXX-XXXX-XXXXX
    Policy = "-".join([
        cleaned_numbers[0][:3],    # First part: XXX (3 digits)
        cleaned_numbers[1][:4],    # Second part: XXXX (4 digits)
        cleaned_numbers[2][:5]     # Third part: XXXXX (5 digits)
    ])
    # Print the formatted policy number
    print("Policy number: {}".format(Policy))
    # Return the formatted policy number
    return Policy

def process_fee(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # ขยายภาพ
                        scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)
                        # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)
                        # แสดงภาพ
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.5, beta=-70)
                        # แสดงภาพ
                        #cv2_imshow( enhanced_image)
                        # แปลงภาพเป็นข้อความโดยใช้ pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        text = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        #print("เบี้ยประกันภัยที่เปลี่ยนแปลง: {}".format(text))
                        if "603" in text:
                            insurance_fee = 600.00
                        elif "904" in text:
                            insurance_fee = 900.00
                        else:
                            insurance_fee = "Not found"
                        print("Insurance premium:", insurance_fee)
                        return insurance_fee

def process_duty(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # ขยายภาพ
                        scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)
                        # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)
                        # แสดงภาพ
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.5, beta=-70)
                        # แสดงภาพ
                        #cv2_imshow( enhanced_image)
                        # แปลงภาพเป็นข้อความโดยใช้ pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        text = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        #print("เบี้ยประกันภัยที่เปลี่ยนแปลง: {}".format(text))
                        if "603" in text:
                            stamp_duty = 3.00
                        elif "904" in text:
                            stamp_duty = 4.00
                        else:
                            insurance_fee = "Not found"
                            stamp_duty = "Not found"
                        print("Stamp duty:", stamp_duty)
                        return stamp_duty

def process_Balanc(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # ขยายภาพ
                        scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                        # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(image, d=9, sigmaColor=75, sigmaSpace=75)
                        # ขยายภาพ
                        scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
                        #cv2_imshow(enhanced_image)
                        # ใช้ EasyOCR เพื่ออ่านข้อความ
                        # reader = easyocr.Reader(['en', 'th'], gpu= True)
                        reader = pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        result = reader.image_to_string(enhanced_image, config=custom_config)
                        # ใช้ Regular Expression เพื่อคัดแยกเฉพาะตัวเลข 0-9 และเครื่องหมาย /
                        filtered_result = re.findall(r'\d+\.\d+', result)
                        # พิมพ์ผลลัพธ์ที่คัดแยกแล้ว
                        print("Balanc: ", end='')
                        for Balanc in filtered_result:
                            print(Balanc)
                        return Balanc

def process_VAT(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # ขยายภาพ
                        scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                        # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(image, d=9, sigmaColor=75, sigmaSpace=75)
                        # ขยายภาพ
                        scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
                        #cv2_imshow(enhanced_image)
                        # ใช้ EasyOCR เพื่ออ่านข้อความ
                        # reader = easyocr.Reader(['en', 'th'], gpu= True)
                        reader = pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        result = reader.image_to_string(enhanced_image, config=custom_config)
                        # ใช้ Regular Expression เพื่อคัดแยกเฉพาะตัวเลข 0-9 และเครื่องหมาย /
                        filtered_result = re.findall(r'\d+\.\d+', result)
                        # พิมพ์ผลลัพธ์ที่คัดแยกแล้ว
                        print("VAT: ", end='')
                        for VAT in filtered_result:
                            print(VAT)
                        return VAT

def process_total(filename, rotated_image, x1, y1, x2, y2):
                        cropped_image = rotated_image[y1:y2, x1:x2]
                        # ขยายขนาดภาพ
                        scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
                        width = int(cropped_image.shape[1] * scale_percent / 100)
                        height = int(cropped_image.shape[0] * scale_percent / 100)
                        dim = (width, height)
                        resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
                        # แปลงภาพเป็นภาพสีเทา
                        gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)
                        # กรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
                        filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)
                        # ปรับสีให้เห็นชัดขึ้น
                        enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.5, beta=-70)
                        # กำหนดพารามิเตอร์สำหรับ pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        # แปลงภาพเป็นข้อความโดยใช้ pytesseract
                        Total = pytesseract.image_to_string(enhanced_image, lang='eng', config=custom_config)
                        #cv2_imshow(cropped_image)
                        #print("{}".format(Total_text))
                        # คัดลอกเฉพาะตัวเลขและจุดทศนิยมสองตำแหน่ง
                        import re
                        matches = re.findall(r'\b\d+\.\d+\b', Total)
                        if matches:
                            Total = matches[0]
                        else:
                            Total = "ไม่พบข้อมูล"
                        print("Total: {}".format(Total))
                        return Total

# รับข้อมูลจากผู้ใช้
input_data = input(" BKI-001, BKI-002, BKI-003: ")

# กำหนดตำแหน่งและขนาดของพิกัดสำหรับแต่ละประเภทข้อมูล
if input_data == "BKI-001":
  # ระบุโฟลเดอร์ที่มีไฟล์ภาพ
  folder_path = 'test_BKI_JPG'
  # ดึงรายการของไฟล์ภาพในโฟลเดอร์
  image_files = os.listdir(folder_path)
  # ตัวแปรเก็บจำนวนไฟล์ที่ถูกโหลด
  loaded_count = 0
  # วนลูปผ่านไฟล์ภาพแต่ละไฟล์
  for filename in image_files:
      # ตั้งค่าที่อยู่ของไฟล์ภาพ
      image_path = os.path.join(folder_path, filename)

      # โหลดภาพ
      image = cv2.imread(image_path)

      # ตรวจสอบว่าภาพถูกโหลดเข้ามาหรือไม่
      if image is not None:
          loaded_count += 1
          # ประมวลผลภาพ
          rotated_image = process_image(image)

          # ดำเนินการประมวลผลต่อไปหากภาพถูกหมุน
          if rotated_image is not None:
              # ประมวลผลประเภทข้อมูล
              TYPE = process_TYPE(rotated_image, filename, BKI_001_x1_TYPE, BKI_001_y1_TYPE, BKI_001_x2_TYPE, BKI_001_y2_TYPE )
              # ประมวลผลวันที่
              DATE = process_DATE(filename, rotated_image, BKI_001_x1_DATE, BKI_001_y1_DATE, BKI_001_x2_DATE, BKI_001_y2_DATE)
              # ประมวลผลชื่อ
              NAME = process_NAME(filename, rotated_image, BKI_001_x1_NAME, BKI_001_y1_NAME, BKI_001_x2_NAME, BKI_001_y2_NAME )
              # ประมวลผลเลขกรมธรรม์
              Policy = process_Policy(filename, rotated_image, BKI_001_x1_Policy, BKI_001_y1_Policy, BKI_001_x2_Policy, BKI_001_y2_Policy)
              # ประมวลผลเบี้ยประกันภัยที่เปลี่ยนแปลง(insurance_fee)
              insurance_fee = process_fee(filename, rotated_image, BKI_001_x1_fee_duty, BKI_001_y1_fee_duty, BKI_001_x2_fee_duty, BKI_001_y2_fee_duty )
              # ประมวลผลเบี้ยประกันภัยที่เปลี่ยนแปลง(stamp_duty)
              stamp_duty = process_duty(filename, rotated_image, BKI_001_x1_fee_duty, BKI_001_y1_fee_duty, BKI_001_x2_fee_duty, BKI_001_y2_fee_duty )
              # ประมวลผล ผลต่าง
              Balanc= process_Balanc(filename, rotated_image, BKI_001_x1_fee_duty, BKI_001_y1_fee_duty, BKI_001_x2_fee_duty, BKI_001_y2_fee_duty )
              # ประมวลผล VAT
              VAT= process_VAT(filename, rotated_image, BKI_001_x1_VAT, BKI_001_y1_VAT, BKI_001_x2_VAT, BKI_001_y2_VAT  )
              # ประมวลผล total
              total= process_total(filename, rotated_image, BKI_001_x1_total, BKI_001_y1_total, BKI_001_x2_total, BKI_001_y2_total  )
              # สร้าง DataFrame
              data = {
                'ประเภท': [TYPE],
                'วันที่': [DATE],
                'ชื่อ': [NAME],
                'เลขกรมธรรม์': [Policy],
                'เบี้ยประกันภัยที่เปลี่ยนแปลง': [insurance_fee],
                'อากรแสตมป์': [stamp_duty],
                'Balanc': [Balanc],
                'VAT': [VAT],
                'รวมเป็นเงิน': [ total]
                        }
              # สร้าง DataFrame
              df = pd.DataFrame(data)
              # สลับแถวและคอลัมน์
              df_transposed = df.T
              # ระบุที่อยู่ของไฟล์ Excel ที่ต้องการสร้าง
              excel_file = ("{}.xlsx".format(filename))
              # กำหนด path ของไฟล์ Excel ที่ต้องการบันทึก
              directory_path = 'excel_BKI'
              excel_path = os.path.join(directory_path,excel_file)
              # บันทึกไฟล์ Excel ไปยัง Google Drive
              #df.to_excel(excel_path, index=False)
              df.T.to_excel(excel_path, header=False)
              # แสดงที่อยู่ของไฟล์ Excel ที่ถูกบันทึก
              print("Excel file :", excel_path)

      else:
        print(f"ไม่สามารถโหลดไฟล์ภาพ {filename} ได้")

elif input_data == "BKI-002":
  print("ข้อมูลที่ป้อนไม่ถูกต้อง")

elif input_data == "BKI-003":
  print("ข้อมูลที่ป้อนไม่ถูกต้อง")

else:
    print("ข้อมูลที่ป้อนไม่ถูกต้อง")