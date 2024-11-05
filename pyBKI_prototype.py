import os
import cv2
import numpy as np
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import re
import sys
import pandas as pd
import matplotlib.pyplot as plt

def main(args):
    if args:  # ตรวจสอบว่ามี arguments ที่ถูกส่งเข้ามาหรือไม่
        text_choice = args[0]  # เก็บค่าตัวแรกใน args ไว้ในตัวแปร first_choice
        print(f"Returned choice: {text_choice}")
        return text_choice  # ส่งค่าตัวแรกใน args กลับออกจากฟังก์ชัน main()

def read_file_based_on_choice(text_choice):
    file_path_pixel = f'C:\\Users\\ASUS\\Desktop\\OCR\\pixels location\\{text_choice}.xlsx'
    return file_path_pixel

def process_data(text_choice):
    result = read_file_based_on_choice(text_choice)

    if result is not None:
        df = pd.read_excel(result)  # อ่านไฟล์ Excel จาก file_path ที่ได้

        if not df.empty:
            data = {}
            #print(df.head())

            # เลือกข้อมูลแต่ละแถวและกำหนดให้กับตัวแปรต่างๆ
            for i in range(12):  # วนลูป 12 ครั้ง
                selected_data = df.iloc[i, :]
                data[f'x1_{i}'] = selected_data['x1']
                data[f'y1_{i}'] = selected_data['y1']
                data[f'x2_{i}'] = selected_data['x2']
                data[f'y2_{i}'] = selected_data['y2']


            # กำหนดค่าตัวแปรตามที่ต้องการ
            x1_TYPE, y1_TYPE, x2_TYPE, y2_TYPE = data['x1_0'], data['y1_0'], data['x2_0'], data['y2_0']
            x1_DATE, y1_DATE, x2_DATE, y2_DATE = data['x1_1'], data['y1_1'], data['x2_1'], data['y2_1']
            x1_NAME, y1_NAME, x2_NAME, y2_NAME = data['x1_2'], data['y1_2'], data['x2_2'], data['y2_2']
            x1_Policy, y1_Policy, x2_Policy, y2_Policy = data['x1_3'], data['y1_3'], data['x2_3'], data['y2_3']
            x1_prem, y1_prem, x2_prem, y2_prem = data['x1_4'], data['y1_4'], data['x2_4'], data['y2_4']
            x1_stamp, y1_stamp, x2_stamp, y2_stamp = data['x1_5'], data['y1_5'], data['x2_5'], data['y2_5']
            x1_Balanc, y1_Balanc, x2_Balanc, y2_Balanc = data['x1_6'], data['y1_6'], data['x2_6'], data['y2_6']
            x1_VAT, y1_VAT, x2_VAT, y2_VAT = data['x1_7'], data['y1_7'], data['x2_7'], data['y2_7']
            x1_total, y1_total, x2_total, y2_total = data['x1_8'], data['y1_8'], data['x2_8'], data['y2_8']
            rotate = data['x1_9']
            folder_PDF_path = data['x1_10']
            folder_OUTPUT_path = data['x1_11']
            

            return (x1_TYPE, y1_TYPE, x2_TYPE, y2_TYPE,
                    x1_DATE, y1_DATE, x2_DATE, y2_DATE,
                    x1_NAME, y1_NAME, x2_NAME, y2_NAME,
                    x1_Policy, y1_Policy, x2_Policy, y2_Policy,
                    x1_prem, y1_prem, x2_prem, y2_prem,
                    x1_stamp, y1_stamp, x2_stamp, y2_stamp,
                    x1_Balanc, y1_Balanc, x2_Balanc, y2_Balanc,
                    x1_VAT, y1_VAT, x2_VAT, y2_VAT,
                    x1_total, y1_total, x2_total, y2_total,
                    rotate,folder_PDF_path,folder_OUTPUT_path
                    )

        else:
            print("ไฟล์ Excel ไม่มีข้อมูล")
    else:
        print("ไม่สามารถอ่านไฟล์ Excel ได้")

# ฟังก์ชันปรับขอบของกรอบ
def adjust_contours(cnts, left_top_padding=0, left_bottom_padding=0, right_top_padding=0, right_bottom_padding=0):
    adjusted_cnts = cnts.copy()
    
    # Adjust each point based on the padding values
    adjusted_cnts[0][0][1] += right_top_padding       # Top-right
    adjusted_cnts[3][0][1] += left_top_padding         # Top-left
    adjusted_cnts[1][0][1] += left_bottom_padding      # Bottom-left
    adjusted_cnts[2][0][1] += right_bottom_padding     # Bottom-right
    
    return adjusted_cnts

def process_image(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Apply GaussianBlur to reduce noise
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

    # Apply Canny edge detection
    edges = cv2.Canny(blurred, 75, 200)
    kernel = np.ones((5, 5), np.uint8)
    dilate = cv2.dilate(edges, kernel, iterations=2)

    # Find contours
    contours, _ = cv2.findContours(dilate.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)

    cnts = None
    for contour in contours:
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.05 * peri, True)
        if len(approx) == 4:
            cnts = approx
            break

    if cnts is None:
        print("Contour with 4 points not found.")
        return None

    # Adjust the contour by adding padding (customize as needed)
    cnts = adjust_contours(cnts, left_top_padding=-250, left_bottom_padding=0, right_top_padding=-250, right_bottom_padding=0)

    cv2.drawContours(image, [cnts], -1, (0, 255, 0), 3)

    def order_points(pts):
        rect = np.zeros((4, 2), dtype="float32")
        s = pts.sum(axis=1)
        rect[0] = pts[np.argmin(s)]  # Top-left point
        rect[2] = pts[np.argmax(s)]  # Bottom-right point

        diff = np.diff(pts, axis=1)
        rect[1] = pts[np.argmin(diff)]  # Top-right point
        rect[3] = pts[np.argmax(diff)]  # Bottom-left point
        return rect

    def four_point_transform(image, pts):
        rect = order_points(pts)
        (tl, tr, br, bl) = rect

        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        dst = np.array([
            [0, 0],
            [maxWidth - 1, 0],
            [maxWidth - 1, maxHeight - 1],
            [0, maxHeight - 1]], dtype="float32")

        M = cv2.getPerspectiveTransform(rect, dst)
        rotated_image = cv2.warpPerspective(image, M, (maxWidth, maxHeight))

        return rotated_image

    # Apply perspective transformation on the image
    rotated_image = four_point_transform(image, cnts.reshape(4, 2))

    if rotated_image is not None:
        # Show the rotated_image result
        cv2.imshow("Rotated Image", rotated_image)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

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

def process_DATE_num(filename, rotated_image, x1, y1, x2, y2):
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

def process_DATE_Mtext(filename, rotated_image, x1, y1, x2, y2):
    cropped_image = rotated_image[y1:y2, x1:x2]
    # แปลงภาพเป็นภาพสีเทา
    gray = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2GRAY)
    # ทำการกรองภาพแบบ Bilateral Filtering เพื่อลด noise และเพิ่มความคมชัดของภาพ
    filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)
    # ขยายภาพ
    scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(filtered_image.shape[1] * scale_percent / 100)
    height = int(filtered_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(filtered_image, dim, interpolation=cv2.INTER_AREA)
    # ปรับสีให้เห็นชัดขึ้น
    enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
    
    # ใช้ pytesseract เพื่ออ่านข้อความ
    custom_config = r'--oem 3 --psm 6'
    DAY = pytesseract.image_to_string(enhanced_image, lang='tha+eng', config=custom_config)
    return DAY

def process_NAME(filename, rotated_image, x1, y1, x2, y2):
    # ตัดภาพตามพิกัดที่กำหนด
    cropped_image = rotated_image[y1:y2, x1:x2]
    # ขยายขนาดภาพ
    scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)
    # ปรับสีให้เห็นชัดขึ้น
    enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)
    plt.figure(figsize=(10, 5))
    plt.subplot(1, 2, 1)
    plt.imshow(cv2.cvtColor(cropped_image, cv2.COLOR_BGR2RGB))
    plt.title('Cropped Image')
    plt.show()
    # ใช้ pytesseract เพื่ออ่านข้อความ (ชื่อ)
    custom_config = r'--oem 3 --psm 6'
    detected_text = pytesseract.image_to_string(enhanced_image, lang='tha+eng', config=custom_config)
    print(detected_text)
    # ตรวจสอบคำที่ได้จาก pytesseract
    valid_prefixes = ["นาย", "นางสาว", "นาง", "คุณ", "Mr", "Mrs", "Ms", "Miss", "บริษัท","น.ส.","น.,ส,."]
    found_prefixes = []
    # ตรวจสอบและเก็บคำที่พบ
    for prefix in valid_prefixes:
        if prefix in detected_text:
            found_prefixes.append(prefix)
            # เพิ่มเงื่อนไขเพิ่มเติมเมื่อพบคำนำหน้าเป็น "บริษัท"
            if prefix == "บริษัท":
                # หาคำที่เจอหลัง "บริษัท"
                index = detected_text.find(prefix) + len(prefix)
                remaining_text = detected_text[index:].strip()
                # ตัดผลลัพธ์ที่พบหลัง " " และเอาเฉพาะชื่อบริษัท
                remaining_text = remaining_text.split()[0]
                NAME = "{} {}".format(prefix, remaining_text)
                print("Name :", NAME)
                return NAME

    # หาตำแหน่งของคำนำหน้าในกรณีที่ไม่เป็น "บริษัท"
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
    text = pytesseract.image_to_string(cropped_image, lang='eng', config=custom_config)
    # Extract all sequences of digits from the OCR result
    #filtered_result = re.findall(r'\b[A-Za-z]+\.[A-Za-z]+\.[0-9-]+\b', text)
    # Concatenate all digits into a single string
    num_policy = ''.join(text)
    return num_policy

def check_policy_BKI(num_policy):
    cleaned_numbers = []
    for num in num_policy.split('-'):
        if num.endswith('10') or num.endswith('1'):
            num = num[:-2]
        cleaned_numbers.append(num)
    
    # Process the first part to extract last 3 digits
    if len(cleaned_numbers) > 0:
        first_part = cleaned_numbers[0][-3:]  # Take last 3 digits
    else:
        first_part = 'XXX'
    # Second part: XXXXX (5 digits) or 'XXXXX' if not found
    second_part = cleaned_numbers[1][:5] if len(cleaned_numbers) > 1 else 'XXXXX'
    
    # Third part: XXXXXX (6 digits) or 'XXXXXX' if not found
    third_part = cleaned_numbers[2][:6] if len(cleaned_numbers) > 2 else 'XXXXXX'
    
    # Check if any part is missing and mark it as "(ข้อมูลผิดพลาด)"
    parts = [
        first_part if first_part else '(ข้อมูลผิดพลาด)',
        second_part if second_part else '(ข้อมูลผิดพลาด)',
        third_part if third_part else '(ข้อมูลผิดพลาด)'
    ]
    
    # Format the policy number as XXX-XXXX-XXXXX
    Policy = "-".join(parts)
    print("Policy number: {}".format(Policy))
    return Policy

def check_policy_MIT_MT(num_policy):
    # Remove any spaces or special characters that should not be part of the output
    num_policy = num_policy.replace(" ", "").replace("_", "")

    # Split the policy number by '.'
    parts = num_policy.split('.')

    # Check for "TV" at the beginning
    if len(parts) < 2 or parts[0] != "TV":
        return f"{num_policy} ไม่มี 'TV' ที่คาดหวัง"

    # Check for "VMI" after "TV"
    if len(parts) < 3 or parts[1] != "VMI":
        return f"{num_policy} ไม่มี 'VMI' ที่คาดหวังหลัง 'TV'"

    # Find the third part (7 digits or less)
    digits = re.search(r'\d{1,7}', parts[2])
    if not digits:
        return f"{num_policy} ไม่มีตัวเลข 7 ตัวหลัง 'VMI'"

    # Construct the expected format
    Policy = f"TV.VMI.{digits.group(0)}"

    return Policy

def process_pream(filename, rotated_image, x1, y1, x2, y2):
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
                        insurance_fee = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        print("insurance_fee:",format(insurance_fee))
                        return insurance_fee

def process_stamp(filename, rotated_image, x1, y1, x2, y2):
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
                        enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.50, beta=-70)
                        # แสดงภาพ
                        #cv2_imshow( enhanced_image)
                        # แปลงภาพเป็นข้อความโดยใช้ pytesseract
                        custom_config = r'--oem 3 --psm 6'
                        stamp_duty = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        print("stamp_duty:",format(stamp_duty))
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
                        Balanc = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        print("Balanc:",format(Balanc))
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
                        VAT = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        print("VAT:",format(VAT))
                        return VAT

def process_total(filename, rotated_image, x1, y1, x2, y2):
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
                        total = pytesseract.image_to_string(enhanced_image, lang='eng',config=custom_config)
                        print("total:",format(total))
                        return total

def adjust_ocr_value(ocr_result):
    # การกรองอักขระที่ไม่ใช่ตัวเลข, จุลภาค, และจุลภาคใหญ่ออก
    valid_chars = '0123456789.,-'  # Include comma (,) as a valid character
    filtered_result = ''.join(char for char in ocr_result if char in valid_chars)

    # การแทนที่เครื่องหมาย "," เป็น "."
    filtered_result = filtered_result.replace(',', '.')

    # การจัดการกับจุดทศนิยม (หากมีมากกว่า 1 จุดทศนิยม)
    decimal_count = filtered_result.count('.')
    if decimal_count > 1:
        last_decimal_index = filtered_result.rfind('.')
        decimal_part = filtered_result[last_decimal_index + 1:]
        integer_part = filtered_result[:last_decimal_index].replace('.', '', decimal_count - 1)
        filtered_result = f"{integer_part}.{decimal_part}"

    # การลบเครื่องหมายลบที่ติดกับจุดทศนิยม
    if '.' in filtered_result and '-' in filtered_result:
        minus_index = filtered_result.find('-')
        filtered_result = filtered_result[:minus_index] + filtered_result[minus_index+1:]

    # การส่งคืนผลลัพธ์
    return filtered_result.strip()

if __name__ == "__main__": 

    text_choice = main(sys.argv[1:])
    #print("Returned choice:", text_choice)
    processed_data = process_data(text_choice)
    
    if processed_data is not None:
        (x1_TYPE, y1_TYPE, x2_TYPE, y2_TYPE,
        x1_DATE, y1_DATE, x2_DATE, y2_DATE,
        x1_NAME, y1_NAME, x2_NAME, y2_NAME,
        x1_Policy, y1_Policy, x2_Policy, y2_Policy,
        x1_prem, y1_prem, x2_prem, y2_prem,
        x1_stamp, y1_stamp, x2_stamp, y2_stamp,
        x1_Balanc, y1_Balanc, x2_Balanc, y2_Balanc,
        x1_VAT, y1_VAT, x2_VAT, y2_VAT,
        x1_total, y1_total, x2_total, y2_total,
        rotate,folder_PDF_path,folder_OUTPUT_path) = processed_data
        #print({},folder_OUTPUT_path)

        folder_path = f'{folder_PDF_path}'
        image_files = os.listdir(folder_path)
        loaded_count = 0
        
        for filename in image_files:
            image_path = os.path.join(folder_path, filename)
            image = cv2.imread(image_path)
            
            if image is not None:
                loaded_count += 1
                
                # หมุนภาพถ้าเป็น Y
                if rotate == "Y" or rotate == "y":
                    rotated_image = process_image(image)
                    print(f"rotate : {rotate}ES")
                else:
                    rotated_image = image  
                    print(f"rotate : NO")
                
                # Process each type of data
                TYPE = process_TYPE(rotated_image, filename, x1_TYPE, y1_TYPE, x2_TYPE, y2_TYPE)
                if text_choice == "BKI-test":
                    DATE = process_DATE_num(filename, rotated_image, x1_DATE, y1_DATE, x2_DATE, y2_DATE)
                elif text_choice == "MIT-MT":
                    DATE = process_DATE_Mtext(filename, rotated_image, x1_DATE, y1_DATE, x2_DATE, y2_DATE)
                #print(DATE)
                NAME = process_NAME(filename, rotated_image, x1_NAME, y1_NAME, x2_NAME, y2_NAME)
                Policy = process_Policy(filename, rotated_image, x1_Policy, y1_Policy, x2_Policy, y2_Policy)
                if text_choice == "BKI-test":
                    num_policy = Policy  # Assign the value returned from process_Policy
                    Policy = check_policy_BKI(num_policy)  # Pass num_policy to check_policy_BKI
                elif text_choice == "MIT-MT":
                    num_policy = Policy  # Assign the value returned from process_Policy
                    Policy = check_policy_MIT_MT(num_policy)  # Pass num_policy to check_policy_BKI 

                insurance_fee = process_pream(filename, rotated_image, x1_prem, y1_prem, x2_prem, y2_prem,)
                pream = adjust_ocr_value(insurance_fee)
                
                stamp_duty = process_stamp(filename, rotated_image, x1_stamp, y1_stamp, x2_stamp, y2_stamp)
                stamp = adjust_ocr_value(stamp_duty)
                
                Balanc_num = process_Balanc(filename, rotated_image, x1_Balanc, y1_Balanc, x2_Balanc, y2_Balanc) 
                Balanc = adjust_ocr_value(Balanc_num)
                
                VAT_num = process_VAT(filename, rotated_image, x1_VAT, y1_VAT, x2_VAT, y2_VAT)
                VAT = adjust_ocr_value(VAT_num)
                
                total_num = process_total(filename, rotated_image, x1_total, y1_total, x2_total, y2_total)
                total = adjust_ocr_value(total_num)


                # Create DataFrame and save to Excel
                data = {
                    'ประเภท': [TYPE],
                    'วันที่': [DATE],
                    'ชื่อ': [NAME],
                    'เลขกรมธรรม์': [Policy],
                    'เบี้ยประกันภัยที่เปลี่ยนแปลง': [pream],
                    'อากรแสตมป์': [stamp],
                    'Balanc': [Balanc],
                    'VAT': [VAT],
                    'รวมเป็นเงิน': [total]
                }
                
                df = pd.DataFrame(data)
                df.T.to_excel(f'{folder_OUTPUT_path}\{filename}.xlsx', header=False)
                print("Excel file saved successfully.")
            print("_____________________________________________________")
    
    else:
        print("ไม่สามารถดึงข้อมูลจากไฟล์ Excel ได้")
