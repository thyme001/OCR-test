![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303574825708814367/image.png?ex=672c4029&is=672aeea9&hm=3a5b85ae2a5cdfaffef769dbdf953b377eefc5b6755ea9ce05a28b9caff3ff38&)
`main()` ทำงานดังนี้:

1.  **รับ Argument**: ฟังก์ชันนี้รับพารามิเตอร์ `args` ซึ่งเป็น list ของ arguments ที่ส่งเข้ามาในฟังก์ชัน `main()` (เช่น ข้อความหรือตัวเลขหลายค่าในรูปของ list)
    
2.  **ตรวจสอบ Arguments**: `if args:` เช็คว่ามี arguments ใน list `args` หรือไม่ ถ้า `args` มีค่าหรือไม่เป็นค่าว่าง (`[]`) จะเข้าสู่บรรทัดต่อไป
    
3.  **เก็บค่า Argument แรก**:
    
    -   `text_choice = args[0]` เก็บค่าตัวแรกใน `args` ไว้ในตัวแปร `text_choice`
4.  **แสดงผลและส่งค่ากลับ**:
    
    -   `print(f"Returned choice: {text_choice}")` แสดงข้อความ "Returned choice: " พร้อมกับค่า `text_choice`
    -   `return text_choice` ส่งค่าของ `text_choice` ออกไปจากฟังก์ชัน ซึ่งทำให้ค่าตัวแรกของ arguments ถูกส่งกลับไปให้ตัวเรียกใช้ฟังก์ชัน `main()`

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303575844068921417/image.png?ex=672c411c&is=672aef9c&hm=46507a61e039bddf28becb4f0d31ff65e7cd93dae7001c85da8e2b0dadc261dd&)
ฟังก์ชัน `read_file_based_on_choice(text_choice)` ทำงานดังนี้:

1.  **สร้างเส้นทางไฟล์**:
    
    -   ฟังก์ชันนี้สร้าง path ของไฟล์ Excel โดยใช้ค่า `text_choice` ที่ส่งเข้ามา
    -   `file_path_pixel = f'C:\\Users\\ASUS\\Desktop\\OCR\\pixels location\\{text_choice}.xlsx'` สร้างเส้นทางไปยังไฟล์ `.xlsx` โดยเติมค่า `text_choice` ลงในชื่อไฟล์
    -   ตัวอย่างเช่น หาก `text_choice` เป็น `"data1"`, เส้นทางไฟล์จะเป็น: `C:\Users\ASUS\Desktop\OCR\pixels location\data1.xlsx`
2.  **ส่งคืนเส้นทางไฟล์**:
    
    -   ฟังก์ชันจะส่งค่าของ `file_path_pixel` กลับไป ซึ่งเป็น path ของไฟล์ Excel ที่ตั้งอยู่ในโฟลเดอร์ `"pixels location"` บน Desktop ของผู้ใช้

ฟังก์ชันนี้ใช้สร้าง path ไฟล์ตามชื่อที่ผู้ใช้ระบุ เพื่อให้ง่ายต่อการเข้าถึงไฟล์เฉพาะในระบบ

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303576192812843069/image.png?ex=672c416f&is=672aefef&hm=c7628b5d48d5c01e0c19c883826b8317e5e7aa37e2951da90096e6ca285323a4&)
ฟังก์ชัน `process_data(text_choice)` ทำงานดังนี้:

1.  **อ่านไฟล์ Excel**:
    
    -   ฟังก์ชัน `process_data` เริ่มต้นด้วยการเรียกใช้ `read_file_based_on_choice(text_choice)` เพื่อรับเส้นทางไฟล์ Excel ที่จะถูกเปิด
    -   ถ้าไฟล์นั้นไม่เป็น `None` (หมายถึงสามารถหาไฟล์ได้), ฟังก์ชันจะใช้ `pd.read_excel(result)` เพื่ออ่านข้อมูลจากไฟล์นั้นลงใน DataFrame (ตัวแปร `df`)
2.  **ตรวจสอบว่า DataFrame ว่างหรือไม่**:
    
    -   ถ้า DataFrame (`df`) ไม่ว่าง (ไม่ใช่ empty), ฟังก์ชันจะดำเนินการต่อไป
    -   ฟังก์ชันจะเลือกข้อมูลจากแถวที่ 0-11 ของ `df` โดยใช้การวนลูป `for i in range(12)` และนำค่าจากคอลัมน์ `x1`, `y1`, `x2`, และ `y2` ในแต่ละแถวมาบันทึกลงใน dictionary `data`
3.  **กำหนดค่าตัวแปร**:
    
    -   หลังจากดึงข้อมูลมาแล้ว, ฟังก์ชันจะกำหนดค่าตัวแปรต่างๆ เช่น `x1_TYPE`, `y1_TYPE`, `x2_TYPE`, `y2_TYPE` เป็นต้น โดยการใช้ข้อมูลจาก `data` ตามลำดับแถวต่างๆ
4.  **ส่งคืนค่าตัวแปร**:
    
    -   ฟังก์ชันจะส่งคืนค่าของตัวแปรต่างๆ ที่ได้จาก DataFrame เช่น ข้อมูลจากคอลัมน์ `x1`, `y1`, `x2`, `y2` สำหรับแต่ละแถว รวมถึงค่า `rotate`, `folder_PDF_path`, และ `folder_OUTPUT_path`
5.  **กรณีข้อมูลไม่ถูกต้อง**:
    
    -   ถ้าไม่สามารถอ่านไฟล์ Excel ได้หรือ DataFrame ว่าง, ฟังก์ชันจะพิมพ์ข้อความว่า "ไม่สามารถอ่านไฟล์ Excel ได้" หรือ "ไฟล์ Excel ไม่มีข้อมูล"

ฟังก์ชันนี้ใช้สำหรับการประมวลผลข้อมูลจากไฟล์ Excel และแยกค่าออกมาเป็นตัวแปรที่ใช้งานในโปรแกรมได้

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303576585655423006/image.png?ex=672c41cd&is=672af04d&hm=ca395472d1846f41013582f1cca331fa5e06ba826feeadb8a9e00d401838adec&)
ฟังก์ชัน `adjust_contours` มีไว้เพื่อปรับขอบของกรอบ (Contours) ตามค่าการปรับแต่ง (padding) ที่กำหนดไว้ในแต่ละมุม ซึ่งทำงานดังนี้:

1.  **รับพารามิเตอร์**:
    
    -   `cnts`: Contour ของกรอบที่ต้องการปรับขอบ โดยปกติจะเป็น `numpy array` ที่เก็บพิกัดของมุมต่างๆ
    -   `left_top_padding`, `left_bottom_padding`, `right_top_padding`, `right_bottom_padding`: ค่าการปรับตำแหน่งในแต่ละมุมของกรอบ
2.  **คัดลอกข้อมูล**:
    
    -   `adjusted_cnts = cnts.copy()` ทำการคัดลอกข้อมูลของกรอบ `cnts` เพื่อไม่ให้เปลี่ยนแปลงข้อมูลต้นฉบับโดยตรง
3.  **ปรับค่าพิกัดแต่ละมุม**:
    
    -   `adjusted_cnts[0][0][1] += right_top_padding`: ปรับตำแหน่งในแนวตั้ง (y-coordinate) ของมุมขวาบน
    -   `adjusted_cnts[3][0][1] += left_top_padding`: ปรับตำแหน่งของมุมซ้ายบน
    -   `adjusted_cnts[1][0][1] += left_bottom_padding`: ปรับตำแหน่งของมุมซ้ายล่าง
    -   `adjusted_cnts[2][0][1] += right_bottom_padding`: ปรับตำแหน่งของมุมขวาล่าง
4.  **ส่งค่ากลับ**:
    
    -   ส่งค่าของ `adjusted_cnts` ที่มีการปรับขอบเรียบร้อยแล้วออกจากฟังก์ชัน

### การใช้งาน

ฟังก์ชันนี้ช่วยในการปรับขอบกรอบโดยเพิ่มระยะในแนวตั้งให้กับแต่ละมุมตามค่าที่กำหนดสำหรับใช้งานปรับแต่งขอบของกรอบให้แม่นยำ

![enter image description here](https://media.discordapp.net/attachments/1263395471833960476/1303577055706873878/image.png?ex=672c423d&is=672af0bd&hm=8cf9c4797c21744905ce42841dd71dcb6c90e33506c333079a062ffa3aab5eb3&=&format=webp&quality=lossless&width=528&height=662)
ฟังก์ชัน `process_image` ใช้สำหรับค้นหากรอบ (contour) ที่ต้องการในภาพ จากนั้นปรับขอบและใช้การแปลง perspective เพื่อให้ได้ภาพที่ถูกตัดตามกรอบที่ต้องการ โดยมีขั้นตอนการทำงานดังนี้:

1.  **แปลงภาพเป็น Grayscale**:
   
    `gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)` 
    
    แปลงภาพให้เป็นโทนสีเทาเพื่อลดข้อมูลในภาพ ทำให้ง่ายต่อการประมวลผลขอบ (edge detection)
    
2.  **ลด noise ด้วย Gaussian Blur**:
    `blurred = cv2.GaussianBlur(gray, (5, 5), 0)` 
    
    ใช้ Gaussian Blur เพื่อทำให้ภาพเนียนขึ้น ลดความผิดพลาดในการตรวจจับขอบจาก noise
    
3.  **ตรวจจับขอบด้วย Canny Edge Detection**:
    
    `edges = cv2.Canny(blurred, 75, 200)` 
    
    ใช้ Canny edge detection เพื่อตรวจจับขอบในภาพ
    
4.  **ขยายขอบ (Dilate)**:
    
    `kernel = np.ones((5, 5), np.uint8)
    dilate = cv2.dilate(edges, kernel, iterations=2)` 
    
    ขยายขอบของพื้นที่ที่ตรวจจับได้ ทำให้เส้นขอบชัดเจนขึ้น
    
5.  **หากรอบ (Contours)**:
    
    `contours, _ = cv2.findContours(dilate.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)` 
    
    หาขอบเขต (contour) ของภาพที่ผ่านการประมวลผล โดยเรียงลำดับขนาดจากใหญ่ไปเล็ก
    
6.  **เลือก Contour ที่มีจุด 4 จุด**:
    
    `for contour in contours:
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.05 * peri, True)
        if len(approx) == 4:
            cnts = approx
            break` 
    
    ค้นหา contour ที่มีจุดโค้งเว้า (approximation) จำนวน 4 จุด เพื่อนำมาใช้เป็นกรอบ (เหมาะกับการหากรอบรูปสี่เหลี่ยม)
    
7.  **ปรับขอบ Contour ด้วย `adjust_contours`**:
    
    `cnts = adjust_contours(cnts, left_top_padding=-250, left_bottom_padding=0, right_top_padding=-250, right_bottom_padding=0)` 
    
    ใช้ฟังก์ชัน `adjust_contours` เพื่อปรับขอบแต่ละด้านตามค่าที่กำหนดใน padding
    
8.  **จัดเรียงจุดในกรอบให้ถูกต้อง**:
    
    `def order_points(pts):
        # Function code to order points` 
    
    กำหนดฟังก์ชัน `order_points` สำหรับจัดเรียงจุดในกรอบให้ตรงตามลำดับที่ต้องการ (บนซ้าย, บนขวา, ล่างขวา, ล่างซ้าย)
    
9.  **แปลง Perspective ด้วย `four_point_transform`**:
    
    `def four_point_transform(image, pts):
        # Function code to apply perspective transform` 
    
    ใช้ฟังก์ชัน `four_point_transform` เพื่อแปลง perspective ของภาพให้อยู่ในมุมมองตรงตามกรอบที่เลือก
    
10.  **แสดงภาพที่ถูกแปลงแล้ว**:
    
    `cv2.imshow("Rotated Image", rotated_image)
    cv2.waitKey(0)
    cv2.destroyAllWindows()` 
    
    แสดงภาพที่ถูกแปลงให้ผู้ใช้ดู
    
11.  **ส่งผลลัพธ์กลับจากฟังก์ชัน**:
    
    `return rotated_image` 
    

### สรุป:

ฟังก์ชัน `process_image` ทำการตรวจจับขอบ ปรับขอบ และแปลง perspective ของภาพโดยอิงจากกรอบที่ค้นหาได้

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303579386166906890/image.png?ex=672c4468&is=672af2e8&hm=2a81df5d7bee9e830c2525e7d09d7a13fd3b26996f84f65b532ce3bba3bfa32a&)
ฟังก์ชัน `process_TYPE` นี้มีการทำงานเพื่อประมวลผลภาพที่ถูกส่งเข้ามา โดยมีขั้นตอนหลักดังนี้:

1.  **กำหนดค่าความกว้างและความสูงใหม่**:
    
    `new_width = 100
    new_height = 100` 
    
    กำหนดความกว้างและความสูงใหม่สำหรับการปรับขนาดภาพที่ต้องการ
    
2.  **ปรับขนาดรูปภาพ**:
    
    `resized_image = cv2.resize(rotated_image, (new_width, new_height))` 
    
    ใช้ `cv2.resize` เพื่อเปลี่ยนขนาดภาพตามค่าความกว้างและความสูงที่ตั้งไว้
    
3.  **ตัดรูปภาพ (Crop)**:

    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
    ตัดเฉพาะส่วนที่ต้องการโดยกำหนดพิกัด `(x1, y1)` ถึง `(x2, y2)`
    
4.  **เพิ่มความคมชัดให้รูป**:
 
    `enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)` 
    
    ใช้ `convertScaleAbs` เพื่อเพิ่มความคมชัดของภาพโดยการปรับค่า `alpha` และ `beta` เพื่อปรับคอนทราสต์และความสว่างของภาพ
    
5.  **เรียกใช้ OCR เพื่อดึงข้อความจากภาพที่ถูกตัด**:
 
    `custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(cropped_image, lang='eng', config=custom_config)` 
    
    ใช้ `pytesseract` เพื่อทำ OCR บนภาพที่ถูกตัด และดึงข้อความออกมา
    
6.  **ค้นหาตำแหน่งของเครื่องหมายวงเล็บเปิดและปิด**:

    `opening_brackets = ["(", "{", "["]
    closing_brackets = [")", "}", "]"]
    opening_indexes = [text.find(bracket) for bracket in opening_brackets if text.find(bracket) != -1]
    closing_indexes = [text.rfind(bracket) for bracket in closing_brackets if text.rfind(bracket) != -1]` 
    
    ตรวจสอบข้อความที่ OCR ว่ามีเครื่องหมายวงเล็บเปิด (`(`, `{`, `[`) และวงเล็บปิด (`)`, `}`, `]`) อยู่หรือไม่ และดึงตำแหน่งของเครื่องหมายเหล่านั้น
    
7.  **แยกข้อความระหว่างเครื่องหมายวงเล็บ**:
   
    `if 'opening_index' in locals() and 'closing_index' in locals() and opening_index != -1 and closing_index != -1:
        extracted_text = text[opening_index + 1 : closing_index]
        TYPE = extracted_text
        print("type:", TYPE)
        return TYPE
    else:
        print("ไม่พบเครื่องหมายเปิดและปิดคำสั่งในข้อความ")` 
    
    ถ้าพบเครื่องหมายวงเล็บเปิดและปิด จะดึงข้อความระหว่างวงเล็บออกมาและเก็บค่าไว้ใน `TYPE`
    
8.  **คืนค่า `TYPE`**: ฟังก์ชันจะคืนค่าข้อความที่ดึงออกมา (`TYPE`) หรือแจ้งเตือนถ้าไม่พบเครื่องหมายวงเล็บในข้อความ
    

### หมายเหตุ

-   ค่าของ `alpha` และ `beta` ในการปรับความคมชัดอาจต้องปรับให้เหมาะสมกับแต่ละภาพ เพื่อให้ OCR ทำงานได้ดีที่สุด
-   พารามิเตอร์ `x1`, `y1`, `x2`, `y2` ควรกำหนดให้สอดคล้องกับตำแหน่งของข้อความในภาพ

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303584507215089715/image.png?ex=672c492d&is=672af7ad&hm=acec35ba6ffb46ead832f97d47753c6937d3a4b784b9aedf0e73d84bf721f39b&)
ฟังก์ชัน `process_DATE_num` นี้ทำงานเพื่อดึงวันที่จากรูปภาพที่ถูกสแกน โดยมีขั้นตอนต่าง ๆ ดังนี้:

1.  **ตัดรูปภาพตามพิกัด**:
   
    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
    ทำการตัดรูปภาพให้เป็นส่วนที่ต้องการจาก `rotated_image` โดยกำหนดพิกัด `(x1, y1)` ถึง `(x2, y2)` ซึ่งเป็นตำแหน่งที่ต้องการในภาพ
    
2.  **แปลงภาพเป็นสีเทา**:
  
    `gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)` 
    
    ใช้ฟังก์ชัน `cv2.cvtColor` เพื่อแปลงภาพเป็นสีเทา (`GRAY`) เพื่อให้ OCR ทำงานได้ดีกว่า
    
3.  **กรองภาพแบบ Bilateral Filtering**:

    `filtered_image = cv2.bilateralFilter(image, d=9, sigmaColor=75, sigmaSpace=75)` 
    
    ใช้ `cv2.bilateralFilter` เพื่อกรองภาพ ลด noise และทำให้ภาพชัดขึ้นเพื่อให้ OCR ทำงานได้ดีขึ้น
    
4.  **ขยายขนาดภาพ**:
  
    `scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)` 
    
    ขยายขนาดของรูปภาพที่ตัดแล้วให้มีขนาดใหญ่ขึ้น โดยกำหนดเปอร์เซ็นต์การขยาย
    
5.  **ปรับความคมชัดของภาพ**:

    `enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)` 
    
    ปรับคอนทราสต์และความสว่างของภาพเพื่อทำให้ข้อความชัดเจนขึ้น
    
6.  **ใช้ OCR เพื่อดึงข้อความจากภาพ**:
   
    `reader = pytesseract
    custom_config = r'--oem 3 --psm 6'
    result = reader.image_to_string(enhanced_image, config=custom_config)` 
    
    ใช้ `pytesseract` เพื่อดึงข้อความจากภาพที่ถูกปรับขนาดและปรับปรุงแล้ว
    
7.  **กรองเฉพาะตัวเลขและเครื่องหมาย "/"**:

    `filtered_result = re.findall(r'[\d/]+', result)` 
    
    ใช้ Regular Expression (RegEx) เพื่อค้นหาตัวเลขและเครื่องหมาย `/` ในข้อความที่ OCR ดึงออกมา
    
8.  **ใช้ RegEx เพื่อค้นหาวันที่**:

    `pattern = r'(0[1-9]|1[0-9]|2[0-9]|3[01])/(0[1-9]|1[012])/(\d{4})'
    matches = re.findall(pattern, result)` 
    
    ใช้ RegEx เพื่อตรวจสอบว่ามีข้อความที่ตรงกับรูปแบบวันที่ในรูปแบบ `DD/MM/YYYY` หรือไม่
    
9.  **พิมพ์และส่งคืนวันที่**:
  
    `if not matches:
        print("Not found")
    else:
        for match in matches:
            DAY = "{}/{}/{}".format(match[0], match[1], match[2])
            print("DATE:", DAY)
        return DAY` 
    
    ถ้าพบวันที่ในข้อความ จะพิมพ์วันที่และส่งคืนค่า `DAY` ที่พบ
    

### หมายเหตุ

-   ควรตรวจสอบว่า `image` ที่ใช้ในฟังก์ชันนั้นถูกต้อง (ในกรณีนี้ใช้ `rotated_image` แทน `image` ภายในฟังก์ชัน) และค่า `cropped_image` ได้รับการปรับแต่งก่อนส่งเข้า OCR
-   การใช้ `pytesseract` สามารถปรับการตั้งค่า (เช่น `--psm 6` หรือ `--oem 3`) ขึ้นอยู่กับรูปแบบของเอกสารที่ต้องการดึงข้อมูล

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303585076151455804/image.png?ex=672c49b5&is=672af835&hm=f4054e8933d0e485a5b441501a97170ff52216f07a2f48d6993552988ad8665c&)
ฟังก์ชัน `process_DATE_Mtext` นี้มีขั้นตอนการทำงานเพื่อตรวจจับและดึงข้อความ (ในที่นี้คือวันที่) จากรูปภาพที่ถูกสแกน โดยมีการใช้ `pytesseract` ในการแปลงภาพเป็นข้อความและทำการปรับปรุงคุณภาพของภาพเพื่อให้ได้ผลลัพธ์ที่ดีขึ้น ก่อนที่ข้อความจะถูกดึงออกมา ขั้นตอนดังนี้:

1.  **ตัดรูปตามพิกัด**:

    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
    ฟังก์ชันนี้จะตัดรูปภาพจากพิกัดที่กำหนด โดยใช้ค่าพิกัด `(x1, y1)` ถึง `(x2, y2)` ซึ่งเป็นตำแหน่งของพื้นที่ที่ต้องการจากภาพที่หมุนแล้ว
    
2.  **แปลงภาพเป็นสีเทา**:
  
    `gray = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2GRAY)` 
    
    ใช้ `cv2.cvtColor` เพื่อแปลงภาพที่ตัดแล้วให้เป็นภาพสีเทา (grayscale) ซึ่งช่วยให้ OCR ทำงานได้ดีกว่าในภาพที่มีการแยกแยะรายละเอียด
    
3.  **กรองภาพเพื่อลด Noise ด้วย Bilateral Filtering**:

    `filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)` 
    
    ใช้ `cv2.bilateralFilter` เพื่อทำการกรองภาพและลดสัญญาณรบกวน (noise) ซึ่งจะช่วยทำให้ OCR สามารถอ่านตัวอักษรได้ง่ายขึ้น โดยรักษาความคมชัดของภาพ
    
4.  **ขยายขนาดภาพ**:

    `scale_percent = 170  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(filtered_image.shape[1] * scale_percent / 100)
    height = int(filtered_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(filtered_image, dim, interpolation=cv2.INTER_AREA)` 
    
    ขยายขนาดของภาพที่ผ่านการกรองแล้วให้มีขนาดใหญ่ขึ้น ซึ่งจะช่วยเพิ่มความละเอียดในการอ่านตัวอักษร
    
5.  **ปรับความคมชัดของภาพ**:

    `enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)` 
    
    ปรับคอนทราสต์และความสว่างของภาพให้ดีขึ้น ซึ่งจะทำให้ข้อความในภาพดูชัดเจนมากขึ้น
    
6.  **ใช้ `pytesseract` เพื่อดึงข้อความจากภาพ**:

    `custom_config = r'--oem 3 --psm 6'
    DAY = pytesseract.image_to_string(enhanced_image, lang='tha+eng', config=custom_config)` 
    
    ใช้ `pytesseract` เพื่อทำการ OCR จากภาพที่ได้รับการปรับปรุง โดยการตั้งค่า `--oem 3` และ `--psm 6` ช่วยให้สามารถใช้ OCR แบบเต็มรูปแบบและเลือกโหมดการประมวลผลที่เหมาะสมสำหรับการอ่านข้อความในภาพ
    
7.  **ส่งคืนผลลัพธ์**:

    `return DAY` 
    
    ผลลัพธ์ที่ได้จาก OCR จะถูกส่งคืนเป็นตัวแปร `DAY` ซึ่งเป็นข้อความที่ถูกดึงออกจากภาพ
    

### หมายเหตุ

-   ฟังก์ชันนี้สามารถใช้ได้ทั้งภาษาไทย (`'tha'`) และภาษาอังกฤษ (`'eng'`) โดยสามารถตั้งค่าภาษาในการทำ OCR ผ่าน `lang='tha+eng'`
-   อาจต้องปรับปรุงขั้นตอนการปรับแต่งภาพ (เช่น การปรับค่าความคมชัดหรือการกรองภาพ) ตามลักษณะของภาพจริงที่ใช้เพื่อให้ได้ผลลัพธ์ที่ดีที่สุด


![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303585712607723581/image.png?ex=672c4a4d&is=672af8cd&hm=aa36d6b46f9d4e31b04d75be55c991952a1aef968d801f7fa99d6800829f5b73&)
ฟังก์ชัน `process_NAME` นี้มีขั้นตอนการทำงานเพื่อดึงข้อมูลชื่อจากภาพที่ได้รับการตัดและประมวลผลแล้ว โดยใช้ OCR (Optical Character Recognition) ผ่าน `pytesseract` ขั้นตอนต่างๆ ที่เกิดขึ้นในฟังก์ชันมีดังนี้:

### ขั้นตอนการทำงาน:

1.  **ตัดภาพตามพิกัดที่กำหนด**:
 
    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
    ฟังก์ชันจะตัดภาพตามพิกัดที่ได้รับ (`x1, y1, x2, y2`) ซึ่งระบุพื้นที่ที่ต้องการดึงออกจากภาพที่หมุนแล้ว
    
2.  **ขยายขนาดภาพ**:
  
    `scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)` 
    
    ขยายขนาดภาพที่ตัดมาให้มีขนาดใหญ่ขึ้นเพื่อช่วยเพิ่มความชัดเจนในการอ่านข้อความ
    
3.  **ปรับสีให้เห็นชัดขึ้น**:
   
    `enhanced_image = cv2.convertScaleAbs(resized_image, alpha=1.495, beta=-70)` 
    
    ปรับความคมชัด (contrast) และความสว่าง (brightness) ของภาพ เพื่อให้ OCR สามารถตรวจจับข้อความได้ดีขึ้น
    
4.  **แสดงภาพที่ถูกตัด**:
   
    `plt.figure(figsize=(10, 5))
    plt.subplot(1, 2, 1)
    plt.imshow(cv2.cvtColor(cropped_image, cv2.COLOR_BGR2RGB))
    plt.title('Cropped Image')
    plt.show()` 
    
    ใช้ `matplotlib` ในการแสดงผลภาพที่ถูกตัดให้ดู
    
5.  **ใช้ `pytesseract` ในการดึงข้อความ**:
    
    `detected_text = pytesseract.image_to_string(enhanced_image, lang='tha+eng', config=custom_config)
    print(detected_text)` 
    
    ใช้ `pytesseract` เพื่อแปลงภาพที่ปรับปรุงแล้วเป็นข้อความ โดยใช้ทั้งภาษาไทย (`'tha'`) และภาษาอังกฤษ (`'eng'`) และการตั้งค่า `custom_config` เพื่อปรับการทำงานของ OCR
    
6.  **ตรวจสอบคำนำหน้า (Prefix)**:
 
    `valid_prefixes = ["นาย", "นางสาว", "นาง", "คุณ", "Mr", "Mrs", "Ms", "Miss", "บริษัท","น.ส.","น.,ส,."]
    found_prefixes = []
    for prefix in valid_prefixes:
        if prefix in detected_text:
            found_prefixes.append(prefix)` 
    
    ฟังก์ชันจะตรวจสอบว่าในข้อความที่ตรวจจับมา (จาก OCR) มีคำนำหน้าใดบ้างที่อยู่ในรายการ `valid_prefixes` ที่กำหนดไว้ เช่น "นาย", "นางสาว", "บริษัท" เป็นต้น
    
7.  **จัดการกรณีพิเศษสำหรับ "บริษัท"**:
 
    `if prefix == "บริษัท":
        index = detected_text.find(prefix) + len(prefix)
        remaining_text = detected_text[index:].strip()
        remaining_text = remaining_text.split()[0]
        NAME = "{} {}".format(prefix, remaining_text)
        print("Name :", NAME)
        return NAME` 
    
    ถ้า OCR พบคำว่า "บริษัท" ฟังก์ชันจะตัดข้อความที่ตามหลังออกมาและใช้เฉพาะชื่อบริษัท
    
8.  **สร้างชื่อจากคำนำหน้าที่พบ**:
 
    `if found_prefixes:
        NAME = "{} {}".format(found_prefixes[0], remaining_text)
        print("Name :", NAME)
    else:
        NAME = None` 
    
    ถ้ามีคำนำหน้าที่พบ เช่น "นาย", "คุณ", หรือ "Mr" จะนำมาผสมกับชื่อที่เหลือ (หลังจากคำนำหน้า) เพื่อสร้างชื่อเต็ม
    
9.  **ส่งคืนชื่อที่ได้**:

    `return NAME` 
    
    ฟังก์ชันจะส่งคืนชื่อที่ถูกสร้างขึ้น หากไม่มีชื่อหรือไม่พบคำนำหน้า จะส่งคืนค่า `None`
    

### หมายเหตุ:

-   ฟังก์ชันนี้รองรับการตรวจจับทั้งภาษาไทยและภาษาอังกฤษ แต่จะต้องปรับคำที่พบได้ให้เหมาะสมกับกรณีที่ใช้
-   หากพบคำนำหน้า "บริษัท" ฟังก์ชันจะจัดการข้อความพิเศษนั้นโดยตัดชื่อบริษัทออกจากข้อความ


![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303586172219686972/image.png?ex=672c4aba&is=672af93a&hm=0e077822028bb382daf1c4af07632b4fa2c689ca9e53baab014cd2dd2d04069d&)
ฟังก์ชัน `process_Policy` นี้จะทำการประมวลผลภาพที่ตัดมาเพื่อดึงข้อมูลเกี่ยวกับหมายเลขกรมธรรม์ (Policy Number) จากข้อความในภาพที่ได้รับการตัดและประมวลผลแล้วด้วย OCR (Optical Character Recognition) ผ่าน `pytesseract` ขั้นตอนที่ฟังก์ชันทำงานมีดังนี้:

### ขั้นตอนการทำงาน:

1.  **ตัดภาพตามพิกัดที่กำหนด**:
    
    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
    ฟังก์ชันจะตัดภาพตามพิกัดที่ได้รับ (`x1, y1, x2, y2`) ซึ่งเป็นพื้นที่ที่ต้องการดึงออกจากภาพที่หมุนแล้ว
    
2.  **ตั้งค่าคอนฟิกสำหรับ `pytesseract`**:
      
    `custom_config = r'--oem 3 --psm 6'` 
    
    กำหนดการตั้งค่าการประมวลผล OCR ผ่าน `pytesseract` โดยใช้ `--oem 3` เพื่อเลือก OCR Engine Mode ที่ดีที่สุด และ `--psm 6` ซึ่งเป็นโหมดการจัดรูปแบบบล็อกข้อความ
    
3.  **ใช้ `pytesseract` ในการแปลงภาพเป็นข้อความ**:

    `text = pytesseract.image_to_string(cropped_image, lang='eng', config=custom_config)` 
    
    ฟังก์ชันนี้จะใช้ `pytesseract` เพื่อแปลงภาพที่ตัดมาเป็นข้อความที่สามารถอ่านได้ โดยกำหนดให้ OCR ตรวจจับภาษาอังกฤษ (`'eng'`)
    
4.  **การแยกข้อมูลหมายเลขกรมธรรม์**:
  
    `num_policy = ''.join(text)` 
    
    เนื่องจากการใช้ `pytesseract` อาจจะทำให้ข้อความที่ได้มีลักษณะเป็นชุดของตัวอักษรหรือเลขที่ไม่ต้องการ ฟังก์ชันนี้จะทำการเชื่อมต่อ (concatenate) ข้อความทั้งหมดที่ได้รับจาก OCR เพื่อให้ได้หมายเลขกรมธรรม์
    
5.  **ส่งคืนผลลัพธ์**:
 
    `return num_policy` 
    
    ฟังก์ชันจะส่งคืนหมายเลขกรมธรรม์ (หรือข้อความทั้งหมดที่ได้จาก OCR)
    

### หมายเหตุ:

-   หากต้องการดึงหมายเลขกรมธรรม์ที่มีรูปแบบเฉพาะ (เช่น `A.B.1234-5678`) สามารถใช้ `re.findall()` และ `regex` เพื่อแยกเฉพาะหมายเลขกรมธรรม์ออกมาได้จากผลลัพธ์ที่ได้จาก OCR.
-   การใช้ `''.join(text)` อาจทำให้ข้อมูลทั้งหมดที่ OCR ตรวจจับได้ถูกเชื่อมต่อกันเป็นข้อความยาวๆ โดยไม่มีการกรองหรือตรวจสอบรูปแบบของข้อมูล



![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303586614118711317/image.png?ex=672c4b24&is=672af9a4&hm=3474f515562828598e556306c284f80f1fd7b2fe0ce35b68d36e8441e95e2848&)

ฟังก์ชัน `check_policy_BKI` นี้จะทำการตรวจสอบและทำความสะอาดหมายเลขกรมธรรม์ โดยมีการปรับแต่งหมายเลขให้ถูกต้องตามรูปแบบที่ต้องการ หากพบข้อมูลที่ไม่ครบหรือผิดปกติ ฟังก์ชันจะแสดงข้อความเตือนและคืนหมายเลขกรมธรรม์ที่จัดรูปแบบใหม่

### การทำงานของฟังก์ชัน:

1.  **การทำความสะอาดหมายเลขกรมธรรม์**:
    
    -   ฟังก์ชันจะทำการตัดทอนตัวเลขที่ลงท้ายด้วย `10` หรือ `1` ออกจากหมายเลขในแต่ละส่วน (ที่แยกโดยเครื่องหมาย `-`):
   
    `if num.endswith('10') or num.endswith('1'):
        num = num[:-2]` 
    
    ตัวอย่างเช่น ถ้าหมายเลขคือ `12310`, จะตัด `10` ออกเหลือ `123`
    
2.  **การจัดการหมายเลขส่วนแรก**:
    
    -   ฟังก์ชันจะดึงเฉพาะ 3 ตัวสุดท้ายจากหมายเลขส่วนแรก:
 
    `first_part = cleaned_numbers[0][-3:]  # Take last 3 digits` 
    
    ถ้าไม่มีข้อมูลในส่วนนี้ จะใช้ค่า `XXX` แทน:
   
    `if len(cleaned_numbers) > 0:
        first_part = cleaned_numbers[0][-3:]
    else:
        first_part = 'XXX'` 
    
3.  **การจัดการหมายเลขส่วนที่สอง**:
    
    -   ฟังก์ชันจะตรวจสอบว่าในส่วนที่สอง (ซึ่งต้องมี 5 ตัว) มีข้อมูลหรือไม่ ถ้ามีจะเลือก 5 ตัวแรก ถ้าไม่มีจะใช้ค่า `XXXXX`:
 
    `second_part = cleaned_numbers[1][:5] if len(cleaned_numbers) > 1 else 'XXXXX'` 
    
4.  **การจัดการหมายเลขส่วนที่สาม**:
    
    -   ฟังก์ชันจะตรวจสอบว่าในส่วนที่สาม (ต้องมี 6 ตัว) มีข้อมูลหรือไม่ ถ้ามีจะเลือก 6 ตัวแรก ถ้าไม่มีจะใช้ค่า `XXXXXX`:
   
    `third_part = cleaned_numbers[2][:6] if len(cleaned_numbers) > 2 else 'XXXXXX'` 
    
5.  **การตรวจสอบข้อมูลที่หายไป**:
    
    -   หากหมายเลขส่วนใดส่วนหนึ่งไม่มีข้อมูล จะมีการแทนที่ด้วยข้อความ `(ข้อมูลผิดพลาด)`:
   
    `parts = [
        first_part if first_part else '(ข้อมูลผิดพลาด)',
        second_part if second_part else '(ข้อมูลผิดพลาด)',
        third_part if third_part else '(ข้อมูลผิดพลาด)'
    ]` 
    
6.  **การจัดรูปแบบหมายเลขกรมธรรม์**:
    
    -   หมายเลขกรมธรรม์จะถูกรวมกันเป็นรูปแบบ `XXX-XXXX-XXXXX` โดยใช้เครื่องหมาย `-` คั่นระหว่างแต่ละส่วน:

    `Policy = "-".join(parts)` 
    
7.  **การพิมพ์และส่งคืนผลลัพธ์**:
    
    -   ฟังก์ชันจะแสดงหมายเลขกรมธรรม์ที่ได้ และส่งคืนหมายเลขในรูปแบบที่จัดเรียงใหม่:
    
    `print("Policy number: {}".format(Policy))
    return Policy` 
    

### ตัวอย่างการทำงาน:

#### กรณีที่ข้อมูลครบถ้วน:

-   ถ้า `num_policy = "12345-67890-123456"` ผลลัพธ์จะเป็น:
 
    `Policy number: 345-67890-123456` 
    

#### กรณีที่ข้อมูลบางส่วนหายไป:

-   ถ้า `num_policy = "12345-67890"` ผลลัพธ์จะเป็น:
    
    `Policy number: 345-67890-XXXXXX` 
    

#### กรณีที่ข้อมูลผิดรูปแบบ:

-   ถ้า `num_policy = "123-456-"` ผลลัพธ์จะเป็น:
   
    `Policy number: 123-456-XXXXXX` 
    

ฟังก์ชันนี้ช่วยให้หมายเลขกรมธรรม์ถูกจัดรูปแบบให้ตรงกับข้อกำหนดที่คาดหวังและสามารถตรวจจับข้อมูลที่ขาดหายไปได้


![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303587544813932564/image.png?ex=672c4c02&is=672afa82&hm=f4a90025f324f9f0126951d63eccf83f4acdf9849002beac510ad90b39ffc63c&)

ฟังก์ชัน `check_policy_MIT_MT` นี้ใช้เพื่อให้การตรวจสอบและจัดรูปแบบหมายเลขกรมธรรม์ (policy number) ที่มีรูปแบบเฉพาะ โดยมีการตรวจสอบดังนี้:

### การทำงานของฟังก์ชัน:

1.  **การลบช่องว่างและอักขระพิเศษ**: ฟังก์ชันเริ่มต้นโดยการลบช่องว่าง (`" "`) และอักขระ `_` ออกจากหมายเลขกรมธรรม์:

    `num_policy = num_policy.replace(" ", "").replace("_", "")` 
    
    สิ่งนี้ช่วยให้มั่นใจได้ว่าไม่มีอักขระที่ไม่จำเป็นอยู่ในหมายเลขกรมธรรม์ก่อนที่จะดำเนินการตรวจสอบต่อไป
    
2.  **การแบ่งหมายเลขกรมธรรม์**: หมายเลขกรมธรรม์จะถูกแบ่งโดยจุด (`"."`):
 
    `parts = num_policy.split('.')` 
    
3.  **การตรวจสอบว่าเริ่มต้นด้วย "TV"**: ฟังก์ชันจะตรวจสอบว่า `parts[0]` คือ `"TV"` หรือไม่ ถ้าไม่ใช่ จะส่งคืนข้อความแสดงข้อผิดพลาด:

    `if len(parts) < 2 or parts[0] != "TV":
        return f"{num_policy} ไม่มี 'TV' ที่คาดหวัง"` 
    
4.  **การตรวจสอบว่า "VMI" ตามหลัง "TV"**: ถ้า `parts[1]` ไม่ใช่ `"VMI"` ฟังก์ชันจะส่งคืนข้อความที่แจ้งว่าไม่มี `"VMI"` หลังจาก `"TV"`:
 
    `if len(parts) < 3 or parts[1] != "VMI":
        return f"{num_policy} ไม่มี 'VMI' ที่คาดหวังหลัง 'TV'"` 
    
5.  **การค้นหาตัวเลขที่มีความยาวไม่เกิน 7 ตัว**: ฟังก์ชันใช้ `re.search` เพื่อตรวจหาตัวเลขที่มีความยาวไม่เกิน 7 ตัวในส่วนที่สาม (หลัง `"VMI"`):
 
    `digits = re.search(r'\d{1,7}', parts[2])
    if not digits:
        return f"{num_policy} ไม่มีตัวเลข 7 ตัวหลัง 'VMI'"` 
    
6.  **การจัดรูปแบบหมายเลขกรมธรรม์**: ถ้าทุกส่วนถูกต้อง ฟังก์ชันจะสร้างหมายเลขกรมธรรม์ในรูปแบบ `TV.VMI.XXXXXXX` (ซึ่ง `XXXXXXX` คือหมายเลขที่พบ):

    `Policy = f"TV.VMI.{digits.group(0)}"` 
    
7.  **การส่งคืนหมายเลขกรมธรรม์ที่จัดรูปแบบ**: ฟังก์ชันจะส่งคืนหมายเลขกรมธรรม์ที่จัดรูปแบบเรียบร้อยแล้ว:
 
    `return Policy` 
    

### ตัวอย่างการทำงาน:

#### กรณีที่ข้อมูลถูกต้อง:

-   หาก `num_policy = "TV.VMI.1234567"`, ผลลัพธ์จะเป็น:

    `Policy = "TV.VMI.1234567"` 
    

#### กรณีที่ไม่มี `"TV"`:

-   หาก `num_policy = "VMI.1234567"`, ผลลัพธ์จะเป็น:

    `"VMI.1234567 ไม่มี 'TV' ที่คาดหวัง"` 
    

#### กรณีที่ไม่มี `"VMI"` หลังจาก `"TV"`:

-   หาก `num_policy = "TV.VM.1234567"`, ผลลัพธ์จะเป็น:

    `"TV.VM.1234567 ไม่มี 'VMI' ที่คาดหวังหลัง 'TV'"` 
    

#### กรณีที่ไม่มีตัวเลขหลัง `"VMI"`:

-   หาก `num_policy = "TV.VMI.abcdef"`, ผลลัพธ์จะเป็น:

    `"TV.VMI.abcdef ไม่มีตัวเลข 7 ตัวหลัง 'VMI'"` 
    

ฟังก์ชันนี้ช่วยให้ตรวจสอบหมายเลขกรมธรรม์ที่มีรูปแบบเฉพาะ และถ้าหากข้อมูลไม่ตรงตามที่คาดหวัง ก็จะมีข้อความแสดงข้อผิดพลาด


![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303588303513325618/image.png?ex=672c4cb7&is=672afb37&hm=e7ed33ec5d151b3bb110b1422c158bcd6bb57dd0014eeedf2e78620f2f334233&)

ฟังก์ชัน `process_pream` นี้ใช้สำหรับการประมวลผลภาพเพื่อตรวจจับและแปลงข้อความ (OCR) จากภาพที่มีข้อมูลเกี่ยวกับ "ค่าธรรมเนียมประกัน" หรือข้อมูลอื่นๆ ที่ต้องการ โดยมีขั้นตอนการทำงานดังนี้:

### ขั้นตอนการทำงานของฟังก์ชัน:

1.  **ตัดภาพตามพิกัดที่กำหนด**: ฟังก์ชันเริ่มต้นด้วยการตัดภาพจาก `rotated_image` โดยใช้พิกัด `x1, y1, x2, y2` ที่ระบุไว้ เพื่อได้ส่วนที่ต้องการของภาพ:

    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
2.  **ขยายขนาดภาพ**: ขนาดของภาพที่ตัดออกมาจะถูกขยายโดยใช้สเกล 200%:

    `scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)` 
    
3.  **แปลงภาพเป็นภาพสีเทา**: ภาพที่ขยายแล้วจะถูกแปลงเป็นภาพสีเทา (grayscale) เพื่อลดความซับซ้อนในการประมวลผล:
    
    python
    
    คัดลอกโค้ด
    
    `gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)` 
    
4.  **ทำการกรองภาพด้วย Bilateral Filtering**: การกรองด้วย Bilateral Filter ช่วยลด noise ในภาพและเพิ่มความคมชัดของภาพ:

    `filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)` 
    
5.  **ปรับความคมชัดของภาพ**: ใช้ `convertScaleAbs` เพื่อปรับความคมชัดของภาพให้เห็นชัดขึ้น:
    
    `enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.5, beta=-70)` 
    
6.  **การแปลงภาพเป็นข้อความ (OCR)**: ใช้ `pytesseract` เพื่อแปลงภาพเป็นข้อความ โดยใช้การตั้งค่าพิเศษ `custom_config`:
   
    `custom_config = r'--oem 3 --psm 6'
    insurance_fee = pytesseract.image_to_string(enhanced_image, lang='eng', config=custom_config)` 
    
7.  **แสดงผลลัพธ์**: ฟังก์ชันจะพิมพ์ข้อความที่ได้จาก OCR และส่งคืนค่าผลลัพธ์:
     
    `print("insurance_fee:", format(insurance_fee))
    return insurance_fee` 
    
### หมายเหตุ:
-   หากข้อมูลในภาพมีลักษณะที่ไม่ชัดเจนหรือมีการเบลอมาก ฟังก์ชันนี้อาจต้องการการปรับแต่งเพิ่มเติม เช่น การปรับค่า `alpha`, `beta` หรือ `sigma` เพื่อให้ได้ผลลัพธ์ที่ดีที่สุดจากการแปลงข้อความ (OCR).

![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303588987306250240/image.png?ex=672c4d5a&is=672afbda&hm=23f7abea3f42f2cb0c80c640db2379fe2fb7bb9af3cb811d48406f8c079cbf5b&)
ฟังก์ชัน `process_(Balance,VAT,total)` นี้ใช้สำหรับการประมวลผลภาพเพื่อตรวจจับและแปลงข้อความ (OCR) จากภาพที่มีข้อมูลเกี่ยวกับ "ค่าธรรมเนียมประกัน" หรือข้อมูลอื่นๆ ที่ต้องการ โดยมีขั้นตอนการทำงานดังนี้:

### ขั้นตอนการทำงานของฟังก์ชัน:

1.  **ตัดภาพตามพิกัดที่กำหนด**: ฟังก์ชันเริ่มต้นด้วยการตัดภาพจาก `rotated_image` โดยใช้พิกัด `x1, y1, x2, y2` ที่ระบุไว้ เพื่อได้ส่วนที่ต้องการของภาพ:

    `cropped_image = rotated_image[y1:y2, x1:x2]` 
    
2.  **ขยายขนาดภาพ**: ขนาดของภาพที่ตัดออกมาจะถูกขยายโดยใช้สเกล 200%:

    `scale_percent = 200  # เปอร์เซ็นต์ของการขยายภาพ
    width = int(cropped_image.shape[1] * scale_percent / 100)
    height = int(cropped_image.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized_image = cv2.resize(cropped_image, dim, interpolation=cv2.INTER_AREA)` 
    
3.  **แปลงภาพเป็นภาพสีเทา**: ภาพที่ขยายแล้วจะถูกแปลงเป็นภาพสีเทา (grayscale) เพื่อลดความซับซ้อนในการประมวลผล:
    
    python
    
    คัดลอกโค้ด
    
    `gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)` 
    
4.  **ทำการกรองภาพด้วย Bilateral Filtering**: การกรองด้วย Bilateral Filter ช่วยลด noise ในภาพและเพิ่มความคมชัดของภาพ:

    `filtered_image = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)` 
    
5.  **ปรับความคมชัดของภาพ**: ใช้ `convertScaleAbs` เพื่อปรับความคมชัดของภาพให้เห็นชัดขึ้น:
    
    `enhanced_image = cv2.convertScaleAbs(filtered_image, alpha=1.5, beta=-70)` 
    
6.  **การแปลงภาพเป็นข้อความ (OCR)**: ใช้ `pytesseract` เพื่อแปลงภาพเป็นข้อความ โดยใช้การตั้งค่าพิเศษ `custom_config`:
   
    `custom_config = r'--oem 3 --psm 6'
    insurance_fee = pytesseract.image_to_string(enhanced_image, lang='eng', config=custom_config)` 
    
7.  **แสดงผลลัพธ์**: ฟังก์ชันจะพิมพ์ข้อความที่ได้จาก OCR และส่งคืนค่าผลลัพธ์:
     
    `print("insurance_(Balance,VAT,total):", format(insurance_(Balance,VAT,total))
    return insurance_(Balance,VAT,total)` 
    
### หมายเหตุ:
-   หากข้อมูลในภาพมีลักษณะที่ไม่ชัดเจนหรือมีการเบลอมาก ฟังก์ชันนี้อาจต้องการการปรับแต่งเพิ่มเติม เช่น การปรับค่า `alpha`, `beta` หรือ `sigma` เพื่อให้ได้ผลลัพธ์ที่ดีที่สุดจากการแปลงข้อความ (OCR).
- code ของ ฟังก์ชัน `process_(Balance,VAT,total)` และ ฟังก์ชัน `process_pream`
คือมีการทำงานที่เหมือนกัน


![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303590099531071528/image.png?ex=672c4e63&is=672afce3&hm=6234ccf820e7b2c2481900d4045453cec77e7e491c502ebdd516033723687373&)

ฟังก์ชัน `adjust_ocr_value` นี้ถูกออกแบบมาเพื่อปรับค่าผลลัพธ์จากการแปลงภาพเป็นข้อความ (OCR) ให้เป็นตัวเลขที่ถูกต้อง โดยทำการกรองและปรับค่าผลลัพธ์ในหลายขั้นตอน ดังนี้:

### ขั้นตอนการทำงานของฟังก์ชัน:

1.  **กรองตัวอักษรที่ไม่ใช่ตัวเลขหรือเครื่องหมายที่เกี่ยวข้อง**: ฟังก์ชันจะกรองเฉพาะตัวเลข, จุลภาค (`,`), จุดทศนิยม (`.`), และเครื่องหมายลบ (`-`) ออกจากข้อความ:

    `valid_chars = '0123456789.,-'  # กำหนดอักขระที่เป็นตัวเลข, จุดทศนิยม, จุลภาค, และเครื่องหมายลบ
    filtered_result = ''.join(char for char in ocr_result if char in valid_chars)` 
    
2.  **แทนที่เครื่องหมายจุลภาค (`,`) เป็นจุดทศนิยม (.)**: ฟังก์ชันจะทำการแทนที่เครื่องหมายจุลภาค (`,`) ที่พบในผลลัพธ์ให้เป็นจุดทศนิยม (`.`) เพื่อความถูกต้องในการแสดงผลของตัวเลข:

    `filtered_result = filtered_result.replace(',', '.')` 
    
3.  **การจัดการกับจุดทศนิยมที่เกินมา**: หากในข้อความมีจุดทศนิยมมากกว่า 1 จุด ฟังก์ชันจะรักษาจุดทศนิยมแรกและนำค่าทั้งหมดหลังจากจุดทศนิยมสุดท้ายมารวมกัน:
   
    `decimal_count = filtered_result.count('.')
    if decimal_count > 1:
        last_decimal_index = filtered_result.rfind('.')
        decimal_part = filtered_result[last_decimal_index + 1:]
        integer_part = filtered_result[:last_decimal_index].replace('.', '', decimal_count - 1)
        filtered_result = f"{integer_part}.{decimal_part}"` 
    
4.  **การลบเครื่องหมายลบที่ติดกับจุดทศนิยม**: หากมีเครื่องหมายลบ (`-`) อยู่ในค่าที่มีจุดทศนิยม ฟังก์ชันจะลบเครื่องหมายลบที่ไม่เหมาะสมออก:

    `if '.' in filtered_result and '-' in filtered_result:
        minus_index = filtered_result.find('-')
        filtered_result = filtered_result[:minus_index] + filtered_result[minus_index+1:]` 
    
5.  **ส่งคืนผลลัพธ์**: ฟังก์ชันจะคืนค่าผลลัพธ์ที่ถูกปรับแต่งแล้วโดยลบช่องว่างที่ไม่จำเป็นออก:

    `return filtered_result.strip()` 
    

### ตัวอย่างการทำงาน:

สมมุติว่า `ocr_result` คือ `"1,234.56"`, หลังจากการประมวลผลฟังก์ชันจะให้ผลลัพธ์เป็น `"1234.56"`.

หาก `ocr_result` คือ `"123.45.67"`, หลังจากการประมวลผลฟังก์ชันจะให้ผลลัพธ์เป็น `"12345.67"`.

หาก `ocr_result` คือ `"123.45-67"`, หลังจากการประมวลผลฟังก์ชันจะให้ผลลัพธ์เป็น `"123.4567"`.

### หมายเหตุ:

-   ฟังก์ชันนี้ช่วยในการทำความสะอาดผลลัพธ์ OCR ที่อาจมีความผิดพลาด เช่น เครื่องหมายจุลภาคที่ไม่ตรงจุดทศนิยม หรือการปรากฏของเครื่องหมายลบในตำแหน่งที่ไม่เหมาะสม.
-   หากผลลัพธ์ OCR มีรูปแบบที่ซับซ้อนกว่านี้ เช่น การมีตัวอักษรหรือสัญลักษณ์เพิ่มเติมที่ไม่ใช่ตัวเลข อาจต้องปรับฟังก์ชันเพิ่มเติมให้รองรับกรณีเหล่านั้น.



![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303591641289854996/image.png?ex=672c4fd2&is=672afe52&hm=f6cf7372064061a651b81408b28e463ca92a9a5e99bef68996c0ef9677854835&)
![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303591289186553896/image.png?ex=672c4f7e&is=672afdfe&hm=d07992462cd84a1b476c9a32ef9518de2ef625076b4cfb6d571bb067628a27ac&)
![enter image description here](https://cdn.discordapp.com/attachments/1263395471833960476/1303591440361979944/image.png?ex=672c4fa2&is=672afe22&hm=c9e375c1a4649f062e02b52a27681f2845b82b639c11082ca00119e2f32af8e2&)

นี้เป็นส่วนสคริปต์ที่ใช้ในการประมวลผลและดึงข้อมูลจากไฟล์ภาพ โดยมีขั้นตอน ดังนี้:

### ขั้นตอนหลักและคำอธิบาย:

1.  **การตั้งค่าเบื้องต้น**:
    
    -   ฟังก์ชัน `main()` จะอ่านพารามิเตอร์ที่รับจากบรรทัดคำสั่ง (command-line arguments) ที่ส่งเข้ามาในสคริปต์
    -   ฟังก์ชัน `process_data()` จะประมวลผลพารามิเตอร์เหล่านั้นและส่งคืนค่าต่างๆ เช่น พิกัด (coordinates) สำหรับการประมวลผลภาพ
    
    `text_choice = main(sys.argv[1:])
    processed_data = process_data(text_choice)` 
    
2.  **การประมวลผลไฟล์ภาพ**:
    
    -   หลังจากดึงข้อมูลแล้ว (เช่น พิกัดของพื้นที่ที่ต้องการจากภาพ) จะมีการระบุโฟลเดอร์ที่เก็บภาพ (อาจจะเป็นไฟล์ PDF หรือไฟล์ภาพ) และทำการประมวลผล
    -   สคริปต์จะวนลูปผ่านไฟล์ในโฟลเดอร์นั้น โดยใช้ `cv2.imread()` ในการโหลดภาพ

    `folder_path = f'{folder_PDF_path}'
    image_files = os.listdir(folder_path)` 
    
3.  **การจัดการการหมุนภาพ**:
    
    -   ถ้าค่าของ `rotate` เป็น `"Y"` หรือ `"y"` จะมีการหมุนภาพโดยใช้ฟังก์ชัน `process_image()` หากไม่ใช่ค่าดังกล่าว ภาพจะไม่ถูกหมุน
    
    `if rotate == "Y" or rotate == "y":
        rotated_image = process_image(image)
    else:
        rotated_image = image` 
    
4.  **การดึงข้อมูลจากภาพ**:
    
    -   ข้อมูลต่างๆ เช่น `TYPE`, `DATE`, `NAME`, `Policy`, `insurance_fee`, ฯลฯ ถูกดึงออกมาจากภาพโดยใช้ฟังก์ชันที่เกี่ยวข้อง (เช่น `process_TYPE`, `process_DATE_num`, ฯลฯ)
    -   ฟังก์ชัน `adjust_ocr_value()` ถูกใช้ในการทำความสะอาดและจัดรูปแบบผลลัพธ์จาก OCR (เช่น การเปลี่ยนเครื่องหมายจุลภาคเป็นจุดทศนิยม การจัดการจุดทศนิยมที่มากเกินไป ฯลฯ)
   
    `TYPE = process_TYPE(rotated_image, filename, x1_TYPE, y1_TYPE, x2_TYPE, y2_TYPE)
    DATE = process_DATE_num(filename, rotated_image, x1_DATE, y1_DATE, x2_DATE, y2_DATE)
    NAME = process_NAME(filename, rotated_image, x1_NAME, y1_NAME, x2_NAME, y2_NAME)
    Policy = process_Policy(filename, rotated_image, x1_Policy, y1_Policy, x2_Policy, y2_Policy)` 
    
5.  **การปรับค่า OCR**:
    
    -   ค่าที่ถูกดึงมา เช่น `insurance_fee`, `stamp_duty`, และ `Balanc_num` จะถูกทำความสะอาดด้วยฟังก์ชัน `adjust_ocr_value()` เพื่อให้มั่นใจว่าผลลัพธ์จาก OCR ถูกต้อง เช่น การเปลี่ยนจุลภาคเป็นจุดทศนิยม การจัดการกับจุดทศนิยมที่ไม่ถูกต้อง ฯลฯ
   
    `pream = adjust_ocr_value(insurance_fee)
    stamp = adjust_ocr_value(stamp_duty)
    Balanc = adjust_ocr_value(Balanc_num)
    VAT = adjust_ocr_value(VAT_num)
    total = adjust_ocr_value(total_num)` 
    
6.  **การบันทึกข้อมูลลงใน Excel**:
    
    -   สร้างพจนานุกรม (dictionary) เพื่อเก็บข้อมูลที่ดึงมา
    -   พจนานุกรมนี้จะถูกแปลงเป็น Pandas DataFrame และทำการบันทึกลงในไฟล์ Excel
    
    `data = {
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
    df.T.to_excel(f'{folder_OUTPUT_path}\\{filename}.xlsx', header=False)` 
    
7.  **ผลลัพธ์สุดท้าย**:
    
    -   หลังจากการประมวลผลภาพแต่ละไฟล์ สคริปต์จะแสดงข้อความยืนยันว่าไฟล์ Excel ถูกบันทึกสำเร็จ
    -   เมื่อประมวลผลไฟล์ทั้งหมดเสร็จสิ้น จะมีการพิมพ์ข้อความบอกว่าเสร็จสิ้นการทำงานแล้ว
    
    `print("Excel file saved successfully.")
    print("_____________________________________________________")` 
    

### หมายเหตุ:

1.  **การจัดการข้อผิดพลาด**:
    
    -  อาจต้องการเพิ่มการจัดการข้อผิดพลาด (เช่น `try-except`) รอบๆ การทำงานที่เกี่ยวข้องกับการอ่านไฟล์และการประมวลผล OCR เพื่อป้องกันไม่ให้สคริปต์หยุดทำงานเมื่อเกิดข้อผิดพลาดกับไฟล์หรือกระบวนการ OCR
2.  **การปรับปรุงการจัดการเส้นทางไฟล์**:
    
    -   `folder_OUTPUT_path` ใช้เครื่องหมาย `\` เป็นตัวแบ่งโฟลเดอร์ ซึ่งอาจทำให้เกิดปัญหาบนระบบที่ไม่ใช่ Windows อาจใช้ `os.path.join()` เพื่อให้แน่ใจว่าโค้ดสามารถทำงานได้บนทุกแพลตฟอร์ม
    
    `output_file_path = os.path.join(folder_OUTPUT_path, f"{filename}.xlsx")` 
    
3.  **การอ่านไฟล์ภาพอย่างมีประสิทธิภาพ**:
    
    -   ถ้าภาพมีขนาดใหญ่ อาจจะใช้วิธีการลดขนาดภาพหรือแปลงเป็นภาพขาวดำก่อนการประมวลผล เพื่อประหยัดหน่วยความจำและเพิ่มความเร็วในการทำงาน
4.  **การบันทึกข้อมูล (Logging)**:
    
    -   แทนที่จะแสดงข้อความโดยตรงในคอนโซล อาจพิจารณาใช้การบันทึกข้อมูล (logging) ซึ่งจะช่วยให้คุณสามารถควบคุมการแสดงผลและติดตามการทำงานของสคริปต์ได้ง่ายขึ้นในกรณีที่ทำงานในระบบที่ซับซ้อน

### สรุป:

โค้ดนี้ออกแบบมาเพื่อดึงข้อมูลจากภาพ ทำความสะอาดผลลัพธ์จาก OCR และบันทึกข้อมูลที่ได้ลงในไฟล์ Excel 

