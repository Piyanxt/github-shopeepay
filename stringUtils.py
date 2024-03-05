import openpyxl
from openpyxl.styles import Font, Alignment
from datetime import datetime, timedelta
import base64

def replace_All(str):
	str = str.replace(" ", "")
	str = str.replace('\n', ' ').replace('\r', '')
	str = str.replace(" ", "")
	str = str.replace('"', '')
	str = str.replace("[", "")
	str = str.replace("]", "")
	return str

# def compare_String(str,str2):
# 	str = str.replace(" ", "")
# 	str = str.replace('\n', ' ').replace('\r', '')
# 	str = str.replace(" ", "")
# 	str = str.replace('"', '')
# 	str = str.replace("[", "")
# 	str = str.replace("]", "")
# 	str2 = str2.replace(" ", "")
# 	str2 = str2.replace('\n', ' ').replace('\r', '')
# 	str2 = str2.replace(" ", "")
# 	str2 = str2.replace('"', '')
# 	str2 = str2.replace("[", "")
# 	str2 = str2.replace("]", "")
# 	if str == str2:	
# 		return "True"
# 	else:
# 		return "False"

def compareString(str,str2):
	str = str.replace('\n', ' ').replace('\r', '')
	str = str.replace(" ", "")
	str2 = str2.replace('\n', ' ').replace('\r', '')
	str2 = str2.replace(" ", "")
	print("Value of str: ", str)
	print("Value of str2: ", str2)
	if str == str2:	
		return "True"
	else:
		return "False"

def editFile(fld,sheet,file,value,font_color=None,alignment_value=None):  
	xfile = openpyxl.load_workbook(file)
	sheets = xfile[sheet]
	sheets[fld] = value
	font = None
	alignment = None
	if   font_color == "Text-Green":
		font = Font(name="Tahoma", size="10", color="1E8449", bold=None)  # สีเขียว
	elif font_color == "Text-Red":
		font = Font(name="Tahoma", size="10", color="FF0000", bold=None)  # สีแดง
	else :
		font = Font(name="Tahoma", size="10", color="000000", bold=None)  # สีดำ
	if alignment_value is not None:
		if alignment_value == "center":
			alignment = Alignment(horizontal="center")
	else:
		alignment = Alignment(horizontal="left")
	cell = sheets[fld]
	cell.value = value
	if font:
		cell.font = font
	if alignment:
		cell.alignment = alignment
	xfile.save(file)


def thai_date_to_iso(input_date):
	month_map = {
        "มกราคม": "01", "กุมภาพันธ์": "02", "มีนาคม": "03",
        "เมษายน": "04", "พฤษภาคม": "05", "มิถุนายน": "06",
        "กรกฎาคม": "07", "สิงหาคม": "08", "กันยายน": "09",
        "ตุลาคม": "10", "พฤศจิกายน": "11", "ธันวาคม": "12"
    }
	date_parts = input_date.split()
	day = date_parts[0]
	month = month_map[date_parts[1]]
	year = str(int(date_parts[2]) - 543)
	formatted_date = f"{year}-{month}-{day}"
	return formatted_date

def add_days_to_thai_date(input_date, days_to_add):
    formatted_date = thai_date_to_iso(input_date)
    # แปลงวันที่ในรูปแบบ ISO เป็นวัตถุ datetime
    input_date_obj = datetime.strptime(formatted_date, "%Y-%m-%d")
	# แปลงจำนวนวันที่ส่งมาบวกเป็นจำนวนเต็ม
    days_to_add = int(days_to_add)
    output_date_obj = input_date_obj + timedelta(days=days_to_add)
    output_date = output_date_obj.strftime("%d %B %Y")
    year = output_date.split()[-1]
    year_with_era = str(int(year) + 543)
    output_date = output_date.replace(year, year_with_era)
    return output_date

def replace_Month(str):
	str = str.replace("January", "มกราคม")
	str = str.replace("February", "กุมภาพันธ์")
	str = str.replace("March", "มีนาคม")
	str = str.replace("April", "เมษายน")
	str = str.replace("May", "พฤษภาคม")
	str = str.replace("June", "มิถุนายน")
	str = str.replace("July", "กรกฎาคม")
	str = str.replace("August", "สิงหาคม")
	str = str.replace("September", "กันยายน")
	str = str.replace("October", "ตุลาคม")
	str = str.replace("November", "พฤศจิกายน")
	str = str.replace("December", "ธันวาคม")
	return str

def replace_Date(date_string):
    months = {
        'Jan': 'มกราคม',
        'Feb': 'กุมภาพันธ์',
        'Mar': 'มีนาคม',
        'Apr': 'เมษายน',
        'May': 'พฤษภาคม',
        'Jun': 'มิถุนายน',
        'Jul': 'กรกฎาคม',
        'Aug': 'สิงหาคม',
        'Sep': 'กันยายน',
        'Oct': 'ตุลาคม',
        'Nov': 'พฤศจิกายน',
        'Dec': 'ธันวาคม',
    }
    # แปลงข้อความเป็นวันที่
    date_obj = datetime.strptime(date_string, '%b %d %Y %I:%M%p')
    # สร้างรูปแบบใหม่
    day = str(date_obj.day).zfill(2)  # เพิ่มเลข 0 ข้างหน้าเมื่อความยาวเป็น 1
    formatted_date = f"{day} {months[date_obj.strftime('%b')]} {date_obj.year + 543}"
    return formatted_date