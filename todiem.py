import os
import pandas as pd
from glob import glob
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from io import BytesIO
import pdfplumber
import re

def load_excel_data(excel_path):
    try:
        df = pd.read_excel(excel_path)
        df['Mã SV'] = df['Mã SV'].astype(str).str.strip()
        grades = dict(zip(df['Mã SV'], df['Điểm']))
        print(f"Số lượng sinh viên trong file Excel: {len(grades)}")
        return grades
    except Exception as e:
        print(f"Lỗi khi đọc file Excel: {str(e)}")
        return None

def get_user_input_info():
    info = {
        "supervisor1": input("Nhập tên cán bộ coi thi 1: "),
        "supervisor2": input("Nhập tên cán bộ coi thi 2: "),
        "grader1": input("Nhập tên giảng viên chấm thi 1: "),
        "grader2": input("Nhập tên giảng viên chấm thi 2: ")
    }
    return info

def convert_to_text(score):
    if pd.isna(score):
        return "Vắng"
    if score < 0 or score > 10:
        return "Không hợp lệ"
    integer_part = int(score)
    decimal_part = score - integer_part
    number_words = ["Không", "Một", "Hai", "Ba", "Bốn", "Năm", "Sáu", "Bảy", "Tám", "Chín"]
    if integer_part == 10:
        return "Mười"
    if decimal_part == 0:
        return number_words[integer_part]
    elif round(decimal_part, 1) == 0.5:
        return f"{number_words[integer_part]} rưỡi"

def find_grade_column(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            for word in words:
                if "Điểm" in word['text']:
                    return word['x0'], word['top'], word['bottom']
    return None

def extract_student_positions(pdf_path):
    student_positions = {}
    student_id_pattern = r'\b[1-9]\d{8}\b'
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            words = page.extract_words()
            for word in words:
                text = word['text']
                if re.match(student_id_pattern, text):
                    x0 = word['x0']
                    y0 = word['top']
                    student_positions[text] = (x0, y0, page_num)
    print(f"Tổng số mã sinh viên tìm thấy trong PDF: {len(student_positions)}")
    return student_positions

def add_grade_to_pdf(input_pdf, output_pdf, grades, total_students, info):
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()
    
    # Đăng ký font
    pdfmetrics.registerFont(TTFont('arial', 'arial.ttf'))
    
    grade_column = find_grade_column(input_pdf)
    if grade_column is None:
        print("Không tìm thấy cột 'Điểm' trong PDF.")
        return
    
    column_x, column_top, column_bottom = grade_column
    student_positions = extract_student_positions(input_pdf)
    
    # Tính số sinh viên vắng dựa trên mã SV trong PDF và điểm NaN
    absent_count = sum(1 for ma_sv in student_positions if ma_sv in grades and pd.isna(grades[ma_sv]))
    print(f"Số lượng sinh viên vắng trong bảng điểm PDF: {absent_count}")
    
    total_pages = len(pdf_reader.pages)
    for page_num in range(total_pages):
        page = pdf_reader.pages[page_num]
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("arial", 9)
        
        # Thêm thông tin tổng số sinh viên và cán bộ vào trang đầu tiên
        if page_num == 0:
            can.drawString(125, 93.5, f"{total_students}")
            can.drawString(125, 75, f"{absent_count}")
            can.drawCentredString(380, 52, f"{info['supervisor1']}")
            can.drawCentredString(380, 15, f"{info['supervisor2']}")
            can.drawCentredString(505, 52, f"{info['grader1']}")
            can.drawCentredString(505, 15, f"{info['grader2']}")
        
        for ma_sv_pdf, position in student_positions.items():
            if position[2] - 1 != page_num:
                continue
                
            ma_sv_pdf = str(ma_sv_pdf).strip()
            x = column_x - 12
            y = position[1]
            y_adjusted = 832 - y
            
            if ma_sv_pdf in grades:
                score = grades[ma_sv_pdf]
                grade_text = convert_to_text(score)
                can.drawString(x, y_adjusted, grade_text)
                
                if score is not None and score >= 0 and score <= 10:
                    circle_x = column_x + 44.5 + int(score) * 16.7
                    circle_y = y_adjusted + 3
                    circle_diameter = 4 * 2.8346
                    radius = circle_diameter / 2
                    can.setFillColor(colors.black)
                    can.circle(circle_x, circle_y, radius, fill=1)
                    
                    if round(score - int(score), 1) == 0.5:
                        additional_circle_x = column_x + 229
                        additional_circle_y = y_adjusted + 3
                        can.circle(additional_circle_x, additional_circle_y, radius, fill=1)
            else:
                print(f"Mã SV {ma_sv_pdf} không có trong file Excel.")
                if ma_sv_pdf in grades and pd.isna(grades[ma_sv_pdf]):
                    grade_text = "Vắng"
                    can.drawString(x, y_adjusted, grade_text)
        
        can.save()
        packet.seek(0)
        
        try:
            new_pdf_reader = PdfReader(BytesIO(packet.read()))
            if len(new_pdf_reader.pages) > 0:
                page.merge_page(new_pdf_reader.pages[0])
        except Exception as e:
            print(f"Lỗi khi tạo trang mới: {str(e)}")
        
        pdf_writer.add_page(page)
    
    try:
        with open(output_pdf, "wb") as output_file:
            pdf_writer.write(output_file)
        print(f"Hoàn thành! Kiểm tra file {output_pdf}")
    except Exception as e:
        print(f"Lỗi khi ghi file PDF: {str(e)}")

def process_files(folder, keyword, info):
    excel_file = os.path.join(folder, f'grade_{keyword}.xlsx')
    pdf_files = [f for f in os.listdir(folder) if keyword in f and f.endswith('.pdf')]
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder, pdf_file)
        output_pdf = os.path.join(folder, f'output_{pdf_file}')
        grades = load_excel_data(excel_file)
        if grades is None:
            continue
        student_positions = extract_student_positions(pdf_path)
        total_students = len(student_positions)
        add_grade_to_pdf(pdf_path, output_pdf, grades, total_students, info)

def main():
    current_dir = os.getcwd()
    
    # Nhập thông tin cán bộ và giảng viên một lần
    info = get_user_input_info()
    
    # Tạo các file Excel từ file đầu vào
    xlsx_files = glob("*.xlsx")
    if not xlsx_files:
        print("Không tìm thấy file .xlsx nào.")
        return
    
    input_file = xlsx_files[0]
    df = pd.read_excel(input_file)
    
    for grade_type, column in [
        ('qt', 'Điểm quá trình'),
        ('gk', 'Điểm giữa kỳ'),
        ('ck', 'Điểm cuối kỳ')
    ]:
        if 'StudentID' in df.columns and column in df.columns:
            grade_df = df[['StudentID', column]].copy()
            grade_df.columns = ['Mã SV', 'Điểm']
            grade_file = f'grade_{grade_type}.xlsx'
            grade_df.to_excel(grade_file, index=False)
            print(f"Đã tạo file {grade_file} thành công.")
    
    # Đổi tên file PDF
    for file_name in os.listdir(current_dir):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(current_dir, file_name)
            with open(file_path, "rb") as pdf_file:
                reader = PdfReader(pdf_file)
                content = ""
                for page in reader.pages:
                    content += page.extract_text() or ""
                
                new_file_name = file_name
                for keyword in ['quá trình', 'giữa kỳ', 'cuối kỳ']:
                    if keyword in content.lower():
                        suffix = {'quá trình': '_qt', 'giữa kỳ': '_gk', 'cuối kỳ': '_ck'}[keyword]
                        new_file_name = new_file_name.replace(".pdf", f"{suffix}.pdf")
                        break
                
                if new_file_name != file_name:
                    new_file_path = os.path.join(current_dir, new_file_name)
                    os.rename(file_path, new_file_Path)
                    print(f"Đã đổi tên: {file_name} thành {new_file_name}")
    
    # Xử lý các file
    for keyword in ['qt', 'gk', 'ck']:
        process_files(current_dir, keyword, info)
    
    # Xóa file không cần thiết
    for file in os.listdir(current_dir):
        file_path = os.path.join(current_dir, file)
        if (file.endswith('.pdf') and 'output' not in file.lower()) or \
           file.endswith('.xlsx'):
            os.remove(file_path)
            print(f"Đã xóa file: {file}")

if __name__ == "__main__":
    main()
