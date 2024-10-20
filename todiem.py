import os
import pandas as pd
from glob import glob

# Tìm file .xlsx đầu vào trong thư mục hiện tại
xlsx_files = glob("*.xlsx")

if len(xlsx_files) == 0:
    print("Không tìm thấy file .xlsx nào trong thư mục hiện tại.")
else:
    input_file = xlsx_files[0]  # Lấy file .xlsx đầu tiên
    print(f"Đang xử lý file: {input_file}")

    # Đọc file Excel đầu vào
    df = pd.read_excel(input_file)

    # Tạo file grade_qt.xlsx với cột "Mã SV" và "Điểm"
    if 'StudentID' in df.columns and 'Điểm quá trình' in df.columns:
        grade_qt = df[['StudentID', 'Điểm quá trình']].copy()
        grade_qt.columns = ['Mã SV', 'Điểm']
        grade_qt_file = 'grade_qt.xlsx'
        grade_qt.to_excel(grade_qt_file, index=False)
        print(f"Đã tạo file {grade_qt_file} thành công.")
    else:
        print("Không tìm thấy cột 'StudentID' hoặc 'Điểm quá trình' trong file đầu vào.")

    # Tạo file grade_gk.xlsx với cột "Mã SV" và "Điểm"
    if 'StudentID' in df.columns and 'Điểm giữa kỳ' in df.columns:
        grade_gk = df[['StudentID', 'Điểm giữa kỳ']].copy()
        grade_gk.columns = ['Mã SV', 'Điểm']
        grade_gk_file = 'grade_gk.xlsx'
        grade_gk.to_excel(grade_gk_file, index=False)
        print(f"Đã tạo file {grade_gk_file} thành công.")
    else:
        print("Không tìm thấy cột 'StudentID' hoặc 'Điểm giữa kỳ' trong file đầu vào.")

    # Tạo file grade_ck.xlsx với cột "Mã SV" và "Điểm"
    if 'StudentID' in df.columns and 'Điểm cuối kỳ' in df.columns:
        grade_ck = df[['StudentID', 'Điểm cuối kỳ']].copy()
        grade_ck.columns = ['Mã SV', 'Điểm']
        grade_ck_file = 'grade_ck.xlsx'
        grade_ck.to_excel(grade_ck_file, index=False)
        print(f"Đã tạo file {grade_ck_file} thành công.")
    else:
        print("Không tìm thấy cột 'StudentID' hoặc 'Điểm cuối kỳ' trong file đầu vào.")


import os
from PyPDF2 import PdfReader

def rename_pdf_files(folder_path):
    # Duyệt qua tất cả các file PDF trong thư mục
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)

            # Đọc nội dung file PDF
            with open(file_path, "rb") as pdf_file:
                reader = PdfReader(pdf_file)
                content = ""
                for page in reader.pages:
                    content += page.extract_text()  # Trích xuất nội dung từ mỗi trang

                # Kiểm tra các từ khóa và đổi tên file tương ứng
                new_file_name = file_name
                if "quá trình" in content:
                    new_file_name = new_file_name.replace(".pdf", "_qt.pdf")
                elif "giữa kỳ" in content:
                    new_file_name = new_file_name.replace(".pdf", "_gk.pdf")
                elif "cuối kỳ" in content:
                    new_file_name = new_file_name.replace(".pdf", "_ck.pdf")

                # Đổi tên file nếu cần
                if new_file_name != file_name:
                    new_file_path = os.path.join(folder_path, new_file_name)
                    os.rename(file_path, new_file_path)
                    print(f"Đã đổi tên: {file_name} thành {new_file_name}")

# Gọi hàm với đường dẫn đến thư mục chứa các file PDF
folder_path = "./"  # Thay thế bằng đường dẫn tới thư mục của bạn
rename_pdf_files(folder_path)

# @title Viết và tô điểm
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from io import BytesIO
import pdfplumber
import re
import os

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

def convert_to_text(score):
    if pd.isna(score):  # Kiểm tra nếu ô điểm trống
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
    student_id_pattern = r'\b\d{9}\b'

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

def add_grade_to_pdf(input_pdf, output_pdf, grades):
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()

    # Đăng ký font Times
    pdfmetrics.registerFont(TTFont('arial', '/content/todiem/arial.ttf'))

    grade_column = find_grade_column(input_pdf)
    if grade_column is None:
        print("Không tìm thấy cột 'Điểm' trong PDF.")
        return

    column_x, column_top, column_bottom = grade_column
    print(f"Tọa độ cột 'Điểm': x={column_x}, y_top={column_top}")

    student_positions = extract_student_positions(input_pdf)
    if not student_positions:
        print("Không tìm thấy mã sinh viên nào trong PDF.")
        return

    total_pages = len(pdf_reader.pages)
    for page_num in range(total_pages):
        page = pdf_reader.pages[page_num]
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("arial", 9)

        for ma_sv_pdf, position in student_positions.items():
            if position[2] - 1 != page_num:  # Kiểm tra xem mã SV có thuộc trang hiện tại không
                continue

            ma_sv_pdf = str(ma_sv_pdf).strip()
            x = column_x - 12  # Vị trí x để ghi điểm
            y = position[1]
            y_adjusted = 832 - y  # Điều chỉnh vị trí y

            if ma_sv_pdf in grades:
                score = grades[ma_sv_pdf]
                grade_text = convert_to_text(score)

                # Ghi điểm vào PDF
                can.drawString(x, y_adjusted, grade_text)

                # Vẽ hình tròn chỉ khi mã SV có trong file Excel và có điểm
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

                # Chỉ ghi "Vắng" nếu mã SV có trong file Excel nhưng không có điểm
                if ma_sv_pdf in grades and pd.isna(grades[ma_sv_pdf]):
                    grade_text = "Vắng"
                    can.drawString(x, y_adjusted, grade_text)
                    # Nếu đã ghi "Vắng", không vẽ hình tròn
                    continue

        can.save()
        packet.seek(0)

        try:
            new_pdf_reader = PdfReader(BytesIO(packet.read()))
            if len(new_pdf_reader.pages) > 0:
                page.merge_page(new_pdf_reader.pages[0])
            else:
                print(f"Không có nội dung mới để thêm vào trang {page_num + 1}")
        except Exception as e:
            print(f"Lỗi khi tạo trang mới cho trang {page_num + 1}: {str(e)}")

        pdf_writer.add_page(page)

    try:
        with open(output_pdf, "wb") as output_file:
            pdf_writer.write(output_file)
        print(f"Hoàn thành! Kiểm tra file {output_pdf}")
    except Exception as e:
        print(f"Lỗi khi ghi file PDF: {str(e)}")

def process_files(folder, keyword):
    # Tạo đường dẫn cho các file Excel và PDF
    excel_file = os.path.join(folder, f'{keyword}.xlsx')
    pdf_files = [f for f in os.listdir(folder) if keyword in f and f.endswith('.pdf')]

    # Chạy qua từng file PDF có chứa keyword
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder, pdf_file)
        output_pdf = os.path.join(folder, f'output_{pdf_file}')

        grades = load_excel_data(excel_file)
        if grades is None:
            continue

        add_grade_to_pdf(pdf_path, output_pdf, grades)

def main():
    current_dir = os.getcwd()  # Lấy thư mục hiện tại

    # Tìm file Excel chứa "qt", "gk" và "ck" trong thư mục hiện tại
    excel_file_qt = None
    excel_file_gk = None
    excel_file_ck = None
    files = os.listdir(current_dir)  # Lấy danh sách file trong thư mục hiện tại

    for file in files:
        if 'qt' in file.lower() and file.endswith('.xlsx'):
            excel_file_qt = os.path.join(current_dir, file)
            print(f"Tìm thấy file Excel chứa 'qt': {excel_file_qt}")
            break  # Chỉ cần 1 file, nên dừng lại

    for file in files:
        if 'gk' in file.lower() and file.endswith('.xlsx'):
            excel_file_gk = os.path.join(current_dir, file)
            print(f"Tìm thấy file Excel chứa 'gk': {excel_file_gk}")
            break  # Chỉ cần 1 file, nên dừng lại

    for file in files:
        if 'ck' in file.lower() and file.endswith('.xlsx'):
            excel_file_ck = os.path.join(current_dir, file)
            print(f"Tìm thấy file Excel chứa 'ck': {excel_file_ck}")
            break  # Chỉ cần 1 file, nên dừng lại

    if excel_file_qt is None:
        print("Không tìm thấy file Excel chứa 'qt'.")
        return

    # Lần lượt xử lý các file PDF có chứa "qt" trong thư mục hiện tại
    pdf_files_qt = [f for f in files if 'qt' in f.lower() and f.endswith('.pdf')]

    for pdf_file in pdf_files_qt:
        input_pdf = os.path.join(current_dir, pdf_file)
        output_pdf = os.path.join(current_dir, f'output_{pdf_file}')  # Tạo tên file đầu ra
        print(f"Đang xử lý file PDF: {input_pdf} với Excel: {excel_file_qt}")
        grades = load_excel_data(excel_file_qt)
        if grades is not None:
            add_grade_to_pdf(input_pdf, output_pdf, grades)

    # Lần lượt xử lý các file PDF có chứa "gk" trong thư mục hiện tại
    if excel_file_gk is not None:
        pdf_files_gk = [f for f in files if 'gk' in f.lower() and f.endswith('.pdf')]

        for pdf_file in pdf_files_gk:
            input_pdf = os.path.join(current_dir, pdf_file)
            output_pdf = os.path.join(current_dir, f'output_{pdf_file}')  # Tạo tên file đầu ra
            print(f"Đang xử lý file PDF: {input_pdf} với Excel: {excel_file_gk}")
            grades = load_excel_data(excel_file_gk)
            if grades is not None:
                add_grade_to_pdf(input_pdf, output_pdf, grades)

    # Lần lượt xử lý các file PDF có chứa "ck" trong thư mục hiện tại
    if excel_file_ck is not None:
        pdf_files_ck = [f for f in files if 'ck' in f.lower() and f.endswith('.pdf')]

        for pdf_file in pdf_files_ck:
            input_pdf = os.path.join(current_dir, pdf_file)
            output_pdf = os.path.join(current_dir, f'output_{pdf_file}')  # Tạo tên file đầu ra
            print(f"Đang xử lý file PDF: {input_pdf} với Excel: {excel_file_ck}")
            grades = load_excel_data(excel_file_ck)
            if grades is not None:
                add_grade_to_pdf(input_pdf, output_pdf, grades)

if __name__ == "__main__":
    main()

import os

def delete_unnecessary_files():
    current_dir = os.getcwd()  # Lấy thư mục hiện tại
    files = os.listdir(current_dir)  # Lấy danh sách tất cả các file trong thư mục hiện tại

    for file in files:
        file_path = os.path.join(current_dir, file)

        # Nếu là file PDF nhưng không chứa chữ "output" thì xóa
        if file.endswith('.pdf') and 'output' not in file.lower():
            os.remove(file_path)
        # Nếu là file Excel thì xóa
        elif file.endswith('.xlsx'):
            os.remove(file_path)
        # Nếu là file arial.ttf thì xóa
        elif file.lower() == 'arial.ttf':
            os.remove(file_path)
# Gọi hàm để thực hiện xóa file
delete_unnecessary_files()
