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
        
        # --- BẮT ĐẦU ĐOẠN SỬA ---
        def chuan_hoa_mssv(value):
            try:
                # Nếu là NaN/None
                if pd.isna(value): return ""
                
                # Ép về float trước (xử lý cả trường hợp int, float, str dạng số)
                # Sau đó ép về int để cắt bỏ phần thập phân (.0)
                # Cuối cùng ép về str
                return str(int(float(value)))
            except:
                # Nếu không thể ép kiểu số (vd: mã có chữ cái), chỉ cắt khoảng trắng
                return str(value).strip()

        # Áp dụng hàm chuẩn hóa cho cột Mã SV
        df['Mã SV'] = df['Mã SV'].apply(chuan_hoa_mssv)
        # --- KẾT THÚC ĐOẠN SỬA ---

        grades = dict(zip(df['Mã SV'], df['Điểm']))
        
        # In ra để kiểm tra
        print(f"Số lượng sinh viên trong file Excel: {len(grades)}")
        if len(grades) > 0:
            print(f"Ví dụ mã SV sau khi chuẩn hóa (Excel): {list(grades.keys())[0]}")
            
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
    pdfmetrics.registerFont(TTFont('arial', '/content/todiem/arial.ttf'))
    
    grade_column = find_grade_column(input_pdf)
    if grade_column is None:
        print("Không tìm thấy cột 'Điểm' trong PDF.")
        return
    
    column_x, column_top, column_bottom = grade_column
    student_positions = extract_student_positions(input_pdf)
    
    # Tính số sinh viên vắng dựa trên mã SV trong PDF và điểm NaN
    absent_count = sum(1 for ma_sv in student_positions if ma_sv in grades and pd.isna(grades[ma_sv]))
    print(f"Số lượng sinh viên vắng trong bảng điểm PDF: {absent_count}")
    # Tính số số SV dự thi
    present_count = total_students - absent_count
    total_pages = len(pdf_reader.pages)
    for page_num in range(total_pages):
        page = pdf_reader.pages[page_num]
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("arial", 9)
        
        # Thêm thông tin tổng số sinh viên và cán bộ vào trang đầu tiên
        if page_num == 0:
            can.drawString(125, 93.5, f"{present_count}")
            can.drawString(125, 75, f"{absent_count}")
            can.drawCentredString(380, 49, f"{info['supervisor1']}")
            can.drawCentredString(380, 18, f"{info['supervisor2']}")
            can.drawCentredString(505, 49, f"{info['grader1']}")
            can.drawCentredString(505, 18, f"{info['grader2']}")
        
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
                    os.rename(file_path, new_file_path)
    
    # Xử lý các file
    for keyword in ['qt', 'gk', 'ck']:
        process_files(current_dir, keyword, info)

    
    # Scale factor
    scale = 0.99

    for filename in os.listdir():
        if filename.endswith(".pdf") and "output" in filename:
            reader = PdfReader(filename)
            writer = PdfWriter()

            for page in reader.pages:
                # Thu nhỏ trang bằng PyPDF2
                page.scale_by(scale)
                writer.add_page(page)

            new_filename = f"scaled_{filename}"
            with open(new_filename, "wb") as f_out:
                writer.write(f_out)

            print(f"Đã tạo: {new_filename}")

    # @title Ghép thành 1 file pdf (qt, gk, ck nếu có)
    # --- ĐÃ XÓA IMPORT OS, GLOB Ở ĐÂY ---
    from PyPDF2 import PdfMerger # Dòng này để lại hoặc đưa lên đầu file cũng được, nhưng ở đây không gây lỗi biến os.
    
    # Thư mục chứa các file PDF
    input_folder = "."
    
    # Lấy danh sách file PDF có dạng scaled_output_*
    # Lưu ý: glob đã import ở đầu file rồi, nên dùng glob.glob bình thường
    all_pdfs = glob(os.path.join(input_folder, "scaled_output_*.pdf")) 
    
    # Tách thành 3 nhóm: qt, gk, ck
    qt_files = sorted([f for f in all_pdfs if "qt" in os.path.basename(f).lower()])
    gk_files = sorted([f for f in all_pdfs if "gk" in os.path.basename(f).lower()])
    ck_files = sorted([f for f in all_pdfs if "ck" in os.path.basename(f).lower()])
    
    # Ghép theo thứ tự: qt → gk → ck
    pdf_files = qt_files + gk_files + ck_files
    
    # Chỉ ghép nếu có ít nhất 1 file
    if pdf_files:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
    
        output_file = "merged_qt_gk_ck.pdf"
        merger.write(output_file)
        merger.close()
    
        print(f"Đã ghép {len(pdf_files)} file (qt → gk → ck) thành: {output_file}")
    else:
        print("Không tìm thấy file nào để ghép!")
    
    # Xóa file không cần thiết
    for file in os.listdir(current_dir):
        file_path = os.path.join(current_dir, file)
        
        # --- LƯU Ý LOGIC XÓA FILE (Tôi đã sửa lại giúp bạn chỗ logic AND vô lý) ---
        # Logic cũ của bạn: file.endswith('qt.xlsx') AND file.endswith('gk.xlsx')... -> Không bao giờ xảy ra
        # Logic sửa: Dùng OR để xóa các file tạm
        is_temp_pdf = (file.endswith('.pdf') and 'scaled' not in file.lower() and file != "merged_qt_gk_ck.pdf")
        is_temp_xlsx = (file.endswith('grade_qt.xlsx') or file.endswith('grade_gk.xlsx') or file.endswith('grade_ck.xlsx'))
        
        if is_temp_pdf or is_temp_xlsx:
            try:
                os.remove(file_path)
            except:
                pass

if __name__ == "__main__":
    main()
