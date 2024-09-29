from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        # First Row: Logo (left), Title (center), Contact Info (right)
        self.set_font('DejaVu', 'B', 10)
        
        # Left column - Logo
        self.image('Images/logo.jpg', 10, self.get_y(), 20)  # Adjusted size for smaller logo
        
        # Center column - Title
        self.set_xy(40, self.get_y() + 5)  # Adjusted top margin
        self.set_font('DejaVu', 'B', 12)
        self.set_text_color(255, 102, 0)  # Set title color to orange
        
        # Title 1
        self.cell(130, 8, 'CBS TRUNG TÂM NGOẠI NGỮ', 0, 1, 'C')
        
        # Title 2 right after the first one (no space)
        self.cell(190, 8, 'CON ONG THÔNG MINH', 0, 1, 'C')
        
        # Right column - Contact Information
        self.set_xy(160, self.get_y() - 16)  # Adjust position for contact info
        self.set_font('DejaVu', '', 6)
        self.set_text_color(0, 0, 0)  # Black color for text
        self.multi_cell(40, 5, '18 ĐƯỜNG SỐ 2, P5, Q8, TPHCM\n077 900 7218 – 0938 893 479', 1, 'C')

        # Add some space between titles and the next section
        self.ln(10)

    def footer(self):
        # Footer with page number
        self.set_y(-15)
        self.set_font('DejaVu', 'I', 8)
        self.cell(0, 10, 'Trang ' + str(self.page_no()), 0, 0, 'C')

def initialize(pdf):
    # Initialize PDF
    pdf.set_auto_page_break(auto=True, margin=10)

    # Add DejaVu font for Vietnamese text support
    pdf.add_font('DejaVu', '', 'PDF\\dejavu-sans\\ttf\\DejaVuSans.ttf')
    pdf.add_font('DejaVu', 'B', 'PDF\\dejavu-sans\\ttf\\DejaVuSans-Bold.ttf')
    pdf.add_font('DejaVu', 'I', 'PDF\\dejavu-sans\\ttf\\DejaVuSans-Oblique.ttf')  # Italic font

    # Add a page
    pdf.add_page()

    # Main Title
    pdf.set_font('DejaVu', 'B', 12)
    pdf.set_text_color(255, 0, 0)  # Red color for the main title
    pdf.cell(0, 10, 'KẾT QUẢ KIỂM TRA CUỐI KHÓA', new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(5)


def exam_data(pdf, level, address, exam_date, stage, exam_type, exam_time, main_teacher, examiner_teacher, 
              study_date, study_time):
    # Adjust font size for the table
    pdf.set_font('DejaVu', '', 8)

    # Table - Exam Details (6 columns, 4 rows)
    pdf.set_fill_color(250, 191, 143)  # Light orange background for header
    pdf.set_text_color(0, 0, 0)  # Black color for the main title

    # Row 1
    #pdf.cell(width, height, '', border=0, ln=0, align=align, fill=fill)
    pdf.cell(25, 8, 'Cấp độ thi:', 1, 0, 'C', True)  # Widened column
    pdf.cell(40, 8, level, 1, 0, 'C')
    pdf.cell(25, 8, 'Địa điểm thi:', 1, 0, 'C', True)
    pdf.cell(40, 8, address, 1, 0, 'C')
    pdf.cell(25, 8, 'Ngày thi:', 1, 0, 'C', True)
    pdf.cell(35, 8, exam_date, 1, 1, 'C')


    # Row 2
    pdf.cell(25, 8, 'Giai đoạn:', 1, 0, 'C', True)
    pdf.cell(40, 8,  stage, 1, 0, 'C')
    pdf.cell(25, 8, 'Bài thi:', 1, 0, 'C', True)
    pdf.cell(40, 8, exam_type, 1, 0, 'C')
    pdf.cell(25, 8, 'Giờ thi:', 1, 0, 'C', True)
    pdf.cell(35, 8, exam_time, 1, 1, 'C')

    # Row 3
    pdf.cell(25, 8, 'GV chính:', 1, 0, 'C', True)
    pdf.cell(40, 8, main_teacher, 1, 0, 'C')
    pdf.cell(25, 8, 'GV gác thi:', 1, 0, 'C', True)
    pdf.cell(40, 8, examiner_teacher, 1, 0, 'C')
    pdf.cell(25, 8, 'Ngày học:', 1, 0, 'C', True)
    pdf.cell(35, 8, study_date, 1, 1, 'C')

    # Row 4
    pdf.cell(25, 8, '', 0, 0, 'C', False)
    pdf.cell(40, 8, '', 0, 0, 'C', False)
    pdf.cell(25, 8, 'Giờ học:', 1, 0, 'C', True)
    pdf.cell(40, 8, study_time, 1, 0, 'C')
    pdf.cell(25, 8, '', 0, 0, 'C', False)
    pdf.cell(35, 8, '', 0, 1, 'C', False)

    pdf.ln(5)


def student_data(pdf, name, birth, class_no):
    # Student Info Footer (6 columns, 1 row)
    pdf.set_font('DejaVu', '', 10)
    pdf.set_text_color(255, 0, 0)  # Red color for the main title
    pdf.set_fill_color(255, 255, 255)
    pdf.cell(30, 10, 'Họ và tên HS:', 1, 0, 'L', True)
    pdf.cell(65, 10, name, 1, 0, 'L')
    pdf.cell(25, 10, 'Ngày sinh:', 1, 0, 'L')
    pdf.cell(30, 10, birth, 1, 0, 'L')
    pdf.cell(15, 10, 'Lớp:', 1, 0, 'L')
    pdf.cell(20, 10, class_no, 1, 1, 'L')

    pdf.ln(5)


def create_table(pdf, xt, yt, zt, xf, yf, zf, stage, listening, reading, speaking):
    pdf.set_font('DejaVu', 'B', 8)
    pdf.set_text_color(xt, yt, zt)  # Red color for the main title
    pdf.set_fill_color(xf, yf, zf)

    pdf.cell(30, 8, stage, 1, 0, 'C', True)
    pdf.cell(165, 8, '', 0, 1, 'C', False)

    pdf.cell(30, 8, 'Điểm nghe', 1, 0, 'C', True)
    pdf.cell(35, 8, str(listening) + '/20', 1, 0, 'C')
    pdf.cell(30, 8, 'Điểm đọc viết', 1, 0, 'C', True)
    pdf.cell(35, 8, str(reading) + '/20', 1, 0, 'C')
    pdf.cell(30, 8, 'Điểm nói', 1, 0, 'C', True)
    pdf.cell(35, 8, str(speaking) + '/20', 1, 1, 'C')

    total_grade = listening + reading + speaking
    percent = total_grade / 60 * 100

    pdf.cell(30, 8, 'Điểm tổng cộng:', 1, 0, 'C', True)
    pdf.cell(65, 8, str(total_grade) + '/60', 1, 0, 'C')
    pdf.cell(65, 8, 'Phần trăm điểm đạt', 1, 0, 'C')
    pdf.cell(35, 8, str(percent) + '%', 1, 1, 'C')

    pdf.cell(30, 8, 'Nhận xét:', 1, 0, 'C', True)
    pdf.cell(165, 8, 'Giỏi hơn anh Dũng là cái chắc', 1, 1, 'C')

    pdf.ln(5)


def return_value(pdf, level, address, exam_date, stage, exam_type, exam_time, main_teacher, examiner_teacher, study_date, study_time, 
                 name, birth, class_no, 
                 stage1, listening1, reading1, speaking1, stage2, listening2, reading2, speaking2, stage3, listening3, reading3, speaking3, 
                 stage4, listening4, reading4, speaking4, stage5, listening5, reading5, speaking5):
    # Exam information
    exam_data(pdf, level, address, exam_date, stage, exam_type, exam_time, main_teacher, examiner_teacher, study_date, study_time)
    
    # Student information
    student_data(pdf, name, birth, class_no)
    
    # Grade in every stage
    create_table(pdf, 0, 0, 0, 239, 210, 209, stage1, listening1, reading1, speaking1)
    create_table(pdf, 0, 0, 0, 250, 191, 143, stage2, listening2, reading2, speaking2)
    create_table(pdf, 0, 0, 0, 189, 210, 246, stage3, listening3, reading3, speaking3)
    create_table(pdf, 0, 0, 0, 194, 214, 155, stage4, listening4, reading4, speaking4)
    create_table(pdf, 0, 0, 0, 178, 161, 199, stage5, listening5, reading5, speaking5)


import os
from datetime import datetime

def create_file():
    # Initialize pdf
    pdf = PDF('P', 'mm', 'A4')
    initialize(pdf)
    
    # Examination information
    return_value(pdf, 'M', 'C.HB406','11/09/2024','23','Reading','15h30','Nguyễn Trung Sơn','Phạm Hoàng Dũng','M-W-F','19h20', 
                 'Phạm Ngọc Quỳnh Nhi', '15/06/2003', '8A11', 
                 'Giai đoạn 1', 20, 20, 20, 'Giai đoạn 2', 20, 20, 20, 'Giai đoạn 3', 20, 20, 20, 
                 'Giai đoạn 4', 20, 20, 20, 'Giai đoạn 5', 20, 20, 20)

    # Lấy thời gian hiện tại
    now = datetime.now()

    # Định dạng thời gian theo giờ phút giây ngày tháng năm
    time_string = now.strftime("%H%M%S-%d%m%Y")

    # Tạo tên file với thời gian hiện tại
    output_file = f'PhamHoangDung-{time_string}.pdf'

    # Đường dẫn đến thư mục Downloads
    download_path = os.path.join(os.path.expanduser('~'), 'Downloads', output_file)

    # Lưu file PDF vào thư mục Downloads
    pdf.output(download_path)

    print(f'File saved in: {download_path}')

    print(f"PDF created successfully: {output_file}")



create_file()