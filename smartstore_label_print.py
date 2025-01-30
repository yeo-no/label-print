import pandas as pd
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import traceback
import subprocess  # 디렉토리 열기 및 포커스용

# -----------------------------------------------------------------------------
# 텍스트 줄바꿈 (제한 없는 기본 함수)
# -----------------------------------------------------------------------------
def wrap_text_by_width(pdf, text, max_width):
    words = text.split()
    lines = []
    current_line = ""
    
    for w in words:
        if not current_line:
            current_line = w
        else:
            test_line = current_line + " " + w
            if pdf.get_string_width(test_line) <= max_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = w
    if current_line:
        lines.append(current_line)
    return lines

# -----------------------------------------------------------------------------
# 말줄임표 처리, 최대 줄 제한 함수
# -----------------------------------------------------------------------------
def wrap_text_by_width_with_limit(pdf, text, max_width, max_lines=2):
    words = text.split()
    lines = []
    current_line = ""

    for w in words:
        if not current_line:
            current_line = w
        else:
            test_line = current_line + " " + w
            if pdf.get_string_width(test_line) <= max_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = w
                if len(lines) == max_lines:  # 최대 줄에 도달
                    break

    # 마지막 줄 추가
    if current_line and len(lines) < max_lines:
        lines.append(current_line)

    # 실제로 텍스트가 잘린 경우만 말줄임표 붙이기
    used_words_count = len(" ".join(lines).split())
    total_words_count = len(words)
    if len(lines) == max_lines and used_words_count < total_words_count:
        last_line = lines[-1]
        while pdf.get_string_width(last_line + "...") > max_width and len(last_line) > 0:
            last_line = last_line[:-1]
        lines[-1] = last_line + "..."

    return lines

# -----------------------------------------------------------------------------
# 말줄임표 + 행걸이 들여쓰기
# -----------------------------------------------------------------------------
def print_hanging_text_with_limit(pdf, x, y, max_width, prefix, text, line_height=5, max_lines=2):
    pdf.set_xy(x, y)
    prefix_w = pdf.get_string_width(prefix)
    indent_width = max_width - prefix_w

    lines = wrap_text_by_width_with_limit(pdf, text, indent_width, max_lines)

    if lines:
        first_line = lines[0]
    else:
        first_line = ""

    pdf.cell(prefix_w, line_height, prefix, border=0, ln=0)
    pdf.cell(indent_width, line_height, first_line, border=0, ln=1)

    for line in lines[1:]:
        pdf.set_x(x + prefix_w)
        pdf.cell(indent_width, line_height, line, border=0, ln=1)

    return pdf.get_y()

# -----------------------------------------------------------------------------
# PDF 생성 함수
# -----------------------------------------------------------------------------
def generate_label_pdf(input_excel, output_pdf, row_count, font_size):
    file_ext = os.path.splitext(input_excel)[1].lower()

    # 엑셀 파일 읽기
    if file_ext == ".xls":
        df = pd.read_excel(input_excel, engine="xlrd", skiprows=1, header=0)
    elif file_ext == ".xlsx":
        df = pd.read_excel(input_excel, engine="openpyxl", skiprows=1, header=0)
    else:
        raise ValueError("지원되지 않는 파일 형식입니다. .xls 또는 .xlsx 파일을 사용하세요.")

    # 컬럼 체크
    product_column = '상품명'
    quantity_column = '수량'
    option_column = '옵션정보'
    order_no_column = '주문번호'
    if any(col not in df.columns for col in [order_no_column, product_column, quantity_column, option_column]):
        raise ValueError("엑셀 파일에 필요한 컬럼(주문번호, 상품명, 수량, 옵션정보)이 없습니다.")

    # 주문번호별 순번
    unique_order_numbers = df[order_no_column].unique()
    order_number_mapping = {order_no: idx + 1 for idx, order_no in enumerate(unique_order_numbers)}
    df['순번'] = df[order_no_column].map(order_number_mapping)

    # PDF 객체 생성
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    # 폰트 설정
    pdf.add_font('MalgunGothic', '', r"C:\Windows\Fonts\malgun.ttf", uni=True)
    pdf.set_font("MalgunGothic", size=font_size)

    # 라벨 배치 옵션
    labels_per_row = 2
    margin_x = 8
    margin_y = 5
    start_x = margin_x
    start_y = margin_y
    internal_margin = 3
    line_height = 5
    page_height = 297 - margin_y  # A4 세로 크기에서 약간의 여백

    # row_count를 사용해 라벨 높이 추정
    label_height = (page_height / row_count) - margin_y  
    label_width = 90  # 라벨 폭(고정)

    x = start_x
    y = start_y

    for _, row in df.iterrows():
        order_no = row[order_no_column]
        seq = row['순번']
        product_name = str(row[product_column]).replace('\n', ' ').replace('\r', ' ')
        quantity = str(row[quantity_column])
        option_info = str(row[option_column]) if not pd.isna(row[option_column]) else None

        # 페이지 넘김
        if y + label_height > 297 - margin_y:
            pdf.add_page()
            x = start_x
            y = start_y

        # 라벨 테두리
        pdf.rect(x, y, label_width, label_height)

        # 텍스트 출력
        text_x = x + internal_margin
        text_y = y + internal_margin
        text_width = label_width - 2 * internal_margin

        # No. XX
        pdf.set_xy(text_x, text_y)
        pdf.cell(text_width, line_height, f"No. {seq:02}", border='B', ln=1, align='C')
        pdf.cell(0, 2, '', ln=1)

        # 상품명
        curr_y = print_hanging_text_with_limit(
            pdf=pdf,
            x=text_x,
            y=pdf.get_y(),
            max_width=text_width,
            prefix="상품명: ",
            text=product_name,
            line_height=line_height,
            max_lines=2
        )

        # 옵션정보
        if option_info:
            curr_y = print_hanging_text_with_limit(
                pdf=pdf,
                x=text_x,
                y=curr_y,
                max_width=text_width,
                prefix="옵션정보: ",
                text=option_info,
                line_height=line_height,
                max_lines=2
            )

        # 수량
        pdf.set_xy(text_x, curr_y)
        pdf.cell(text_width, line_height, f"수량: {quantity}", border=0, ln=1)

        # 다음 라벨 위치
        if x + 2 * label_width + margin_x > pdf.w:
            x = start_x
            y += label_height + margin_y
        else:
            x += label_width + margin_x

    pdf.output(output_pdf)
    return output_pdf

def check_filename_collision(file_path):
    """
    파일이 존재하면 (1), (2) ... 를 붙여가며 유니크한 파일명을 생성.
    """
    base, ext = os.path.splitext(file_path)
    counter = 1
    new_path = file_path
    while os.path.exists(new_path):
        new_path = f"{base} ({counter}){ext}"
        counter += 1
    return new_path

class DualInputDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("라벨 설정")
        self.geometry("300x150")
        self.resizable(False, False)

        # 창을 화면 중앙에 배치
        self.update_idletasks()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)

        tk.Label(self, text="Row Count (default: 7)", font=("맑은 고딕", 9), anchor='w').grid(
            row=0, column=0, sticky='w', padx=10, pady=10)
        tk.Label(self, text="Font Size (default: 10)", font=("맑은 고딕", 9), anchor='w').grid(
            row=1, column=0, sticky='w', padx=10, pady=10)

        self.row_count_var = tk.StringVar()
        self.font_size_var = tk.StringVar()

        tk.Entry(self, textvariable=self.row_count_var, font=("맑은 고딕", 10)).grid(
            row=0, column=1, padx=10, pady=10)
        tk.Entry(self, textvariable=self.font_size_var, font=("맑은 고딕", 10)).grid(
            row=1, column=1, padx=10, pady=10)

        tk.Button(self, text="확인", command=self.on_confirm, font=("맑은 고딕", 10)).grid(
            row=2, column=0, columnspan=2, pady=15)

        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_confirm(self):
        try:
            # 사용자가 입력하지 않은 경우 기본값 설정
            row_count = int(self.row_count_var.get().strip()) if self.row_count_var.get().strip() else 7
            font_size = int(self.font_size_var.get().strip()) if self.font_size_var.get().strip() else 10
        except:
            messagebox.showerror("입력 오류", "정수 형태로 입력하세요.")
            return

        if row_count < 1:
            messagebox.showerror("입력 오류", "라벨 '행' 수는 최소 1 이상이어야 합니다.")
            return
        if font_size < 6 or font_size > 30:
            messagebox.showerror("입력 오류", "폰트 사이즈는 6~30 사이여야 합니다.")
            return

        self.result = (row_count, font_size)
        self.destroy()

    def on_close(self):
        self.result = None
        self.destroy()

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

import subprocess

def open_directory(file_path):
    target_path = os.path.normpath(file_path)  # 경로 변환
    cmd = f'explorer /select,"{target_path}"'  # 명령 생성

    print("Debug: target_path =", target_path)
    print("Debug: CMD =", cmd)

    try:
        subprocess.run(cmd, shell=True)
    except Exception as e:
        print(f"Error opening directory: {str(e)}")

def main():
    root = tk.Tk()
    root.withdraw()

    dialog = DualInputDialog(root)
    dialog.grab_set()
    root.wait_window(dialog)

    if not dialog.result:
        print("라벨 설정 취소.")
        return

    row_count, font_size = dialog.result
    excel_file = select_excel_file()
    if not excel_file:
        print("엑셀 파일이 선택되지 않았습니다.")
        return

    output_pdf = os.path.splitext(excel_file)[0] + "_labels.pdf"
    output_pdf = check_filename_collision(output_pdf)

    try:
        pdf_path = generate_label_pdf(excel_file, output_pdf, row_count, font_size)
        open_directory(pdf_path)  # 디렉토리 열기
        messagebox.showinfo("변환 성공", f"PDF 생성 성공!\n파일 경로:\n{pdf_path}")
    except Exception as e:
        err_text = traceback.format_exc()
        messagebox.showerror("오류 발생", f"PDF 생성 중 오류:\n{err_text}")

if __name__ == "__main__":
    main()