import os
import win32com.client
import pdfplumber
from pywinauto import Application
import re
import time
import pyautogui
import threading 


def open_hwp_file(hwp_file, output_pdf):
    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule('')
    hwp.Open(hwp_file)  # HWP 파일 열기
    time.sleep(2)
    hwp.SaveAs(output_pdf, "PDF")  # PDF로 저장
    time.sleep(2)
    # pyautogui.moveTo(1920, 0)
    # pyautogui.click(1920, 0, clicks=1)
    # hwp.Clear()
    hwp.Quit()  # 한글 종료
    return 

def convert_hwp_to_pdf_windows(hwp_file, output_pdf):
    print(f"Converting HWP file to PDF: {hwp_file} (Windows)")
    try:
        pyautogui.moveTo(941, 570)
        thread = threading.Thread(target=open_hwp_file, args=(hwp_file, output_pdf))
        thread.start()
        time.sleep(1)
        pyautogui.click(clicks=3, interval=1)
        thread.join()
        return output_pdf
    except Exception as e:
        print(f"Failed to convert HWP to PDF: {hwp_file} due to {e}")
        return None

# DOCX 파일을 PDF로 변환하는 함수 (Windows용)
def convert_docx_to_pdf_windows(docx_file, output_pdf):
    try:
        print(f"Converting DOCX file to PDF: {docx_file} (Windows)")
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_file)
        doc.SaveAs(output_pdf, FileFormat=17)  # 17은 PDF 포맷
        doc.Close()
        word.Quit()
        return output_pdf
    except Exception as e:
        print(f"Failed to convert DOCX to PDF: {docx_file} due to {e}")
        return None

# PPTX 파일을 PDF로 변환하는 함수 (Windows용)
def convert_pptx_to_pdf_windows(pptx_file, output_pdf):
    try:
        print(f"Converting PPTX file to PDF: {pptx_file} (Windows)")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        ppt = powerpoint.Presentations.Open(pptx_file, WithWindow=False)
        ppt.SaveAs(output_pdf, 32)  # 32는 PDF 포맷
        ppt.Close()
        powerpoint.Quit()
        return output_pdf
    except Exception as e:
        print(f"Failed to convert PPTX to PDF: {pptx_file} due to {e}")
        return None

# XLSX 파일을 PDF로 변환하는 함수 (Windows용)
def convert_xlsx_to_pdf_windows(xlsx_file, output_pdf):
    try:
        print(f"Converting XLSX file to PDF: {xlsx_file} (Windows)")
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(xlsx_file)
        workbook.ExportAsFixedFormat(0, output_pdf)  # 0은 PDF 포맷
        workbook.Close(False)
        excel.Quit()
        return output_pdf
    except Exception as e:
        print(f"Failed to convert XLSX to PDF: {xlsx_file} due to {e}")
        return None

# PDF 파일에서 텍스트와 표를 추출하는 함수 (pdfplumber 사용)
def extract_text_from_pdf(pdf_file):
    try:
        print(f"Extracting text from PDF file: {pdf_file}")
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text()  # 텍스트 추출
                tables = page.extract_tables()  # 테이블 추출
                if tables:
                    for table in tables:
                        table_text = "\n".join(["\t".join(row) for row in table])
                        text += "\n" + table_text + "\n"
        return text
    except Exception as e:
        print(f"Failed to extract text from PDF file: {pdf_file} due to {e}")
        return None

# 파일의 형식에 따라 적절한 PDF 변환 후 텍스트 및 표 추출
def extract_text_from_file(file_path):
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"  # 변환된 PDF 경로 설정

    if file_path.endswith('.hwp'):
        pdf_file = convert_hwp_to_pdf_windows(file_path, pdf_path)
        print("추출 중 \n")
        if pdf_file:
            return extract_text_from_pdf(pdf_file)  # 변환된 PDF에서 텍스트와 표 추출

    elif file_path.endswith('.docx'):
        pdf_file = convert_docx_to_pdf_windows(file_path, pdf_path)
        if pdf_file:
            return extract_text_from_pdf(pdf_file)

    elif file_path.endswith('.pptx'):
        pdf_file = convert_pptx_to_pdf_windows(file_path, pdf_path)
        if pdf_file:
            return extract_text_from_pdf(pdf_file)

    elif file_path.endswith('.xlsx'):
        pdf_file = convert_xlsx_to_pdf_windows(file_path, pdf_path)
        if pdf_file:
            return extract_text_from_pdf(pdf_file)

    elif file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    
    else:
        print(f"Unsupported file format: {file_path}")
        return None

# 파일 이름 중복 시 "(1)", "(2)" 추가하는 함수
def make_unique_file_path(file_path):
    base, ext = os.path.splitext(file_path)
    counter = 1
    new_file_path = file_path
    
    # 파일이 존재하면 "(1)", "(2)"를 붙인 새로운 이름 생성
    while os.path.exists(new_file_path):
        new_file_path = f"{base}({counter}){ext}"
        counter += 1
    
    return new_file_path

# 텍스트 파일을 저장하는 함수
def save_extracted_text(output_root, original_file_path, extracted_text):
    relative_path = os.path.relpath(original_file_path, start=original_root)
    new_file_path = os.path.join(output_root, relative_path)
    new_file_path = os.path.splitext(new_file_path)[0] + '_text.txt'
    
    # 파일 이름 중복 처리
    # new_file_path = make_unique_file_path(new_file_path)
    
    new_dir = os.path.dirname(new_file_path)
    os.makedirs(new_dir, exist_ok=True)
    with open(new_file_path, 'w', encoding='utf-8') as f:
        f.write(extracted_text)
    print(f"Saved extracted text to: {new_file_path}")

# 폴더를 순회하며 파일에서 텍스트를 추출하고 저장하는 함수
def traverse_and_extract_and_save(root_folder, output_root):
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            file_path = os.path.join(root, file)
            print(f"Processing file: {file_path}")
            extracted_text = extract_text_from_file(file_path)  # 파일에서 텍스트 추출
            if extracted_text:
                save_extracted_text(output_root, file_path, extracted_text)
            else:
                print(f"No text extracted from: {file_path}")

# 원본 폴더 경로와 새로운 폴더 경로
original_root = os.path.abspath('../data_test/data3_9')  # 절대 경로로 변환
output_root = os.path.abspath('../data_test/data3_9_extract_text')  # 절대 경로로 변환

# 파일을 탐색하며 텍스트를 추출하고 저장
print(f"Starting extraction process for folder: {original_root}")
traverse_and_extract_and_save(original_root, output_root)
print(f"Extraction process completed. Files saved to: {output_root}")
