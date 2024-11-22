import os
import re
from openpyxl import Workbook

def extract_jpg_paths_from_txt(txt_file):
    """Trích xuất các đường dẫn tệp .jpg từ một file txt"""
    jpg_paths = []
    with open(txt_file, 'r', encoding='utf-8') as file:
        content = file.read()
        # Biểu thức chính quy để tìm các đường dẫn có đuôi .jpg
        jpg_paths = re.findall(r'[A-Za-z0-9_./\\-]+\.jpg', content)
    return jpg_paths

def save_to_excel(jpg_paths, output_file):
    """Lưu các đường dẫn tệp .jpg vào file Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "JPG Files"
    
    # Thêm tiêu đề
    ws.append(["File Path"])
    
    # Thêm dữ liệu
    for path in jpg_paths:
        ws.append([path])
    
    # Lưu file Excel
    wb.save(output_file)

def scan_folder_for_txt_files(folder_path):
    """Quét thư mục và tìm tất cả các file .txt"""
    txt_files = []
    # Duyệt qua các file trong thư mục
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".txt"):
                txt_files.append(os.path.join(root, file))
    return txt_files

if __name__ == "__main__":
    # Nhập đường dẫn thư mục cần quét
    folder_path = input("C:/thiepcuoi/thiepcuoionline_files_backup_2/")
    output_file = "jpg_paths.xlsx"  # Tên file Excel xuất ra

    # Lấy danh sách các file .txt trong thư mục
    txt_files = scan_folder_for_txt_files(folder_path)
    
    # Danh sách để lưu tất cả các đường dẫn .jpg
    all_jpg_paths = []

    # Quét từng file .txt và trích xuất các đường dẫn .jpg
    for txt_file in txt_files:
        jpg_paths = extract_jpg_paths_from_txt(txt_file)
        all_jpg_paths.extend(jpg_paths)

    # Lưu các đường dẫn vào file Excel
    save_to_excel(all_jpg_paths, output_file)
    
    print(f"Danh sách các path .jpg đã được lưu vào file {output_file}.")
