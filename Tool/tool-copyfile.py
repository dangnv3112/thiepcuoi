import openpyxl
import shutil
import os

def copy_excel_to_multiple_files(source_file_path, sheet_name, output_names):
    # Kiểm tra file gốc có tồn tại hay không
    if not os.path.isfile(source_file_path):
        print(f"File '{source_file_path}' không tồn tại.")
        return
    
    print(f"Đang mở file: {source_file_path}")
    
    # Mở file Excel gốc
    wb = openpyxl.load_workbook(source_file_path)
    print(f"Đã mở file: {source_file_path}")
    
    # Kiểm tra sheet có tồn tại hay không
    if sheet_name not in wb.sheetnames:
        print(f"Sheet '{sheet_name}' không tồn tại trong file '{source_file_path}'.")
        return
    
    # Lặp qua danh sách các tên file đầu ra và tạo các file tương ứng
    for output_name in output_names:
        new_file_path = os.path.join(os.path.dirname(source_file_path), f'{output_name}')
        shutil.copyfile(source_file_path, new_file_path)
        print(f"Đã tạo file: {new_file_path}")

# Đường dẫn tới file Excel gốc và tên sheet cần copy
source_file_path = 'D:/2024-2025/NangCapOLT_Dasan/CHECKLIST_tem_new.xlsx'  # Đường dẫn tuyệt đối
sheet_name = 'Checklist'

# Danh sách các tên file bạn muốn tạo
output_names = ["D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_1.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_2.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_3.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_4.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_5.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_6.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_7.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_8.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_9.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_10.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_11.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_12.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_13.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_14.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_15.xlsx",
"D:/2024-2025/NangCapOLT_Dasan/Checklist_CR_16.xlsx"

]  # Tùy chọn tên file

copy_excel_to_multiple_files(source_file_path, sheet_name, output_names)
