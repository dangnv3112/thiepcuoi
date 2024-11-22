import os
import urllib.parse

# Thư mục chứa tệp
directory = "C:/thiepcuoi/thiepcuoionline_files/"

# Lặp qua các tệp trong thư mục
for filename in os.listdir(directory):
    # Giải mã tên tệp
    decoded_filename = urllib.parse.unquote(filename)

    # Đường dẫn cũ và mới
    old_path = os.path.join(directory, filename)
    new_path = os.path.join(directory, decoded_filename)

    # Đổi tên tệp
    try:
        os.rename(old_path, new_path)
        print(f"Đã đổi tên: {filename} -> {decoded_filename}")
    except Exception as e:
        print(f"Lỗi khi đổi tên {filename}: {e}")
