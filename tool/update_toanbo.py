import os

def replace_words_in_file(input_file, output_file, replacements):
    """
    Đọc tệp văn bản, thay thế các từ dựa trên danh sách thay thế và lưu kết quả vào tệp mới.

    :param input_file: Đường dẫn đến tệp đầu vào (TXT)
    :param output_file: Đường dẫn đến tệp đầu ra (TXT)
    :param replacements: Từ điển {từ_cũ: từ_mới} để thay thế
    """
    try:
        # Đọc nội dung từ tệp đầu vào
        with open(input_file, 'r', encoding='utf-8') as file:
            content = file.read()

        # Thay thế các từ dựa trên từ điển replacements
        for old_word, new_word in replacements.items():
            content = content.replace(old_word, new_word)

        # Ghi nội dung đã chỉnh sửa vào tệp đầu ra
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(content)

        print(f"Đã thay thế và lưu kết quả vào: {output_file}")
    except FileNotFoundError:
        print(f"Không tìm thấy tệp: {input_file}")
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")


def replace_words_in_all_files(input_dir, output_dir, replacements):
    """
    Lặp qua tất cả các tệp trong thư mục, thay thế các từ và lưu kết quả vào thư mục đầu ra.

    :param input_dir: Thư mục chứa các tệp đầu vào
    :param output_dir: Thư mục lưu các tệp đã thay thế
    :param replacements: Từ điển {từ_cũ: từ_mới} để thay thế
    """
    try:
        # Lặp qua tất cả các tệp trong thư mục đầu vào
        for filename in os.listdir(input_dir):
            input_file_path = os.path.join(input_dir, filename)
            
            # Chỉ thay thế trong các tệp văn bản (txt)
            if filename.endswith('.txt'):
                # Đường dẫn tệp đầu ra
                output_file_path = os.path.join(output_dir, filename)

                # Gọi hàm để thay thế từ
                replace_words_in_file(input_file_path, output_file_path, replacements)

    except Exception as e:
        print(f"Đã xảy ra lỗi khi thay thế trong thư mục: {e}")


# Sử dụng chương trình
if __name__ == "__main__":
    # Thư mục chứa các tệp đầu vào và đầu ra
    input_directory = "C:/thiepcuoi/updated_files/"  # Thư mục chứa các tệp văn bản gốc
    output_directory = "C:/thiepcuoi/updated_files_new/"  # Thư mục lưu các tệp đã thay thế

    # Tạo thư mục đầu ra nếu chưa tồn tại
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Danh sách thay thế: {từ_cũ: từ_mới}
    replacements_dict = {
        f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E5%25BE.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02470.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_012.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02500.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_002.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02567.jpg",
f"thiepcuoionline_files/-%25F0%259F%2592%2599_5_%25E6%2591%2584%25E5%25BD%25B1%25E5%25B8%2588%25F0%259F%2593%25B8%25E6%259.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02726.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_016.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02748.jpg",
f"thiepcuoionline_files/-%25F0%259F%2592%2599_3_%25E6%2591%2584%25E5%25BD%25B1%25E5%25B8%2588%25F0%259F%2593%25B8%25E6%259.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02772.jpg",
f"thiepcuoionline_files/-%25F0%259F%2592%2599_11_%25E6%2591%2584%25E5%25BD%25B1%25E5%25B8%2588%25F0%259F%2593%25B8%25E6%25.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02796.jpg",
f"thiepcuoionline_files/-%25F0%259F%2592%2599_10_%25E6%2591%2584%25E5%25BD%25B1%25E5%25B8%2588%25F0%259F%2593%25B8%25E6%25.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03119.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_011.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02814.jpg",
f"thiepcuoionline_files/-%25F0%259F%2592%2599_5_%25E6%2591%2584%25E5%25BD%25B1%25E5%25B8%2588%25F0%259F%2593%25B8%25_002.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02843.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_009.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02881.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_006.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02956.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_014.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03074.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_005.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03163.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_004.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03215.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_003.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03258.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_008.jpg":"C:/thiepcuoi/fileanhcuoi/H2H03377.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_010.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02500.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_017.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02726.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_015.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02764.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_013.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02796.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_007.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02814.jpg",
f"thiepcuoionline_files/%25E6%258D%25AE%25E8%25AF%25B4%25E8%25BF%2599%25E7%25A7%258D%25E9%25A3%258E%25E6%25A0%25BC%25E_018.jpg":"C:/thiepcuoi/fileanhcuoi/H2H02881.jpg"}

    # Gọi hàm để thay thế trong tất cả các tệp văn bản
    replace_words_in_all_files(input_directory, output_directory, replacements_dict)
