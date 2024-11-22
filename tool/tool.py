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

# Sử dụng chương trình
if __name__ == "__main__":
    # Đường dẫn tệp đầu vào và đầu ra
    input_file_path = "C:/thiepcuoi/thiepcuoi_new.txt"  # Tệp văn bản gốc
    output_file_path = "C:/thiepcuoi/thiepcuoi_new_2.txt"  # Tệp sau khi thay thế

    # Danh sách thay thế: {từ_cũ: từ_mới}
    replacements_dict = {f"Thanh Phong":"Vê Đăng",
f"Như Quỳnh":"Minh Thư",
f">24<":">15<",
f"Thứ ba":"CN",
f"11:00":"11:30",
f"10/09/1997":"31/12/2000",
f"02/09/2000":"H2H02796.jpg",
f"09/2012":"08/2012",
f"Đại học Sư phạm":"Học Viện Hàng Không"}

    # Gọi hàm để thay thế
    replace_words_in_file(input_file_path, output_file_path, replacements_dict)
