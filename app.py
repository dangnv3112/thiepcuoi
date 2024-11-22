from flask import Flask, redirect, request, jsonify
import qrcode
from io import BytesIO
from PIL import Image

app = Flask(__name__)

# URL mặc định
current_url = "https://weddingday.com.vn/mau-thiep/thiep-cuoi-hien-dai-i01/"

@app.route('/')
def home():
    return jsonify({
        "message": "Welcome to the Dynamic QR Code Generator!",
        "current_url": current_url,
        "instructions": {
            "/qrcode": "Generate a QR code pointing to the current URL.",
            "/update?url=new_url": "Update the current URL.",
        }
    })

@app.route('/update')
def update_url():
    """
    Cập nhật URL đích.
    """
    global current_url
    new_url = request.args.get('url')
    if new_url:
        current_url = new_url
        return jsonify({"message": "URL updated successfully!", "new_url": current_url})
    return jsonify({"error": "No URL provided. Use ?url=new_url"}), 400

@app.route('/redirect')
def dynamic_redirect():
    """
    Chuyển hướng đến URL hiện tại.
    """
    return redirect(current_url)

@app.route('/qrcode')
def generate_qr():
    """
    Tạo mã QR.
    """
    qr_url = request.url_root + "redirect"  # URL của endpoint chuyển hướng
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_url)
    qr.make(fit=True)

    # Lưu QR code vào bộ nhớ
    img = qr.make_image(fill_color="black", back_color="white")
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)

    # Trả ảnh QR về trình duyệt
    return app.response_class(buffer, mimetype='image/png')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
