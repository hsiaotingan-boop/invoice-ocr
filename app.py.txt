from flask import Flask, request, send_file, render_template_string
from PIL import Image
import pytesseract
import pandas as pd
import re

# ⚠️ 如果你是 Windows，請把下面這行打開，並確認路徑正確
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

app = Flask(__name__)

# 簡單的網頁畫面
HTML = """
<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <title>發票拍照轉 Excel</title>
</head>
<body style="text-align:center; font-family:Arial; margin-top:50px;">
    <h2>📸 發票拍照 → Excel</h2>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="photo" accept="image/*" capture="camera" required>
        <br><br>
        <button type="submit">上傳並產生 Excel</button>
    </form>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        # 取得上傳的照片
        file = request.files["photo"]
        img = Image.open(file.stream)

        # OCR 讀取文字
        text = pytesseract.image_to_string(img, lang="chi_tra")

        # 用簡單方式抓資料
        invoice_no = re.search(r"[A-Z]{2}\d{8}", text)
        total = re.search(r"(總計|合計)\s*([0-9]+)", text)
        tax = re.search(r"稅額\s*([0-9]+)", text)

        # 整理成表格
        data = {
            "發票號碼":
