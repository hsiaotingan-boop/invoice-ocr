from flask import Flask, request, send_file, render_template_string
from PIL import Image
import pytesseract
import pandas as pd
import re
import os

app = Flask(__name__)

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
        file = request.files["photo"]
        img = Image.open(file.stream)

        text = pytesseract.image_to_string(img, lang="chi_tra")

        invoice_no = re.search(r"[A-Z]{2}\d{8}", text)
        total = re.search(r"(總計|合計)\s*([0-9]+)", text)
        tax = re.search(r"稅額\s*([0-9]+)", text)

        data = {
            "發票號碼": [invoice_no.group() if invoice_no else ""],
            "總金額": [total.group(2) if total else ""],
            "稅額": [tax.group(1) if tax else ""]
        }

        df = pd.DataFrame(data)
        output = "invoice.xlsx"
        df.to_excel(output, index=False)

        return send_file(output, as_attachment=True)

    return render_template_string(HTML)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


