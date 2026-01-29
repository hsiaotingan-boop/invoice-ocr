from flask import Flask, request, send_file, render_template_string
from PIL import Image
import pytesseract
import pandas as pd
import re
import os
import io
import uuid
from datetime import datetime

app = Flask(__name__)

# =========================
# 漂亮一點的前端頁面
# =========================
HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>發票拍照 → Excel</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial; background:#0b1220; margin:0; color:#e8eefc;}
    .wrap{max-width:900px; margin:40px auto; padding:0 16px;}
    .card{background:#111b32; border:1px solid #1f2b4a; border-radius:16px; padding:22px; box-shadow:0 10px 30px rgba(0,0,0,.25);}
    h1{margin:0 0 6px; font-size:26px;}
    .sub{opacity:.8; margin:0 0 18px; line-height:1.5;}
    .tips{background:#0d1630; border:1px dashed #2b3b66; border-radius:12px; padding:12px 14px; margin:14px 0 18px; font-size:14px; opacity:.9;}
    .row{display:flex; gap:12px; flex-wrap:wrap; align-items:center;}
    input[type=file]{background:#0d1630; border:1px solid #2b3b66; color:#e8eefc; padding:10px; border-radius:12px; width:min(520px, 100%);}
    button{background:#5b8cff; border:none; color:#06112b; padding:12px 16px; border-radius:12px; font-weight:700; cursor:pointer;}
    button:hover{filter:brightness(1.05);}
    .foot{margin-top:14px; font-size:13px; opacity:.7;}
    .badge{display:inline-block; padding:6px 10px; border-radius:999px; background:#0d1630; border:1px solid #2b3b66; font-size:12px; opacity:.9;}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="badge">📸 多張發票一鍵整理</div>
      <h1>發票拍照 → Excel</h1>
      <p class="sub">選擇（或拍照）多張發票，上傳後會自動抓：發票號碼、總金額、稅額，以及「品項 / 數量 / 單價 / 小計」，並產出 Excel。</p>

      <div class="tips">
        小提醒：拍照越清楚越準（光線充足、不要歪、不要糊）。不同店家版型差很大，品項欄位解析若怪怪的也正常，之後我可以幫你再針對版型強化。
      </div>

      <form method="post" enctype="multipart/form-data">
        <div class="row">
          <input type="file" name="photos" accept="image/*" capture="camera" multiple required>
          <button type="submit">上傳並下載 Excel</button>
        </div>
      </form>

      <div class="foot">Render 免費方案第一次開啟可能會慢一點（冷啟動），屬正常現象。</div>
    </div>
  </div>
</body>
</html>
"""

# =========================
# 解析：抓發票號碼/金額/稅額 + 品項表（簡易通用版）
# 由於各店家格式差異很大，這裡用「最常見排列」做抓取：
#   品名  數量  單價  金額
# =========================
def parse_invoice_text(text: str):
    # 基本欄位
    invoice_no = re.search(r"[A-Z]{2}\d{8}", text)
    total = re.search(r"(總計|合計)\s*([0-9]+)", text)
    tax = re.search(r"稅額\s*([0-9]+)", text)

    invoice_no = invoice_no.group() if invoice_no else ""
    total = total.group(2) if total else ""
    tax = tax.group(1) if tax else ""

    # 品項明細（盡量排除總計/合計/稅額等行）
    items = []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # 常見：品名 2 50 100（以空白分隔）
    pat1 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")
    # 有些會是：品名 2x50 100 或 2*50
    pat2 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)[xX\*](\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")

    skip_keywords = ("總計", "合計", "稅額", "小計", "找零", "現金", "信用卡", "電子支付", "應付", "收款", "折扣")

    for ln in lines:
        if any(k in ln for k in skip_keywords):
            continue

        m = pat2.match(ln)
        if not m:
            m = pat1.match(ln)

        if m:
            name = m.group(1).strip()
            qty = m.group(2)
            unit = m.group(3)
            amt = m.group(4)

            # 避免把雜訊當品項：品名太短或全是符號就跳過
            if len(name) < 2:
                continue

            items.append({
                "品項": name,
                "數量": qty,
                "單價": unit,
                "金額": amt
            })

    return {
        "發票號碼": invoice_no,
        "總金額": total,
        "稅額": tax,
        "items": items
    }

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "GET":
        return render_template_string(HTML)

    files = request.files.getlist("photos")
    if not files:
        return "沒有收到檔案，請重新上傳", 400

    invoice_rows = []
    item_rows = []

    for idx, f in enumerate(files, start=1):
        img = Image.open(f.stream)
        text = pytesseract.image_to_string(img, lang="chi_tra")

        parsed = parse_invoice_text(text)

        inv_no = parsed["發票號碼"] or f"(unknown-{idx})"
        invoice_rows.append({
            "序號": idx,
            "發票號碼": parsed["發票號碼"],
            "總金額": parsed["總金額"],
            "稅額": parsed["稅額"]
        })

        for it in parsed["items"]:
            item_rows.append({
                "發票序號": idx,
                "發票號碼": parsed["發票號碼"],
                "品項": it["品項"],
                "數量": it["數量"],
                "單價": it["單價"],
                "金額": it["金額"]
            })

    df_inv = pd.DataFrame(invoice_rows)
    df_items = pd.DataFrame(item_rows)

    # 產生 Excel（用記憶體，不寫入硬碟）
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="invoices", index=False)
        df_items.to_excel(writer, sheet_name="items", index=False)

    output.seek(0)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"invoices_{ts}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

