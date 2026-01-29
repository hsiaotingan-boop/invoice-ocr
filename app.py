from flask import Flask, request, send_file, render_template_string
from PIL import Image
import pytesseract
import pandas as pd
import re
import os
import io
from datetime import datetime

app = Flask(__name__)

# =========================
# 漂亮 + 進階多張處理 UI
# - 拍照：一次一張，但可一直加進清單
# - 相簿：一次多張加入清單
# - 最後一鍵產出 Excel
# =========================
HTML = r"""
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
    .sub{opacity:.85; margin:0 0 18px; line-height:1.6;}
    .tips{background:#0d1630; border:1px dashed #2b3b66; border-radius:12px; padding:12px 14px; margin:14px 0 18px; font-size:14px; opacity:.9;}
    .row{display:flex; gap:12px; flex-wrap:wrap; align-items:center;}
    .btn{background:#5b8cff; border:none; color:#06112b; padding:12px 16px; border-radius:12px; font-weight:800; cursor:pointer;}
    .btn.secondary{background:#0d1630; color:#e8eefc; border:1px solid #2b3b66;}
    .btn:disabled{opacity:.5; cursor:not-allowed;}
    input[type=file]{background:#0d1630; border:1px solid #2b3b66; color:#e8eefc; padding:10px; border-radius:12px; width:min(520px, 100%);}
    .foot{margin-top:14px; font-size:13px; opacity:.7;}
    .list{margin-top:16px; background:#0d1630; border:1px solid #2b3b66; border-radius:12px; padding:12px;}
    .list h3{margin:0 0 10px; font-size:14px; opacity:.85;}
    .chips{display:flex; gap:8px; flex-wrap:wrap;}
    .chip{background:#111b32; border:1px solid #2b3b66; border-radius:999px; padding:6px 10px; font-size:12px; opacity:.95;}
    .actions{display:flex; gap:10px; flex-wrap:wrap; margin-top:12px;}
    .badge{display:inline-block; padding:6px 10px; border-radius:999px; background:#0d1630; border:1px solid #2b3b66; font-size:12px; opacity:.9; margin-bottom:10px;}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="badge">📸 拍照可累積 + 🖼️ 相簿可多選</div>
      <h1>發票拍照 → Excel</h1>
      <p class="sub">
        ✅ <b>拍照模式</b>：一次拍一張，但可以一直拍（會累積在下面清單）<br>
        ✅ <b>相簿多選</b>：一次選多張加入清單<br>
        最後按「上傳並下載 Excel」會把清單內全部照片一起處理（Excel 內有兩張表：invoices / items）。
      </p>

      <div class="tips">
        小提醒：拍照越清楚越準（光線充足、不要歪、不要糊）。不同店家版型差很大，品項欄位解析若怪怪的也正常，之後可以再針對常見版型強化。
      </div>

      <form id="uploadForm">
        <!-- A：拍照（iPhone 會一次一張，但可多次加入清單） -->
        <div class="row">
          <input id="cameraInput" type="file" accept="image/*" capture="environment">
          <button type="button" class="btn secondary" id="addCameraBtn">📸 拍照加入清單</button>
        </div>

        <div style="height:10px"></div>

        <!-- B：相簿多選（一次多張） -->
        <div class="row">
          <input id="galleryInput" type="file" accept="image/*" multiple>
          <button type="button" class="btn secondary" id="addGalleryBtn">🖼️ 相簿多選加入清單</button>
        </div>

        <div class="list">
          <h3>已加入清單：<span id="count">0</span> 張</h3>
          <div class="chips" id="chips"></div>

          <div class="actions">
            <button type="submit" class="btn" id="submitBtn" disabled>⬇️ 上傳並下載 Excel</button>
            <button type="button" class="btn secondary" id="clearBtn">🧹 清空清單</button>
          </div>
        </div>

        <div class="foot">
          Render 免費方案第一次開啟可能會慢一點（冷啟動），屬正常現象。
        </div>
      </form>
    </div>
  </div>

<script>
  const cameraInput = document.getElementById('cameraInput');
  const galleryInput = document.getElementById('galleryInput');
  const chips = document.getElementById('chips');
  const countEl = document.getElementById('count');
  const submitBtn = document.getElementById('submitBtn');
  const clearBtn = document.getElementById('clearBtn');
  const addCameraBtn = document.getElementById('addCameraBtn');
  const addGalleryBtn = document.getElementById('addGalleryBtn');
  const form = document.getElementById('uploadForm');

  // DataTransfer：用來累積多次選取/拍照的檔案
  const dt = new DataTransfer();

  function refreshUI() {
    chips.innerHTML = '';
    for (const f of dt.files) {
      const div = document.createElement('div');
      div.className = 'chip';
      div.textContent = f.name || 'photo';
      chips.appendChild(div);
    }
    countEl.textContent = dt.files.length;
    submitBtn.disabled = dt.files.length === 0;
  }

  function addFiles(fileList) {
    for (const f of fileList) dt.items.add(f);
    refreshUI();
  }

  addCameraBtn.addEventListener('click', () => {
    if (cameraInput.files && cameraInput.files.length > 0) {
      addFiles(cameraInput.files);
      cameraInput.value = ""; // 讓下次還能再拍/再選同一張
    } else {
      alert("請先拍一張照片（或選一張）");
    }
  });

  addGalleryBtn.addEventListener('click', () => {
    if (galleryInput.files && galleryInput.files.length > 0) {
      addFiles(galleryInput.files);
      galleryInput.value = "";
    } else {
      alert("請先從相簿選照片（可多選）");
    }
  });

  clearBtn.addEventListener('click', () => {
    while (dt.items.length) dt.items.remove(0);
    refreshUI();
  });

  // 送出：用 fetch 上傳檔案，拿回 blob 直接下載
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    if (dt.files.length === 0) {
      alert("清單是空的，請先加入照片");
      return;
    }

    const formData = new FormData();
    for (const f of dt.files) formData.append('photos', f);

    submitBtn.disabled = true;
    const oldText = submitBtn.textContent;
    submitBtn.textContent = "處理中…（可能需要一點時間）";

    try {
      const res = await fetch('/upload', { method: 'POST', body: formData });
      if (!res.ok) {
        const t = await res.text();
        throw new Error(t || '上傳失敗');
      }

      const blob = await res.blob();

      // 嘗試從 header 拿檔名（若拿不到就用預設）
      let filename = "invoices.xlsx";
      const cd = res.headers.get('Content-Disposition');
      if (cd) {
        const m = /filename="([^"]+)"/.exec(cd);
        if (m && m[1]) filename = m[1];
      }

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      submitBtn.textContent = oldText;
      submitBtn.disabled = dt.files.length === 0;
    } catch (err) {
      alert("錯誤：" + err.message);
      submitBtn.textContent = oldText;
      submitBtn.disabled = dt.files.length === 0;
    }
  });

  refreshUI();
</script>
</body>
</html>
"""

# =========================
# OCR + 解析（通用簡易版）
# 1) 發票號碼：AB12345678
# 2) 總計/合計：抓數字
# 3) 稅額：抓數字
# 4) 品項（簡易規則）：常見 "品名  數量  單價  金額"
# =========================
def parse_invoice_text(text: str):
    invoice_no_m = re.search(r"[A-Z]{2}\d{8}", text)
    total_m = re.search(r"(總計|合計)\s*([0-9]+)", text)
    tax_m = re.search(r"稅額\s*([0-9]+)", text)

    invoice_no = invoice_no_m.group() if invoice_no_m else ""
    total = total_m.group(2) if total_m else ""
    tax = tax_m.group(1) if tax_m else ""

    items = []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # 品名 2 50 100
    pat1 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")
    # 品名 2x50 100 或 2*50
    pat2 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)[xX\*](\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")

    skip_keywords = ("總計", "合計", "稅額", "小計", "找零", "現金", "信用卡", "電子支付", "應付", "收款", "折扣", "發票", "統編")

    for ln in lines:
        if any(k in ln for k in skip_keywords):
            continue

        m = pat2.match(ln) or pat1.match(ln)
        if not m:
            continue

        name = m.group(1).strip()
        qty = m.group(2)
        unit = m.group(3)
        amt = m.group(4)

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


@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML)


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("photos")
    if not files:
        return "沒有收到檔案（photos），請重新上傳", 400

    invoice_rows = []
    item_rows = []

    for idx, f in enumerate(files, start=1):
        try:
            img = Image.open(f.stream)
        except Exception:
            return f"第 {idx} 張圖片無法讀取，請換一張更清楚的照片", 400

        text = pytesseract.image_to_string(img, lang="chi_tra")
        parsed = parse_invoice_text(text)

        invoice_rows.append({
            "序號": idx,
            "發票號碼": parsed["發票號碼"],
            "總金額": parsed["總金額"],
            "稅額": parsed["稅額"],
        })

        for it in parsed["items"]:
            item_rows.append({
                "發票序號": idx,
                "發票號碼": parsed["發票號碼"],
                "品項": it["品項"],
                "數量": it["數量"],
                "單價": it["單價"],
                "金額": it["金額"],
            })

    df_inv = pd.DataFrame(invoice_rows)
    df_items = pd.DataFrame(item_rows)

    # 產生 Excel（記憶體，不落地）
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="invoices", index=False)
        df_items.to_excel(writer, sheet_name="items", index=False)
    output.seek(0)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"invoices_{ts}.xlsx"

    # 讓瀏覽器知道這是 Excel
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

