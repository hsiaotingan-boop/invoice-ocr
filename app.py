from flask import Flask, request, send_file, render_template_string, jsonify
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import pytesseract
import pandas as pd
import re
import os
import io
import threading
import uuid
from datetime import datetime

app = Flask(__name__)

HTML = r"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>發票拍照 → Excel</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial; background:#0b1220; margin:0; color:#e8eefc;}
    .wrap{max-width:920px; margin:40px auto; padding:0 16px;}
    .card{background:#111b32; border:1px solid #1f2b4a; border-radius:18px; padding:22px; box-shadow:0 10px 30px rgba(0,0,0,.25);}
    h1{margin:0 0 6px; font-size:26px;}
    .sub{opacity:.85; margin:0 0 18px; line-height:1.6;}
    .tips{background:#0d1630; border:1px dashed #2b3b66; border-radius:12px; padding:12px 14px; margin:14px 0 18px; font-size:14px; opacity:.9;}
    .row{display:flex; gap:12px; flex-wrap:wrap; align-items:center;}
    .btn{background:#5b8cff; border:none; color:#06112b; padding:12px 16px; border-radius:12px; font-weight:800; cursor:pointer;}
    .btn.secondary{background:#0d1630; color:#e8eefc; border:1px solid #2b3b66;}
    .btn:disabled{opacity:.5; cursor:not-allowed;}
    input[type=file]{background:#0d1630; border:1px solid #2b3b66; color:#e8eefc; padding:10px; border-radius:12px; width:min(560px, 100%);}
    .foot{margin-top:14px; font-size:13px; opacity:.7;}
    .list{margin-top:16px; background:#0d1630; border:1px solid #2b3b66; border-radius:12px; padding:12px;}
    .list h3{margin:0 0 10px; font-size:14px; opacity:.85;}
    .chips{display:flex; gap:8px; flex-wrap:wrap;}
    .chip{background:#111b32; border:1px solid #2b3b66; border-radius:999px; padding:6px 10px; font-size:12px; opacity:.95;}
    .actions{display:flex; gap:10px; flex-wrap:wrap; margin-top:12px;}
    .badge{display:inline-block; padding:6px 10px; border-radius:999px; background:#0d1630; border:1px solid #2b3b66; font-size:12px; opacity:.9; margin-bottom:10px;}
    .status{margin-top:10px; font-size:14px; opacity:.9;}
    .barWrap{margin-top:10px; width:100%; height:10px; background:#0b1220; border:1px solid #2b3b66; border-radius:999px; overflow:hidden;}
    .bar{height:100%; width:0%; background:#5b8cff;}
    .small{font-size:12px; opacity:.75; margin-top:6px;}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="badge">📸 拍照可累積 + 🖼️ 相簿可多選 + ⏱️ 進度顯示 + 🧾 OCR Debug</div>
      <h1>發票拍照 → Excel</h1>
      <p class="sub">
        ✅ 拍照一次一張，但可以一直加進清單<br>
        ✅ 相簿可一次多選多張<br>
        ✅ 進度會顯示「正在處理第 X / N 張…」<br>
        ✅ Excel 會多一張表 <b>ocr_text</b>（讓你看 OCR 到底讀到什麼）
      </p>

      <div class="tips">
        小提醒：反光/糊/歪都會讓 OCR 讀不到。盡量正、清楚、光線足。<br>
        如果 invoices 空白，請看 Excel 的 ocr_text 表，就知道 OCR 有沒有讀到關鍵字。
      </div>

      <form id="uploadForm">
        <div class="row">
          <input id="cameraInput" type="file" accept="image/*" capture="environment">
          <button type="button" class="btn secondary" id="addCameraBtn">📸 拍照加入清單</button>
        </div>

        <div style="height:10px"></div>

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

          <div class="status" id="status"></div>
          <div class="barWrap" id="barWrap" style="display:none;">
            <div class="bar" id="bar"></div>
          </div>
          <div class="small" id="small"></div>
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

  const statusEl = document.getElementById('status');
  const barWrap = document.getElementById('barWrap');
  const bar = document.getElementById('bar');
  const smallEl = document.getElementById('small');

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
      cameraInput.value = "";
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
    statusEl.textContent = "";
    smallEl.textContent = "";
    barWrap.style.display = "none";
    bar.style.width = "0%";
  });

  function setProgress(current, total, msg) {
    statusEl.textContent = msg || "";
    if (total && total > 0) {
      barWrap.style.display = "block";
      const pct = Math.floor((current / total) * 100);
      bar.style.width = pct + "%";
      smallEl.textContent = `進度：${current}/${total}（${pct}%）`;
    }
  }

  async function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }

  async function pollProgress(jobId) {
    while (true) {
      const res = await fetch(`/progress/${jobId}`);
      if (!res.ok) {
        const t = await res.text();
        throw new Error(t || "讀取進度失敗");
      }
      const data = await res.json();

      if (data.status === "processing") {
        setProgress(data.current, data.total, data.message);
      } else if (data.status === "done") {
        setProgress(data.total, data.total, "✅ 完成！準備下載 Excel…");
        return;
      } else if (data.status === "error") {
        throw new Error(data.error || "處理失敗");
      }
      await sleep(500);
    }
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    if (dt.files.length === 0) {
      alert("清單是空的，請先加入照片");
      return;
    }

    submitBtn.disabled = true;
    const oldText = submitBtn.textContent;
    submitBtn.textContent = "上傳中…";

    barWrap.style.display = "block";
    bar.style.width = "0%";
    statusEl.textContent = "上傳中…";
    smallEl.textContent = "";

    try {
      const formData = new FormData();
      for (const f of dt.files) formData.append('photos', f);

      const startRes = await fetch('/start', { method: 'POST', body: formData });
      if (!startRes.ok) {
        const t = await startRes.text();
        throw new Error("錯誤：" + t);
      }
      const startData = await startRes.json();
      const jobId = startData.job_id;

      submitBtn.textContent = "處理中…";
      await pollProgress(jobId);

      // iOS Safari 有時候不喜歡程式 click 下載，改成直接導向下載
      window.location.href = `/download/${jobId}`;

      submitBtn.textContent = oldText;
      submitBtn.disabled = dt.files.length === 0;
    } catch (err) {
      alert(err.message);
      submitBtn.textContent = oldText;
      submitBtn.disabled = dt.files.length === 0;
      statusEl.textContent = "❌ 發生錯誤";
    }
  });

  refreshUI();
</script>
</body>
</html>
"""

jobs = {}
jobs_lock = threading.Lock()

def preprocess_image(img: Image.Image, max_width: int = 1600) -> Image.Image:
    img = img.convert("RGB")
    w, h = img.size
    if w > max_width:
        new_h = int(h * (max_width / w))
        img = img.resize((max_width, new_h), Image.LANCZOS)

    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Contrast(img).enhance(1.8)
    img = ImageEnhance.Sharpness(img).enhance(1.3)

    # 二值化（加速 + 更像黑白掃描）
    threshold = 150
    img = img.point(lambda x: 255 if x > threshold else 0)

    # 輕微銳化
    img = img.filter(ImageFilter.SHARPEN)
    return img

def normalize_money(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")

def parse_invoice_text(text: str):
    # 允許 AB 12 345678 或 AB-12345678
    invoice_no_m = re.search(r"([A-Z]{2})\s*[-]?\s*(\d{8})", text)
    invoice_no = ""
    if invoice_no_m:
        invoice_no = invoice_no_m.group(1) + invoice_no_m.group(2)

    # 金額允許逗號
    total_m = re.search(r"(總計|合計)\s*[:：]?\s*([0-9,]+)", text)
    tax_m = re.search(r"稅額\s*[:：]?\s*([0-9,]+)", text)

    total = normalize_money(total_m.group(2)) if total_m else ""
    tax = normalize_money(tax_m.group(1)) if tax_m else ""

    items = []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    pat1 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")
    pat2 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)[xX\*](\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")

    skip_keywords = ("總計","合計","稅額","小計","找零","現金","信用卡","電子支付",
                     "應付","收款","折扣","發票","統編","載具","交易","日期","時間")

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

        items.append({"品項": name, "數量": qty, "單價": unit, "金額": amt})

    return {"發票號碼": invoice_no, "總金額": total, "稅額": tax, "items": items}

def build_excel_bytes(inv_rows, item_rows, ocr_rows):
    df_inv = pd.DataFrame(inv_rows)
    df_items = pd.DataFrame(item_rows)
    df_ocr = pd.DataFrame(ocr_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="invoices", index=False)
        df_items.to_excel(writer, sheet_name="items", index=False)
        df_ocr.to_excel(writer, sheet_name="ocr_text", index=False)
    output.seek(0)
    return output.getvalue()

def worker_process(job_id: str, images_bytes_list):
    total = len(images_bytes_list)
    inv_rows, item_rows, ocr_rows = [], [], []

    try:
        for i, img_bytes in enumerate(images_bytes_list, start=1):
            with jobs_lock:
                jobs[job_id]["status"] = "processing"
                jobs[job_id]["current"] = i
                jobs[job_id]["total"] = total
                jobs[job_id]["message"] = f"正在處理第 {i} / {total} 張…（OCR 辨識中）"

            img = Image.open(io.BytesIO(img_bytes))
            img = preprocess_image(img)

            # OCR：中+英（發票號碼常需要英文）
            config = "--oem 3 --psm 6"
            text = pytesseract.image_to_string(img, lang="chi_tra+eng", config=config)

            parsed = parse_invoice_text(text)

            inv_rows.append({
                "序號": i,
                "發票號碼": parsed["發票號碼"],
                "總金額": parsed["總金額"],
                "稅額": parsed["稅額"],
            })

            for it in parsed["items"]:
                item_rows.append({
                    "發票序號": i,
                    "發票號碼": parsed["發票號碼"],
                    "品項": it["品項"],
                    "數量": it["數量"],
                    "單價": it["單價"],
                    "金額": it["金額"],
                })

            ocr_rows.append({
                "序號": i,
                "發票號碼(解析結果)": parsed["發票號碼"],
                "OCR文字(前2000字)": (text[:2000] if text else "")
            })

        with jobs_lock:
            jobs[job_id]["message"] = "正在產生 Excel…"

        excel_bytes = build_excel_bytes(inv_rows, item_rows, ocr_rows)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"invoices_{ts}.xlsx"

        with jobs_lock:
            jobs[job_id]["status"] = "done"
            jobs[job_id]["excel_bytes"] = excel_bytes
            jobs[job_id]["filename"] = filename
            jobs[job_id]["current"] = total
            jobs[job_id]["total"] = total
            jobs[job_id]["message"] = "✅ 完成"

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["error"] = str(e)
            jobs[job_id]["message"] = "❌ 發生錯誤"

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML)

@app.route("/start", methods=["POST"])
def start():
    files = request.files.getlist("photos")
    if not files:
        return "沒有收到檔案（photos），請重新上傳", 400

    images_bytes_list = [f.read() for f in files]

    job_id = uuid.uuid4().hex
    with jobs_lock:
        jobs[job_id] = {
            "status": "processing",
            "current": 0,
            "total": len(images_bytes_list),
            "message": "任務已建立，準備開始…",
            "error": "",
            "excel_bytes": b"",
            "filename": ""
        }

    t = threading.Thread(target=worker_process, args=(job_id, images_bytes_list), daemon=True)
    t.start()

    return jsonify({"job_id": job_id})

@app.route("/progress/<job_id>", methods=["GET"])
def progress(job_id):
    with jobs_lock:
        job = jobs.get(job_id)

    if not job:
        return jsonify({"status": "error", "error": "找不到任務（可能已過期）"}), 404

    if job["status"] == "error":
        return jsonify({"status": "error", "error": job.get("error", "未知錯誤")})

    if job["status"] == "done":
        return jsonify({
            "status": "done",
            "current": job.get("current", 0),
            "total": job.get("total", 0),
            "message": job.get("message", "")
        })

    return jsonify({
        "status": "processing",
        "current": job.get("current", 0),
        "total": job.get("total", 0),
        "message": job.get("message", "")
    })

@app.route("/download/<job_id>", methods=["GET"])
def download(job_id):
    with jobs_lock:
        job = jobs.get(job_id)

    if not job or job.get("status") != "done" or not job.get("excel_bytes"):
        return "檔案尚未準備好，請稍後再試", 400

    excel_bytes = job["excel_bytes"]
    filename = job["filename"] or "invoices.xlsx"

    with jobs_lock:
        jobs.pop(job_id, None)

    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
