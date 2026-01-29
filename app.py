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

# =========================
# 前端：漂亮 UI + 兩種加入方式 + 進度顯示
# 送出流程：
# 1) POST /start 取得 job_id
# 2) 前端輪詢 GET /progress/<job_id> 顯示「正在處理第 X/N」
# 3) 完成後下載 GET /download/<job_id>
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
      <div class="badge">📸 拍照可累積 + 🖼️ 相簿可多選 + ⏱️ 進度顯示</div>
      <h1>發票拍照 → Excel</h1>
      <p class="sub">
        ✅ <b>拍照模式</b>：一次拍一張，但可以一直拍（會累積在下面清單）<br>
        ✅ <b>相簿多選</b>：一次選多張加入清單<br>
        ✅ 下載 Excel 會包含：<b>invoices（摘要）</b> / <b>items（品項明細）</b><br>
        ✅ 系統會自動縮圖 + 增強對比（通常更快也更準）
      </p>

      <div class="tips">
        小提醒：拍照越清楚越準（光線充足、不要歪、不要糊）。不同店家版型差很大，品項欄位解析若怪怪的也正常，之後可以再針對常見版型強化。
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

  // DataTransfer：累積多次拍照/選取
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
      if (!res.ok) throw new Error("讀取進度失敗");
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

  // 送出：先 /start，拿 job_id，再輪詢，再下載
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

      // 1) start
      const startRes = await fetch('/start', { method: 'POST', body: formData });
      if (!startRes.ok) {
        const t = await startRes.text();
        throw new Error(t || "啟動任務失敗");
      }
      const startData = await startRes.json();
      const jobId = startData.job_id;

      // 2) poll
      submitBtn.textContent = "處理中…";
      await pollProgress(jobId);

      // 3) download
      const a = document.createElement('a');
      a.href = `/download/${jobId}`;
      a.click();

      submitBtn.textContent = oldText;
      submitBtn.disabled = dt.files.length === 0;
    } catch (err) {
      alert("錯誤：" + err.message);
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

# =========================
# 記憶體中的工作狀態（短時間用，Render free OK）
# jobs[job_id] = {
#   status: "processing"|"done"|"error",
#   current: int,
#   total: int,
#   message: str,
#   error: str,
#   excel_bytes: bytes,
#   filename: str
# }
# =========================
jobs = {}
jobs_lock = threading.Lock()


# =========================
# OCR 前預處理：縮圖 + 對比 + 灰階（加速 & 更穩）
# - max_width: 最大寬度（太大會慢）
# - autocontrast + contrast 增強 + 略銳化
# =========================
def preprocess_image(img: Image.Image, max_width: int = 1600) -> Image.Image:
    # 轉成 RGB 避免某些模式出錯
    img = img.convert("RGB")

    # 縮圖（只在太大時縮）
    w, h = img.size
    if w > max_width:
        new_h = int(h * (max_width / w))
        img = img.resize((max_width, new_h), Image.LANCZOS)

    # 灰階
    img = ImageOps.grayscale(img)

    # 自動拉對比（去霧）
    img = ImageOps.autocontrast(img)

    # 再加一點對比
    img = ImageEnhance.Contrast(img).enhance(1.6)

    # 輕微銳化
    img = img.filter(ImageFilter.SHARPEN)

    return img


# =========================
# 解析：基本欄位 + 品項表（簡易通用版）
# 品項抓法：常見 "品名  數量  單價  金額"
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

    pat1 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")
    pat2 = re.compile(r"^(.+?)\s+(\d+(?:\.\d+)?)[xX\*](\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)$")

    skip_keywords = ("總計", "合計", "稅額", "小計", "找零", "現金", "信用卡", "電子支付",
                     "應付", "收款", "折扣", "發票", "統編", "載具", "交易", "日期", "時間")

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


def build_excel_bytes(invoice_rows, item_rows):
    df_inv = pd.DataFrame(invoice_rows)
    df_items = pd.DataFrame(item_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="invoices", index=False)
        df_items.to_excel(writer, sheet_name="items", index=False)
    output.seek(0)
    return output.getvalue()


def worker_process(job_id: str, images_bytes_list):
    total = len(images_bytes_list)
    invoice_rows = []
    item_rows = []

    try:
        for i, img_bytes in enumerate(images_bytes_list, start=1):
            with jobs_lock:
                jobs[job_id]["status"] = "processing"
                jobs[job_id]["current"] = i
                jobs[job_id]["total"] = total
                jobs[job_id]["message"] = f"正在處理第 {i} / {total} 張…（OCR 辨識中）"

            img = Image.open(io.BytesIO(img_bytes))
            img = preprocess_image(img)

            text = pytesseract.image_to_string(img, lang="chi_tra")
            parsed = parse_invoice_text(text)

            invoice_rows.append({
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

        with jobs_lock:
            jobs[job_id]["message"] = "正在產生 Excel…"

        excel_bytes = build_excel_bytes(invoice_rows, item_rows)
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

    # 把檔案讀成 bytes，避免 thread 裡面讀 stream 出問題
    images_bytes_list = []
    for f in files:
        images_bytes_list.append(f.read())

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

    # 下載後把 job 清掉（避免記憶體累積）
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
