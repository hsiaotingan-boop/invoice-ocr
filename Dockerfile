FROM python:3.11-slim

# 安裝 OCR 需要的 tesseract + 繁中語言包
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-chi-tra \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

COPY . /app/

# Render 會提供環境變數 PORT
CMD ["python", "app.py"]
