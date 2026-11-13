FROM python:3.12-slim

# Tesseract OCR (imagenes y PDF escaneados) con idioma español e inglés
RUN apt-get update && apt-get install -y --no-install-recommends \
        tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render/Railway/etc. definen PORT; por defecto 8001
ENV PORT=8001
CMD ["sh", "-c", "uvicorn main:app --app-dir backend --host 0.0.0.0 --port ${PORT}"]
