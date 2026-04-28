# مهووس v31 — Dockerfile لـ Cloud Run
FROM python:3.11-slim

# متغيرات البيئة العامة
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# تثبيت اعتماديات Python (بدون build-essential — كل الحزم لها wheels)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# نسخ الكود (الاسم الفعلي → app.py داخل الصورة)
COPY mahwous_app.py app.py

# مستخدم غير-root للأمان
RUN useradd --create-home --shell /bin/bash app && chown -R app:app /app
USER app

# Cloud Run يُمرّر PORT عبر متغير البيئة
ENV PORT=8080
EXPOSE 8080

# health check (اختياري — Cloud Run يُتحقق تلقائياً من PORT)
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:${PORT}/_stcore/health').read()" || exit 1

# تشغيل Streamlit
CMD streamlit run app.py \
    --server.port=$PORT \
    --server.address=0.0.0.0 \
    --server.headless=true \
    --server.enableCORS=false \
    --server.enableXsrfProtection=false \
    --server.maxUploadSize=500 \
    --browser.gatherUsageStats=false
