# Бэкенд сервиса проверки документов: FastAPI + uvicorn (main:app), порт 8081.
# Системные пакеты: LibreOffice и poppler — конвертация DOCX/PPTX → PDF → PNG для OCR-пути.
# GPU-стек (requirements-gpu.txt) в образ не входит: layout и OCR работают по HTTP.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-impress poppler-utils \
        fonts-dejavu fonts-liberation curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Зависимости отдельным слоем — пересобираются только при изменении requirements.txt.
COPY requirements.txt ./
RUN pip install -r requirements.txt

COPY . .

# Каталоги данных монтируются томами (docker-compose.yml); создаём на случай запуска без них.
RUN mkdir -p logs_result uploads

EXPOSE 8081
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD curl -sf http://127.0.0.1:8081/api/health || exit 1

CMD ["python", "-m", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8081"]
