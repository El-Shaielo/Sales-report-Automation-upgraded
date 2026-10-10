FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        fonts-dejavu-core \
        fonts-liberation \
        libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY automatizar_cortes_de_venta_oficial_mejorado.py .
COPY "Plantilla_Cortes_de_ventas_Oficial.docx" .

EXPOSE 10000

CMD ["sh", "-c", "uvicorn automatizar_cortes_de_venta_oficial_mejorado:app --host 0.0.0.0 --port ${PORT:-10000}"]
