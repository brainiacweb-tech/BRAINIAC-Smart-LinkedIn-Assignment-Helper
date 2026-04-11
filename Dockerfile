FROM python:3.11-slim

# Install LibreOffice + dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        fonts-liberation \
        && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 10000

CMD gunicorn app:app --timeout 120 --workers 1 --bind "0.0.0.0:${PORT:-10000}"
