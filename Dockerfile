FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && playwright install --with-deps chromium

COPY crawler_core.py app.py ./
COPY templates ./templates

ENV PORT=10000
EXPOSE 10000

# 1 worker bắt buộc: trạng thái job nằm trong bộ nhớ (JOBS dict)
CMD gunicorn --workers 1 --threads 8 --timeout 120 --bind 0.0.0.0:$PORT app:app
