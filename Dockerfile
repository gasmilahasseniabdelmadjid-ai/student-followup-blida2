FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1
RUN apt-get update && apt-get install -y --no-install-recommends fonts-dejavu && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
RUN mkdir -p /app/data/receipts
EXPOSE 8080
CMD ["sh","-c","exec gunicorn --workers 2 --threads 2 --timeout 120 --graceful-timeout 30 --access-logfile - --error-logfile - --bind 0.0.0.0:8080 app:app"]
