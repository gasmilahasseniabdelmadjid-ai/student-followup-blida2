FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 PORT=8080
RUN useradd --create-home --uid 10001 appuser
WORKDIR /app
COPY scientific_publications/requirements.txt ./requirements.txt
RUN pip install -r requirements.txt
COPY scientific_publications/ ./
RUN mkdir -p /app/data/receipts /app/data/covers && chown -R appuser:appuser /app
USER appuser
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=10s --retries=3 CMD python -c "import os,urllib.request; urllib.request.urlopen('http://127.0.0.1:'+os.environ.get('PORT','8080')+'/healthz', timeout=5)"
CMD ["sh","-c","exec gunicorn --workers=2 --threads=2 --timeout=120 --graceful-timeout=30 --access-logfile=- --error-logfile=- --bind 0.0.0.0:\${PORT:-8080} app:app"]