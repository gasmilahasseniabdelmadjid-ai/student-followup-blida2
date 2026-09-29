# دليل النشر الإنتاجي
1. اضبط SECRET_KEY وADMIN_PASSWORD كمتغيرات سرية.
2. DATA_DIR=/app/data.
3. اربط Volume دائم على /app/data.
4. Healthcheck: /healthz.
5. أنشئ نطاقاً عاماً بعد نجاح النشر.
6. فعّل HTTPS على النطاق الإنتاجي.
