# Blida 2 Scientific Publications Platform

منصة إنتاجية لجمع بيانات النشر العلمي لأساتذة وطلبة الدكتوراه بجامعة البليدة 2.

- العربية / الفرنسية / الإنجليزية
- عدة مقالات في إرسال واحد
- تصنيف A+ / A / B / C / غير مصنف
- Scopus / Web of Science / ASJP / Google Scholar / DOAJ / أخرى / غير مفهرس
- وصل PDF يحتوي البيانات والباركود وQR للتحقق
- صفحة تحقق عامة
- لوحة إدارة وتصدير CSV
- SQLite مع WAL ومجلد بيانات دائم
- حماية CSRF والجلسات الآمنة وترويسات أمنية
- Docker / Docker Compose
- Healthcheck على /healthz

## Railway
اضبط SECRET_KEY وADMIN_PASSWORD وDATA_DIR=/app/data.
اربط Volume دائم على /app/data.
اضبط Healthcheck على /healthz وأنشئ نطاقاً عاماً بعد نجاح النشر.
