# Student Follow-up — University of Blida 2

مستودع بناء تطبيق **بطاقة متابعة الطلبة — جامعة البليدة 2**.

## البناء
يتم بناء APK على GitHub Actions باستخدام Node.js 22 وJava 21 وAndroid API 36.

قبل التشغيل، أضف في **Settings → Secrets and variables → Actions**:
- SFC_KEYSTORE_B64
- SFC_KEY_ALIAS
- SFC_KEYSTORE_PASSWORD
- SFC_KEY_PASSWORD

لا ترسل قيم الأسرار داخل المحادثة.

بعد ذلك افتح **Actions → Build Direct APK → Run workflow**.

<!-- CI build pipeline updated -->

<!-- Production signing verification trigger: 2026-09-27 -->
