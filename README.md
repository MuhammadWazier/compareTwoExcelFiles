# مقارنة ملفات Excel

## المتطلبات

- Node.js و npm

## التثبيت والتشغيل

بعد تنزيل المشروع أو استنساخه من GitHub، افتح الطرفية داخل مجلد المشروع ونفّذ:

```bash
npm ci
npm run dev
```

يعتمد `npm ci` على `package-lock.json` لتثبيت الإصدارات المطلوبة. لا حاجة إلى رفع `node_modules` إلى GitHub.

لإنشاء نسخة الإنتاج:

```bash
npm run build
```

تُنشأ ملفات الإنتاج في مجلد `dist`.

## النشر على GitHub Pages

يُبنى الموقع ويُنشر تلقائيًا عند الدفع إلى فرع `main` عبر GitHub Actions.

1. افتح إعدادات المستودع على GitHub.
2. انتقل إلى **Settings > Pages**.
3. ضمن **Build and deployment**، اختر **GitHub Actions** كمصدر النشر.
4. ادفع تغييراتك إلى `main`، ثم تابع حالة النشر من تبويب **Actions**.

رابط الموقع:
https://muhammadwazier.github.io/compareTwoExcelFiles/
