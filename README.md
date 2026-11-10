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
