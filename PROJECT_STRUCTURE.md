# 📁 بنية المشروع

## الملفات الرئيسية

### 🔷 مكتبات التصدير

#### `table-pdf-exporter-advanced.js` (المحرك الرئيسي)
المكتبة الأساسية التي تتعامل مع كل شيء:

```javascript
class AdvancedTablePDFExporter
├── init()                                   // تهيئة المتصفح
├── login()                                  // تسجيل الدخول
├── selectDecisionNumber()                  // إدخال رقم القرار
├── openFirstForm()                         // فتح الاستمارة الأولى
├── navigateToClassificationCriteria()     // الانتقال للتصنيف
├── extractRedBoxTable()                   // استخراج الجدول الأحمر
├── exportTableAsCleanPDF()                // تصدير PDF
├── captureTableScreenshot()               // لقطة شاشة
├── scrollTableAndLoadData()               // تحميل البيانات
├── exportBatch()                          // معالجة دفعة قرارات
├── close()                                // إغلاق المتصفح
└── run()                                  // تشغيل كامل العملية
```

#### `table-pdf-exporter.js` (نسخة بسيطة)
إذا أردت نسخة أبسط بدون الميزات المتقدمة.

---

### 🟢 أدوات واجهة المستخدم

#### `exporter-cli.js` (واجهة تفاعلية)
```bash
$ node exporter-cli.js

1. يطلب اسم المستخدم
2. يطلب كلمة المرور
3. يطلب أرقام القرارات (واحد في السطر)
4. يعرض النتائج والأخطاء
```

#### `debug-exporter.js` (أداة التصحيح)
```bash
$ node debug-exporter.js admin password 21

يوضح كل خطوة بالتفصيل:
- يعرض المتصفح على الشاشة
- يسجل كل الطلبات والاستجابات
- يحفظ لقطات شاشة في كل مرحلة
- يساعد في استكشاف الأخطاء
```

---

### 📚 الأمثلة والتوثيق

#### `example-usage.js`
4 أمثلة عملية:
1. **الاستخدام الأساسي** - `node example-usage.js 1`
2. **إعدادات مخصصة** - `node example-usage.js 2`
3. **خطوات يدوية** - `node example-usage.js 3`
4. **معالجة دفعة** - `node example-usage.js 4`

#### `README_AR.md`
توثيق شامل بالعربية عن:
- المميزات
- المتطلبات
- طرق الاستخدام
- استكشاف الأخطاء

#### `QUICK_START.md`
دليل بدء سريع:
- البدء في 2 دقيقة
- أمثلة فورية
- حل المشاكل الشائعة

---

### ⚙️ ملفات التكوين

#### `package.json`
```json
{
  "name": "star1-",
  "dependencies": {
    "puppeteer": "^21.0.0"
  },
  "scripts": {
    "export-pdf": "node table-pdf-exporter-advanced.js"
  }
}
```

#### `.gitignore`
```
node_modules
exports/
debug-exports/
```

---

## 🔄 تدفق البيانات

```
exporter-cli.js (واجهة تفاعلية)
        ↓
AdvancedTablePDFExporter (المحرك)
        ↓
Puppeteer (التحكم بالمتصفح)
        ↓
exports/ (ملفات PDF + صور)
```

---

## 📊 مثال: معالجة 3 قرارات

```javascript
// المدخلات
decisions = [
  { decisionNumber: '21', name: 'قرار_21' },
  { decisionNumber: '22', name: 'قرار_22' },
  { decisionNumber: '23', name: 'قرار_23' }
]

// العملية
for each decision:
  1. selectDecisionNumber('21')
     └─ يدخل الرقم 21 في حقل البحث
     
  2. openFirstForm()
     └─ يفتح أول استمارة من 3000
     
  3. navigateToClassificationCriteria()
     └─ ينقر على "معايير التصنيف"
     
  4. scrollTableAndLoadData()
     └─ يحمل كل 3000 سجل
     
  5. extractRedBoxTable()
     └─ يجد الجدول الأحمر
     
  6. exportTableAsCleanPDF()
     └─ يحفظ PDF
     
  7. captureTableScreenshot()
     └─ يحفظ PNG

// المخرجات
exports/
├── قرار_21_1691234567890.pdf
├── قرار_21_1691234567890.png
├── قرار_22_1691234567899.pdf
├── قرار_22_1691234567899.png
├── قرار_23_1691234567908.pdf
└── قرار_23_1691234567908.png
```

---

## 🛠️ التطوير والتعديل

### إضافة خطوة جديدة

مثال: إضافة دالة للتحقق من الأخطاء

```javascript
// في table-pdf-exporter-advanced.js
async validateData() {
  console.log('✓ التحقق من البيانات...');
  
  const isValid = await this.page.evaluate(() => {
    const table = document.querySelector('table');
    return table && table.querySelectorAll('tr').length > 0;
  });
  
  if (!isValid) {
    throw new Error('البيانات غير صحيحة');
  }
  
  console.log('✅ البيانات صحيحة');
}

// ثم استدعها في exportBatch():
await this.validateData();
```

### تخصيص الألوان والطراز

في `extractRedBoxTable()`:
```javascript
// البحث عن ألوان أخرى:
if (style.borderColor.includes('blue')) { /* أزرق */ }
if (style.borderColor.includes('green')) { /* أخضر */ }
if (style.borderColor.includes('red')) { /* أحمر */ }
```

---

## 📈 معلومات الأداء

| العملية | الوقت |
|--------|------|
| تسجيل الدخول | 2-3 ثانية |
| إدخال رقم القرار | 1-2 ثانية |
| فتح الاستمارة | 2-3 ثانية |
| الانتقال للتصنيف | 1-2 ثانية |
| تحميل 3000 سجل | 5-10 ثواني |
| تصدير PDF | 2-3 ثواني |
| **المجموع (قرار واحد)** | **~15-25 ثانية** |
| **3 قرارات** | **~45-75 ثانية** |

---

## 🔒 الأمان

- كلمات المرور لا تُحفظ
- لا توجد اتصالات خارجية
- جميع العمليات محلية
- البيانات تُحفظ فقط كـ PDF و PNG

---

## 🐛 استكشاف الأخطاء

```bash
# 1. شغل debug mode:
node debug-exporter.js admin pass123 21

# 2. شاهد لقطات الشاشة في:
ls debug-exports/

# 3. أخطاء شائعة:
# - "لم يتم العثور على الجدول" → الرقم خاطئ
# - "فشل تسجيل الدخول" → البيانات خاطئة
# - "Connection timeout" → إنترنت بطيء
```

---

## 📞 الدعم

إذا واجهت مشاكل:
1. اقرأ `QUICK_START.md`
2. جرب `debug-exporter.js`
3. تحقق من `README_AR.md`
4. اطلب المساعدة مع لقطات الشاشة من `debug-exports/`

---

**آخر تحديث: 2024 | جميع الحقوق محفوظة**
