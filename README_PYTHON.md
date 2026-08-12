# 🚀 أداة تصدير معايير التصنيف إلى PDF - Python

نسخة Python من أداة تصدير الجداول من منصة Heritage Survey (STAR1) إلى ملفات PDF.

## ✨ المميزات

- ✅ تسجيل دخول أوتوماتيكي
- ✅ إدخال رقم قرار التسجيل تلقائياً
- ✅ فتح الاستمارات والانتقال لمعايير التصنيف
- ✅ تصدير الجدول فقط (بدون باقي الصفحة)
- ✅ حفظ PDF باسم الموقع الفعلي
- ✅ واجهة تفاعلية سهلة
- ✅ معالجة دفعية (batch processing)

## 📋 المتطلبات

- Python 3.8+
- Playwright

## 🔧 التثبيت

```bash
# 1. تثبيت المتطلبات
pip install -r requirements_python.txt

# 2. تثبيت المتصفح (مرة واحدة فقط)
playwright install chromium
```

## ▶️ الاستخدام

### الطريقة 1️⃣: واجهة تفاعلية (الأفضل)

```bash
python exporter_cli.py
```

**الخطوات:**
1. أدخل اسم المستخدم
2. أدخل كلمة المرور
3. أدخل أرقام القرارات (واحد في السطر)
4. اضغط Enter مرتين للانتهاء

**مثال:**
```
👤 اسم المستخدم: admin
🔑 كلمة المرور: ••••••••
📋 أدخل أرقام قرارات التسجيل:
   (اضغط Enter مرتين للانتهاء)

رقم قرار التسجيل #1: 21
رقم قرار التسجيل #2: 22
رقم قرار التسجيل #3: 23
رقم قرار التسجيل #4: 

✅ سيتم تصدير 3 قرار
هل تريد المتابعة؟ (نعم/لا): نعم
```

### الطريقة 2️⃣: استخدام برنامج Python

```python
import asyncio
from table_pdf_exporter_advanced import AdvancedTablePDFExporter

async def main():
    decisions = [
        {'name': 'قرار_1', 'decisionNumber': '21'},
        {'name': 'قرار_2', 'decisionNumber': '22'},
    ]

    exporter = AdvancedTablePDFExporter('admin', 'password')
    results = await exporter.run(decisions)

    for r in results:
        if r['status'] == 'success':
            print(f"✅ {r['siteName']}: {r['pdfPath']}")
        else:
            print(f"❌ {r['name']}: {r['error']}")

asyncio.run(main())
```

## 📊 المخرجات

الملفات المُصدَّرة توجد في مجلد `./exports/`:

```
exports/
├── مكة_المكرمة_1691234567890.pdf
├── مكة_المكرمة_1691234567890.png
├── الرياض_1691234567890.pdf
├── الرياض_1691234567890.png
└── ...
```

## ⚙️ تخصيص الإعدادات

```python
exporter = AdvancedTablePDFExporter(
    username='admin',
    password='password',
    options={
        'baseUrl': 'https://...',
        'outputDir': './my-exports',
        'headless': False,  # عرض المتصفح على الشاشة
        'navigationTimeout': 60000,  # 60 ثانية
    }
)
```

## 🛠️ تعديل المتغيرات

### تغيير مجلد المخرجات

```python
exporter = AdvancedTablePDFExporter(username, password, {
    'outputDir': './custom-exports'
})
```

### عرض المتصفح على الشاشة

```python
exporter = AdvancedTablePDFExporter(username, password, {
    'headless': False
})
```

### زيادة المهلة الزمنية

```python
exporter = AdvancedTablePDFExporter(username, password, {
    'navigationTimeout': 90000  # 90 ثانية
})
```

## 📝 مثال كامل

```python
import asyncio
from table_pdf_exporter_advanced import AdvancedTablePDFExporter

async def main():
    # البيانات
    decisions = [
        {'name': 'قرار_21', 'decisionNumber': '21'},
        {'name': 'قرار_22', 'decisionNumber': '22'},
    ]

    # إنشاء المصدِّر
    exporter = AdvancedTablePDFExporter(
        username='admin',
        password='password123',
        options={
            'outputDir': './exports',
            'headless': True,
        }
    )

    # تشغيل التصدير
    try:
        results = await exporter.run(decisions)

        # عرض النتائج
        print('\n╔════════════════════════════════════════╗')
        print('║             النتائج النهائية           ║')
        print('╚════════════════════════════════════════╝')

        for r in results:
            if r['status'] == 'success':
                print(f"✅ {r['siteName']}")
                print(f"   PDF: {r['pdfPath']}")
                print(f"   PNG: {r['pngPath']}")
            else:
                print(f"❌ {r['name']}: {r['error']}")

    except Exception as e:
        print(f'❌ خطأ: {e}')

asyncio.run(main())
```

## ❌ حل المشاكل

### مشكلة: "لم يتم العثور على الجدول"

```bash
# جرب مع عرض المتصفح:
python -c "
import asyncio
from table_pdf_exporter_advanced import AdvancedTablePDFExporter

async def main():
    exporter = AdvancedTablePDFExporter('admin', 'pass', {
        'headless': False  # عرض المتصفح
    })
    await exporter.init()
    await exporter.login()
    # افحص الصفحة يدوياً

asyncio.run(main())
"
```

### مشكلة: "انتهاء المهلة الزمنية"

```python
exporter = AdvancedTablePDFExporter(username, password, {
    'navigationTimeout': 90000  # زيادة إلى 90 ثانية
})
```

### مشكلة: لا يعمل Playwright

```bash
# تأكد من تثبيت المتصفح:
playwright install chromium

# أعد تثبيت Playwright:
pip install --upgrade playwright
```

## 📊 معلومات الأداء

| المعيار | التفاصيل |
|--------|---------|
| حجم PDF | ~2-3 MB |
| حجم PNG | ~8-10 MB |
| وقت القرار الواحد | ~20-30 ثانية |
| 5 قرارات | ~2-2.5 دقيقة |
| 10 قرارات | ~4-5 دقائق |

## 🔄 الفرق بين النسخة Python و JavaScript

| الميزة | Python | JavaScript |
|--------|--------|-----------|
| المكتبة | Playwright | Puppeteer |
| الأداء | متقارب | متقارب |
| السهولة | ✅ سهلة جداً | متوسطة |
| التكامل | ✅ أفضل مع Python | Node.js |
| التوثيق | ✅ شامل | شامل |

## 📚 موارد إضافية

- [توثيق Playwright Python](https://playwright.dev/python/)
- [Python AsyncIO](https://docs.python.org/3/library/asyncio.html)
- [الملف الأصلي JavaScript](./table-pdf-exporter-advanced.js)

## 🎯 نصائح

1. **ابدأ بقرار واحد** للتأكد من أن كل شيء يعمل
2. **استخدم headless=False** إذا حصلت مشكلة
3. **اترك وقتاً كافياً** بين القرارات (الأداة تفعل ذلك تلقائياً)
4. **افحص الملفات المصدرة** في مجلد `exports/`

---

**آخر تحديث:** 12 أغسطس 2026
