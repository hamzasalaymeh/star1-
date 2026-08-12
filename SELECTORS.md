# 🎯 الـ Selectors الدقيقة للواجهة

## ✅ تم استخراجها من الـ HTML الفعلي للمنصة

---

## 📋 البحث والفلترة

### 1. حقل رقم قرار التسجيل
```javascript
// Selector
input[name="ctl12$ctl28"]

// HTML الفعلي
<input name="ctl12$ctl28" type="text" class="form-control" />

// الاستخدام
const input = page.$('input[name="ctl12$ctl28"]');
input.type(decisionNumber);
```

---

### 2. زر البحث
```javascript
// Selector
#ctl12_btnSearch

// HTML الفعلي
<input type="submit" name="ctl12$btnSearch" value="بحــــث" 
       id="ctl12_btnSearch" class="btn btn-secondary" />

// الاستخدام
const btn = page.$('#ctl12_btnSearch');
await btn.click();
```

---

### 3. حقول الفلترة الأخرى (اختياري)
```javascript
// الموقع/الكود
input[name="ctl12$ctl04"]

// المنطقة
select[name="ctl12$ctl08"]

// الفرع
select[name="ctl12$ctl12"]

// المحافظة
select[name="ctl12$ctl16"]

// الحي
input[name="ctl12$ctl20"]

// البلدة
input[name="ctl12$ctl24"]

// الخطوة
select[name="ctl12$ctl32"]
```

---

## 📊 الجدول

### الجدول الرئيسي
```javascript
// Selector
#ctl12_GridView1

// HTML الفعلي
<table class="GridView GridGreen" id="ctl12_GridView1" 
       cellspacing="0" rules="all" border="1">
  <tr>
    <th>الكود</th>
    <th>اسم الموقع</th>
    <th>المنطقة</th>
    ...
  </tr>
  <tr>
    <td>90</td>
    <td><a href="DataEdit/90">مكة المكرمة</a></td>
    ...
  </tr>
</table>

// الاستخدام
const table = page.$('#ctl12_GridView1');
```

---

## 🔗 الروابط (استمارات)

### روابط الاستمارات
```javascript
// Selector - أول استمارة
a[href*="DataEdit/"]

// Selector - استمارة محددة
a[href="DataEdit/90"]

// HTML الفعلي
<a href="DataEdit/90">مكة المكرمة</a>
<a href="DataEdit/275">بيت البرزنجي</a>
<a href="DataEdit/512">مبنى رقم 37...</a>

// الاستخدام
const firstForm = page.$('a[href*="DataEdit/"]');
await firstForm.click();
```

---

## ⏳ ملخص الخطوات

```javascript
// 1. إدخال رقم القرار
await page.$('input[name="ctl12$ctl28"]').type('21');

// 2. الضغط على البحث
await page.$('#ctl12_btnSearch').click();

// 3. انتظار تحديث الجدول
await page.waitForTimeout(2000);

// 4. فتح أول استمارة
const firstForm = await page.$('a[href*="DataEdit/"]');
await firstForm.click();

// 5. انتظار تحميل الصفحة
await page.waitForNavigation();

// 6. البحث عن التبويبات (في الصفحة الجديدة)
// - روابط معايير التصنيف
// - الجدول الأحمر (معايير التصنيف)
```

---

## 📝 ملاحظات مهمة

### ASP.NET Forms
هذا نموذج ASP.NET Classic، يستخدم:
- `__VIEWSTATE` - حالة الصفحة
- `__EVENTVALIDATION` - التحقق من الأحداث
- `__doPostBack()` - إرسال النموذج

### الأسماء الديناميكية
أسماء الحقول تبدأ بـ `ctl12$` - قد تتغير في نسخ أخرى من الصفحة

### الخيارات
```html
<select name="ctl12$ctl32">
  <option value="500">تطبيق المعايير</option>
  <option value="600">مراجعة التصنيف</option>
  <option value="700">التسجيل المبدئي</option>
  ...
</select>
```

---

## 🔄 الخطوات التالية

**بعد فتح الاستمارة:**
1. ابحث عن التبويبات (معايير التصنيف)
2. ابحث عن الجدول الأحمر في صفحة معايير التصنيف
3. استخرج البيانات وصدرها كـ PDF

---

**تم التحديث: 12 أغسطس 2026**
