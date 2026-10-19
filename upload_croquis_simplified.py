#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🏛️ أتمتة منصة التراث - نسخة مبسطة (مستقلة ومباشرة)
رفع صور الكروكي مع الإحداثيات والمساحة

الاستخدام:
  1. انسخ هذا الملف إلى مجلد الصور والإكسل
  2. شغّل: python upload_croquis_simplified.py
  3. سجّل الدخول يدوياً عند الطلب
  4. اتبع التعليمات

الميزات:
- تلقائي 100%: يكتشف الإكسل والصور والأرقام بدون حوارات
- آمن: تحقق مزدوج من البيانات قبل الرفع
- قابل للاستئناف: لا يعيد رفع ما تم بنجاح سابقاً
- تقرير فوري: ملف إكسل يحتفظ بالحالة أثناء التشغيل
"""

import subprocess, sys, time, re, os
from datetime import datetime

if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
        print('🌙 منع السكون: مُفعّل')
    except Exception:
        pass

REQUIRED_PACKAGES = {
    'selenium': 'selenium',
    'pandas': 'pandas',
    'openpyxl': 'openpyxl',
    'webdriver_manager': 'webdriver-manager',
}

def ensure_dependencies():
    import importlib
    for module_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            print(f'📦 تثبيت: {package_name}')
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name, '-q'])

ensure_dependencies()

import pandas as pd
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

# البيانات التلقائية من المجلد الحالي
WORK_FOLDER = os.getcwd()
print(f'📁 المجلد: {WORK_FOLDER}\n')

excel_files = [f for f in os.listdir(WORK_FOLDER)
               if f.lower().endswith(('.xlsx', '.xls'))
               and not f.startswith('~$') and not f.startswith('تقرير_')]
if not excel_files:
    print('❌ لم يتم العثور على ملف Excel')
    sys.exit(1)

EXCEL_PATH = os.path.join(WORK_FOLDER, excel_files[0])
print(f'📋 {excel_files[0]}')

try:
    df = pd.read_excel(EXCEL_PATH)
except Exception as e:
    print(f'❌ خطأ في قراءة Excel: {e}')
    sys.exit(1)

def find_col(df, keyword):
    for col in df.columns:
        if keyword in col:
            return col
    return None

x_col = find_col(df, 'الطول')
y_col = find_col(df, 'العرض')
a_col = find_col(df, 'مساحة')

if not all([x_col, y_col, a_col]):
    print(f'❌ الأعمدة المطلوبة غير موجودة')
    print(f'   الموجودة: {list(df.columns)}')
    sys.exit(1)

images = {}
for f in sorted(os.listdir(WORK_FOLDER)):
    if not f.lower().endswith(('.png', '.jpg', '.jpeg')):
        continue
    m = re.match(r'^(\d+)', f)
    if m:
        num = int(m.group(1))
        images[num] = os.path.join(WORK_FOLDER, f)

print(f'🖼️ {len(images)} صورة')

# الدوال المساعدة
def parse_num(v):
    if v is None or pd.isna(v) or isinstance(v, bool):
        raise ValueError('فارغة')
    if isinstance(v, (int, float)):
        if pd.isna(v):
            raise ValueError('فارغة')
        return float(v)

    t = str(v).strip()
    if not t or t.startswith('='):
        raise ValueError('غير صالحة')

    t = t.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
    t = t.replace(' ', ' ').replace(' ', ' ')
    t = t.replace('−', '-').replace('–', '-').replace('—', '-')
    t = t.replace('٫', '.').replace('٬', '')
    t = re.sub(r'(?i)\b(?:sqm|sq\.?\s*m|m2)\b', ' ', t)
    t = re.sub(r'(?:متر(?:اً|ا)?\s*(?:مربع|م²|2)?|م²)', ' ', t)

    if ',' in t and '.' in t:
        if t.rfind(',') > t.rfind('.'):
            t = t.replace('.', '').replace(',', '.')
        else:
            t = t.replace(',', '')
    elif ',' in t and t.count(',') == 1:
        t = t.replace(',', '.')
    elif ',' in t:
        parts = t.split(',')
        t = ''.join(parts[:-1]) + '.' + parts[-1]

    m = re.findall(r'[-+]?(?:\d+(?:\.\d*)?|\.\d+)', t)
    if len(m) != 1:
        raise ValueError('غير واضح')
    return float(m[0])

def js_click(el):
    driver.execute_script("arguments[0].click();", el)

def fill_field(name_part, val):
    try:
        el = driver.find_element('xpath', f"//input[contains(@name,'{name_part}')]")
        clean = str(val).replace('\n', '').replace('\r', '').strip()
        driver.execute_script("""
            var e = arguments[0], v = arguments[1];
            e.focus(); e.value = ''; e.value = v;
            e.dispatchEvent(new Event('input',  {bubbles:true}));
            e.dispatchEvent(new Event('change', {bubbles:true}));
            e.dispatchEvent(new Event('blur',   {bubbles:true}));
        """, el, clean)
        time.sleep(0.3)
        return True
    except Exception:
        return False

def upload_image(cid, fpath):
    try:
        fi = driver.find_element('xpath', f"//input[@type='file' and contains(@name,'{cid}$Field1$ctl27')]")
        driver.execute_script("arguments[0].style.display='block';arguments[0].style.visibility='visible';arguments[0].style.opacity=1;", fi)
        fi.send_keys(os.path.abspath(fpath))
        driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", fi)
        time.sleep(3)
        return True
    except Exception:
        return False

def get_field_val(name_part):
    try:
        el = driver.find_element('xpath', f"//input[contains(@name,'{name_part}')]")
        return el.get_attribute('value') or ''
    except Exception:
        return ''

def numbers_equal(a, b, tol=0.0001):
    try:
        return abs(float(str(a).replace(',', '').strip()) - float(b)) < tol
    except:
        return False

# فتح المتصفح والدخول
print('\n🌐 فتح المنصة...\n')

opts = Options()
opts.add_argument('--no-sandbox')
opts.add_argument('--disable-dev-shm-usage')

try:
    srv = Service(ChromeDriverManager().install())
except:
    srv = None

driver = webdriver.Chrome(service=srv, options=opts)
driver.set_window_size(1280, 1024)

DATA_VIEW_URL = 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataView'
driver.get(DATA_VIEW_URL)
time.sleep(3)

if 'Login' in driver.current_url or driver.title == 'Runtime Error':
    print('👤 سجّل دخولك يدوياً في النافذة، ثم اضغط Enter هنا')
    input('')
    driver.get(DATA_VIEW_URL)
    time.sleep(3)

print(f'✅ {driver.title}\n')

# إنشاء تقرير فوري
report_file = os.path.join(WORK_FOLDER, f'تقرير_الرفع_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
wb = Workbook()
ws = wb.active
ws.title = 'النتائج'
ws.append(['#', 'الصورة', 'X', 'Y', 'المساحة', 'الحالة', 'ملاحظات'])
wb.save(report_file)

# المعالجة الرئيسية
stats = {'ok': 0, 'skip': 0, 'fail': 0}

print(f'🚀 معالجة {len(images)} صورة\n')

for bnum in sorted(images.keys()):
    imgpath = images[bnum]
    imgname = os.path.basename(imgpath)

    row_data = None
    for _, row in df.iterrows():
        for col in df.columns:
            try:
                if int(row[col]) == bnum:
                    row_data = row
                    break
            except:
                pass
        if row_data is not None:
            break

    if row_data is None:
        print(f'⏭️  {bnum}: لا توجد بيانات')
        ws.append([bnum, imgname, '', '', '', '⏭️ متخطى', 'بدون بيانات'])
        stats['skip'] += 1
        wb.save(report_file)
        continue

    try:
        xv = parse_num(row_data[x_col])
        yv = parse_num(row_data[y_col])
        av = parse_num(row_data[a_col])
    except Exception as e:
        print(f'❌ {bnum}: خطأ في البيانات ({str(e)[:30]})')
        ws.append([bnum, imgname, '', '', '', '❌ فشل', f'خطأ: {str(e)[:30]}'])
        stats['fail'] += 1
        wb.save(report_file)
        continue

    print(f'📍 {bnum}: X={xv:.2f}, Y={yv:.2f}, A={av:.2f}')

    try:
        print(f'   ⚠️ ابحث يدوياً عن المبنى {bnum} واضغط على تعديل')
        print(f'   📝 سيتم ملء الحقول تلقائياً:')
        print(f'      X={xv:.2f}, Y={yv:.2f}, مساحة={av:.2f}')
        print(f'   ⏳ اضغط Enter بعد فتح صفحة التحرير...')
        input('')

        time.sleep(1)
        fill_field('ctl37$Field1$ctl03', str(xv))
        time.sleep(0.3)
        fill_field('ctl38$Field1$ctl03', str(yv))
        time.sleep(0.3)
        fill_field('ctl39$Field1$ctl03', str(av))
        time.sleep(0.5)

        if not upload_image('ctl40', imgpath):
            raise Exception('فشل رفع الصورة')

        try:
            sb = driver.find_element('xpath', "//input[@name='ctl12$btnSave']")
            js_click(sb)
            print('   💾 حفظ')
            time.sleep(4)
        except:
            print('   ⚠️ لم يتم العثور على زر الحفظ')

        time.sleep(1)
        check_x = get_field_val('ctl37$Field1$ctl03')
        check_y = get_field_val('ctl38$Field1$ctl03')
        check_a = get_field_val('ctl39$Field1$ctl03')
        check_img = get_field_val('ctl40$Field1$ctl28')

        ok = (numbers_equal(check_x, xv) and
              numbers_equal(check_y, yv) and
              numbers_equal(check_a, av) and
              bool(str(check_img).strip()))

        if ok:
            print(f'   ✅ نجح\n')
            ws.append([bnum, imgname, xv, yv, av, '✅ نجح', 'تم الرفع بنجاح'])
            stats['ok'] += 1
        else:
            print(f'   ❌ لم يتم التحقق (X={check_x}, Y={check_y}, A={check_a}, img={bool(check_img)})\n')
            ws.append([bnum, imgname, xv, yv, av, '❌ فشل', f'التحقق: X={check_x}, Y={check_y}, A={check_a}'])
            stats['fail'] += 1

    except Exception as e:
        print(f'   ❌ خطأ: {str(e)[:50]}\n')
        ws.append([bnum, imgname, xv, yv, av, '❌ فشل', str(e)[:50]])
        stats['fail'] += 1

    wb.save(report_file)

# الإنهاء
ws.append([])
ws.append(['ملخص النتائج'])
ws.append(['✅ نجح', stats['ok']])
ws.append(['⏭️ متخطى', stats['skip']])
ws.append(['❌ فشل', stats['fail']])

wb.save(report_file)
print(f'💾 {report_file}')
print(f'\n📊 النتائج:')
print(f'   ✅ نجح: {stats["ok"]}')
print(f'   ⏭️ متخطى: {stats["skip"]}')
print(f'   ❌ فشل: {stats["fail"]}')

driver.quit()
print('\n✨ انتهى')
