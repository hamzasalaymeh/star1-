# -*- coding: utf-8 -*-
"""
🏛️ أتمتة الإحداثيات (بدون كروكي) — إدخال X / Y / المساحة من ملف Excel مباشرة.

الفكرة: ترفع ملف Excel فيه عمود «Site ID» + «خط الطول X» + «خط العرض Y»
(و«المساحة» اختياري)، فيدخل لكل موقع بصفحته على المنصة ويكتب إحداثياته ويحفظ —
بدون رفع صورة كروكي وبدون مطابقة بالاسم (المطابقة برقم الموقع مباشرة = الأدق).

يكتب فعلياً على المنصة (صفحة استكمال بيانات الميدان / DataEdit/<id>/2)، فيُنصح
بتجربته على عدد محدود أولاً عبر خيار «الحد».

يعمل مستقلاً أو عبر مركز التحكم (يبعث التقدّم ويحترم الإيقاف/الاستئناف).
"""
import os
import sys
import time
import json
import argparse
import subprocess

# جسر مركز التحكم (اختياري)
try:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from agent_runtime import R as _R
except Exception:
    class _NoR:
        enabled = False
        def total(self, *a, **k): pass
        def progress(self, *a, **k): pass
        def log(self, *a, **k): pass
        def done(self, *a, **k): pass
        def should_stop(self): return False
        def wait_if_paused(self): return True
    _R = _NoR()

try:
    import openpyxl
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError as exc:
    raise SystemExit(
        f"مكتبة ناقصة: {getattr(exc,'name','?')}\n"
        "ثبّت مرة واحدة:  py -m pip install selenium webdriver-manager openpyxl"
    ) from exc

BASE = 'https://heritage2.itqan-consultant.com'
INBOX = f'{BASE}/Web/App/Home/Inbox'

# أسماء حقول الإحداثيات في صفحة استكمال بيانات الميدان (مثبتة من أتمتة الكروكي)
F_X, F_Y, F_AREA = 'ctl37$Field1$ctl03', 'ctl38$Field1$ctl03', 'ctl39$Field1$ctl03'
SAVE_BTN = "//input[@name='ctl12$btnSave']"


# ───────────────────────── قراءة Excel ─────────────────────────
def _find_col(headers, *keys):
    for i, h in enumerate(headers):
        hn = str(h or '').strip().lower()
        for k in keys:
            if k.lower() in hn:
                return i
    return None


def num_str(v):
    """تحويل الرقم لنص نظيف (بلا صيغة علمية، بلا أصفار زائدة)."""
    if v is None:
        return ''
    s = str(v).strip()
    if s in ('', '<null>', '<Null>', 'None', 'nan'):
        return ''
    try:
        f = float(s)
    except ValueError:
        return s
    out = f"{f:.6f}".rstrip('0').rstrip('.')
    return out


def read_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    ci_id = _find_col(headers, 'site id', 'siteid', 'رقم الموقع', 'معرف', 'id')
    ci_x = _find_col(headers, 'خط الطول', 'الطول', 'longitude', ' x', 'x ')
    ci_y = _find_col(headers, 'خط العرض', 'العرض', 'latitude', ' y', 'y ')
    ci_area = _find_col(headers, 'مساحة', 'المساحة', 'area')
    ci_name = _find_col(headers, 'اسم الموقع', 'الاسم', 'name')
    if ci_id is None or ci_x is None or ci_y is None:
        raise SystemExit(
            "تعذّر العثور على الأعمدة المطلوبة. لازم يحتوي الملف على:\n"
            "  • عمود رقم الموقع (Site ID)\n"
            "  • عمود «خط الطول X»\n"
            "  • عمود «خط العرض Y»\n"
            f"الأعمدة الموجودة: {headers}"
        )
    rows = []
    for r in range(2, ws.max_row + 1):
        sid = ws.cell(r, ci_id + 1).value
        x = num_str(ws.cell(r, ci_x + 1).value)
        y = num_str(ws.cell(r, ci_y + 1).value)
        area = num_str(ws.cell(r, ci_area + 1).value) if ci_area is not None else ''
        name = ws.cell(r, ci_name + 1).value if ci_name is not None else ''
        if sid is None or not str(sid).strip():
            continue
        sid = str(sid).strip()
        if sid.endswith('.0'):
            sid = sid[:-2]
        if not (x and y):
            continue
        rows.append({'id': sid, 'x': x, 'y': y, 'area': area, 'name': str(name or '')})
    return rows


# ───────────────────────── المتصفح ─────────────────────────
driver = None
wait = None


def launch_chrome():
    global driver, wait
    chrome_paths = [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
    ]
    chrome_exe = next((p for p in chrome_paths if os.path.exists(p)), None)
    if not chrome_exe:
        raise FileNotFoundError('❌ Chrome غير مثبّت في المسارات المعتادة')
    subprocess.run(['taskkill', '/f', '/im', 'chrome.exe'], capture_output=True)
    subprocess.run(['taskkill', '/f', '/im', 'chromedriver.exe'], capture_output=True)
    time.sleep(2)
    subprocess.Popen([
        chrome_exe, '--remote-debugging-port=9222',
        r'--user-data-dir=C:\chrome_debug_session',
        '--no-first-run', '--no-default-browser-check', '--start-maximized'
    ])
    time.sleep(4)
    opts = Options()
    opts.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    wait = WebDriverWait(driver, 15)
    print('✅ تم فتح Chrome والاتصال به')


def wait_for_login(max_wait=600):
    """ينتظر تسجيل الدخول اليدوي في نافذة Chrome (آمن مع مركز التحكم — بلا input)."""
    driver.get(INBOX); time.sleep(3)
    deadline = time.time() + max_wait
    warned = False
    while ('Login' in (driver.current_url or '') or driver.title == 'Runtime Error'):
        if not warned:
            print('👤 سجّل دخولك الآن في نافذة Chrome المفتوحة — سأكمل تلقائياً بعد الدخول…')
            _R.log('بانتظار تسجيل الدخول في نافذة Chrome…', 'warn')
            warned = True
        try:
            driver.get(BASE); time.sleep(5)
            driver.get(INBOX); time.sleep(2)
        except Exception:
            time.sleep(3)
        if time.time() > deadline:
            raise SystemExit('انتهت مهلة انتظار تسجيل الدخول.')
    print('✅ تم تسجيل الدخول')
    _R.log('تم تسجيل الدخول', 'ok')


def js_click(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", el)


def fill_field(name_contains, value):
    try:
        inp = driver.find_element('xpath', f"//input[contains(@name,'{name_contains}')]")
        clean = str(value).replace('\n', '').replace('\r', '').strip()
        driver.execute_script("""
            var el=arguments[0], val=arguments[1];
            el.focus(); el.value=''; el.value=val;
            el.dispatchEvent(new Event('input',{bubbles:true}));
            el.dispatchEvent(new Event('change',{bubbles:true}));
            el.dispatchEvent(new Event('blur',{bubbles:true}));
        """, inp, clean)
        time.sleep(0.4)
        return (inp.get_attribute('value') or '').strip() == clean
    except Exception as e:
        print(f'   ❌ {name_contains}: {str(e)[:50]}')
        return False


def get_val(name_contains, timeout=15):
    end = time.time() + timeout
    while time.time() < end:
        try:
            el = driver.find_element('xpath', f"//input[contains(@name,'{name_contains}')]")
            v = el.get_attribute('value')
            if v and v.strip():
                return v
        except Exception:
            pass
        time.sleep(0.5)
    return ''


def process_site(row):
    """يكتب X/Y/المساحة لموقع واحد ويحفظ ويتحقق. يرجّع True عند النجاح."""
    url = f'{BASE}/Web/App/Pools/DataEdit/{row["id"]}/2'
    for attempt in range(2):
        driver.get(url); time.sleep(3)
        fields = [(F_X, row['x']), (F_Y, row['y'])]
        if row['area']:
            fields.append((F_AREA, row['area']))
        for fname, fval in fields:
            for _ in range(4):
                if fill_field(fname, fval):
                    break
                time.sleep(0.4)
        time.sleep(1.2)
        try:
            js_click(driver.find_element('xpath', SAVE_BTN))
            time.sleep(6)
        except Exception as e:
            print(f'   ❌ حفظ: {str(e)[:50]}')
            continue
        # تحقّق
        driver.get(url); time.sleep(2)
        cx, cy = get_val(F_X), get_val(F_Y)
        try:
            if abs(float(cx) - float(row['x'])) < 0.0001 and abs(float(cy) - float(row['y'])) < 0.0001:
                return True
        except ValueError:
            pass
        print(f'   🔁 تحقق فشل (X={cx}, Y={cy}) — إعادة محاولة' if attempt == 0 else '')
    return False


def main():
    ap = argparse.ArgumentParser(description='إدخال إحداثيات من Excel بدون كروكي')
    ap.add_argument('--excel', default=os.environ.get('HERITAGE_COORDS_EXCEL', ''))
    ap.add_argument('--output', default=os.environ.get('HERITAGE_OUTPUT', '') or 'CoordsRun')
    ap.add_argument('--limit', type=int,
                    default=int(os.environ.get('HERITAGE_COORDS_LIMIT') or 0))
    args = ap.parse_args()

    excel = (args.excel or '').strip()
    if not excel or not os.path.isfile(excel):
        # لو مُرِّر مجلد، نبحث عن أول Excel فيه
        folder = os.environ.get('HERITAGE_FOLDER', '').strip()
        if folder and os.path.isdir(folder):
            cands = [f for f in os.listdir(folder)
                     if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
            if cands:
                excel = os.path.join(folder, cands[0])
    if not excel or not os.path.isfile(excel):
        raise SystemExit('❌ حدّد ملف الإحداثيات (Excel) — لم يُعثر على ملف صالح.')

    rows = read_rows(excel)
    if args.limit and args.limit > 0:
        rows = rows[:args.limit]
    print(f'📄 ملف الإحداثيات: {excel}')
    print(f'📊 عدد المواقع للإدخال: {len(rows)}')
    print('✍️ تنبيه: هذه العملية تكتب فعلياً على المنصة (إحداثيات + مساحة).')

    out_dir = os.path.abspath(args.output)
    os.makedirs(out_dir, exist_ok=True)
    ckpt = os.path.join(out_dir, '.checkpoint_coords.json')
    try:
        done = set(json.load(open(ckpt, encoding='utf-8')).get('done', []))
    except Exception:
        done = set()
    if done:
        print(f'📌 استئناف: {len(done)} موقع مُدخل سابقاً — سيُتخطّى.')

    _R.total(len(rows))
    _R.progress(done=len(done))

    launch_chrome()
    wait_for_login()

    ok = fail = skip = 0
    results = []
    for i, row in enumerate(rows, 1):
        if not _R.wait_if_paused() or _R.should_stop():
            print('⏹ إيقاف آمن — التقدّم محفوظ.')
            break
        _R.progress(current=f'{row["id"]} {row["name"][:28]}')
        if row['id'] in done:
            skip += 1
            _R.progress(skipped=skip, current=row['id'])
            continue
        print(f'\n[{i}/{len(rows)}] موقع {row["id"]}  ({row["name"][:30]})  X={row["x"]} Y={row["y"]}'
              + (f' مساحة={row["area"]}' if row['area'] else ''))
        try:
            good = process_site(row)
        except Exception as e:
            good = False
            print(f'   ❌ خطأ: {str(e)[:60]}')
        if good:
            ok += 1
            done.add(row['id'])
            try:
                json.dump({'done': sorted(done)}, open(ckpt, 'w', encoding='utf-8'), ensure_ascii=False)
            except Exception:
                pass
            print('   ✅ تم الإدخال والحفظ')
            _R.log(f'✅ {row["id"]} تم', 'ok')
        else:
            fail += 1
            results.append(row['id'])
            print('   ❌ لم يُحفظ')
            _R.log(f'❌ {row["id"]} فشل', 'error')
        _R.progress(done=len(done), failed=fail, skipped=skip)

    _R.done(ok=ok, failed=fail, skipped=skip)
    print(f"\n{'='*55}\n  🎉 انتهى\n  ✅ أُدخل: {ok}\n  ⏭️ متخطّى: {skip}\n  ❌ فشل: {fail}")
    if results:
        print('  المواقع الفاشلة: ' + ', '.join(results[:30]) + (' …' if len(results) > 30 else ''))
    print('='*55)
    try:
        driver.quit()
    except Exception:
        pass


if __name__ == '__main__':
    main()
