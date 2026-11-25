# -*- coding: utf-8 -*-
"""
🏛️ أتمتة الإحداثيات (بدون كروكي) — إدخال X / Y / المساحة من ملف Excel مباشرة.

ترفع ملف Excel فيه «اسم الموقع» + «خط الطول X» + «خط العرض Y» (و«المساحة» اختياري)،
وتُدخل رقم القرار. تبحث المنصة بالقرار، تجمع مبانيه (رقم+اسم)، تُطابق كل اسم في
الإكسل بأقرب مبنى على المنصة، ثم تكتب إحداثياته وتحفظ — بدون كروكي.

⚠️ يكتب فعلياً على المنصة (DataEdit/<id>/2). جرّب على عدد محدود أولاً عبر «الحد»،
وراجع نسبة التطابق قبل التوسّع. يعمل مستقلاً أو عبر مركز التحكم.
"""
import os
import re
import sys
import time
import json
import argparse
import subprocess

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
F_X, F_Y, F_AREA = 'ctl37$Field1$ctl03', 'ctl38$Field1$ctl03', 'ctl39$Field1$ctl03'
SAVE_BTN = "//input[@name='ctl12$btnSave']"
MATCH_THRESHOLD = 0.72

driver = None
wait = None


# ───────────────────────── قراءة Excel ─────────────────────────
def _find_col(headers, *keys):
    for i, h in enumerate(headers):
        hn = str(h or '').strip().lower()
        for k in keys:
            if k.lower() in hn:
                return i
    return None


def num_str(v):
    if v is None:
        return ''
    s = str(v).strip()
    if s in ('', '<null>', '<Null>', 'None', 'nan'):
        return ''
    try:
        f = float(s)
    except ValueError:
        return s
    return f"{f:.6f}".rstrip('0').rstrip('.')


def read_rows(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    ci_name = _find_col(headers, 'اسم الموقع', 'اسم المبنى', 'الاسم', 'name')
    ci_x = _find_col(headers, 'الإحداثي x', 'الاحداثي x', 'خط الطول', 'الطول', 'longitude', ' x', 'x ')
    ci_y = _find_col(headers, 'الإحداثي y', 'الاحداثي y', 'خط العرض', 'العرض', 'latitude', ' y', 'y ')
    ci_area = _find_col(headers, 'مساحة', 'المساحة', 'area')
    # عمود رقم المبنى الحقيقي على المنصة (كود الموقع) — لو وُجد نستخدمه مباشرة
    ci_code = _find_col(headers, 'كود الموقع', 'كود', 'رقم المبنى', 'رقم الموقع',
                        'building id', 'site id', 'siteid', 'معرف')
    if ci_x is None or ci_y is None or (ci_name is None and ci_code is None):
        raise SystemExit(
            "لازم يحتوي الملف على عمودي الإحداثيات (X و Y) + (كود الموقع أو اسم الموقع).\n"
            f"الأعمدة الموجودة: {headers}"
        )
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, ci_name + 1).value if ci_name is not None else ''
        x = num_str(ws.cell(r, ci_x + 1).value)
        y = num_str(ws.cell(r, ci_y + 1).value)
        area = num_str(ws.cell(r, ci_area + 1).value) if ci_area is not None else ''
        code = ws.cell(r, ci_code + 1).value if ci_code is not None else ''
        code = str(code or '').strip()
        if code.endswith('.0'):
            code = code[:-2]
        if not (x and y):
            continue
        if (not name or not str(name).strip()) and not code:
            continue
        rows.append({'name': str(name or '').strip(), 'x': x, 'y': y, 'area': area, 'code': code})
    return rows


def codes_are_real(rows):
    """الكود صالح كرقم منصة لو أغلب القيم أرقام من 4 خانات فأكثر (مثل 41234)."""
    vals = [r['code'] for r in rows if r.get('code')]
    if len(vals) < max(1, int(len(rows) * 0.6)):
        return False
    good = sum(1 for v in vals if v.isdigit() and int(v) >= 1000)
    return good >= int(len(vals) * 0.8)


# ───────────────────────── مطابقة الأسماء ─────────────────────────
def normalize_name(s):
    s = str(s).strip()
    s = re.sub(r'[\u064b-\u0652]', '', s)
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    s = s.replace('ة', 'ه').replace('ى', 'ي')
    for w in ['التاريخيه', 'التاريخي', 'قريه', 'قرية', 'موقع', 'منطقه', 'منطقة', 'حي']:
        s = s.replace(w, '')
    s = s.replace('ـ', '')
    return re.sub(r'\s+', '', s).strip()


def name_similarity(a, b):
    from difflib import SequenceMatcher
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def best_match(name, buildings, used):
    best, best_sc = None, 0.0
    for b in buildings:
        if b['id'] in used:
            continue
        sc = name_similarity(name, b['name'])
        if sc > best_sc:
            best_sc, best = sc, b
    return (best, best_sc) if (best and best_sc >= MATCH_THRESHOLD) else (None, best_sc)


# ───────────────────────── المتصفح ─────────────────────────
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


def reconnect():
    """يعيد الربط بنفس نافذة Chrome (debug port) لو انقطعت الجلسة — بلا فتح نافذة جديدة."""
    global driver, wait
    try:
        opts = Options()
        opts.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        wait = WebDriverWait(driver, 15)
        _ = driver.current_url
        return True
    except Exception:
        return False


def wait_for_login(max_wait=1800):
    """ينتظر الدخول اليدوي بدون أي إعادة تحميل (فلا نقاطع المستخدم أثناء الكتابة)."""
    try:
        driver.get(INBOX); time.sleep(3)
    except Exception:
        reconnect()
    deadline = time.time() + max_wait
    warned = False
    while True:
        try:
            url = driver.current_url or ''
            title = driver.title or ''
        except Exception:
            # النافذة أُغلقت أو انقطعت — نحاول إعادة الربط دون إزعاج المستخدم
            print('⚠️ انقطع الاتصال بنافذة Chrome — لا تغلق النافذة. أحاول إعادة الربط…')
            if not reconnect():
                time.sleep(4)
                if time.time() > deadline:
                    raise SystemExit('تعذّر الاتصال بنافذة Chrome (قد تكون أُغلقت).')
            continue
        # نُعتبر مسجّلين عند مغادرة صفحة الدخول/الخطأ
        if 'login' not in url.lower() and title != 'Runtime Error' and 'account' not in url.lower():
            print('✅ تم تسجيل الدخول')
            _R.log('تم تسجيل الدخول', 'ok')
            return
        if not warned:
            print('👤 سجّل دخولك في نافذة Chrome على مهلك — لن أُعيد تحميل الصفحة، وسأكمل تلقائياً بعد الدخول.')
            _R.log('بانتظار تسجيل الدخول (بدون إعادة تحميل)…', 'warn')
            warned = True
        if time.time() > deadline:
            raise SystemExit('انتهت مهلة انتظار تسجيل الدخول.')
        time.sleep(2)


def js_set(name, value):
    el = driver.find_element('xpath', f"//input[@name='{name}']")
    driver.execute_script("""
        var el=arguments[0],val=arguments[1];
        el.value=val;
        el.dispatchEvent(new Event('input',{bubbles:true}));
        el.dispatchEvent(new Event('change',{bubbles:true}));
    """, el, value)


def js_click(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", el)


def collect_platform_buildings(decision):
    """يبحث المنصة برقم القرار ويجمع كل المباني (id + name) عبر الصفحات."""
    driver.get(INBOX); time.sleep(3)
    try:
        js_set('ctl12$ctl28', str(decision))
        time.sleep(0.3)
        js_click(driver.find_element('xpath', "//input[@name='ctl12$btnSearch']"))
        time.sleep(3)
    except Exception as e:
        print(f'⚠️ تعذّر إرسال البحث: {str(e)[:60]}')
    buildings, seen, page = [], set(), 1
    while True:
        time.sleep(2)
        for row in driver.find_elements('xpath', "//table//tr[.//a[@href]]"):
            try:
                link = ''
                for a in row.find_elements('xpath', './/a[@href]'):
                    href = a.get_attribute('href') or ''
                    if 'DataEdit/' in href:
                        link = href
                        name = (a.text or '').strip()
                        break
                m = re.search(r'DataEdit/(\d+)', link)
                if not m:
                    continue
                bid = m.group(1)
                if bid in seen:
                    continue
                if not name or not re.search(r'[\u0600-\u06FFA-Za-z]', name):
                    # الاسم في خلية ثانية
                    for a in row.find_elements('xpath', './/a[@href]'):
                        at = (a.text or '').strip()
                        if at and re.search(r'[\u0600-\u06FF]', at):
                            name = at; break
                seen.add(bid)
                buildings.append({'id': bid, 'name': name})
            except Exception:
                continue
        nxt = driver.find_elements('xpath',
            f"//table//tr//td//a[text()='{page+1}'] | //a[contains(@href,'Page${page+1}')]")
        if nxt:
            js_click(nxt[0]); page += 1
        else:
            break
    return buildings


def fill_field(name_contains, value):
    try:
        inp = driver.find_element('xpath', f"//input[contains(@name,'{name_contains}')]")
        clean = str(value).replace('\n', '').replace('\r', '').strip()
        driver.execute_script("""
            var el=arguments[0],val=arguments[1];
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


def write_coords(bid, x, y, area):
    url = f'{BASE}/Web/App/Pools/DataEdit/{bid}/2'
    for attempt in range(2):
        driver.get(url); time.sleep(3)
        fields = [(F_X, x), (F_Y, y)]
        if area:
            fields.append((F_AREA, area))
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
        driver.get(url); time.sleep(2)
        cx, cy = get_val(F_X), get_val(F_Y)
        try:
            if abs(float(cx) - float(x)) < 0.0001 and abs(float(cy) - float(y)) < 0.0001:
                return True
        except ValueError:
            pass
        if attempt == 0:
            print(f'   🔁 تحقق فشل (X={cx}, Y={cy}) — إعادة محاولة')
    return False


def main():
    ap = argparse.ArgumentParser(description='إدخال إحداثيات من Excel (كود مباشر أو مطابقة بالاسم)')
    ap.add_argument('--excel', default=os.environ.get('HERITAGE_COORDS_EXCEL', ''))
    ap.add_argument('--decision', default=os.environ.get('HERITAGE_DECISION', ''))
    ap.add_argument('--output', default=os.environ.get('HERITAGE_OUTPUT', '') or 'CoordsRun')
    ap.add_argument('--limit', type=int, default=int(os.environ.get('HERITAGE_COORDS_LIMIT') or 0))
    args = ap.parse_args()

    excel = (args.excel or '').strip()
    if (not excel or not os.path.isfile(excel)):
        folder = os.environ.get('HERITAGE_FOLDER', '').strip()
        if folder and os.path.isdir(folder):
            cands = [f for f in os.listdir(folder)
                     if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
            if cands:
                excel = os.path.join(folder, cands[0])
    if not excel or not os.path.isfile(excel):
        raise SystemExit('❌ حدّد ملف الإحداثيات (Excel).')

    rows = read_rows(excel)
    if args.limit and args.limit > 0:
        rows = rows[:args.limit]
    decision = (args.decision or '').strip()

    # اختيار الوضع تلقائياً: كود المنصة المباشر (الأدق) أو المطابقة بالاسم.
    CODE_MODE = codes_are_real(rows)
    print(f'📄 ملف الإحداثيات: {excel}')
    print(f'📊 عدد المواقع: {len(rows)} | الوضع: '
          + ('كود المنصة المباشر ✅' if CODE_MODE else 'مطابقة بالاسم'))
    print('✍️ تنبيه: تكتب فعلياً على المنصة (إحداثيات + مساحة).')

    out_dir = os.path.abspath(args.output)
    os.makedirs(out_dir, exist_ok=True)
    ckpt = os.path.join(out_dir, '.checkpoint_coords.json')
    try:
        done = set(json.load(open(ckpt, encoding='utf-8')).get('done', []))
    except Exception:
        done = set()

    _R.total(len(rows))
    _R.progress(done=len(done))

    platform = []
    if not CODE_MODE:
        if not decision:
            raise SystemExit('❌ وضع المطابقة بالاسم يحتاج رقم القرار. '
                             'أو أضف عمود «كود الموقع» (رقم المبنى على المنصة) لإدخال مباشر بلا رقم قرار.')

    launch_chrome()
    wait_for_login()

    if not CODE_MODE:
        print('🔎 جمع مباني القرار من المنصة للمطابقة…')
        _R.log('جاري جمع مباني القرار للمطابقة…', 'info')
        platform = collect_platform_buildings(decision)
        print(f'🏢 مباني المنصة: {len(platform)}')
        if not platform:
            raise SystemExit('❌ لم تُجمع مباني من المنصة لهذا القرار — تأكد من الرقم وتسجيل الدخول.')

    ok = fail = skip = 0
    used = set()
    report = []   # (اسم الإكسل، رقم/اسم المنصة، ملاحظة، الحالة)
    for i, row in enumerate(rows, 1):
        if not _R.wait_if_paused() or _R.should_stop():
            print('⏹ إيقاف آمن — التقدّم محفوظ.'); break

        # تحديد رقم المبنى على المنصة
        if CODE_MODE:
            bid = row['code']
            note = ''
            disp = f"{bid} ({row['name'][:24]})" if row['name'] else bid
        else:
            match, score = best_match(row['name'], platform, used)
            if not match:
                fail += 1
                report.append((row['name'], '—', f'{score:.0%}', 'لا تطابق'))
                print(f"[{i}/{len(rows)}] ❌ «{row['name'][:30]}» تشابه {score:.0%} — أقل من الحد")
                _R.log(f"لا تطابق: {row['name'][:28]} ({score:.0%})", 'warn')
                _R.progress(failed=fail, current=row['name'][:28]); continue
            used.add(match['id'])
            bid = match['id']
            note = f'{score:.0%}'
            disp = f"«{row['name'][:22]}»→«{match['name'][:22]}» ({score:.0%}) [{bid}]"

        key = str(bid)
        _R.progress(current=disp[:40])
        if key in done:
            skip += 1
            _R.progress(skipped=skip); continue

        print(f"\n[{i}/{len(rows)}] {disp}  X={row['x']} Y={row['y']}"
              + (f" م={row['area']}" if row['area'] else ''))
        try:
            good = write_coords(bid, row['x'], row['y'], row['area'])
        except Exception as e:
            good = False; print(f'   ❌ {str(e)[:60]}')
        if good:
            ok += 1; done.add(key)
            try:
                json.dump({'done': sorted(done)}, open(ckpt, 'w', encoding='utf-8'), ensure_ascii=False)
            except Exception:
                pass
            report.append((row['name'] or bid, bid, note, 'تم'))
            print('   ✅ تم الإدخال والحفظ'); _R.log(f"✅ {disp[:26]} تم", 'ok')
        else:
            fail += 1
            report.append((row['name'] or bid, bid, note, 'فشل حفظ'))
            print('   ❌ لم يُحفظ'); _R.log(f"❌ {disp[:26]} فشل حفظ", 'error')
        _R.progress(done=len(done), failed=fail, skipped=skip)

    # تقرير Excel
    try:
        rep = openpyxl.Workbook(); rs = rep.active; rs.title = 'تقرير الإحداثيات'
        rs.sheet_view.rightToLeft = True
        rs.append(['اسم الموقع', 'رقم المنصة', 'ملاحظة/تطابق', 'الحالة'])
        for a, b, nt, st in report:
            rs.append([a, b, nt, st])
        rpath = os.path.join(out_dir, f'تقرير_الإحداثيات_{decision or "codes"}.xlsx')
        rep.save(rpath)
        print(f'📊 تقرير: {rpath}')
    except Exception as e:
        print(f'⚠️ تعذّر حفظ التقرير: {e}')

    _R.done(ok=ok, failed=fail, skipped=skip)
    print(f"\n{'='*55}\n  🎉 انتهى\n  ✅ أُدخل: {ok}\n  ⏭️ متخطّى: {skip}\n  ❌ فشل/لا تطابق: {fail}\n{'='*55}")
    try:
        driver.quit()
    except Exception:
        pass


if __name__ == '__main__':
    main()
