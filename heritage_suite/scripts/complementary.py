# ============================================================
# 🏛️ أتمتة منصة التراث — (وارد اتمته متعدده - تكميلي)
# البحث عبر: "الوارد" (Inbox)
# الصفحة: بيانات تكميلية (/12) — إحداثيات المضلع + الحقول الثابتة
# يشتغل على عدة قرى دفعة واحدة (مجلد فيه عدة ملفات Excel)
# ============================================================
# الاستخدام:
#   1) تختار مجلد فيه عدة ملفات Excel (كل ملف اسمه = اسم القرية)
#   2) الكود يسألك عن الارتفاع/المطار/المسافة لكل قرية بالترتيب
#   3) يشتغل قرية قرية، ويطلّع تقرير Excel منفصل لكل قرية
#
# اسم ملف Excel = اسم القرية للبحث في المنصة (بدون الامتداد)
# أعمدة الإحداثيات تُكتشف تلقائياً بالكلمة المفتاحية (الطول/العرض)
# ============================================================

import subprocess, sys, time, re, os
from datetime import datetime

for pkg in ['selenium','pandas','openpyxl','webdriver-manager']:
    subprocess.check_call([sys.executable,'-m','pip','install',pkg,'-q'])

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

import tkinter as tk
from tkinter import filedialog, simpledialog

# ── مربعات الحوار ─────────────────────────────────────────
def ask_string(title, prompt):
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    val = simpledialog.askstring(title=title, prompt=prompt, parent=root)
    root.destroy()
    return val

def ask_folder():
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    val = filedialog.askdirectory(title='اختر مجلد القرى (فيه ملفات Excel)')
    root.destroy()
    return val

# ── 1) اختيار المجلد ──────────────────────────────────────
FOLDER_PATH = ask_folder()
if not FOLDER_PATH: raise ValueError('❌ لم يتم اختيار مجلد')
print(f'📁 المجلد: {FOLDER_PATH}')

# ── 2) إيجاد كل ملفات Excel (باستثناء تقارير سابقة) ───────
excel_files = [
    f for f in os.listdir(FOLDER_PATH)
    if f.lower().endswith(('.xlsx', '.xls'))
    and not f.startswith('~$')
    and not f.lower().startswith('report_')
]
if not excel_files:
    raise FileNotFoundError('❌ لا توجد ملفات Excel في المجلد')

excel_files.sort()
print(f'\n📊 وُجد {len(excel_files)} ملف قرية:')
for f in excel_files:
    print(f'   • {f}  →  قرية: "{os.path.splitext(f)[0]}"')

# ── 3) جمع بيانات كل قرية دفعة واحدة في البداية ───────────
villages_config = []
print('\n📝 الآن أدخل بيانات كل قرية...')
for f in excel_files:
    village_name = os.path.splitext(f)[0]
    print(f'\n──── قرية: {village_name} ────')

    elevation    = ask_string(f'الارتفاع | {village_name}',
                              f'قرية: {village_name}\n\nالارتفاع عن سطح البحر (متر):')
    airport_name = ask_string(f'اسم المطار | {village_name}',
                              f'قرية: {village_name}\n\nاسم أقرب مطار:')
    airport_dist = ask_string(f'المسافة | {village_name}',
                              f'قرية: {village_name}\n\nالمسافة لأقرب مطار (كم):')

    # ── تنظيف القيم من أي أسطر جديدة أو مسافات زائدة ────────
    elevation    = (elevation or '').strip()
    airport_name = (airport_name or '').strip()
    airport_dist = (airport_dist or '').strip()

    villages_config.append({
        'file': f,
        'village': village_name,
        'elevation': elevation,
        'airport_name': airport_name,
        'airport_dist': airport_dist,
    })
    print(f'   ✅ {village_name}: ارتفاع={elevation}, مطار={airport_name}, مسافة={airport_dist}')

print(f'\n✅ تم جمع بيانات {len(villages_config)} قرية — سيبدأ التشغيل الآن')

# ── 4) فتح Chrome ─────────────────────────────────────────
chrome_paths = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]
chrome_exe = next((p for p in chrome_paths if os.path.exists(p)), None)
if not chrome_exe: raise FileNotFoundError('❌ Chrome غير موجود')

subprocess.run(['taskkill', '/f', '/im', 'chrome.exe'],       capture_output=True)
subprocess.run(['taskkill', '/f', '/im', 'chromedriver.exe'], capture_output=True)
time.sleep(2)

subprocess.Popen([
    chrome_exe,'--remote-debugging-port=9222',
    r'--user-data-dir=C:\chrome_debug_session',
    '--no-first-run','--no-default-browser-check','--start-maximized'
])
time.sleep(4)

options = Options()
options.add_experimental_option('debuggerAddress','127.0.0.1:9222')
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)
wait = WebDriverWait(driver, 15)
print('✅ Chrome جاهز')

# ── تحسينات عامة: زر إيقاف + إعادة اتصال + تثبيت تاب ──────
import threading as _threading

# ── تثبيت تاب الأتمتة ─────────────────────────────────────
try:
    AUTOMATION_TAB = driver.current_window_handle
except Exception:
    AUTOMATION_TAB = None

def ensure_automation_tab():
    if AUTOMATION_TAB is None:
        return
    try:
        if driver.current_window_handle != AUTOMATION_TAB:
            driver.switch_to.window(AUTOMATION_TAB)
    except Exception:
        try:
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass

def is_connection_alive():
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False

def reconnect_driver(max_attempts=30, wait_between=5):
    """يعيد ربط Selenium بنفس نافذة كروم الشغّالة (بعد نوم الجهاز مثلاً)."""
    global driver, wait, AUTOMATION_TAB
    print('   🔌 محاولة إعادة الاتصال بكروم...')
    for attempt in range(1, max_attempts + 1):
        try:
            _o = Options()
            _o.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=_o)
            wait = WebDriverWait(driver, 15)
            _ = driver.current_url
            try:
                AUTOMATION_TAB = driver.current_window_handle
            except Exception:
                AUTOMATION_TAB = None
            print(f'   ✅ تمت إعادة الاتصال بنجاح (محاولة {attempt})')
            return True
        except Exception as e:
            print(f'   ⏳ محاولة {attempt}/{max_attempts} فشلت — بعد {wait_between}ث... ({str(e)[:40]})')
            time.sleep(wait_between)
    print('   ❌ تعذّر إعادة الاتصال بكروم')
    return False

def ensure_alive():
    if is_connection_alive():
        return True
    print('\n   ⚠️ انقطع الاتصال بكروم (نوم الجهاز؟) — جاري إعادة الاتصال...')
    return reconnect_driver()

# ── زر الإيقاف المؤقت العائم ──────────────────────────────
_pause_flag = {'paused': False, 'stopped_now': False, 'status_widget': None, 'win': None}

def _set_status(text, color='#333333'):
    w = _pause_flag.get('status_widget')
    if w is not None:
        try:
            w.config(text=text, fg=color)
        except Exception:
            pass

def _start_pause_button():
    def run():
        import tkinter as _tk
        win = _tk.Tk()
        win.title('تحكم')
        win.attributes('-topmost', True)
        win.geometry('240x92+40+40')
        win.resizable(False, False)
        try:
            win.attributes('-alpha', 0.95)
        except Exception:
            pass
        def toggle():
            _pause_flag['paused'] = not _pause_flag['paused']
            if _pause_flag['paused']:
                btn.config(text='▶ استئناف', bg='#c0392b', fg='white')
                status.config(text='⏳ سيتوقف بعد المبنى الحالي...', fg='#c0392b')
            else:
                _pause_flag['stopped_now'] = False
                btn.config(text='⏸ إيقاف مؤقت', bg='#27ae60', fg='white')
                status.config(text='▶ شغّال', fg='#27ae60')
        btn = _tk.Button(win, text='⏸ إيقاف مؤقت', command=toggle,
                         bg='#27ae60', fg='white',
                         font=('Arial', 13, 'bold'), height=1)
        btn.pack(fill='x', padx=6, pady=(6, 2))
        status = _tk.Label(win, text='▶ شغّال', font=('Arial', 9, 'bold'), fg='#27ae60')
        status.pack(pady=(0, 4))
        _pause_flag['status_widget'] = status
        _pause_flag['win'] = win
        win.mainloop()
    _threading.Thread(target=run, daemon=True).start()

def close_pause_button():
    """يغلق نافذة الإيقاف تلقائياً عند انتهاء الكود."""
    w = _pause_flag.get('win')
    if w is not None:
        try:
            w.after(0, w.destroy)
        except Exception:
            pass


def wait_if_paused():
    was_paused = False
    while _pause_flag['paused']:
        if not was_paused:
            _pause_flag['stopped_now'] = True
            _set_status('✅ متوقف — تقدر تروح الآن', '#c0392b')
            print('\n⏸️  ✅ توقف فعلي — خلّص المبنى الحالي. تقدر تروح الآن. اضغط "استئناف" لما ترجع.')
            was_paused = True
        time.sleep(0.5)
    if was_paused:
        _set_status('🔌 جاري فحص الاتصال...', '#e67e22')
        print('▶️  تم الاستئناف — نفحص الاتصال بكروم...')
        ensure_alive()
        _set_status('▶ شغّال', '#27ae60')

_start_pause_button()
print('🟢 نافذة التحكم بالإيقاف المؤقت جاهزة (تطفو فوق كل النوافذ)')


# ── 5) الدخول للمنصة (مرة واحدة) ─────────────────────────
driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
time.sleep(3)
print(f'✅ الصفحة: {driver.title}')

if driver.title == 'Runtime Error' or 'Login' in driver.current_url:
    driver.get('https://heritage2.itqan-consultant.com')
    time.sleep(3)
    input('\n👤 سجّل دخولك الآن في نافذة Chrome، ثم اضغط Enter هنا للمتابعة...')
    driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
    time.sleep(3)
    print(f'✅ الصفحة بعد الدخول: {driver.title}')

# ── 6) الدوال المساعدة ────────────────────────────────────
def js_click(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", el)

def fill_field(name_contains, value):
    try:
        inp = driver.find_element('xpath',
            f"//input[contains(@name,'{name_contains}')]")
        # ── تنظيف القيمة من أي \n أو \r قبل الكتابة ─────────
        clean_value = str(value).replace('\n', '').replace('\r', '').strip()
        driver.execute_script("arguments[0].value = '';", inp)
        time.sleep(0.2)
        inp.clear(); time.sleep(0.2)
        inp.send_keys(clean_value); time.sleep(0.2)
        try:
            row_text = inp.find_element('xpath',
                './ancestor::tr[1]').text.strip()[:40]
        except: row_text=''
        print(f'   ✅ "{row_text}" = {clean_value}')
        return True
    except Exception as e:
        print(f'   ❌ "{name_contains}": {str(e)[:50]}')
        return False

def clean_coords(coords):
    if len(coords) >= 2:
        if coords[0]['x']==coords[-1]['x'] and coords[0]['y']==coords[-1]['y']:
            print('   ℹ️ آخر نقطة = أول نقطة — سيتم تخطيها')
            return coords[:-1]
    return coords

def find_col(df, keyword):
    for col in df.columns:
        if keyword in col:
            return col
    return None

def get_all_buildings():
    all_buildings=[]; page_num=1; seen=set()
    while True:
        time.sleep(2)
        # كل الصفوف اللي فيها رابط — ونفلتر لاحقاً على صفوف النتائج فقط
        rows = driver.find_elements('xpath', "//table//tr[.//a[@href]]")
        for row in rows:
            try:
                cells = row.find_elements('xpath', './/td')
                # صف نتيجة حقيقي لازم فيه خلية "كود" رقمية (3+ أرقام) مثل 38436
                # هذا يميّز صفوف النتائج عن روابط القوائم الجانبية والترويسة
                code_num = None
                for cell in cells:
                    ct = (cell.text or '').strip()
                    if ct.isdigit() and len(ct) >= 3:
                        code_num = ct; break
                if code_num is None:
                    continue

                text = row.text.strip()

                # اسم الموقع = أول رابط نصه يحتوي حروف (عربي/لاتيني)
                site_name = ''; link = ''
                for a in row.find_elements('xpath', './/a[@href]'):
                    atext = (a.text or '').strip()
                    if atext and re.search(r'[\u0600-\u06FFA-Za-z]', atext):
                        site_name = atext
                        link = a.get_attribute('href') or ''
                        break
                if not link:
                    try:
                        link = row.find_element('xpath', './/a').get_attribute('href')
                    except Exception:
                        link = ''
                if not link:
                    continue

                # رقم المبنى (إن وُجد "مبنى/موقع رقم X") — للمطابقة بالرقم
                bnum = None
                m = re.search(r'(?:مبن[ىي]|موقع)\s+رقم\s+(\d+)', text)
                if m:
                    bnum = int(m.group(1))

                if link in seen:
                    continue
                seen.add(link)
                all_buildings.append({'number':bnum, 'name':site_name, 'link':link})
            except Exception:
                continue
        next_link = driver.find_elements('xpath',
            f"//table//tr//td//a[text()='{page_num+1}'] | "
            f"//a[contains(@href,'Page${page_num+1}')]")
        if next_link: js_click(next_link[0]); page_num+=1
        else: break
    return all_buildings


# ── أدوات المطابقة بالاسم (لما يكون coded أسماء لا أرقام) ──
def normalize_name(s):
    """توحيد الاسم العربي للمطابقة: يشيل التشكيل، يوحّد الألف/الياء/التاء،
    يشيل كلمات شائعة مختلفة (التاريخي، موقع، قرية...) والمسافات."""
    s = str(s).strip()
    s = re.sub(r'[\u064b-\u0652]', '', s)              # تشكيل
    s = s.replace('أ','ا').replace('إ','ا').replace('آ','ا')
    s = s.replace('ة','ه').replace('ى','ي')
    for w in ['التاريخيه','التاريخي','قريه','قرية','موقع','منطقه','منطقة','حي']:
        s = s.replace(w, '')
    s = s.replace('ـ','')                               # تطويل
    s = re.sub(r'\s+', '', s)                           # كل المسافات
    return s.strip()


def name_similarity(a, b):
    """نسبة تشابه بين اسمين بعد التوحيد (0..1)."""
    from difflib import SequenceMatcher
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def best_name_match(excel_name, platform_buildings, threshold=0.72):
    """يرجّع أفضل مبنى من المنصة يطابق اسم الإكسل (أو None)."""
    best = None; best_score = 0.0
    for b in platform_buildings:
        score = name_similarity(excel_name, b.get('name', ''))
        if score > best_score:
            best_score = score; best = b
    if best is not None and best_score >= threshold:
        return best, best_score
    return None, best_score


def coded_is_numeric(series):
    """يكتشف هل عمود coded أرقام (True) أو أسماء (False)."""
    ok = 0; total = 0
    for v in series.dropna().unique():
        total += 1
        try:
            int(float(str(v).strip()))
            ok += 1
        except Exception:
            pass
    return total > 0 and (ok / total) >= 0.7


def js_fill_verify(el, value, tries=4):
    """
    يكتب القيمة في خانة الـ popup عبر JavaScript (يشتغل حتى لو النافذة مخفية
    أو الشاشة مقفولة)، ثم يقرأ القيمة من الخانة ويتأكد أنها انكتبت فعلاً.
    يرجّع True فقط لو القيمة موجودة فعلياً في الخانة بعد الكتابة.
    """
    target = str(value).replace('\n', '').replace('\r', '').strip()
    for _ in range(tries):
        try:
            driver.execute_script("""
                var el = arguments[0], val = arguments[1];
                el.focus();
                el.value = '';
                el.value = val;
                el.dispatchEvent(new Event('input',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.dispatchEvent(new Event('keyup',  {bubbles:true}));
                el.dispatchEvent(new Event('blur',   {bubbles:true}));
            """, el, target)
            time.sleep(0.25)
            actual = (el.get_attribute('value') or '').strip()
            if actual == target:
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def process_building(building_id, building_num, coords_by_building,
                     elevation, airport_name, airport_dist):
    ensure_automation_tab()
    coords = clean_coords(coords_by_building[building_num])
    base_url = (
        f'https://heritage2.itqan-consultant.com'
        f'/Web/App/Pools/DataEdit/{building_id}/12'
    )
    print(f'\n🏢 مبنى {building_num} | {len(coords)} إحداثي')
    driver.get(base_url); time.sleep(3)

    print('   📝 تعبئة الحقول الثابتة...')
    fill_field('ctl06$Field1$ctl03', elevation)
    fill_field('ctl13$Field1$ctl03', airport_name)
    fill_field('ctl14$Field1$ctl03', airport_dist)

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1.5)
    # أقسام "لا يوجد" (2-8): تُكتب لو الحقل موجود، وتُتجاهل بهدوء لو غير موجود
    # (المواقع المسمّاة أحياناً صفحتها بلا هذه الأقسام)
    for s in ['ctl46','ctl47','ctl48','ctl49','ctl50','ctl51','ctl52']:
        try:
            _els = driver.find_elements('xpath',
                f"//input[contains(@name,'{s}$Field1$ctl03')]")
            if _els:
                fill_field(f'{s}$Field1$ctl03', 'لا يوجد')
                time.sleep(0.2)
        except Exception:
            pass

    driver.execute_script("window.scrollTo(0, 0);"); time.sleep(1)
    try:
        js_click(driver.find_element('xpath',"//input[@name='ctl12$btnSave']"))
        print('   💾 تم حفظ بيانات الاستمارة')
        time.sleep(4)
    except Exception as e:
        print(f'   ❌ خطأ في الحفظ: {e}'); return False

    driver.get(base_url); time.sleep(3)
    print('   🔄 تم إعادة التحميل')

    success_count = 0
    for point_num, coord in enumerate(coords, start=1):
        print(f'   ➕ النقطة {point_num}/{len(coords)}: '
              f'عرض={coord["y"]}, طول={coord["x"]}')
        saved = False; attempts = 0

        while not saved and attempts < 3:
            attempts += 1
            try:
                rows_before = len(driver.find_elements('xpath',
                    "//tr[ancestor::tr[.//span[contains("
                    "text(),'الاحداثيات الجغرافية')]]]"
                    "[.//a[contains(text(),'تعديل')]]"
                ))

                add_btn = driver.find_element('xpath',
                    "//a[contains(@href,'btnAdd') and "
                    "ancestor::tr[.//span[contains("
                    "text(),'الاحداثيات الجغرافية')]]]")
                js_click(add_btn); time.sleep(2)

                modal = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((
                        By.XPATH,
                        "//div[contains(@class,'modal-content')]")))
                time.sleep(0.5)

                f_point = modal.find_element('xpath', ".//input[@type='number']")
                f_lat   = modal.find_element('xpath', "(.//input[@type='text'])[1]")
                f_lon   = modal.find_element('xpath', "(.//input[@type='text'])[2]")

                # ── كتابة القيم عبر JS + تأكيد أنها انكتبت فعلاً ──
                ok_point = js_fill_verify(f_point, point_num)
                ok_lat   = js_fill_verify(f_lat,   coord['y'])
                ok_lon   = js_fill_verify(f_lon,   coord['x'])
                time.sleep(0.2)

                # لو أي خانة ما انكتبت فعلياً، لا نحفظ — نغلق ونعيد المحاولة
                if not (ok_point and ok_lat and ok_lon):
                    print(f'      ⚠️ محاولة {attempts}/3: لم تتأكد كتابة القيم '
                          f'(نقطة={ok_point}, Y={ok_lat}, X={ok_lon}) — إعادة')
                    try:
                        close = driver.find_elements('xpath',
                            "//input[@name="
                            "'ctl12$TemplatePage$DialogTemplate$btnClose']")
                        if close: js_click(close[0])
                        time.sleep(1)
                    except Exception: pass
                    continue

                save_modal = modal.find_element('xpath',
                    ".//input[@name="
                    "'ctl12$TemplatePage$DialogTemplate$btnSave']")
                js_click(save_modal)

                confirmed = False
                for _ in range(20):
                    time.sleep(0.5)
                    try:
                        rows_after = len(driver.find_elements('xpath',
                            "//tr[ancestor::tr[.//span[contains("
                            "text(),'الاحداثيات الجغرافية')]]]"
                            "[.//a[contains(text(),'تعديل')]]"
                        ))
                        if rows_after > rows_before:
                            confirmed = True
                            break
                    except Exception:
                        pass

                if confirmed:
                    print(f'      ✅ تم تأكيد حفظ النقطة {point_num}')
                    saved = True; success_count += 1
                else:
                    print(f'      ⚠️ محاولة {attempts}/3: لم يتم تأكيد الحفظ')
                    try:
                        close = driver.find_elements('xpath',
                            "//input[@name="
                            "'ctl12$TemplatePage$DialogTemplate$btnClose']")
                        if close: js_click(close[0])
                        time.sleep(1)
                    except Exception: pass

            except Exception as e:
                print(f'      ⚠️ محاولة {attempts}/3 فشلت: {str(e)[:50]}')
                try:
                    close = driver.find_elements('xpath',
                        "//input[@name="
                        "'ctl12$TemplatePage$DialogTemplate$btnClose']")
                    if close: js_click(close[0])
                    time.sleep(2)
                except Exception: pass

        if not saved:
            print(f'      ❌ فشل حفظ النقطة {point_num} بعد 3 محاولات')

    print(f'   📊 الإحداثيات: {success_count}/{len(coords)}')
    return success_count == len(coords)

# ── 7) الحلقة الرئيسية على كل قرية ────────────────────────
grand_total = {'villages': 0, 'processed': 0, 'skipped': 0, 'failed': 0}

for v_idx, cfg in enumerate(villages_config, start=1):
    village    = cfg['village']
    excel_path = os.path.join(FOLDER_PATH, cfg['file'])

    print(f'\n{"█"*55}')
    print(f'█  قرية {v_idx}/{len(villages_config)}: {village}')
    print(f'{"█"*55}')

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f'❌ فشل قراءة ملف {cfg["file"]}: {e}')
        continue

    x_col     = find_col(df, 'الطول')
    y_col     = find_col(df, 'العرض')
    coded_col = 'coded'

    if not x_col or not y_col:
        print(f'❌ أعمدة الإحداثيات غير موجودة في {cfg["file"]} — الأعمدة: {list(df.columns)}')
        continue

    # ── اكتشاف: عمود coded أرقام أو أسماء؟ ────────────────
    NAME_MODE = not coded_is_numeric(df[coded_col])
    print('🔤 وضع المطابقة: ' + ('بالاسم (coded أسماء)' if NAME_MODE
                                  else 'بالرقم (coded أرقام)'))

    # coords_by_building: المفتاح رقم (وضع الرقم) أو الاسم الخام (وضع الاسم)
    coords_by_building = {}
    for coded_val, group in df.groupby(coded_col):
        if NAME_MODE:
            key = str(coded_val).strip()
            if not key:
                continue
        else:
            try:
                key = int(coded_val)
            except (ValueError, TypeError):
                continue
        coords = []
        for _, row in group.iterrows():
            coords.append({'x':float(row[x_col]),'y':float(row[y_col])})
        coords_by_building[key] = coords

    print(f'🏢 المواقع في Excel: {len(coords_by_building)}')

    # ── البحث في الوارد (Inbox) ───────────────────────────
    driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
    time.sleep(3)
    ensure_automation_tab(); ensure_alive()
    search = driver.find_element('xpath',"//input[@name='ctl12$ctl04']")
    driver.execute_script("""
        var el=arguments[0], val=arguments[1];
        el.value=val;
        el.dispatchEvent(new Event('input',{bubbles:true}));
        el.dispatchEvent(new Event('change',{bubbles:true}));
    """, search, village)
    time.sleep(0.3)
    _btn = driver.find_element('xpath',"//input[@name='ctl12$btnSearch']")
    driver.execute_script("arguments[0].click();", _btn)
    time.sleep(3)
    print(f'🔍 تم البحث عن: {village}')

    buildings = get_all_buildings()
    print(f'🏢 المباني في المنصة: {len(buildings)}')

    if not buildings:
        print(f'⚠️ لم تُوجد مباني للقرية "{village}" — سيتم تخطيها')
        continue

    results_log = []
    def log(bnum, status, msg=''):
        ts   = datetime.now().strftime('%H:%M:%S')
        icon = {'success':'✅','skip':'⏭️','error':'❌'}[status]
        print(f'[{ts}] {icon} مبنى {bnum}: {msg}')
        results_log.append({
            'رقم المبنى':bnum,'الحالة':status,
            'التفاصيل':msg,'الوقت':ts
        })

    processed=skipped=failed=0

    # ── نبني قائمة العمل: لكل مفتاح في الإكسل، نلقى مبناه في المنصة ──
    #    وضع الرقم: المطابقة برقم المبنى. وضع الاسم: تطابق ذكي بالاسم.
    worklist = []   # كل عنصر: (excel_key, platform_building, label)
    if NAME_MODE:
        used_links = set()
        for excel_name in coords_by_building.keys():
            candidates = [b for b in buildings if b['link'] not in used_links]
            match, score = best_name_match(excel_name, candidates)
            if match:
                used_links.add(match['link'])
                worklist.append((excel_name, match,
                                 f'{excel_name}  →  "{match["name"]}" (تطابق {score:.0%})'))
            else:
                worklist.append((excel_name, None,
                                 f'{excel_name} (أفضل تشابه {score:.0%} — أقل من الحد)'))
    else:
        for b in buildings:
            worklist.append((b['number'], b, None))

    total_items = len(worklist)
    for idx, (excel_key, b, label) in enumerate(worklist):
        # ── نقطة آمنة: إيقاف مؤقت + فحص الاتصال قبل كل مبنى ──
        wait_if_paused()
        ensure_alive()

        disp = excel_key if NAME_MODE else f'مبنى {excel_key}'
        print(f'\n━━━ [{village}] {disp} ({idx+1}/{total_items}) ━━━')
        if label:
            print(f'   🔗 {label}')

        if b is None:
            log(excel_key, 'skip', 'لا يوجد تطابق في المنصة')
            skipped+=1; continue
        if excel_key not in coords_by_building:
            log(excel_key, 'skip', 'غير موجود في Excel')
            skipped+=1; continue

        bid = b['link'].rstrip('/').split('/')[-1]

        try:
            ok = process_building(bid, excel_key, coords_by_building,
                                  cfg['elevation'], cfg['airport_name'], cfg['airport_dist'])
            n  = len(clean_coords(coords_by_building[excel_key]))
            log(excel_key,'success' if ok else 'error',
                f'تم إدخال {n} إحداثي' if ok else 'لم تكتمل')
            if ok: processed+=1
            else:  failed+=1
        except Exception as e:
            emsg = str(e)
            if ('HTTPConnection' in emsg or 'Max retries' in emsg
                    or 'connection' in emsg.lower() or 'refused' in emsg.lower()
                    or not is_connection_alive()):
                print('   ⚠️ يبدو أن الاتصال انقطع — إعادة اتصال ومحاولة ثانية...')
                if reconnect_driver():
                    try:
                        ok = process_building(bid, excel_key, coords_by_building,
                                              cfg['elevation'], cfg['airport_name'], cfg['airport_dist'])
                        n  = len(clean_coords(coords_by_building[excel_key]))
                        log(excel_key,'success' if ok else 'error',
                            (f'تم إدخال {n} إحداثي (بعد إعادة اتصال)') if ok else 'لم تكتمل')
                        if ok: processed+=1
                        else:  failed+=1
                        continue
                    except Exception as e2:
                        log(excel_key,'error',f'فشل حتى بعد إعادة الاتصال: {str(e2)[:40]}')
                        failed+=1; continue
            log(excel_key,'error',emsg[:60]); failed+=1

    print(f'\n{"─"*50}')
    print(f'📊 قرية {village}: ✅ {processed}  ⏭️ {skipped}  ❌ {failed}')
    print(f'{"─"*50}')

    safe_name = re.sub(r'[\\/*?:"<>|]', '_', village)
    report_path = os.path.join(
        FOLDER_PATH,
        f'report_{safe_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )
    pd.DataFrame(results_log).to_excel(report_path, index=False)
    print(f'💾 تقرير القرية: {report_path}')

    grand_total['villages']  += 1
    grand_total['processed'] += processed
    grand_total['skipped']   += skipped
    grand_total['failed']    += failed

# ── 8) التقرير الإجمالي النهائي ───────────────────────────
print(f'\n{"█"*55}')
print(f'█  انتهت جميع القرى')
print(f'█  عدد القرى: {grand_total["villages"]}')
print(f'█  ✅ نجح: {grand_total["processed"]}  '
      f'⏭️ تخطى: {grand_total["skipped"]}  '
      f'❌ فشل: {grand_total["failed"]}')
print(f'{"█"*55}')

# ── إغلاق نافذة الإيقاف تلقائياً بعد انتهاء الكود ──────────
try:
    close_pause_button()
    print('🔚 تم إغلاق نافذة الإيقاف — انتهى العمل')
except Exception:
    pass
