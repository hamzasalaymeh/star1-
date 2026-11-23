# ============================================================
# 🏛️ أتمتة منصة التراث — (كروكي + مساحة + احداثي) — متعدد القرى
# صفحة: "استكمال بيانات الميدان" — حقول X / Y / المساحة + رفع صورة الكروكي
# ============================================================
# الاستخدام:
#   1) تختار مجلد أب واحد — جوّاه مجلدات القرى (كل مجلد فيه Excel + صور)
#      اسم مجلد القرية = اسم القرية للبحث في المنصة
#   2) تدخل رقم قرار التسجيل مرة واحدة (نفسه لكل القرى) — اختياري
#   3) يعرض لك قائمة القرى، وتقدر تحدّد range مباني لكل قرية (اختياري)
#      أمثلة: 3,7,12  أو  50:70  أو  3,7,50:70  — فاضي = كل المباني
#   4) يشتغل قرية قرية، وفي النهاية يطلّع ملف Excel واحد فيه شيت
#      لكل قرية يعرض المباني اللي فشلت أو اتخطّت (المحتاجة متابعة)
# ============================================================

import subprocess, sys, time, re, os, threading
from datetime import datetime

for pkg in ['selenium', 'pandas', 'openpyxl', 'webdriver-manager']:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '-q'])

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import filedialog, simpledialog


def ask_string(title, prompt):
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    val = simpledialog.askstring(title=title, prompt=prompt, parent=root)
    root.destroy()
    return val


def ask_parent_folder():
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    val = filedialog.askdirectory(
        title='اختر المجلد الأب (جوّاه مجلدات القرى)')
    root.destroy()
    return val


def ask_village_folder(n):
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    val = filedialog.askdirectory(
        title=f'اختر مجلد القرية رقم {n} (فيه الإكسل + الصور)')
    root.destroy()
    return val


def ask_yes_no(title, prompt):
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    ans = messagebox.askyesno(title=title, message=prompt)
    root.destroy()
    return ans


def ask_mode():
    """يخيّر بين وضع المجلد الأب أو اختيار مجلدات القرى وحدة وحدة.
    يرجّع True لوضع المجلد الأب، False لوضع اختيار كل قرية على حدة."""
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    ans = messagebox.askyesno(
        'طريقة الاختيار',
        'كيف تريد اختيار القرى؟\n\n'
        '• نعم = مجلد أب واحد (جوّاه مجلدات كل القرى) — تلقائي\n'
        '• لا  = أختار مجلد كل قرية بنفسي (واحدة واحدة)')
    root.destroy()
    return ans


def parse_range(text):
    """يحوّل نص مثل '3,7,50:70' لقائمة أرقام. فاضي → None (كل المباني)."""
    if not text or not text.strip():
        return None
    try:
        out = []
        for part in text.replace('،', ',').split(','):
            part = part.strip()
            if ':' in part:
                a, b = part.split(':')
                out.extend(range(int(a.strip()), int(b.strip()) + 1))
            elif part.isdigit():
                out.append(int(part))
        return sorted(set(out)) if out else None
    except Exception:
        return None


# ── 1) اختيار طريقة الإدخال ثم القرى ──────────────────────
# وضعان: (أ) مجلد أب فيه مجلدات قرى  (ب) اختيار ملفات Excel مباشرة
# في كلا الوضعين: اسم القرية = اسم المجلد اللي فيه ملف الـ Excel،
# والصور بنفس مجلد الملف.
# لو شُغّل من مركز التحكم: المجلد يُحدَّد من اللوحة (HERITAGE_FOLDER) فنتخطّى النوافذ.
_HUB_FOLDER = os.environ.get("HERITAGE_FOLDER", "").strip()
USE_PARENT = True if _HUB_FOLDER else ask_mode()

village_folders = []   # قائمة مجلدات القرى (اسم المجلد = اسم القرية)

if USE_PARENT:
    # ── وضع (أ): مجلد أب واحد جوّاه مجلدات القرى ───────────
    PARENT = _HUB_FOLDER or ask_parent_folder()
    if not PARENT:
        raise ValueError('❌ لم يتم اختيار مجلد')
    print(f'📁 المجلد الأب: {PARENT}')

    for name in sorted(os.listdir(PARENT)):
        full = os.path.join(PARENT, name)
        if not os.path.isdir(full):
            continue
        has_excel = any(
            f.lower().endswith(('.xlsx', '.xls'))
            and not f.startswith('~$')
            and not f.lower().startswith('report_')
            and not f.startswith('تقرير_')
            for f in os.listdir(full)
        )
        if has_excel:
            village_folders.append(full)

    if not village_folders:
        raise FileNotFoundError('❌ لا توجد مجلدات قرى (فيها Excel) داخل المجلد الأب')
    REPORTS_DIR = PARENT

else:
    # ── وضع (ب): اختيار مجلد كل قرية على حدة (مع خيار الإضافة) ──
    while True:
        fld = ask_village_folder(len(village_folders) + 1)
        if fld:
            fld = fld.rstrip('/\\')
            if fld in village_folders:
                print(f'⚠️ المجلد مضاف مسبقاً: {os.path.basename(fld)}')
            else:
                # نتأكد إن فيه ملف Excel
                has_excel = any(
                    f.lower().endswith(('.xlsx', '.xls'))
                    and not f.startswith('~$')
                    and not f.lower().startswith('report_')
                    and not f.startswith('تقرير_')
                    for f in os.listdir(fld)
                )
                if has_excel:
                    village_folders.append(fld)
                    print(f'✅ [{len(village_folders)}] {os.path.basename(fld)}')
                else:
                    print(f'⚠️ لا يوجد ملف Excel في: {os.path.basename(fld)} — لم يُضَف')

        if not ask_yes_no('إضافة قرية',
                          f'تم اختيار {len(village_folders)} قرية حتى الآن.\n\n'
                          'تريد إضافة مجلد قرية آخر؟'):
            break

    if not village_folders:
        raise ValueError('❌ لم يتم اختيار أي مجلد قرية صالح')

    # مجلد التقارير = المجلد الأب لأول قرية
    REPORTS_DIR = os.path.dirname(village_folders[0].rstrip('/\\')) or village_folders[0]

print(f'\n📊 وُجد {len(village_folders)} قرية:')
for i, fld in enumerate(village_folders, 1):
    print(f'   {i}. {os.path.basename(fld.rstrip(chr(47)+chr(92)))}')

# ── 2) رقم قرار التسجيل (مرة واحدة لكل القرى) ──────────────
if _HUB_FOLDER:
    REG_DECISION = os.environ.get("HERITAGE_DECISION", "").strip()
else:
    REG_DECISION = ask_string(
        'رقم قرار التسجيل',
        'رقم قرار التسجيل (نفسه لكل القرى — اتركه فاضي لو ما يلزم):')
REG_DECISION = (REG_DECISION or '').strip()
print(f'📋 رقم قرار التسجيل: {REG_DECISION or "(فاضي)"}')

# ── 3) اختيار range المباني لكل قرية (اختياري بالكامل) ────
villages_config = []

# من المركز: كل المباني (بدون أسئلة نطاقات). يدوياً: نسأل.
want_ranges = False if _HUB_FOLDER else ask_yes_no(
    'تحديد مباني',
    'تريد تحديد نطاق مباني معيّنة لأي قرية؟\n\n'
    '• نعم = يسألك عن كل قرية على حدة\n'
    '• لا  = يشتغل على كل مباني كل القرى (بدون أسئلة)'
)

for fld in village_folders:
    village_name = os.path.basename(fld.rstrip('/\\'))
    filter_buildings = None

    if want_ranges:
        rng_text = ask_string(
            f'مباني القرية | {village_name}',
            f'قرية: {village_name}\n\n'
            'حدّد أرقام/نطاق المباني المطلوبة:\n'
            'مثال: 3, 7, 12\n'
            'مثال: 50:70\n'
            'مثال: 3, 7, 50:70\n\n'
            'اتركه فاضياً للعمل على كل مباني هذه القرية:')
        filter_buildings = parse_range(rng_text)

    villages_config.append({
        'folder': fld,
        'village': village_name,
        'filter': filter_buildings,
    })
    if filter_buildings:
        print(f'   🎯 {village_name}: {filter_buildings}')
    else:
        print(f'   ✅ {village_name}: كل المباني')

print(f'\n✅ تم إعداد {len(villages_config)} قرية — سيبدأ التشغيل الآن')

# (مجلد التقارير REPORTS_DIR تم تحديده في قسم الاختيار أعلاه)

# ── 4) فتح Chrome ─────────────────────────────────────────
chrome_paths = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]
chrome_exe = next((p for p in chrome_paths if os.path.exists(p)), None)
if not chrome_exe:
    raise FileNotFoundError('❌ Chrome غير موجود')

subprocess.run(['taskkill', '/f', '/im', 'chrome.exe'],       capture_output=True)
subprocess.run(['taskkill', '/f', '/im', 'chromedriver.exe'], capture_output=True)
time.sleep(2)

subprocess.Popen([
    chrome_exe, '--remote-debugging-port=9222',
    r'--user-data-dir=C:\chrome_debug_session',
    '--no-first-run', '--no-default-browser-check', '--start-maximized'
])
time.sleep(4)

options = Options()
options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)
wait = WebDriverWait(driver, 15)

# ── تثبيت التاب الخاص بالأتمتة (حتى لو فتحت تابات ثانية) ──
try:
    AUTOMATION_TAB = driver.current_window_handle
    print(f'📌 تم تثبيت تاب الأتمتة')
except Exception:
    AUTOMATION_TAB = None

def ensure_automation_tab():
    """يتأكد أن الأوامر تُنفّذ على تاب الأتمتة، حتى لو فتحت تابات أخرى."""
    if AUTOMATION_TAB is None:
        return
    try:
        if driver.current_window_handle != AUTOMATION_TAB:
            driver.switch_to.window(AUTOMATION_TAB)
    except Exception:
        # لو التاب الأصلي أُغلق، نلتقط أول تاب متاح
        try:
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass


def is_connection_alive():
    """يفحص هل قناة Selenium بكروم لا تزال شغّالة."""
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def reconnect_driver(max_attempts=30, wait_between=5):
    """
    يعيد ربط Selenium بنفس نافذة كروم الشغّالة (بعد نوم الجهاز مثلاً)
    بدون فتح كروم من جديد — فتبقى الجلسة على المنصة محفوظة.
    يحاول حتى max_attempts مرة.
    """
    global driver, wait, AUTOMATION_TAB
    print('   🔌 محاولة إعادة الاتصال بكروم...')
    for attempt in range(1, max_attempts + 1):
        try:
            new_options = Options()
            new_options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=new_options
            )
            wait = WebDriverWait(driver, 15)
            # نتأكد إن الاتصال فعلاً شغّال
            _ = driver.current_url
            try:
                AUTOMATION_TAB = driver.current_window_handle
            except Exception:
                AUTOMATION_TAB = None
            print(f'   ✅ تمت إعادة الاتصال بنجاح (محاولة {attempt})')
            return True
        except Exception as e:
            print(f'   ⏳ محاولة {attempt}/{max_attempts} فشلت — '
                  f'إعادة المحاولة بعد {wait_between}ث... ({str(e)[:40]})')
            time.sleep(wait_between)
    print('   ❌ تعذّر إعادة الاتصال بكروم بعد كل المحاولات')
    return False


def ensure_alive():
    """يفحص الاتصال، ولو منقطع يحاول يعيده. يرجّع True لو أصبح شغّالاً."""
    if is_connection_alive():
        return True
    print('\n   ⚠️ انقطع الاتصال بكروم (نوم الجهاز؟) — جاري إعادة الاتصال...')
    return reconnect_driver()

print('✅ تم فتح Chrome والاتصال به')

# ── زر الإيقاف المؤقت العائم (يطفو فوق كل النوافذ) ─────────
# حالة مشتركة:
#   paused       = المستخدم طلب إيقاف
#   stopped_now  = وصلنا فعلاً لنقطة التوقف الآمنة (خلّص المبنى)
_pause_flag = {'paused': False, 'stopped_now': False, 'status_widget': None, 'win': None}

def _set_status(text, color='#333333'):
    """يحدّث نص الحالة داخل نافذة الزر العائمة."""
    w = _pause_flag.get('status_widget')
    if w is not None:
        try:
            w.config(text=text, fg=color)
        except Exception:
            pass

def _start_pause_button():
    """يشغّل نافذة زر عائمة في thread منفصل — لا تجمّد الأتمتة."""
    def run():
        import tkinter as _tk
        win = _tk.Tk()
        win.title('تحكم')
        win.attributes('-topmost', True)      # دائماً فوق كل النوافذ
        win.geometry('240x92+40+40')          # حجم وموضع ابتدائي (الزاوية)
        win.resizable(False, False)
        try:
            win.attributes('-alpha', 0.95)    # شفافية خفيفة
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

        status = _tk.Label(win, text='▶ شغّال', font=('Arial', 9, 'bold'),
                           fg='#27ae60')
        status.pack(pady=(0, 4))

        _pause_flag['status_widget'] = status
        _pause_flag['win'] = win
        win.mainloop()

    t = threading.Thread(target=run, daemon=True)
    t.start()

def close_pause_button():
    """يغلق نافذة الإيقاف تلقائياً عند انتهاء الكود."""
    w = _pause_flag.get('win')
    if w is not None:
        try:
            w.after(0, w.destroy)
        except Exception:
            pass


def wait_if_paused():
    """يوقف التنفيذ طالما زر الإيقاف مفعّل — يُستدعى عند نقاط آمنة."""
    was_paused = False
    while _pause_flag['paused']:
        if not was_paused:
            # وصلنا فعلاً لنقطة التوقف الآمنة (بعد ما خلّص المبنى الحالي)
            _pause_flag['stopped_now'] = True
            _set_status('✅ متوقف — تقدر تروح الآن', '#c0392b')
            print('\n⏸️  ✅ توقف فعلي — خلّص المبنى الحالي. تقدر تروح الآن. '
                  'اضغط "استئناف" لما ترجع.')
            was_paused = True
        time.sleep(0.5)
    if was_paused:
        _set_status('🔌 جاري فحص الاتصال...', '#e67e22')
        print('▶️  تم الاستئناف — نفحص الاتصال بكروم...')
        ensure_alive()
        _set_status('▶ شغّال', '#27ae60')

_start_pause_button()
print('🟢 نافذة التحكم بالإيقاف المؤقت جاهزة (تطفو فوق كل النوافذ)')

# ── 5) الدوال المساعدة ─────────────────────────────────────
def js_click(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", el)


def find_col(df, keyword):
    for col in df.columns:
        if keyword in col:
            return col
    return None


def fill_field(name_contains, value):
    """
    يكتب القيمة عبر JavaScript مباشرة (لا يعتمد على تركيز النافذة)،
    فيشتغل حتى لو نافذة كروم مخفية أو مغطّاة بنافذة أخرى.
    ثم يطلق أحداث input/change/blur ليسجّل النموذج الديناميكي القيمة.
    """
    try:
        inp = driver.find_element('xpath',
            f"//input[contains(@name,'{name_contains}')]")
        clean_value = str(value).replace('\n', '').replace('\r', '').strip()

        # ── الكتابة + إطلاق الأحداث بالكامل عبر JS (بدون send_keys) ──
        driver.execute_script("""
            var el = arguments[0];
            var val = arguments[1];
            el.focus();
            el.value = '';
            el.value = val;
            el.dispatchEvent(new Event('input',  {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            el.dispatchEvent(new Event('blur',   {bubbles:true}));
        """, inp, clean_value)
        time.sleep(0.5)

        # ── تحقق أن القيمة انكتبت فعلاً ─────────────────────
        actual = inp.get_attribute('value') or ''
        try:
            row_text = inp.find_element('xpath', './ancestor::tr[1]').text.strip()[:40]
        except Exception:
            row_text = ''

        if actual.strip() == clean_value:
            print(f'   ✅ "{row_text}" = {clean_value}')
            return True
        else:
            print(f'   ⚠️ "{row_text}" كُتبت "{actual}" بدل "{clean_value}"')
            return False
    except Exception as e:
        print(f'   ❌ "{name_contains}": {str(e)[:50]}')
        return False


def get_field_value_with_retry(name_contains, timeout=20):
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            el = driver.find_element('xpath', f"//input[contains(@name,'{name_contains}')]")
            val = el.get_attribute('value')
            if val and val.strip():
                return val
        except Exception:
            pass
        time.sleep(0.5)
    return ''


def click_delete_file(ctl_id):
    try:
        delete_link = driver.find_element('xpath',
            f"//a[contains(@href,\"{ctl_id}$Field1$ctl31\")]")
        js_click(delete_link)
        time.sleep(2.5)
        print(f'   🗑️ تم حذف الملف القديم في {ctl_id}')
        return True
    except Exception as e:
        print(f'   ⚠️ لم يتم العثور على رابط الحذف: {str(e)[:50]}')
        return False


def upload_file(ctl_id, file_path, max_wait=30):
    value_field = driver.find_element('xpath',
        f"//input[contains(@name,'{ctl_id}$Field1$ctl28')]")
    before_value = value_field.get_attribute('value') or ''

    if before_value.strip():
        print(f'   ℹ️ يوجد ملف سابق ({before_value[:30]}) — سيتم حذفه أولاً')
        click_delete_file(ctl_id)
        before_value = ''

    file_input = driver.find_element('xpath',
        f"//input[@type='file' and contains(@name,'{ctl_id}$Field1$ctl27')]")
    driver.execute_script(
        "arguments[0].style.display='block'; "
        "arguments[0].style.visibility='visible'; "
        "arguments[0].style.opacity=1; "
        "arguments[0].style.position='static';",
        file_input
    )
    file_input.send_keys(file_path)
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
        file_input
    )
    print(f'   📤 جاري رفع الصورة: {os.path.basename(file_path)} ...')

    elapsed = 0
    uploaded = False
    while elapsed < max_wait:
        time.sleep(0.5)
        elapsed += 0.5
        try:
            current_value = driver.find_element('xpath',
                f"//input[contains(@name,'{ctl_id}$Field1$ctl28')]"
            ).get_attribute('value') or ''
        except Exception:
            current_value = ''
        if current_value.strip() and current_value != before_value:
            print(f'   ✅ تم تأكيد رفع الصورة ({elapsed:.1f}ث) — '
                  f'القيمة: "{current_value[:40]}"')
            uploaded = True
            break

    if not uploaded:
        print(f'   ❌ لم يتم تأكيد رفع الصورة بعد {max_wait} ثانية')
    return uploaded


def get_all_buildings():
    all_buildings = []; page_num = 1; seen = set()
    while True:
        time.sleep(2)
        rows = driver.find_elements('xpath', "//table//tr[.//a[@href]]")
        for row in rows:
            try:
                cells = row.find_elements('xpath', './/td')
                # صف نتيجة حقيقي: فيه خلية كود رقمية (3+ أرقام) مثل 38436
                code_num = None
                for cell in cells:
                    ct = (cell.text or '').strip()
                    if ct.isdigit() and len(ct) >= 3:
                        code_num = ct; break
                if code_num is None:
                    continue

                text = row.text.strip()

                # اسم الموقع = أول رابط نصه يحتوي حروف
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

                bnum = None
                m = re.search(r'(?:مبن[ىي]|موقع)\s+رقم\s+(\d+)', text)
                if m:
                    bnum = int(m.group(1))

                if link in seen:
                    continue
                seen.add(link)
                all_buildings.append({'number': bnum, 'name': site_name, 'link': link})
            except Exception:
                continue
        next_link = driver.find_elements('xpath',
            f"//table//tr//td//a[text()='{page_num+1}'] | "
            f"//a[contains(@href,'Page${page_num+1}')]")
        if next_link:
            js_click(next_link[0]); page_num += 1
        else:
            break
    return all_buildings


# ── أدوات المطابقة بالاسم (لما يكون coded أسماء لا أرقام) ──
def normalize_name(s):
    s = str(s).strip()
    s = re.sub(r'[\u064b-\u0652]', '', s)
    s = s.replace('أ','ا').replace('إ','ا').replace('آ','ا')
    s = s.replace('ة','ه').replace('ى','ي')
    for w in ['التاريخيه','التاريخي','قريه','قرية','موقع','منطقه','منطقة','حي']:
        s = s.replace(w, '')
    s = s.replace('ـ','')
    s = re.sub(r'\s+', '', s)
    return s.strip()


def name_similarity(a, b):
    from difflib import SequenceMatcher
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def best_name_match(excel_name, platform_buildings, threshold=0.72):
    best = None; best_score = 0.0
    for b in platform_buildings:
        score = name_similarity(excel_name, b.get('name', ''))
        if score > best_score:
            best_score = score; best = b
    if best is not None and best_score >= threshold:
        return best, best_score
    return None, best_score


def coded_is_numeric(series):
    ok = 0; total = 0
    for v in series.dropna().unique():
        total += 1
        try:
            int(float(str(v).strip()))
            ok += 1
        except Exception:
            pass
    return total > 0 and (ok / total) >= 0.7


def process_building(building_id, building_num, data_by_building, images_by_building):
    ensure_automation_tab()
    data = data_by_building[building_num]
    base_url = (
        f'https://heritage2.itqan-consultant.com'
        f'/Web/App/Pools/DataEdit/{building_id}/2'
    )
    print(f'\n🏢 مبنى {building_num}')
    driver.get(base_url); time.sleep(3)

    print('   📝 تعبئة X / Y / المساحة...')
    # نكتب كل حقل ونتأكد أنه انكتب فعلاً؛ لو فشل نعيد حتى 4 مرات
    for _fname, _fval in [('ctl37$Field1$ctl03', data['x']),
                          ('ctl38$Field1$ctl03', data['y']),
                          ('ctl39$Field1$ctl03', data['area'])]:
        for _try in range(4):
            if fill_field(_fname, _fval):
                break
            time.sleep(0.4)
    time.sleep(1.5)

    if building_num in images_by_building:
        image_uploaded = upload_file('ctl40', images_by_building[building_num])
        if not image_uploaded:
            print(f'   ⛔ لن يتم الحفظ — فشل تأكيد رفع الصورة')
            return False
    else:
        print(f'   ⚠️ لا توجد صورة مطابقة لمبنى {building_num} — تخطي رفع الصورة')

    try:
        save_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSave']")
        js_click(save_btn)
        print('   💾 تم الضغط على حفظ')
        time.sleep(6)
    except Exception as e:
        print(f'   ❌ خطأ في الحفظ: {e}')
        return False

    driver.get(base_url); time.sleep(2)
    check_x    = get_field_value_with_retry('ctl37$Field1$ctl03')
    check_y    = get_field_value_with_retry('ctl38$Field1$ctl03')
    check_area = get_field_value_with_retry('ctl39$Field1$ctl03')

    try:
        ok = (
            abs(float(check_x) - data['x']) < 0.0001 and
            abs(float(check_y) - data['y']) < 0.0001
        )
    except ValueError:
        ok = False

    print(f'   🔄 بعد إعادة التحميل: X={check_x}, Y={check_y}, مساحة={check_area}')

    if not ok:
        print('   🔁 لم تنجح المحاولة الأولى — إعادة محاولة كاملة...')
        driver.get(base_url); time.sleep(3)
        for _fname, _fval in [('ctl37$Field1$ctl03', data['x']),
                              ('ctl38$Field1$ctl03', data['y']),
                              ('ctl39$Field1$ctl03', data['area'])]:
            for _try in range(4):
                if fill_field(_fname, _fval):
                    break
                time.sleep(0.4)
        time.sleep(1.5)
        try:
            save_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSave']")
            js_click(save_btn)
            print('   💾 تم الضغط على حفظ (محاولة ثانية)')
            time.sleep(6)
        except Exception as e:
            print(f'   ❌ خطأ في الحفظ بالمحاولة الثانية: {e}')
            return False

        driver.get(base_url); time.sleep(2)
        check_x    = get_field_value_with_retry('ctl37$Field1$ctl03')
        check_y    = get_field_value_with_retry('ctl38$Field1$ctl03')
        check_area = get_field_value_with_retry('ctl39$Field1$ctl03')
        try:
            ok = (
                abs(float(check_x) - data['x']) < 0.0001 and
                abs(float(check_y) - data['y']) < 0.0001
            )
        except ValueError:
            ok = False
        print(f'   🔄 بعد المحاولة الثانية: X={check_x}, Y={check_y}, مساحة={check_area}')

    return ok


# ── 6) تسجيل الدخول (مرة واحدة) ───────────────────────────
driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
time.sleep(3)
print(f'📍 الصفحة: {driver.title}')

if driver.title == 'Runtime Error' or 'Login' in driver.current_url:
    driver.get('https://heritage2.itqan-consultant.com')
    time.sleep(3)
    input('\n👤 سجّل دخولك الآن في نافذة Chrome، ثم اضغط Enter هنا للمتابعة...')
    driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
    time.sleep(3)
    print(f'✅ الصفحة بعد الدخول: {driver.title}')

# ── 7) الحلقة الرئيسية على كل قرية ────────────────────────
grand_total = {'villages': 0, 'processed': 0, 'skipped': 0, 'failed': 0}
followup_by_village = {}

# ── استئناف ذكي: نحفظ أكواد المباني المكتملة، فلو توقّف وأعدت التشغيل
#    يتخطّى المكتمل ويكمل من مكانه؛ والمباني الفاشلة (غير محفوظة) تُعاد تلقائياً.
import json as _json
CHECKPOINT_FILE = os.path.join(REPORTS_DIR, '.checkpoint_croqui.json')

def _load_done():
    try:
        with open(CHECKPOINT_FILE, encoding='utf-8') as _f:
            return set(_json.load(_f).get('done', []))
    except Exception:
        return set()

def _save_done(done_set):
    try:
        _tmp = CHECKPOINT_FILE + '.tmp'
        with open(_tmp, 'w', encoding='utf-8') as _f:
            _json.dump({'done': sorted(done_set)}, _f, ensure_ascii=False)
        os.replace(_tmp, CHECKPOINT_FILE)
    except Exception as _e:
        print(f'   ⚠️ تعذّر حفظ نقطة الاستئناف: {_e}')

DONE = _load_done()
if DONE:
    print(f'📌 استئناف: {len(DONE)} مبنى مكتمل سابقاً — سيُتخطّى ويُكمل من مكانه')

for v_idx, cfg in enumerate(villages_config, start=1):
    village          = cfg['village']
    FOLDER_PATH      = cfg['folder']
    FILTER_BUILDINGS = cfg['filter']

    print(f'\n{"█"*55}')
    print(f'█  قرية {v_idx}/{len(villages_config)}: {village}')
    if FILTER_BUILDINGS:
        print(f'█  🎯 مباني محددة: {FILTER_BUILDINGS}')
    print(f'{"█"*55}')

    # ── إيجاد ملف Excel الصحيح (يفحص أعمدة كل ملف) ────────
    excel_files = [
        f for f in os.listdir(FOLDER_PATH)
        if f.lower().endswith(('.xlsx', '.xls'))
        and not f.startswith('~$')
        and not f.lower().startswith('report_')
        and not f.startswith('تقرير_')
    ]
    if not excel_files:
        print(f'❌ لا يوجد ملف Excel في مجلد "{village}"')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'skip',
            'التفاصيل': 'لا يوجد ملف Excel في مجلد القرية',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        continue

    # ── نفحص أعمدة كل ملف ونختار اللي فيه الطول+العرض+المساحة ──
    EXCEL_PATH = None
    df = None
    x_col = y_col = area_col = None
    for candidate in excel_files:
        cpath = os.path.join(FOLDER_PATH, candidate)
        try:
            cdf = pd.read_excel(cpath)
        except Exception:
            continue
        cx = find_col(cdf, 'الطول')
        cy = find_col(cdf, 'العرض')
        ca = find_col(cdf, 'مساحة')
        if cx and cy and ca:
            EXCEL_PATH = cpath
            df = cdf
            x_col, y_col, area_col = cx, cy, ca
            print(f'📊 ملف Excel المطابق: {candidate}')
            break
        else:
            print(f'   ⏭️ تخطي "{candidate}" (لا يحتوي أعمدة الإحداثيات)')

    if EXCEL_PATH is None:
        print(f'❌ لا يوجد ملف Excel فيه أعمدة الطول/العرض/المساحة في "{village}"')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'error',
            'التفاصيل': 'لا يوجد ملف Excel صالح (أعمدة الإحداثيات ناقصة)',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        continue

    coded_col = 'coded'

    # ── اكتشاف: عمود coded أرقام أو أسماء؟ ────────────────
    NAME_MODE = not coded_is_numeric(df[coded_col])
    print('🔤 وضع المطابقة: ' + ('بالاسم (coded أسماء)' if NAME_MODE
                                  else 'بالرقم (coded أرقام)'))

    data_by_building = {}
    for _, row in df.iterrows():
        if NAME_MODE:
            key = str(row[coded_col]).strip()
            if not key:
                continue
        else:
            try:
                key = int(row[coded_col])
            except (ValueError, TypeError):
                continue
        data_by_building[key] = {
            'x': float(row[x_col]),
            'y': float(row[y_col]),
            'area': row[area_col],
        }
    print(f'🏢 المواقع في Excel: {len(data_by_building)}')

    # ── فهرسة الصور داخل مجلد القرية ──────────────────────
    #   وضع الرقم: الصورة يطابقها رقم في اسمها.
    #   وضع الاسم: الصورة يطابقها أقرب اسم موقع (تشابه الاسم).
    image_extensions = ('.png', '.jpg', '.jpeg')
    images_by_building = {}
    _all_images = [f for f in os.listdir(FOLDER_PATH)
                   if f.lower().endswith(image_extensions)]
    if NAME_MODE:
        for f in _all_images:
            base = os.path.splitext(f)[0]
            # نطابق اسم الصورة بأقرب اسم موقع في الإكسل
            best_key = None; best_sc = 0.0
            for key in data_by_building.keys():
                sc = name_similarity(base, key)
                if sc > best_sc:
                    best_sc = sc; best_key = key
            if best_key is not None and best_sc >= 0.6:
                images_by_building[best_key] = os.path.join(FOLDER_PATH, f)
    else:
        for f in _all_images:
            match = re.search(r'\d+', os.path.splitext(f)[0])
            if match:
                bnum = int(match.group())
                images_by_building[bnum] = os.path.join(FOLDER_PATH, f)
    print(f'🖼️  الصور المفهرسة: {len(images_by_building)}')

    # ── البحث عن القرية في الوارد (إدخال عبر JS بدون تركيز) ──
    ensure_automation_tab()
    driver.get('https://heritage2.itqan-consultant.com/Web/App/Home/Inbox')
    time.sleep(3)

    # كتابة اسم القرية عبر JavaScript (يشتغل حتى لو كروم مخفي)
    search_field = driver.find_element('xpath', "//input[@name='ctl12$ctl04']")
    driver.execute_script("""
        var el = arguments[0]; var val = arguments[1];
        el.value = val;
        el.dispatchEvent(new Event('input',  {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, search_field, village)
    time.sleep(0.3)

    # كتابة رقم قرار التسجيل عبر JavaScript (إن وُجد)
    if REG_DECISION:
        reg_field = driver.find_element('xpath', "//input[@name='ctl12$ctl28']")
        driver.execute_script("""
            var el = arguments[0]; var val = arguments[1];
            el.value = val;
            el.dispatchEvent(new Event('input',  {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
        """, reg_field, REG_DECISION)
        time.sleep(0.3)

    # الضغط على بحث عبر JS كذلك
    search_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSearch']")
    driver.execute_script("arguments[0].click();", search_btn)
    time.sleep(3)
    print(f'🔍 تم البحث عن: {village}'
          + (f' | قرار: {REG_DECISION}' if REG_DECISION else ''))

    buildings = get_all_buildings()
    print(f'🏢 المباني في المنصة: {len(buildings)}')

    if not buildings:
        print(f'⚠️ لم تُوجد مباني للقرية "{village}" — سيتم تخطيها')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'skip',
            'التفاصيل': 'لم تُوجد أي مباني لهذه القرية في المنصة',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        continue

    results_log = []
    def log(bnum, status, msg=''):
        ts = datetime.now().strftime('%H:%M:%S')
        icon = {'success': '✅', 'skip': '⏭️', 'error': '❌'}[status]
        print(f'[{ts}] {icon} مبنى {bnum}: {msg}')
        results_log.append({'رقم المبنى': bnum, 'الحالة': status,
                            'التفاصيل': msg, 'الوقت': ts})

    processed = skipped = failed = 0

    # ── نبني قائمة العمل: مطابقة بالرقم أو بالاسم ─────────
    worklist = []   # (excel_key, platform_building, label)
    if NAME_MODE:
        used_links = set()
        for excel_name in data_by_building.keys():
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
            bnum = b['number']
            # فلتر النطاق يطبّق فقط في وضع الرقم
            if FILTER_BUILDINGS and bnum not in FILTER_BUILDINGS:
                continue
            worklist.append((bnum, b, None))

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
            skipped += 1; continue
        if excel_key not in data_by_building:
            log(excel_key, 'skip', 'غير موجود في Excel')
            skipped += 1; continue

        bid = b['link'].rstrip('/').split('/')[-1]

        if bid in DONE:
            print('   ⏭️ مُنجز سابقاً — تخطّي (استئناف)')
            continue

        try:
            ok = process_building(bid, excel_key, data_by_building, images_by_building)
            has_img = '+ صورة' if excel_key in images_by_building else '(بدون صورة)'
            log(excel_key, 'success' if ok else 'error',
                f'تم تحديث X/Y/مساحة {has_img}' if ok else 'فشل التحقق بعد محاولتين')
            if ok:
                processed += 1; DONE.add(bid); _save_done(DONE)
            else:
                failed += 1
        except Exception as e:
            emsg = str(e)
            if ('HTTPConnection' in emsg or 'Max retries' in emsg
                    or 'connection' in emsg.lower() or 'refused' in emsg.lower()
                    or not is_connection_alive()):
                print('   ⚠️ يبدو أن الاتصال انقطع أثناء المبنى — إعادة اتصال ومحاولة ثانية...')
                if reconnect_driver():
                    try:
                        ok = process_building(bid, excel_key, data_by_building, images_by_building)
                        has_img = '+ صورة' if excel_key in images_by_building else '(بدون صورة)'
                        log(excel_key, 'success' if ok else 'error',
                            f'تم تحديث X/Y/مساحة {has_img} (بعد إعادة اتصال)'
                            if ok else 'فشل التحقق بعد محاولتين')
                        if ok:
                            processed += 1; DONE.add(bid); _save_done(DONE)
                        else:
                            failed += 1
                        continue
                    except Exception as e2:
                        log(excel_key, 'error', f'فشل حتى بعد إعادة الاتصال: {str(e2)[:40]}')
                        failed += 1
                        continue
            log(excel_key, 'error', emsg[:60])
            failed += 1

    followup_rows = [r for r in results_log if r['الحالة'] in ('error', 'skip')]
    followup_by_village[village] = followup_rows

    print(f'\n{"─"*50}')
    print(f'📊 قرية {village}: ✅ {processed}  ⏭️ {skipped}  ❌ {failed}')
    if followup_rows:
        nums = [str(r['رقم المبنى']) for r in followup_rows]
        print(f'⚠️ مباني محتاجة متابعة: {", ".join(nums)}')
    else:
        print(f'🎉 كل المباني نجحت — لا شي محتاج متابعة')
    print(f'{"─"*50}')

    grand_total['villages']  += 1
    grand_total['processed'] += processed
    grand_total['skipped']   += skipped
    grand_total['failed']    += failed

# ── 8) التقرير الإجمالي: ملف واحد، شيت لكل قرية ────────────
from openpyxl.styles import Font, PatternFill, Alignment

master_report = os.path.join(
    REPORTS_DIR,
    f'تقرير_المتابعة_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
)

with pd.ExcelWriter(master_report, engine='openpyxl') as writer:
    for village, rows in followup_by_village.items():
        safe_sheet = re.sub(r'[\\/*?:\[\]]', '_', village)[:31]

        if rows:
            df_out = pd.DataFrame(rows)[['رقم المبنى', 'الحالة', 'التفاصيل', 'الوقت']]
            df_out['الحالة'] = df_out['الحالة'].map(
                {'error': '❌ فشل', 'skip': '⏭️ تخطي'}
            ).fillna(df_out['الحالة'])
        else:
            df_out = pd.DataFrame(
                [{'رقم المبنى': '—', 'الحالة': '✅ لا يوجد',
                  'التفاصيل': 'كل المباني نجحت — لا شي محتاج متابعة', 'الوقت': ''}]
            )

        df_out.to_excel(writer, sheet_name=safe_sheet, index=False)

        ws = writer.sheets[safe_sheet]
        ws.sheet_view.rightToLeft = True
        widths = {'A': 12, 'B': 12, 'C': 45, 'D': 12}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        header_fill = PatternFill('solid', fgColor='4472C4')
        header_font = Font(bold=True, color='FFFFFF')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

print(f'\n{"█"*55}')
print(f'█  انتهت جميع القرى')
print(f'█  عدد القرى: {grand_total["villages"]}')
print(f'█  ✅ نجح: {grand_total["processed"]}  '
      f'⏭️ تخطى: {grand_total["skipped"]}  '
      f'❌ فشل: {grand_total["failed"]}')
print(f'█  💾 تقرير المتابعة (شيت لكل قرية): {master_report}')
print(f'{"█"*55}')

# ── إغلاق نافذة الإيقاف تلقائياً بعد انتهاء الكود ──────────
try:
    close_pause_button()
    print('🔚 تم إغلاق نافذة الإيقاف — انتهى العمل')
except Exception:
    pass
