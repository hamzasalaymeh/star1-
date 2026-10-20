# ============================================================
# 🏛️ أتمتة منصة التراث v5.0 — تسجيل دخول تلقائي + فصل القرى المتشابهة
# صفحة: "استكمال بيانات الميدان" — حقول X / Y / المساحة + رفع صورة الكروكي
# ============================================================
# الاستخدام (النسخة المستقرة v4.9):
#   1) تختار مجلد أب واحد — جوّاه مجلدات القرى (كل مجلد فيه Excel + صور)
#      اسم مجلد القرية = اسم القرية للبحث في المنصة
#   2) تدخل رقم قرار التسجيل مرة واحدة (نفسه لكل القرى) — اختياري
#   3) يعرض لك قائمة القرى، وتقدر تحدّد range مباني لكل قرية (اختياري)
#      أمثلة: 3,7,12  أو  50:70  أو  3,7,50:70  — فاضي = كل المباني
#   4) يشتغل قرية قرية، وفي النهاية يطلّع ملف Excel واحد فيه شيت
#      لكل قرية يعرض المباني اللي فشلت أو اتخطّت (المحتاجة متابعة)
# ============================================================

import subprocess, sys, time, re, os, threading, importlib, urllib.request
from datetime import datetime

# إعدادات الاعتمادية. القيم متعمدة لتشغيل طويل من دون الدوران في البحث.
SEARCH_RESULT_TIMEOUT = 50
PAGE_CHANGE_TIMEOUT = 35
PAGE_CHANGE_ATTEMPTS = 5
MAX_RESULT_PAGES = 500
GRID_STABLE_HITS = 3

# ── بيانات تسجيل الدخول التلقائي ───────────────────────────
# يكتبها البرنامج تلقائياً في صفحة الدخول. لو تغيّرت كلمة المرور،
# عدّلها هنا فقط. لو تركتها فاضية أو فشل الدخول التلقائي، يرجع
# البرنامج لطلب تسجيل الدخول اليدوي مرة واحدة (تبقى الجلسة محفوظة).
LOGIN_USERNAME = 'H.qurah'
LOGIN_PASSWORD = '313146'
LOGIN_URL = 'https://heritage2.itqan-consultant.com'

# يمنع Windows من إدخال الجهاز في Sleep أثناء التشغيل الليلي.
if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
        print('🌙 تم تفعيل منع السكون أثناء عمل الأتمتة')
    except Exception:
        pass

REQUIRED_PACKAGES = {
    'selenium': 'selenium',
    'pandas': 'pandas',
    'openpyxl': 'openpyxl',
    'webdriver_manager': 'webdriver-manager',
}


def ensure_dependencies():
    """يثبّت فقط المكتبات الناقصة، بدل إعادة التثبيت في كل تشغيل."""
    for module_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            print(f'📦 تثبيت المكتبة الناقصة: {package_name}')
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name, '-q'])


ensure_dependencies()

import pandas as pd
from openpyxl import Workbook
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


def ask_test_mode():
    """وضع فحص حي: يفتح الصفحات ويتحقق من الحقول من دون كتابة أو رفع أو حفظ."""
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    ans = messagebox.askyesno(
        'وضع التشغيل',
        'هل تريد تشغيل اختبار آمن أولاً؟\n\n'
        '• نعم = فحص البحث والصفحات والحقول فقط، بدون تعديل أو رفع أو حفظ.\n'
        '• لا  = تنفيذ التحديث الفعلي.')
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
USE_PARENT = ask_mode()

village_folders = []   # قائمة مجلدات القرى (اسم المجلد = اسم القرية)

if USE_PARENT:
    # ── وضع (أ): مجلد أب واحد جوّاه مجلدات القرى ───────────
    PARENT = ask_parent_folder()
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
REG_DECISION = ask_string(
    'رقم قرار التسجيل',
    'رقم قرار التسجيل (نفسه لكل القرى — اتركه فاضي لو ما يلزم):')
REG_DECISION = (REG_DECISION or '').strip()
print(f'📋 رقم قرار التسجيل: {REG_DECISION or "(فاضي)"}')

# ── 3) اختيار range المباني لكل قرية (اختياري بالكامل) ────
villages_config = []

# سؤال واحد: تبي تحدد نطاقات أصلاً؟ لو لا → كل المباني لكل القرى
want_ranges = ask_yes_no(
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

# أسماء القرى الشقيقة تُستخدم فقط لفصل الأسماء المتداخلة مثل
# «الماوين» و«الماوين بللحمر»، حتى عند اختيار مجلد واحد يدوياً.
KNOWN_VILLAGE_NAMES = [cfg['village'] for cfg in villages_config]
try:
    for sibling_name in os.listdir(REPORTS_DIR):
        sibling_path = os.path.join(REPORTS_DIR, sibling_name)
        if os.path.isdir(sibling_path) and sibling_name not in KNOWN_VILLAGE_NAMES:
            KNOWN_VILLAGE_NAMES.append(sibling_name)
except Exception:
    pass

print(f'\n✅ تم إعداد {len(villages_config)} قرية — سيبدأ التشغيل الآن')

start_text = ask_string(
    'بدء التشغيل',
    f'رقم القرية التي تريد البدء منها (1 إلى {len(villages_config)})\n'
    'اتركه فارغاً للبدء من الأولى:')
try:
    START_VILLAGE = int((start_text or '1').strip())
except ValueError:
    START_VILLAGE = 1
START_VILLAGE = max(1, min(START_VILLAGE, len(villages_config)))
print(f'▶️ سيبدأ من القرية {START_VILLAGE}: '
      f'{villages_config[START_VILLAGE - 1]["village"]}')

TEST_MODE = ask_test_mode()
print('🧪 وضع اختبار آمن — لن يتم حفظ أي تعديل' if TEST_MODE
      else '🚀 وضع التنفيذ الفعلي — سيتم حفظ التعديلات')

# (مجلد التقارير REPORTS_DIR تم تحديده في قسم الاختيار أعلاه)

# ── 4) فتح Chrome بجلسة مخصّصة ومعزولة ─────────────────────
chrome_paths = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]
chrome_exe = next((p for p in chrome_paths if os.path.exists(p)), None)
if not chrome_exe:
    raise FileNotFoundError('❌ Chrome غير موجود')

DEBUG_PORT = 9333
DEBUG_ADDRESS = f'127.0.0.1:{DEBUG_PORT}'
DEBUG_PROFILE = r'C:\heritage_automation_profile'


def debugger_is_ready():
    try:
        with urllib.request.urlopen(
            f'http://{DEBUG_ADDRESS}/json/version', timeout=1
        ) as response:
            return response.status == 200
    except Exception:
        return False


if not debugger_is_ready():
    # لا نغلق Chrome الخاص بالمستخدم إطلاقاً؛ هذه جلسة منفصلة للأتمتة.
    subprocess.Popen([
        chrome_exe, f'--remote-debugging-port={DEBUG_PORT}',
        f'--user-data-dir={DEBUG_PROFILE}',
        '--no-first-run', '--no-default-browser-check', '--start-maximized'
    ])
    for _ in range(20):
        if debugger_is_ready():
            break
        time.sleep(0.5)
    else:
        raise RuntimeError('❌ لم تبدأ جلسة Chrome المخصصة للأتمتة على المنفذ 9333')

options = Options()
options.add_experimental_option('debuggerAddress', DEBUG_ADDRESS)
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
            new_options.add_experimental_option('debuggerAddress', DEBUG_ADDRESS)
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


def read_grid_snapshot():
    """لقطة واحدة متماسكة للجدول والترقيم، من دون StaleElementReference."""
    return driver.execute_script("""
        var grid = document.getElementById('ctl12_GridView1');
        if (!grid) return null;
        var rows = Array.from(grid.querySelectorAll(':scope > tbody > tr, :scope > tr'));
        if (!rows.length) rows = Array.from(grid.querySelectorAll('tr'));
        var header = rows.find(function(r) { return r.querySelectorAll(':scope > th').length > 0; });
        var pager = rows.find(function(r) {
            return (r.className || '').toLowerCase().indexOf('pagger') >= 0 ||
                   (r.className || '').toLowerCase().indexOf('pager') >= 0;
        });
        var dataRows = rows.filter(function(r) {
            var c = (r.className || '').toLowerCase();
            return c.indexOf('pagger') < 0 && c.indexOf('pager') < 0 &&
                   r.querySelectorAll(':scope > td').length > 0;
        });
        var active = '';
        if (pager) {
            var activeEl = pager.querySelector('span');
            active = activeEl ? (activeEl.innerText || '').trim() : '';
        }
        return {
            headers: header ? Array.from(header.querySelectorAll(':scope > th')).map(function(th) {
                return {text: th.innerText || '', colspan: parseInt(th.getAttribute('colspan') || '1')};
            }) : [],
            rows: dataRows.map(function(row) {
                return {
                    text: row.innerText || '',
                    cells: Array.from(row.querySelectorAll(':scope > td')).map(function(c) {
                        return c.innerText || '';
                    }),
                    links: Array.from(row.querySelectorAll('a[href]')).map(function(a) {
                        return {text: a.innerText || '', href: a.href || '', onclick: a.getAttribute('onclick') || ''};
                    })
                };
            }),
            activePage: active,
            pagerLinks: pager ? Array.from(pager.querySelectorAll('a')).map(function(a) {
                return {text: (a.innerText || '').trim(), href: a.getAttribute('href') || '',
                        onclick: a.getAttribute('onclick') || ''};
            }) : []
        };
    """)


def grid_data_signature(snapshot=None):
    """بصمة صفوف الجدول فقط؛ لا يدخل رقم Pager فيها حتى لا نقبل صفحة قديمة."""
    snapshot = snapshot if snapshot is not None else read_grid_snapshot()
    if not snapshot:
        return ''
    row_parts = []
    for row in snapshot.get('rows', []):
        ids = []
        for link in row.get('links', []):
            m = re.search(r'/DataEdit/(\d+)/?(?:[?#]|$)', str(link.get('href', '')))
            if m:
                ids.append(m.group(1))
        row_text = re.sub(r'\s+', ' ', str(row.get('text', '')).strip())[:500]
        if ids or row_text:
            row_parts.append(','.join(ids) + '::' + row_text)
    return '||'.join(row_parts)[:30000]


def grid_signature(snapshot=None):
    """بصمة كاملة للصفحة: Pager + البيانات بعد استقرارها."""
    snapshot = snapshot if snapshot is not None else read_grid_snapshot()
    if not snapshot:
        return ''
    active = str(snapshot.get('activePage') or '')
    return active + '::' + grid_data_signature(snapshot)


def wait_for_fresh_results(previous_signature, timeout=SEARCH_RESULT_TIMEOUT):
    """ينتظر نتيجة جديدة ومستقرة بعد البحث؛ لا يقبل الجدول القديم الموجود قبل الضغط."""
    end_time = time.time() + timeout
    last_signature = None
    stable_hits = 0
    while time.time() < end_time:
        time.sleep(0.5)
        try:
            snapshot = read_grid_snapshot()
            signature = grid_signature(snapshot)
            has_records = any(
                re.search(r'/DataEdit/\d+', str(link.get('href', '')))
                for row in (snapshot or {}).get('rows', [])
                for link in row.get('links', [])
            )
            if has_records and signature != previous_signature:
                if signature == last_signature:
                    stable_hits += 1
                else:
                    last_signature = signature
                    stable_hits = 1
                if stable_hits >= GRID_STABLE_HITS:
                    return snapshot
        except Exception:
            pass
    return None


def wait_for_grid_change(previous_data_signature, expected_page=None,
                         timeout=PAGE_CHANGE_TIMEOUT):
    """ينتظر تغير صفوف الجدول نفسها واستقرارها؛ تغير رقم Pager وحده لا يكفي."""
    end_time = time.time() + timeout
    last_data_signature = None
    stable_hits = 0
    while time.time() < end_time:
        time.sleep(0.5)
        try:
            snapshot = read_grid_snapshot()
            data_signature = grid_data_signature(snapshot)
            active_page = _english_digits((snapshot or {}).get('activePage', '')).strip()
            page_ok = expected_page is None or active_page == str(expected_page)
            if data_signature and data_signature != previous_data_signature and page_ok:
                if data_signature == last_data_signature:
                    stable_hits += 1
                else:
                    last_data_signature = data_signature
                    stable_hits = 1
                if stable_hits >= GRID_STABLE_HITS:
                    return snapshot
            else:
                last_data_signature = None
                stable_hits = 0
        except Exception:
            pass
    return None


def find_col(df, keyword):
    for col in df.columns:
        if keyword in col:
            return col
    return None


def parse_excel_number(value):
    """يقرأ أرقام Excel حتى لو كانت نصاً أو بأرقام/فواصل عربية أو معها وحدة."""
    if value is None or pd.isna(value) or isinstance(value, bool):
        raise ValueError('قيمة فارغة')
    if isinstance(value, (int, float)):
        number = float(value)
        if pd.isna(number):
            raise ValueError('قيمة فارغة')
        return number

    text = str(value).strip()
    if not text:
        raise ValueError('قيمة فارغة')
    if text.startswith('='):
        raise ValueError('معادلة Excel بلا قيمة محسوبة')
    text = text.translate(str.maketrans(
        '٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹',
        '01234567890123456789'
    ))
    text = (text.replace('\u00a0', ' ').replace('\u202f', ' ')
            .replace('−', '-').replace('–', '-').replace('—', '-')
            .replace('٫', '.').replace('٬', ''))
    text = re.sub(r'(?i)\b(?:sqm|sq\.?\s*m|m2)\b', ' ', text)
    text = re.sub(r'(?:متر(?:اً|ا)?\s*(?:مربع|م²|2)?|م²)', ' ', text)
    text = re.sub(r'(?<=\d)\s+(?=\d)', '', text)

    # تحديد الفاصل العشري وفق آخر فاصلة عند اجتماع النقطة والفاصلة.
    if ',' in text and '.' in text:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '').replace(',', '.')
        else:
            text = text.replace(',', '')
    elif ',' in text:
        if text.count(',') == 1:
            text = text.replace(',', '.')
        else:
            parts = text.split(',')
            text = ''.join(parts[:-1]) + '.' + parts[-1]

    matches = re.findall(r'[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?', text)
    if len(matches) != 1:
        raise ValueError('لا تحتوي رقماً واحداً واضحاً')
    return float(matches[0])


def short_excel_value(value, limit=35):
    text = str(value).replace('\n', ' ').replace('\r', ' ').strip()
    return text if len(text) <= limit else text[:limit - 3] + '...'


def count_valid_coordinate_rows(dataframe, x_column, y_column, area_column):
    """عدد الصفوف التي تحتوي ثلاث قيم قابلة للاستخدام فعلاً."""
    valid_count = 0
    for _, candidate_row in dataframe.iterrows():
        try:
            parse_excel_number(candidate_row[x_column])
            parse_excel_number(candidate_row[y_column])
            parse_excel_number(candidate_row[area_column])
            valid_count += 1
        except (TypeError, ValueError):
            pass
    return valid_count


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


def _english_digits(value):
    return str(value or '').translate(str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789'))


def _simple_arabic_name(value):
    value = _english_digits(value).strip().lower()
    value = re.sub(r'[\u064b-\u065f\u0670]', '', value)
    value = value.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    value = value.replace('ة', 'ه').replace('ى', 'ي').replace('ـ', '')
    for generic_word in ('قريه', 'التراثيه', 'التراثي'):
        value = value.replace(generic_word, '')
    value = re.sub(r'[^\u0600-\u06ffa-z0-9]+', '', value)
    return value


def _arabic_name_tokens(value):
    """كلمات مميزة من اسم القرية، مع تجاهل الكلمات العامة واختلاف الهمزات."""
    value = _english_digits(value).strip().lower()
    value = re.sub(r'[\u064b-\u065f\u0670]', '', value)
    value = value.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    value = value.replace('ة', 'ه').replace('ى', 'ي').replace('ـ', '')

    # توحيد جميع الكتابات الشائعة لـ«بللحمر». هذا مهم لأن المنصة قد تعرضها
    # «بلحمر / بالحمر / بل الحمر»، بينما يكون اسم المجلد «بللحمر».
    value = re.sub(
        r'(?<![\u0600-\u06ffa-z0-9])(?:بل|بال|ب)\s+(?:ال)?حمر'
        r'(?![\u0600-\u06ffa-z0-9])',
        ' بللحمر ', value
    )
    value = re.sub(
        r'(?<![\u0600-\u06ffa-z0-9])(?:بللحمر|بلحمر|بالحمر|باللحمر|بلاحمر)'
        r'(?![\u0600-\u06ffa-z0-9])',
        ' بللحمر ', value
    )
    words = re.findall(r'[\u0600-\u06ffa-z0-9]+', value)
    ignored = {'قريه', 'التراثيه', 'التراثي', 'موقع', 'منطقه', 'حي',
               'في', 'من', 'الى', 'ال', 'مبني', 'رقم', 'الرقم',
               'كود', 'الكود'}
    return [word for word in words if word not in ignored and len(word) >= 3]


def _site_village_fragment(site_name):
    """يأخذ جزء اسم القرية بعد «في قرية» ويتجاهل وصف المبنى الذي يسبقه."""
    text = re.sub(r'\s+', ' ', str(site_name or '').strip())
    connectors = list(re.finditer(
        r'(?:^|\s)(?:في\s+(?:(?:قرية|قريه)\s+)?|(?:بقرية|بقريه)\s+)',
        text
    ))
    if connectors:
        return text[connectors[-1].end():].strip()
    return text


def _village_tokens_equivalent(site_tokens, wanted_tokens, threshold=0.86):
    """نفس عدد الكلمات المميزة، مع سماح محدود لاختلاف إملائي داخل الكلمة."""
    from difflib import SequenceMatcher
    site_remaining = list(dict.fromkeys(site_tokens))
    wanted_unique = list(dict.fromkeys(wanted_tokens))
    if not site_remaining or len(site_remaining) != len(wanted_unique):
        return False
    for wanted_word in sorted(wanted_unique, key=len, reverse=True):
        scored = [
            (SequenceMatcher(None, wanted_word, site_word).ratio(), index)
            for index, site_word in enumerate(site_remaining)
        ]
        best_score, best_index = max(scored, default=(0.0, -1))
        if best_score < threshold:
            return False
        site_remaining.pop(best_index)
    return not site_remaining


def _village_subset_score(site_tokens, candidate_tokens, threshold=0.86):
    """درجة وجود هوية قرية داخل اسم أطول؛ كل كلمة تُطابق كلمة واحدة فقط."""
    from difflib import SequenceMatcher
    site_remaining = list(dict.fromkeys(site_tokens))
    candidate_unique = list(dict.fromkeys(candidate_tokens))
    if not candidate_unique or len(candidate_unique) > len(site_remaining):
        return None
    scores = []
    for candidate_word in sorted(candidate_unique, key=len, reverse=True):
        choices = [
            (SequenceMatcher(None, candidate_word, site_word).ratio(), index)
            for index, site_word in enumerate(site_remaining)
        ]
        best_score, best_index = max(choices, default=(0.0, -1))
        if best_score < threshold:
            return None
        scores.append(best_score)
        site_remaining.pop(best_index)
    average = sum(scores) / len(scores)
    return average, len(site_remaining)


def _belongs_to_village(site_name, village):
    """يقبل وصفاً أطول، ثم يختار أطول هوية قرية معروفة تطابق النتيجة."""
    site_fragment = _site_village_fragment(site_name)
    site_tokens = _arabic_name_tokens(site_fragment)
    wanted_tokens = _arabic_name_tokens(village)

    # حاجز قطعي للحالة الخطرة: «الماوين» و«الماوين بللحمر» قريتان مختلفتان.
    # لا نعتمد هنا على وجود اسم القرية الأخرى ضمن المجلدات المختارة.
    site_has_ballahmar = 'بللحمر' in site_tokens
    wanted_has_ballahmar = 'بللحمر' in wanted_tokens
    if site_has_ballahmar != wanted_has_ballahmar:
        return False

    target_match = _village_subset_score(site_tokens, wanted_tokens)
    if target_match is None:
        return False

    target_identity = tuple(sorted(set(wanted_tokens)))
    target_rank = (len(target_identity), target_match[0], -target_match[1])
    best_competitor_rank = None
    for known_name in globals().get('KNOWN_VILLAGE_NAMES', []):
        known_tokens = _arabic_name_tokens(known_name)
        known_identity = tuple(sorted(set(known_tokens)))
        if not known_identity or known_identity == target_identity:
            continue
        known_match = _village_subset_score(site_tokens, known_tokens)
        if known_match is None:
            continue

        # لو اسم القرية المطلوبة جزء من اسم قرية أطول معروفة، والصف يطابق
        # الاسم الأطول، يُرفض فوراً قبل أي ترتيب أو تشابه تقريبي.
        if (set(target_identity) < set(known_identity)
                and set(known_identity).issubset(set(site_tokens))):
            return False
        known_rank = (len(known_identity), known_match[0], -known_match[1])
        if best_competitor_rank is None or known_rank > best_competitor_rank:
            best_competitor_rank = known_rank

    # التعادل يُرفض للأمان؛ لا ندخل استمارة إذا بقي اسم القرية ملتبساً.
    return best_competitor_rank is None or target_rank > best_competitor_rank


def _expanded_headers(snapshot):
    expanded = []
    for item in snapshot.get('headers', []):
        if isinstance(item, dict):
            text = str(item.get('text', '')).strip()
            try:
                span = max(1, int(item.get('colspan', 1)))
            except Exception:
                span = 1
        else:
            text = str(item).strip(); span = 1
        expanded.extend([text] * span)
    return expanded


def _extract_page_buildings(snapshot, village):
    """يستخرج سجلات الصفحة، ويصفي كل صف لا يطابق القرار أو القرية نفسها."""
    headers = _expanded_headers(snapshot)
    code_index = next((i for i, h in enumerate(headers) if 'الكود' in h), None)
    name_index = next((i for i, h in enumerate(headers)
                       if 'اسم' in h and 'الموقع' in h), None)
    wanted_decision = normalize_decimal_digits(REG_DECISION) if REG_DECISION else ''
    records = []; rejected_decision = 0; rejected_village = 0

    for row in snapshot.get('rows', []):
        try:
            cells = [str(c).strip() for c in row.get('cells', [])]
            links = row.get('links', [])
            record_id = None; record_link = ''
            for item in links:
                href = str(item.get('href', '')).strip()
                match = re.search(r'/DataEdit/(\d+)/?(?:[?#]|$)', href)
                if match:
                    record_id = match.group(1); record_link = href; break
            if not record_id:
                continue

            # القرار يؤخذ من خلية كاملة مطابقة تماماً، لا من رقم الصفحة أو مدة الأيام.
            if wanted_decision:
                exact_values = {normalize_decimal_digits(c) for c in cells}
                if wanted_decision not in exact_values:
                    rejected_decision += 1
                    continue

            site_code = ''
            if code_index is not None and code_index < len(cells):
                site_code = _english_digits(cells[code_index]).strip()
            if not site_code.isdigit():
                site_code = next((_english_digits(c).strip() for c in cells
                                  if _english_digits(c).strip().isdigit()
                                  and len(_english_digits(c).strip()) >= 3), '')

            site_name = ''
            if name_index is not None and name_index < len(cells):
                site_name = cells[name_index]
            if not site_name:
                # في هذه الشبكة: الكود أول خلية واسم الموقع ثاني خلية.
                if len(cells) > 1:
                    site_name = cells[1]
                else:
                    site_name = str(row.get('text', '')).strip()
            if not _belongs_to_village(site_name, village):
                rejected_village += 1
                continue

            number_text = _english_digits(site_name + ' ' + str(row.get('text', '')))
            number_match = re.search(r'(?:مبن[ىي]|موقع)\s*رقم\s*(\d+)', number_text)
            if not number_match:
                # لا نستخدم كود المنصة (مثل 39694) كرقم مبنى؛ هذا كان يسبب ربطاً خاطئاً.
                continue
            building_number = int(number_match.group(1))
            records.append({
                'id': record_id,
                'number': building_number,
                'name': site_name,
                'link': record_link,
                'site_code': site_code,
                'registration_decision': wanted_decision,
            })
        except Exception:
            continue
    return records, rejected_decision, rejected_village


def _next_page_argument(snapshot):
    """يقرأ أمر الصفحة التالية من Pager الحقيقي الظاهر في GridView."""
    try:
        current = int(_english_digits(snapshot.get('activePage', '')).strip())
    except Exception:
        current = None
    numeric = []
    next_group = None
    for link in snapshot.get('pagerLinks', []):
        raw = str(link.get('href', '')) + ' ' + str(link.get('onclick', ''))
        match = re.search(r"Page\$(\d+|Next)", raw, re.I)
        if not match:
            continue
        token = match.group(1)
        if token.lower() == 'next':
            next_group = 'Page$Next'
        else:
            page_value = int(token)
            if current is None or page_value > current:
                numeric.append(page_value)
    if numeric:
        wanted = current + 1 if current is not None else min(numeric)
        return f'Page${wanted}' if wanted in numeric else f'Page${min(numeric)}'
    return next_group


def _move_to_results_page(argument, previous_data_signature):
    """يضغط Pager نفسه؛ عند تعثر النقرة يستدعي نفس PostBack ثم ينتظر التغير الحقيقي."""
    for attempt in range(1, PAGE_CHANGE_ATTEMPTS + 1):
        try:
            sent = driver.execute_script("""
                var arg = arguments[0];
                var grid = document.getElementById('ctl12_GridView1');
                if (!grid) return false;
                var links = Array.from(grid.querySelectorAll('a'));
                var target = links.find(function(a) {
                    var raw = (a.getAttribute('href') || '') + ' ' + (a.getAttribute('onclick') || '');
                    return raw.indexOf(arg) >= 0;
                });
                if (target) { target.click(); return true; }
                if (typeof __doPostBack === 'function') {
                    __doPostBack('ctl12$GridView1', arg); return true;
                }
                return false;
            """, argument)
            if sent:
                page_match = re.search(r'Page\$(\d+)$', argument, re.I)
                expected_page = int(page_match.group(1)) if page_match else None
                new_snapshot = wait_for_grid_change(
                    previous_data_signature, expected_page=expected_page
                )
                if new_snapshot:
                    return new_snapshot
        except Exception:
            pass
        print(f'   ⚠️ صفحة النتائج لم تتغير — إعادة انتقال {attempt}/{PAGE_CHANGE_ATTEMPTS}')
        time.sleep(2)
    raise RuntimeError(f'تعذر فتح صفحة النتائج المطلوبة ({argument})')


def get_all_buildings(village, expected_codes=None, first_snapshot=None):
    """يجمع صفحات القرية ويعيد السجلات مع إثبات اكتمال قراءة الـPager."""
    expected = {int(x) for x in (expected_codes or []) if str(x).strip().isdigit()}
    by_id = {}; seen_page_signatures = set()
    snapshot = first_snapshot or read_grid_snapshot()
    scan_complete = False
    stop_reason = 'لم يكتمل الانتقال بين الصفحات'

    for page_counter in range(1, MAX_RESULT_PAGES + 1):
        if not snapshot:
            raise RuntimeError('اختفى جدول النتائج أثناء قراءة الصفحات')
        signature = grid_signature(snapshot)
        if not signature or signature in seen_page_signatures:
            stop_reason = 'تكررت صفحة النتائج قبل الوصول إلى النهاية'
            break
        seen_page_signatures.add(signature)

        page_records, wrong_decision, wrong_village = _extract_page_buildings(snapshot, village)
        for record in page_records:
            by_id.setdefault(record['id'], record)
        found_numbers = {item['number'] for item in by_id.values()}
        active_page = snapshot.get('activePage') or page_counter
        print(f'   📄 صفحة النتائج {active_page}: جُمِع {len(found_numbers)} مبنى فريد')
        if wrong_decision:
            print(f'      🛡️ تجاهل {wrong_decision} صفاً لا يطابق قرار التسجيل {REG_DECISION}')
        if wrong_village:
            print(f'      🛡️ تجاهل {wrong_village} صفاً لا يطابق اسم القرية')

        if expected and expected.issubset(found_numbers):
            print(f'   ✅ تم العثور على جميع أرقام Excel ({len(expected)}/{len(expected)})')
            scan_complete = True
            stop_reason = 'ظهرت جميع أرقام Excel'
            break

        next_argument = _next_page_argument(snapshot)
        if not next_argument:
            print('   ℹ️ وصلنا إلى آخر صفحة نتائج')
            # أكثر من صفحة ثم غياب زر التالي = نهاية موثوقة. في القرية الصغيرة
            # ذات الصفحة الواحدة نثق بها فقط عندما ظهرت كل أرقام Excel؛ وإلا
            # نسمح بصيغة بحث أخرى لمعالجة أعطال المنصة المؤقتة.
            scan_complete = (page_counter > 1 or not expected or
                             expected.issubset(found_numbers))
            stop_reason = ('تم الوصول إلى آخر صفحة نتائج'
                           if scan_complete
                           else 'صفحة واحدة ناقصة مقارنةً بـExcel')
            break
        print(f'   ➡️ الانتقال إلى {next_argument.replace("Page$", "صفحة ")}')
        try:
            snapshot = _move_to_results_page(
                next_argument, grid_data_signature(snapshot)
            )
        except Exception as exc:
            # نعيد ما جُمع من هذه الصيغة؛ الصيغ التالية والبحث الدقيق يكملان الناقص.
            print(f'   ⚠️ تعذر الانتقال للصفحة التالية: {str(exc)[:90]}')
            print('   ↩️ سيتم الاحتفاظ بالصفحات المجموعة ومتابعة البحث بصيغة أخرى')
            stop_reason = f'تعذر الانتقال للصفحة التالية: {str(exc)[:90]}'
            break
    else:
        raise RuntimeError(f'تجاوز عدد صفحات الأمان ({MAX_RESULT_PAGES})')

    records = sorted(by_id.values(), key=lambda item: (item['number'], int(item['id'])))
    return {
        'records': records,
        'complete': scan_complete,
        'pages': len(seen_page_signatures),
        'reason': stop_reason,
    }


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


IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.jfif', '.webp',
                    '.bmp', '.tif', '.tiff')
SKETCH_WORDS = ('كروكي', 'croquis', 'sketch', 'مخطط', 'رسم')


def discover_image_files(folder_path):
    """يجد الصور داخل مجلد القرية وكل المجلدات الفرعية."""
    found = []
    for current_root, dir_names, file_names in os.walk(folder_path):
        dir_names[:] = [name for name in dir_names
                        if not name.startswith('.') and name != '__pycache__']
        for file_name in file_names:
            if file_name.startswith('~$'):
                continue
            if file_name.lower().endswith(IMAGE_EXTENSIONS):
                found.append(os.path.abspath(os.path.join(current_root, file_name)))
    return sorted(set(found))


def _image_number_candidates(image_path, folder_path, expected_numbers):
    """يستنتج رقم المبنى من اسم الصورة أو أقرب مجلد، بدرجات أمان مختلفة."""
    expected = {int(number) for number in expected_numbers}
    relative = os.path.relpath(image_path, folder_path)
    parts = list(os.path.normpath(relative).split(os.sep))
    stem = os.path.splitext(parts[-1])[0]
    scores = {}

    def add(number_text, score):
        try:
            number = int(_english_digits(number_text))
        except (TypeError, ValueError):
            return
        if number in expected:
            scores[number] = max(scores.get(number, 0), score)

    normalized_stem = _english_digits(stem)
    if re.fullmatch(r'\s*0*(\d+)\s*', normalized_stem):
        add(re.fullmatch(r'\s*0*(\d+)\s*', normalized_stem).group(1), 120)

    # الرقم بعد آخر «_» هو رقم المبنى في أسماء مثل «الغرة 1_5» → 5.
    # أولوية عالية حتى لا يلتبس مع الرقم الموجود في اسم القرية («1»).
    trailing_underscore = re.search(r'_\s*0*(\d+)\s*$', normalized_stem)
    if trailing_underscore:
        add(trailing_underscore.group(1), 118)

    explicit_pattern = r'(?:مبن[ىي]|موقع|كروكي|رقم|building|site|sketch)\s*[-_#:]?\s*0*(\d+)'
    for match in re.finditer(explicit_pattern, normalized_stem, re.I):
        add(match.group(1), 115)

    # رقم وحيد صالح داخل اسم الصورة، مثل: village_17_croquis.jpg
    stem_numbers = {int(value) for value in re.findall(r'(?<!\d)(\d+)(?!\d)', normalized_stem)}
    valid_stem_numbers = stem_numbers & expected
    if len(valid_stem_numbers) == 1:
        add(next(iter(valid_stem_numbers)), 95)

    # إن كانت الصورة باسم "كروكي" داخل مجلد "17" أو "مبنى 17".
    for distance, parent_part in enumerate(reversed(parts[:-1]), 1):
        parent_text = _english_digits(parent_part)
        exact_parent = re.fullmatch(r'\s*0*(\d+)\s*', parent_text)
        if exact_parent:
            add(exact_parent.group(1), 110 - min(distance, 5))
            break
        explicit_parent = re.search(explicit_pattern, parent_text, re.I)
        if explicit_parent:
            add(explicit_parent.group(1), 105 - min(distance, 5))
            break

    relative_lower = relative.lower()
    sketch_bonus = 30 if any(word in relative_lower for word in SKETCH_WORDS) else 0
    return {number: score + sketch_bonus for number, score in scores.items()}


def index_numeric_images(image_paths, folder_path, expected_numbers):
    """يربط كل رقم Excel بصورة واحدة مؤكدة، ويترك الحالات الملتبسة بلا ربط."""
    candidates_by_number = {}
    unmatched = []
    conflicting_number = []
    for image_path in image_paths:
        candidates = _image_number_candidates(image_path, folder_path, expected_numbers)
        if not candidates:
            unmatched.append(image_path)
            continue
        best_score = max(candidates.values())
        best_numbers = [number for number, score in candidates.items()
                        if score == best_score]
        if len(best_numbers) != 1:
            conflicting_number.append(image_path)
            continue
        number = best_numbers[0]
        candidates_by_number.setdefault(number, []).append((best_score, image_path))

    indexed = {}
    ambiguous = {}
    for number, candidates in candidates_by_number.items():
        candidates.sort(key=lambda item: (-item[0], item[1].lower()))
        best_score = candidates[0][0]
        best = [path for score, path in candidates if score == best_score]
        if len(best) == 1:
            indexed[number] = best[0]
        else:
            # أكثر من ملف بنفس قوة المطابقة: لا نخاطر برفع كروكي لمبنى آخر.
            ambiguous[number] = best
    return indexed, unmatched, conflicting_number, ambiguous


def numbers_match(actual, expected, tolerance=0.0001):
    """مطابقة رقمية تتحمل الفواصل العادية، وتمنع تقرير نجاح غير حقيقي."""
    try:
        actual_num = float(str(actual).replace(',', '').strip())
        expected_num = float(expected)
        return abs(actual_num - expected_num) < tolerance
    except (TypeError, ValueError):
        return False


def open_building_form(base_url, attempts=4):
    """يفتح تبويب /2 ويتأكد أن نموذج المبنى نفسه اكتمل قبل أي كتابة."""
    last_error = ''
    for attempt in range(1, attempts + 1):
        ensure_automation_tab()
        ensure_alive()
        try:
            driver.get(base_url)
            end_time = time.time() + 35
            while time.time() < end_time:
                if ('Login' in driver.current_url or driver.title == 'Runtime Error'):
                    raise RuntimeError('انتهت جلسة الدخول أو ظهرت Runtime Error')
                x_fields = driver.find_elements(
                    'xpath', "//input[contains(@name,'ctl37$Field1$ctl03')]"
                )
                save_buttons = driver.find_elements(
                    'xpath', "//input[@name='ctl12$btnSave']"
                )
                if x_fields and save_buttons:
                    return True
                time.sleep(0.5)
            last_error = 'لم تكتمل حقول نموذج المبنى خلال 35 ثانية'
        except Exception as exc:
            last_error = str(exc)
        print(f'   ⚠️ تعذر فتح نموذج المبنى — محاولة {attempt}/{attempts}: {last_error[:80]}')
        if not is_connection_alive():
            reconnect_driver()
        time.sleep(3)
    raise RuntimeError(last_error or 'تعذر فتح نموذج المبنى')


def process_building(building_id, building_num, data_by_building, images_by_building):
    ensure_automation_tab()
    data = data_by_building[building_num]
    base_url = (
        f'https://heritage2.itqan-consultant.com'
        f'/Web/App/Pools/DataEdit/{building_id}/2'
    )
    print(f'\n🏢 مبنى {building_num}')
    open_building_form(base_url)

    if TEST_MODE:
        required = ['ctl37$Field1$ctl03', 'ctl38$Field1$ctl03',
                    'ctl39$Field1$ctl03', 'ctl12$btnSave']
        missing = []
        for field_name in required:
            selector = (f"//input[contains(@name,'{field_name}')]"
                        if field_name != 'ctl12$btnSave'
                        else "//input[@name='ctl12$btnSave']")
            if not driver.find_elements('xpath', selector):
                missing.append(field_name)
        if building_num in images_by_building and not driver.find_elements(
            'xpath', "//input[@type='file' and contains(@name,'ctl40$Field1$ctl27')]"
        ):
            missing.append('ctl40$Field1$ctl27')
        if missing:
            print(f'   ❌ فشل الاختبار: الحقول غير موجودة: {", ".join(missing)}')
            return False
        print('   🧪 نجح الاختبار: الصفحة والحقول المطلوبة موجودة — بلا تعديل')
        return True

    # تشغيل قابل للاستئناف: ما تم بنجاح في تشغيل سابق لا نعيد رفعه أو حذفه.
    def _value_now(field_name):
        try:
            return driver.find_element(
                'xpath', f"//input[contains(@name,'{field_name}')]"
            ).get_attribute('value') or ''
        except Exception:
            return ''
    existing_x = _value_now('ctl37$Field1$ctl03')
    existing_y = _value_now('ctl38$Field1$ctl03')
    existing_area = _value_now('ctl39$Field1$ctl03')
    existing_image = _value_now('ctl40$Field1$ctl28')
    already_complete = (
        numbers_match(existing_x, data['x']) and
        numbers_match(existing_y, data['y']) and
        numbers_match(existing_area, data['area']) and
        bool(str(existing_image).strip())
    )
    if already_complete:
        print('   ♻️ الموقع مكتمل ومتحقق منه من تشغيل سابق — لا حاجة لإعادة الرفع')
        return True

    # غياب الملف المحلي ليس عطلاً مؤقتاً؛ لا نكتب الحقول ولا نعيد فتح الصفحة 4 مرات.
    if building_num not in images_by_building:
        print(f'   ⛔ لا توجد صورة كروكي محلية مطابقة لمبنى {building_num}')
        return False

    print('   📝 تعبئة X / Y / المساحة...')
    # نكتب كل حقل ونتأكد أنه انكتب فعلاً؛ لو فشل نعيد حتى 4 مرات
    fields_ok = True
    for _fname, _fval in [('ctl37$Field1$ctl03', data['x']),
                          ('ctl38$Field1$ctl03', data['y']),
                          ('ctl39$Field1$ctl03', data['area'])]:
        field_ok = False
        for _try in range(4):
            if fill_field(_fname, _fval):
                field_ok = True
                break
            time.sleep(0.4)
        fields_ok = fields_ok and field_ok
    if not fields_ok:
        print('   ⛔ لن يتم الحفظ — تعذّرت تعبئة حقل واحد أو أكثر')
        return False
    time.sleep(1.5)

    image_uploaded = upload_file('ctl40', images_by_building[building_num])
    if not image_uploaded:
        print(f'   ⛔ لن يتم الحفظ — فشل تأكيد رفع الصورة')
        return False

    try:
        save_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSave']")
        js_click(save_btn)
        print('   💾 تم الضغط على حفظ')
        time.sleep(6)
    except Exception as e:
        print(f'   ❌ خطأ في الحفظ: {e}')
        return False

    open_building_form(base_url)
    check_x    = get_field_value_with_retry('ctl37$Field1$ctl03')
    check_y    = get_field_value_with_retry('ctl38$Field1$ctl03')
    check_area = get_field_value_with_retry('ctl39$Field1$ctl03')
    check_image = get_field_value_with_retry('ctl40$Field1$ctl28')

    try:
        ok = (numbers_match(check_x, data['x']) and
              numbers_match(check_y, data['y']) and
              numbers_match(check_area, data['area']) and
              bool(str(check_image).strip()))
    except Exception:
        ok = False

    print(f'   🔄 بعد إعادة التحميل: X={check_x}, Y={check_y}, مساحة={check_area}, '
          f'كروكي={"موجود" if check_image else "مفقود"}')

    if not ok:
        print('   🔁 لم تنجح المحاولة الأولى — إعادة محاولة كاملة...')
        open_building_form(base_url)
        for _fname, _fval in [('ctl37$Field1$ctl03', data['x']),
                              ('ctl38$Field1$ctl03', data['y']),
                              ('ctl39$Field1$ctl03', data['area'])]:
            for _try in range(4):
                if fill_field(_fname, _fval):
                    break
                time.sleep(0.4)
        time.sleep(1.5)
        if not upload_file('ctl40', images_by_building[building_num]):
            print('   ❌ فشل تأكيد رفع الكروكي في المحاولة الثانية')
            return False
        try:
            save_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSave']")
            js_click(save_btn)
            print('   💾 تم الضغط على حفظ (محاولة ثانية)')
            time.sleep(6)
        except Exception as e:
            print(f'   ❌ خطأ في الحفظ بالمحاولة الثانية: {e}')
            return False

        open_building_form(base_url)
        check_x    = get_field_value_with_retry('ctl37$Field1$ctl03')
        check_y    = get_field_value_with_retry('ctl38$Field1$ctl03')
        check_area = get_field_value_with_retry('ctl39$Field1$ctl03')
        check_image = get_field_value_with_retry('ctl40$Field1$ctl28')
        try:
            ok = (numbers_match(check_x, data['x']) and
                  numbers_match(check_y, data['y']) and
                  numbers_match(check_area, data['area']) and
                  bool(str(check_image).strip()))
        except Exception:
            ok = False
        print(f'   🔄 بعد المحاولة الثانية: X={check_x}, Y={check_y}, مساحة={check_area}, '
              f'كروكي={"موجود" if check_image else "مفقود"}')

    return ok


# ── 6) تسجيل الدخول (تلقائي مع رجوع يدوي عند الحاجة) ───────
# حقول البحث ctl04 وctl28 موجودة في DataView، لا في Inbox.
DATA_VIEW_URL = 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataView'


def is_logged_in():
    """يتأكد أننا داخل المنصة فعلاً: صفحة DataView مع حقل البحث جاهز."""
    try:
        driver.get(DATA_VIEW_URL)
        time.sleep(3)
        if driver.title == 'Runtime Error' or 'Login' in driver.current_url:
            return False
        # حقل بحث الموقع لا يظهر إلا بعد الدخول الصحيح.
        return bool(driver.find_elements('xpath', "//input[@name='ctl12$ctl04']"))
    except Exception:
        return False


def auto_login():
    """
    يملأ اسم المستخدم وكلمة المرور تلقائياً في صفحة الدخول ويضغط دخول.
    يكتشف الحقول بعدة طرق (name/id/placeholder/type) حتى لو تغيّر شكل
    الصفحة قليلاً. يرجّع True لو نجح الدخول فعلاً (ظهرت DataView).
    """
    if not (LOGIN_USERNAME and LOGIN_PASSWORD):
        return False
    print('🔐 محاولة تسجيل الدخول التلقائي...')
    try:
        driver.get(LOGIN_URL)
        time.sleep(3)

        # ننتظر ظهور حقل كلمة المرور (المؤشر الأوثق لصفحة الدخول).
        password_input = None
        end_time = time.time() + 25
        while time.time() < end_time:
            candidates = driver.find_elements('xpath', "//input[@type='password']")
            candidates = [c for c in candidates if c.is_displayed()]
            if candidates:
                password_input = candidates[0]
                break
            time.sleep(0.5)
        if password_input is None:
            print('   ⚠️ لم أجد حقل كلمة المرور في صفحة الدخول')
            return False

        # حقل اسم المستخدم: نفضّل حقلاً باسم/معرّف/تلميح يدل على المستخدم،
        # وإلا أول حقل نصي ظاهر قبل حقل كلمة المرور.
        user_input = None
        keyword_xpath = (
            "//input[(@type='text' or @type='email' or not(@type)) and ("
            "contains(translate(@name,'USERNAMELOGINE','usernamelogine'),'user') or "
            "contains(translate(@id,'USERNAMELOGINE','usernamelogine'),'user') or "
            "contains(translate(@name,'LOGIN','login'),'login') or "
            "contains(translate(@id,'LOGIN','login'),'login') or "
            "contains(@name,'Email') or contains(@id,'Email') or "
            "contains(@placeholder,'اسم') or contains(@placeholder,'مستخدم') or "
            "contains(@placeholder,'User') or contains(@placeholder,'user'))]"
        )
        matches = [e for e in driver.find_elements('xpath', keyword_xpath) if e.is_displayed()]
        if matches:
            user_input = matches[0]
        else:
            text_inputs = driver.find_elements(
                'xpath', "//input[@type='text' or @type='email' or not(@type)]"
            )
            text_inputs = [e for e in text_inputs if e.is_displayed()]
            if text_inputs:
                user_input = text_inputs[0]
        if user_input is None:
            print('   ⚠️ لم أجد حقل اسم المستخدم في صفحة الدخول')
            return False

        # الكتابة الفعلية (send_keys أوثق من JS في نماذج الدخول).
        try:
            user_input.clear()
        except Exception:
            pass
        user_input.send_keys(LOGIN_USERNAME)
        try:
            password_input.clear()
        except Exception:
            pass
        password_input.send_keys(LOGIN_PASSWORD)
        print(f'   ⌨️ تم إدخال اسم المستخدم: {LOGIN_USERNAME}')
        time.sleep(0.5)

        # زر الدخول: submit أو زر فيه كلمة دخول/login، وإلا Enter.
        login_button = None
        button_xpath = (
            "//input[@type='submit'] | //button[@type='submit'] | "
            "//input[@type='button'][contains(@value,'دخول') or contains(@value,'Login') "
            "or contains(@value,'login') or contains(@value,'تسجيل')] | "
            "//button[contains(.,'دخول') or contains(.,'Login') or contains(.,'login') "
            "or contains(.,'تسجيل')]"
        )
        buttons = [b for b in driver.find_elements('xpath', button_xpath) if b.is_displayed()]
        if buttons:
            login_button = buttons[0]

        if login_button is not None:
            try:
                driver.execute_script("arguments[0].click();", login_button)
            except Exception:
                login_button.click()
        else:
            from selenium.webdriver.common.keys import Keys as _Keys
            password_input.send_keys(_Keys.RETURN)
        print('   ⏳ تم الضغط على دخول — جاري التحقق...')
        time.sleep(4)

        if is_logged_in():
            print('   ✅ نجح تسجيل الدخول التلقائي')
            return True
        print('   ⚠️ لم يكتمل الدخول التلقائي (بيانات خاطئة أو شكل صفحة مختلف)')
        return False
    except Exception as exc:
        print(f'   ⚠️ تعثر الدخول التلقائي: {str(exc)[:80]}')
        return False


# المنطق: لو الجلسة محفوظة من قبل ندخل مباشرة؛ وإلا نجرب التلقائي؛
# وإلا نطلب الدخول اليدوي مرة واحدة (يبقى محفوظاً في البروفايل).
driver.get(DATA_VIEW_URL)
time.sleep(3)
print(f'📍 الصفحة: {driver.title}')

if driver.title == 'Runtime Error' or 'Login' in driver.current_url:
    if not auto_login():
        driver.get(LOGIN_URL)
        time.sleep(3)
        input('\n👤 الدخول التلقائي لم ينجح — سجّل دخولك يدوياً في نافذة '
              'Chrome، ثم اضغط Enter هنا للمتابعة...')
    driver.get(DATA_VIEW_URL)
    time.sleep(3)
    print(f'✅ الصفحة بعد الدخول: {driver.title}')
else:
    print('✅ الجلسة محفوظة من قبل — لا حاجة لإعادة تسجيل الدخول')

# ── 7) الحلقة الرئيسية على كل قرية ────────────────────────
grand_total = {'villages': 0, 'processed': 0, 'skipped': 0, 'failed': 0}
followup_by_village = {}
all_results_by_village = {}

# التقرير ينشأ فوراً ويُحفظ كل 10 مواقع، حتى لا يضيع عند أي توقف مفاجئ.
master_report = os.path.join(
    REPORTS_DIR,
    f'تقرير_المتابعة_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
)
live_workbook = Workbook()
live_sheet = live_workbook.active
live_sheet.title = 'سجل المواقع'
live_sheet.append(['القرية', 'رقم المبنى/الكود', 'الحالة', 'التفاصيل', 'الوقت'])
live_saved_rows = {'count': 0}
live_workbook.save(master_report)
print(f'💾 تم إنشاء تقرير مباشر: {master_report}')


def save_live_report():
    """حفظ ذري بعد كل موقع؛ تعطل العملية لا يترك ملف Excel نصف مكتوب."""
    temp_report = master_report + '.saving.xlsx'
    try:
        live_workbook.save(temp_report)
        os.replace(temp_report, master_report)
        return True
    except Exception as exc:
        print(f'   ⚠️ تعذر تحديث التقرير الآن (سيحاول بعد الموقع التالي): {str(exc)[:70]}')
        try:
            if os.path.exists(temp_report):
                os.remove(temp_report)
        except Exception:
            pass
        return False


def record_live_result(village, building, status, details, timestamp):
    """يحفظ سجلًا مباشرًا؛ الملف موجود حتى لو توقف البرنامج قبل النهاية."""
    status_text = {'success': '✅ ناجح', 'skip': '⏭️ متخطى',
                   'error': '❌ فشل'}.get(status, status)
    live_sheet.append([village, building, status_text, details, timestamp])
    live_saved_rows['count'] += 1
    if save_live_report():
        print('   💾 تم حفظ نتيجة الموقع في تقرير المتابعة')


def normalize_decimal_digits(value):
    """يوحّد الأرقام العربية والإنجليزية للمقارنة الصارمة."""
    return str(value or '').strip().translate(str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789'))


def search_fields_match_request(search_value):
    """يتحقق من حقلي البحث فقط؛ مطابقة كل صف للقرار والقرية تتم أثناء جمعه."""
    try:
        current_village = driver.find_element(
            'xpath', "//input[@name='ctl12$ctl04']"
        ).get_attribute('value') or ''
        if current_village.strip() != search_value.strip():
            print(f'   🛑 قيمة الموقع تغيّرت في الصفحة: "{current_village}"')
            return False
        if REG_DECISION:
            current_decision = driver.find_element(
                'xpath', "//input[@name='ctl12$ctl28']"
            ).get_attribute('value') or ''
            if normalize_decimal_digits(current_decision) != normalize_decimal_digits(REG_DECISION):
                print(f'   🛑 قرار التسجيل في الصفحة غير مطابق: "{current_decision}"')
                return False
        return True
    except Exception as e:
        print(f'   🛑 تعذر التحقق من فلاتر البحث: {str(e)[:70]}')
        return False


def _search_query_variants(village):
    """ينشئ صيغ بحث من الأخص إلى الأوسع لمعالجة اختلاف اسم المجلد عن المنصة."""
    original = re.sub(r'\s+', ' ', str(village).strip())
    core = re.sub(r'[_\-–—]+', ' ', original)
    core = re.sub(
        r'(^|\s)(?:قرية|قريه|التراثية|التراثيه|التراثي)(?=\s|$)',
        ' ', core, flags=re.I
    )
    core = re.sub(r'\s+', ' ', core).strip()
    variants = []

    def add(value):
        value = re.sub(r'\s+', ' ', str(value or '').strip())
        if value and value not in variants:
            variants.append(value)

    add(original)
    add(core)
    for word in core.split():
        normalized = _simple_arabic_name(word)
        if len(normalized) >= 4 and normalized not in {'قريه', 'التراثيه'}:
            add(word)
    return variants


def _clean_exact_village_anchor(value):
    """ينظف وصف القرية للبحث عن مبنى، مع إبقاء كلمة «قرية» المهمة للمنصة."""
    text = re.sub(r'[_\-–—]+', ' ', str(value or '').strip())
    text = re.sub(
        r'(^|\s)(?:التراثية|التراثيه|التراثي)(?=\s|$)',
        ' ', text, flags=re.I
    )
    text = re.sub(r'\s+', ' ', text).strip(' ،,.-')
    return text


def _exact_village_anchors(village, collected_records):
    """يستخرج اسم القرية كما كتبته المنصة من النتائج التي جُمعت بالفعل.

    مثال: «مبنى رقم 7 في قرية المعدونة» يعطينا «قرية المعدونة».
    بهذه الطريقة لا يتحول البحث الدقيق إلى «في المعدونة» ولا نخمن الاسم.
    """
    counts = {}
    for record in collected_records:
        site_name = re.sub(r'\s+', ' ', str(record.get('name', '')).strip())
        phrase = ''

        # نحافظ على كلمة «قرية/قريه» كما ظهرت في اسم الموقع بالمنصة.
        matches = list(re.finditer(
            r'(?:^|\s)في\s+((?:قرية|قريه)\s+.+)$', site_name
        ))
        if matches:
            phrase = matches[-1].group(1)
        else:
            attached = list(re.finditer(
                r'(?:^|\s)ب(?:قرية|قريه)\s+(.+)$', site_name
            ))
            if attached:
                phrase = 'قرية ' + attached[-1].group(1)

        phrase = _clean_exact_village_anchor(phrase)
        if phrase and _belongs_to_village(phrase, village):
            counts[phrase] = counts.get(phrase, 0) + 1

    # الأكثر تكراراً في نتائج المنصة أولاً؛ وعند التعادل نأخذ الأكثر تحديداً.
    anchors = [
        phrase for phrase, _ in sorted(
            counts.items(), key=lambda item: (-item[1], -len(item[0]), item[0])
        )
    ]

    # خطة احتياطية من اسم المجلد: نحذف الوصف العام فقط ولا نحذف كلمة «قرية».
    fallback = _clean_exact_village_anchor(village)
    if fallback and fallback not in anchors:
        anchors.append(fallback)

    # بعض المجلدات اسمها يبدأ خطأً بـ «قرية قرية»؛ نجرب الصيغة الطبيعية أيضاً.
    natural_fallback = re.sub(
        r'^(?:(?:قرية|قريه)\s+){2,}', 'قرية ', fallback
    ).strip()
    if natural_fallback and natural_fallback not in anchors:
        anchors.append(natural_fallback)
    return anchors


def _search_once(search_value, village, expected_codes=None,
                 full_attempts=2, result_timeout=SEARCH_RESULT_TIMEOUT):
    """ينفذ صيغة بحث واحدة ويعيد كل صفحاتها بعد ثبات الجدول."""
    ensure_automation_tab()
    for full_attempt in range(1, full_attempts + 1):
        driver.get(DATA_VIEW_URL)
        time.sleep(3)
        search_field = driver.find_element('xpath', "//input[@name='ctl12$ctl04']")
        print('📍 حقل البحث المختار: الموقع أو كود')
        driver.execute_script("""
            var el = arguments[0]; var val = arguments[1];
            el.value = val;
            el.dispatchEvent(new Event('input',  {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
        """, search_field, search_value)
        time.sleep(0.3)
        if REG_DECISION:
            reg_field = driver.find_element('xpath', "//input[@name='ctl12$ctl28']")
            driver.execute_script("""
                var el = arguments[0]; var val = arguments[1];
                el.value = val;
                el.dispatchEvent(new Event('input',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
            """, reg_field, REG_DECISION)
            time.sleep(0.3)

        if not search_fields_match_request(search_value):
            print('   ⚠️ لم تثبت قيم البحث في الحقول — إعادة تعبئتها')
            continue

        # نلتقط الجدول القديم قبل الضغط حتى لا نقبل نتائج قرية سابقة.
        old_signature = grid_signature()
        # ضغطتان كحد أقصى من الصفحة نفسها، والثانية فقط إذا لم يصل أي جدول جديد.
        for click_attempt in range(1, 3):
            search_btn = driver.find_element('xpath', "//input[@name='ctl12$btnSearch']")
            driver.execute_script("arguments[0].click();", search_btn)
            print(f'🔍 ضغط بحث {click_attempt}/2: {search_value}'
                  + (f' | قرار: {REG_DECISION}' if REG_DECISION else ''))
            print(f'   ⏳ انتظار النتيجة الجديدة حتى {result_timeout} ثانية...')
            first_snapshot = wait_for_fresh_results(old_signature, timeout=result_timeout)
            if first_snapshot and search_fields_match_request(search_value):
                print('   ✅ ظهرت النتائج — تم إيقاف ضغط البحث وبدأ جمع الصفحات')
                scan_result = get_all_buildings(
                    village, expected_codes=expected_codes,
                    first_snapshot=first_snapshot
                )
                buildings = scan_result['records']
                if buildings:
                    return scan_result
                print('   ⚠️ الجدول ظهر لكن لم يحتوِ صفوفاً مطابقة للقرية والقرار')
                # الجدول وصل فعلاً؛ إعادة ضغط الزر نفسه لن تغيّر هويّة الصفوف.
                return scan_result
            if click_attempt < 2:
                print('   ⚠️ لم يصل جدول جديد — انتظار 8 ثوانٍ ثم ضغط البحث مرة ثانية')
                time.sleep(8)

        print(f'   ⚠️ لم تصل نتيجة صالحة — إعادة فتح صفحة البحث '
              f'({full_attempt}/{full_attempts})')

    return {
        'records': [], 'complete': False, 'pages': 0,
        'reason': 'لم يصل جدول نتائج جديد'
    }


def _merge_search_results(merged, records, expected):
    """يدمج النتائج من صيغ بحث مختلفة، ولا يسمح برقم خارج Excel."""
    for record in records:
        if expected and record.get('number') not in expected:
            continue
        merged.setdefault(record['id'], record)


def search_buildings_for_village(village, expected_codes=None):
    """يجمع القرية مرة كاملة؛ عند بلوغ آخر صفحة يبدأ التنفيذ بلا بحث دائري."""
    expected = {int(x) for x in (expected_codes or []) if str(x).strip().isdigit()}
    variants = _search_query_variants(village)
    merged = {}
    authoritative_records = None
    village_identity = ' + '.join(_arabic_name_tokens(village)) or village
    print(f'🔐 هوية القرية الصارمة: {village_identity}')
    if 'الماوين' in _arabic_name_tokens(village):
        wanted_branch = ('الماوين بللحمر'
                         if 'بللحمر' in _arabic_name_tokens(village)
                         else 'الماوين فقط (منع بللحمر)')
        print(f'🛡️ فصل القرى المتشابهة مفعّل: {wanted_branch}')
    print(f'🧭 صيغ البحث التلقائية: {" | ".join(variants)}')

    for variant_index, search_value in enumerate(variants, 1):
        found_numbers = {item['number'] for item in merged.values()}
        remaining = expected - found_numbers if expected else set()
        if expected and not remaining:
            break
        print(f'\n🧭 صيغة البحث {variant_index}/{len(variants)}: "{search_value}"')
        try:
            scan_result = _search_once(
                search_value, village,
                expected_codes=(remaining if expected else None),
                full_attempts=2
            )
            records = scan_result['records']
        except Exception as exc:
            records = []
            scan_result = {
                'records': [], 'complete': False, 'pages': 0,
                'reason': str(exc)
            }
            print(f'   ⚠️ تعثرت هذه الصيغة، سيكمل بما بعدها: {str(exc)[:100]}')
            if not is_connection_alive():
                reconnect_driver()
        _merge_search_results(merged, records, expected)
        found_numbers = {item['number'] for item in merged.values()}
        if expected:
            print(f'   📦 الإجمالي بعد الدمج: {len(found_numbers)}/{len(expected)} من أرقام Excel')
        elif merged:
            break

        # إذا وصلنا فعلياً إلى آخر صفحة، فهذه هي قائمة القرية الكاملة.
        # اختلاف عددها عن Excel يعني فرق بيانات، وليس سبباً لإعادة البحث.
        if records and scan_result.get('complete'):
            authoritative_records = records
            print(f'   ✅ اكتملت قراءة كل صفحات القرية ({scan_result.get("pages", 0)} صفحة)')
            print('   ▶️ لن يُعاد البحث عن القرية — سيبدأ تنفيذ الاستمارات مباشرة')
            break
        if records:
            print(f'   ⚠️ القراءة غير مكتملة: {scan_result.get("reason", "سبب غير معروف")}')
            print('   ↩️ سيتم تجربة صيغة البحث التالية فقط لإكمال الصفحات الناقصة')

    found_numbers = {item['number'] for item in merged.values()}
    if expected:
        missing = sorted(expected - found_numbers)
        platform_numbers = {
            item['number'] for item in (authoritative_records or [])
        }
        platform_only = sorted(platform_numbers - expected)
        if missing:
            preview = ', '.join(map(str, missing[:20]))
            suffix = ' ...' if len(missing) > 20 else ''
            print(f'⚠️ فرق بيانات: {len(missing)} رقم Excel غير ظاهر في جميع صفحات المنصة: '
                  f'{preview}{suffix}')
            print('   ⏭️ ستُسجل هذه الأرقام في التقرير، وسيبدأ تنفيذ الأرقام المطابقة فوراً')
        if platform_only:
            preview = ', '.join(map(str, platform_only[:20]))
            suffix = ' ...' if len(platform_only) > 20 else ''
            print(f'ℹ️ أرقام ظهرت في المنصة وغير موجودة في Excel: {preview}{suffix}')
        else:
            if not missing:
                print(f'✅ اكتملت قائمة القرية: {len(found_numbers)}/{len(expected)} رقم')
    return sorted(merged.values(), key=lambda item: (item['number'], int(item['id'])))

for v_idx, cfg in enumerate(villages_config[START_VILLAGE - 1:], start=START_VILLAGE):
    village          = cfg['village']
    FOLDER_PATH      = cfg['folder']
    FILTER_BUILDINGS = cfg['filter']

    print(f'\n{"█"*55}')
    print(f'█  قرية {v_idx}/{len(villages_config)}: {village}')
    if FILTER_BUILDINGS:
        print(f'█  🎯 مباني محددة: {FILTER_BUILDINGS}')
    print(f'{"█"*55}')

    # ── إيجاد ملف/شيت Excel الصحيح حسب عدد الصفوف الصالحة ──
    excel_files = []
    for excel_root, excel_dirs, excel_names in os.walk(FOLDER_PATH):
        excel_dirs[:] = [name for name in excel_dirs
                         if not name.startswith('.') and name != '__pycache__']
        for excel_name in excel_names:
            if (excel_name.lower().endswith(('.xlsx', '.xls'))
                    and not excel_name.startswith('~$')
                    and not excel_name.lower().startswith('report_')
                    and not excel_name.startswith('تقرير_')):
                excel_files.append(os.path.join(excel_root, excel_name))
    excel_files = sorted(set(excel_files))
    if not excel_files:
        print(f'❌ لا يوجد ملف Excel في مجلد "{village}"')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'skip',
            'التفاصيل': 'لا يوجد ملف Excel في مجلد القرية',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        all_results_by_village[village] = followup_by_village[village]
        _issue = followup_by_village[village][0]
        record_live_result(village, _issue['رقم المبنى'], _issue['الحالة'],
                           _issue['التفاصيل'], _issue['الوقت'])
        live_workbook.save(master_report)
        grand_total['villages'] += 1; grand_total['skipped'] += 1
        continue

    # نفحص كل الشيتات، ولا نختار ملفاً فارغاً لمجرد أن أسماء أعمدته صحيحة.
    EXCEL_PATH = None
    df = None
    x_col = y_col = area_col = None
    EXCEL_SHEET = None
    excel_candidates = []
    for cpath in excel_files:
        relative_excel = os.path.relpath(cpath, FOLDER_PATH)
        try:
            excel_book = pd.ExcelFile(cpath)
        except Exception as exc:
            print(f'   ⏭️ تعذر فتح "{relative_excel}": {str(exc)[:70]}')
            continue
        for sheet_name in excel_book.sheet_names:
            try:
                cdf = pd.read_excel(cpath, sheet_name=sheet_name)
            except Exception:
                continue
            cx = find_col(cdf, 'الطول')
            cy = find_col(cdf, 'العرض')
            ca = find_col(cdf, 'مساحة')
            if not (cx and cy and ca):
                continue
            valid_count = count_valid_coordinate_rows(cdf, cx, cy, ca)
            print(f'   📋 مرشح Excel: "{relative_excel}" | شيت "{sheet_name}" '
                  f'| صفوف صالحة {valid_count}/{len(cdf)}')
            excel_candidates.append({
                'path': cpath, 'relative': relative_excel, 'sheet': sheet_name,
                'df': cdf, 'x': cx, 'y': cy, 'area': ca,
                'valid': valid_count, 'rows': len(cdf),
            })
        try:
            excel_book.close()
        except Exception:
            pass

    if excel_candidates:
        selected_excel = max(
            excel_candidates,
            key=lambda item: (item['valid'], item['rows'],
                              -item['relative'].count(os.sep), item['relative'])
        )
        EXCEL_PATH = selected_excel['path']
        EXCEL_SHEET = selected_excel['sheet']
        df = selected_excel['df']
        x_col, y_col, area_col = (
            selected_excel['x'], selected_excel['y'], selected_excel['area']
        )
        print(f'📊 ملف Excel المختار: {selected_excel["relative"]} '
              f'| شيت: {EXCEL_SHEET} | صفوف صالحة: '
              f'{selected_excel["valid"]}/{selected_excel["rows"]}')

    if EXCEL_PATH is None:
        print(f'❌ لا يوجد ملف Excel فيه أعمدة الطول/العرض/المساحة في "{village}"')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'error',
            'التفاصيل': 'لا يوجد ملف Excel صالح (أعمدة الإحداثيات ناقصة)',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        all_results_by_village[village] = followup_by_village[village]
        _issue = followup_by_village[village][0]
        record_live_result(village, _issue['رقم المبنى'], _issue['الحالة'],
                           _issue['التفاصيل'], _issue['الوقت'])
        live_workbook.save(master_report)
        grand_total['villages'] += 1; grand_total['failed'] += 1
        continue

    coded_col = next(
        (col for col in df.columns if str(col).strip().lower() == 'coded'),
        None
    )
    if coded_col is None:
        code_candidates = [
            col for col in df.columns
            if ('code' in str(col).strip().lower() or 'كود' in str(col))
        ]
        if len(code_candidates) == 1:
            coded_col = code_candidates[0]
            print(f'ℹ️ تم اعتماد عمود الكود: {coded_col}')
        else:
            print(f'❌ لا يوجد عمود coded/كود واضح في ملف "{os.path.basename(EXCEL_PATH)}"')
            followup_by_village[village] = [{
                'رقم المبنى': '—', 'الحالة': 'error',
                'التفاصيل': 'لا يوجد عمود coded أو كود واضح في Excel',
                'الوقت': datetime.now().strftime('%H:%M:%S')}]
            all_results_by_village[village] = followup_by_village[village]
            _issue = followup_by_village[village][0]
            record_live_result(village, _issue['رقم المبنى'], _issue['الحالة'],
                               _issue['التفاصيل'], _issue['الوقت'])
            live_workbook.save(master_report)
            grand_total['villages'] += 1; grand_total['failed'] += 1
            continue

    # ── اكتشاف: عمود coded أرقام أو أسماء؟ ────────────────
    NAME_MODE = not coded_is_numeric(df[coded_col])
    print('🔤 وضع المطابقة: ' + ('بالاسم (coded أسماء)' if NAME_MODE
                                  else 'بالرقم (coded أرقام)'))
    print(f'📌 أعمدة Excel: الكود="{coded_col}" | X="{x_col}" | '
          f'Y="{y_col}" | المساحة="{area_col}"')

    data_by_building = {}
    invalid_excel_rows = []
    for excel_row_number, (_, row) in enumerate(df.iterrows(), start=2):
        raw_key = row[coded_col]
        if NAME_MODE:
            if pd.isna(raw_key):
                continue
            key = str(raw_key).strip()
            if not key or key.lower() == 'nan':
                continue
        else:
            try:
                key_number = parse_excel_number(raw_key)
                if abs(key_number - round(key_number)) > 0.000001:
                    raise ValueError('الكود ليس عدداً صحيحاً')
                key = int(round(key_number))
            except (ValueError, TypeError):
                continue

        raw_x = row[x_col]; raw_y = row[y_col]; raw_area = row[area_col]
        try:
            x_value = parse_excel_number(raw_x)
            y_value = parse_excel_number(raw_y)
            area_value = parse_excel_number(raw_area)
        except (TypeError, ValueError) as exc:
            details = (
                f'صف Excel {excel_row_number}: قيمة غير رقمية — '
                f'X="{short_excel_value(raw_x)}"، '
                f'Y="{short_excel_value(raw_y)}"، '
                f'المساحة="{short_excel_value(raw_area)}" '
                f'({str(exc)[:45]})'
            )
            print(f'   ⚠️ تخطي "{key}": {details}')
            invalid_excel_rows.append({'key': key, 'details': details})
            continue
        if key in data_by_building:
            print(f'   ⚠️ تكرار الكود "{key}" في Excel — سيتم اعتماد آخر صف')
        data_by_building[key] = {
            'x': x_value,
            'y': y_value,
            'area': area_value,
        }
    print(f'🏢 المواقع في Excel: {len(data_by_building)} صالحة'
          + (f' | {len(invalid_excel_rows)} غير صالحة' if invalid_excel_rows else ''))

    # لا نبحث في المنصة بقائمة Excel فارغة ولا نعلن نجاحاً وهمياً.
    if not data_by_building:
        print('❌ لم توجد أي صفوف Excel صالحة — لن يبدأ تعديل المنصة لهذه القرية')
        now_text = datetime.now().strftime('%H:%M:%S')
        issues = [
            {'رقم المبنى': item['key'], 'الحالة': 'error',
             'التفاصيل': item['details'], 'الوقت': now_text}
            for item in invalid_excel_rows
        ] or [{
            'رقم المبنى': '—', 'الحالة': 'error',
            'التفاصيل': 'ملف Excel لا يحتوي أي صف صالح للإحداثيات والمساحة',
            'الوقت': now_text
        }]
        followup_by_village[village] = issues
        all_results_by_village[village] = issues
        for issue in issues:
            record_live_result(village, issue['رقم المبنى'], issue['الحالة'],
                               issue['التفاصيل'], issue['الوقت'])
        grand_total['villages'] += 1
        grand_total['failed'] += len(issues)
        save_live_report()
        continue

    # ── فهرسة الصور داخل مجلد القرية وكل المجلدات الفرعية ──
    #   وضع الرقم: من اسم الصورة أو اسم مجلد المبنى.
    #   وضع الاسم: من كامل المسار النسبي للصورة.
    images_by_building = {}
    _all_images = discover_image_files(FOLDER_PATH)
    print(f'🔎 ملفات الصور المكتشفة داخل القرية ومجلداتها: {len(_all_images)}')
    if NAME_MODE:
        _name_scores = {}
        for image_path in _all_images:
            relative_label = os.path.splitext(
                os.path.relpath(image_path, FOLDER_PATH)
            )[0]
            # نطابق اسم الصورة بأقرب اسم موقع في الإكسل
            best_key = None; best_sc = 0.0
            for key in data_by_building.keys():
                sc = name_similarity(relative_label, key)
                if sc > best_sc:
                    best_sc = sc; best_key = key
            if (best_key is not None and best_sc >= 0.6 and
                    best_sc > _name_scores.get(best_key, 0.0)):
                images_by_building[best_key] = image_path
                _name_scores[best_key] = best_sc
    else:
        (images_by_building, _unmatched_images,
         _conflicting_images, _ambiguous_images) = index_numeric_images(
            _all_images, FOLDER_PATH, set(data_by_building.keys())
        )
        if _ambiguous_images:
            preview = ', '.join(map(str, sorted(_ambiguous_images)[:15]))
            print(f'   ⚠️ صور متعددة متساوية لنفس الرقم؛ لن يتم اختيار صورة عشوائياً: {preview}')
        if _conflicting_images:
            print(f'   ⚠️ {len(_conflicting_images)} صورة تحمل أكثر من رقم محتمل وتم تجاهلها للأمان')
    print(f'🖼️  الصور المطابقة للمباني: {len(images_by_building)}/{len(data_by_building)}')
    _missing_local_images = sorted(
        key for key in data_by_building if key not in images_by_building
    )
    if _missing_local_images:
        preview = ', '.join(map(str, _missing_local_images[:25]))
        suffix = ' ...' if len(_missing_local_images) > 25 else ''
        print(f'⚠️ مبانٍ بلا صورة كروكي محلية: {preview}{suffix}')

    # ── البحث عن القرية مع إعادة المحاولة التلقائية ─────────
    buildings = None; village_search_error = ''
    for village_attempt in range(1, 4):
        try:
            expected_codes = None if NAME_MODE else set(data_by_building.keys())
            buildings = search_buildings_for_village(village, expected_codes)
            break
        except Exception as e:
            village_search_error = str(e)
            print(f'⚠️ تعذرت قراءة قرية "{village}" — محاولة {village_attempt}/3: '
                  f'{village_search_error[:90]}')
            if not is_connection_alive():
                reconnect_driver()
            time.sleep(3)
    if buildings is None:
        msg = f'تم تجاوز القرية بعد 3 محاولات لقراءة نتائجها: {village_search_error[:110]}'
        print(f'❌ {msg}')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'skip', 'التفاصيل': msg,
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        all_results_by_village[village] = followup_by_village[village]
        _issue = followup_by_village[village][0]
        record_live_result(village, _issue['رقم المبنى'], _issue['الحالة'],
                           _issue['التفاصيل'], _issue['الوقت'])
        live_workbook.save(master_report)
        grand_total['villages'] += 1; grand_total['skipped'] += 1
        continue
    print(f'🏢 المباني في المنصة: {len(buildings)}')

    if not buildings:
        print(f'⚠️ لم تُوجد مباني للقرية "{village}" — سيتم تخطيها')
        followup_by_village[village] = [{
            'رقم المبنى': '—', 'الحالة': 'skip',
            'التفاصيل': 'لم تُوجد أي مباني لهذه القرية في المنصة',
            'الوقت': datetime.now().strftime('%H:%M:%S')}]
        all_results_by_village[village] = followup_by_village[village]
        _issue = followup_by_village[village][0]
        record_live_result(village, _issue['رقم المبنى'], _issue['الحالة'],
                           _issue['التفاصيل'], _issue['الوقت'])
        live_workbook.save(master_report)
        grand_total['villages'] += 1; grand_total['skipped'] += 1
        continue

    results_log = []
    def log(bnum, status, msg=''):
        ts = datetime.now().strftime('%H:%M:%S')
        icon = {'success': '✅', 'skip': '⏭️', 'error': '❌'}[status]
        print(f'[{ts}] {icon} مبنى {bnum}: {msg}')
        results_log.append({'رقم المبنى': bnum, 'الحالة': status,
                            'التفاصيل': msg, 'الوقت': ts})
        all_results_by_village[village] = results_log
        record_live_result(village, bnum, status, msg, ts)

    processed = skipped = failed = 0

    # أي صف Excel غير صالح يظهر في التقرير حتى لو كانت بقية الصفوف صالحة.
    for invalid_item in invalid_excel_rows:
        invalid_key = invalid_item['key']
        if (FILTER_BUILDINGS and isinstance(invalid_key, int)
                and invalid_key not in FILTER_BUILDINGS):
            continue
        log(invalid_key, 'error', invalid_item['details'])
        failed += 1

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
        platform_by_number = {}
        for b in buildings:
            bnum = b['number']
            if bnum in platform_by_number and platform_by_number[bnum]['id'] != b['id']:
                print(f'   ⚠️ أكثر من سجل للمبنى {bnum} — اعتماد أول سجل مطابق فقط')
                continue
            platform_by_number[bnum] = b
        # Excel هو مصدر قائمة العمل؛ لا ندخل أي رقم ظهر في المنصة وليس في الملف.
        for excel_key in sorted(data_by_building):
            if FILTER_BUILDINGS and excel_key not in FILTER_BUILDINGS:
                continue
            platform_building = platform_by_number.get(excel_key)
            if platform_building is None:
                worklist.append((excel_key, None, 'غير موجود بعد قراءة جميع صفحات النتائج'))
            else:
                worklist.append((excel_key, platform_building,
                                 f'كود المنصة {platform_building.get("site_code", "—")}'))

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
            log(excel_key, 'error', 'لم يظهر الرقم بعد قراءة جميع صفحات نتائج القرية')
            failed += 1; continue
        if excel_key not in data_by_building:
            log(excel_key, 'skip', 'غير موجود في Excel')
            skipped += 1; continue

        # حماية ثانية قبل فتح النموذج: حتى لو تسرب صف من نتائج البحث لأي سبب،
        # لا يُسمح له بالوصول إلى دالة الكتابة والحفظ.
        if not _belongs_to_village(b.get('name', ''), village):
            wrong_name = str(b.get('name', '')).strip() or 'اسم غير معروف'
            print(f'   🛑 حماية الهوية: السجل يتبع قرية أخرى: "{wrong_name}"')
            log(excel_key, 'error',
                f'تم منع التعديل لأن اسم موقع المنصة لا يطابق القرية: {wrong_name}')
            failed += 1; continue

        bid = b['id']

        ok = False; last_error = ''
        has_local_image = excel_key in images_by_building
        max_building_attempts = 4 if (has_local_image or TEST_MODE) else 1
        for building_attempt in range(1, max_building_attempts + 1):
            try:
                ok = process_building(bid, excel_key, data_by_building, images_by_building)
                if ok:
                    break
                last_error = ('لا توجد صورة كروكي محلية مطابقة'
                              if not has_local_image
                              else 'لم يثبت الحفظ أو رفع الصورة')
            except Exception as exc:
                last_error = str(exc)
                if not is_connection_alive():
                    reconnect_driver()
            if building_attempt < max_building_attempts:
                print(f'   🔁 إعادة المبنى كاملاً '
                      f'({building_attempt + 1}/{max_building_attempts}) بعد 5 ثوانٍ...')
                time.sleep(5)

        if ok:
            success_msg = ('نجح فحص الصفحة والحقول بلا تعديل'
                           if TEST_MODE else 'تم تحديث X/Y/مساحة + صورة والتحقق بعد الحفظ')
            log(excel_key, 'success', success_msg)
            processed += 1
        else:
            if not has_local_image:
                log(excel_key, 'error',
                    'لا توجد صورة كروكي محلية مطابقة — تم التجاوز دون تكرار الصفحة')
            else:
                log(excel_key, 'error',
                    f'فشل بعد {max_building_attempts} محاولات: {last_error[:90]}')
            failed += 1

    followup_rows = [r for r in results_log if r['الحالة'] in ('error', 'skip')]
    followup_by_village[village] = followup_rows
    all_results_by_village[village] = results_log

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
    live_workbook.save(master_report)

# ── 8) التقرير الإجمالي: كل المواقع + ملخص القرى ───────────
from openpyxl.styles import Font, PatternFill, Alignment

live_workbook.save(master_report)
live_workbook.close()

with pd.ExcelWriter(master_report, engine='openpyxl') as writer:
    report_rows = []
    summary_rows = []
    for village, rows in all_results_by_village.items():
        success = sum(r['الحالة'] == 'success' for r in rows)
        skipped = sum(r['الحالة'] == 'skip' for r in rows)
        failed = sum(r['الحالة'] == 'error' for r in rows)
        summary_rows.append({
            'القرية': village, 'ناجح': success, 'متخطى': skipped,
            'فشل': failed, 'الإجمالي': len(rows)
        })
        for row in rows:
            report_rows.append({
                'القرية': village,
                'رقم المبنى/الكود': row['رقم المبنى'],
                'الحالة': {'success': '✅ ناجح', 'skip': '⏭️ متخطى',
                           'error': '❌ فشل'}.get(row['الحالة'], row['الحالة']),
                'التفاصيل': row['التفاصيل'], 'الوقت': row['الوقت']
            })

    pd.DataFrame(report_rows, columns=[
        'القرية', 'رقم المبنى/الكود', 'الحالة', 'التفاصيل', 'الوقت'
    ]).to_excel(writer, sheet_name='سجل المواقع', index=False)
    pd.DataFrame(summary_rows, columns=[
        'القرية', 'ناجح', 'متخطى', 'فشل', 'الإجمالي'
    ]).to_excel(writer, sheet_name='ملخص القرى', index=False)

    header_fill = PatternFill('solid', fgColor='4472C4')
    header_font = Font(bold=True, color='FFFFFF')
    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        ws.sheet_view.rightToLeft = True
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for column_cells in ws.columns:
            letter = column_cells[0].column_letter
            longest = max(len(str(cell.value or '')) for cell in column_cells)
            ws.column_dimensions[letter].width = min(max(longest + 2, 12), 55)

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

if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
    except Exception:
        pass
