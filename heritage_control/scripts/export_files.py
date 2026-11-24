#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""مشروع التراث العمراني — تصدير ملفات المواقع والصور.
حوّل من Notebook إلى برنامج طرفية قابل للاستئناف.
"""
import asyncio, os, re, time, json, argparse, getpass, tempfile, sys

# جسر مركز التحكم (اختياري): يبعث التقدّم ويحترم الإيقاف/الاستئناف.
# لو شُغّل السكربت وحده (بدون المركز) يتحوّل إلى صامت/ممرّر فلا يتأثر السلوك.
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
    import tkinter as tk
    from tkinter import simpledialog, messagebox
except Exception:
    tk = None
    simpledialog = None
    messagebox = None
try:
    import requests
    from bs4 import BeautifulSoup
    from tqdm import tqdm
    from playwright.async_api import async_playwright
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError as exc:
    missing = getattr(exc, "name", "إحدى المكتبات المطلوبة")
    raise SystemExit(
        f"المكتبة المطلوبة غير مثبتة: {missing}\n"
        "نفّذ الأمر التالي مرة واحدة ثم أعد تشغيل البرنامج:\n"
        "py -m pip install requests beautifulsoup4 tqdm playwright openpyxl\n"
        "py -m playwright install chromium"
    ) from exc


# ===== CELL 1 =====
# ── ثوابت ──
BASE_URL           = "https://heritage2.itqan-consultant.com"
LOGIN_URL          = f"{BASE_URL}/Web/App/Home/Login"
DATAVIEW_URL       = f"{BASE_URL}/Web/App/Pools/DataView"
DELAY_REQUESTS     = 1.5
DELAY_PDF          = 2.0
MAX_RETRIES        = 5
RETRY_WAIT         = 15
INTERNET_CHECK_URL = LOGIN_URL
ROWS_PER_PAGE      = 25
SESSION_COOKIE     = ""
AUTH_COOKIES       = {}
USERNAME = ""
PASSWORD = ""
OUTPUT_DIR = ""
PROGRESS_FILE = ""
DECISION = ""


# ===== CELL 2 =====
# ============================================================
#  الخلية 3 - الدوال المساعدة + تسجيل الدخول
# ============================================================

def sanitize(name: str, max_len: int = 100) -> str:
    name = str(name or "").strip()
    name = re.sub(r'[\\/*?:"<>|]', '-', name)
    name = re.sub(r'\s+', ' ', name)
    name = name.rstrip('. ')
    name = name[:max_len].rstrip('. ')
    if name.upper() in {'CON', 'PRN', 'AUX', 'NUL',
                        *(f'COM{i}' for i in range(1, 10)),
                        *(f'LPT{i}' for i in range(1, 10))}:
        name = '_' + name
    return name or "بدون اسم"

def artifact_stem(building: dict) -> str:
    """اسم فريد وآمن يمنع دمج مبنيين متشابهين في الاسم."""
    return sanitize(f"{building.get('name', 'بدون اسم')} [{building.get('id', 'بدون رقم')}]", 140)

def make_dirs(base: str, decision: str) -> dict:
    root = os.path.join(base, f"قرار ترشيح رقم {decision}")
    dirs = {
        'root':         root,
        'nominations':  os.path.join(root, "قرارات الترشيح"),
        'registration': os.path.join(root, "استمارة التسجيل"),
        'summary':      os.path.join(root, "ملخص الاستمارة"),
        'photos':       os.path.join(root, "الصور"),
    }
    for d in dirs.values():
        os.makedirs(d, exist_ok=True)
    print(f"📁 تم إنشاء المجلدات:\n   {root}")
    return dirs

def load_progress(path: str) -> dict:
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise RuntimeError("ملف التقدم تالف: المحتوى ليس كائناً JSON")
        data.setdefault('completed', [])
        data.setdefault('failed', [])
        data.setdefault('buildings', [])
        data['completed'] = [str(x) for x in data['completed']]
        data['failed'] = [
            {**item, 'id': str(item.get('id', ''))}
            for item in data['failed'] if isinstance(item, dict)
        ]
        done = len(data['completed'])
        fail = len(data['failed'])
        print(f"📂 استئناف: ✅ {done} مكتمل | ❌ {fail} فاشل")
        return data
    print("📂 بدء جديد")
    return {'completed': [], 'failed': [], 'buildings': []}

def save_progress(path: str, data: dict):
    """حفظ ذري: لا يترك ملف JSON ناقصاً إذا توقف البرنامج أثناء الكتابة."""
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix="progress_", suffix=".tmp", dir=parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

def wait_for_internet(max_wait: int = 600):
    elapsed = 0
    while elapsed < max_wait:
        try:
            requests.get(INTERNET_CHECK_URL, timeout=5)
            print("✅ الإنترنت عاد")
            return True
        except requests.RequestException:
            print(f"⏳ انقطع الإنترنت - انتظار... ({elapsed}s)")
            time.sleep(30)
            elapsed += 30
    return False

def response_is_login(resp) -> bool:
    """كشف انتهاء الجلسة سواءً حصل redirect أو عادت صفحة الدخول بحالة 200."""
    url = (getattr(resp, 'url', '') or '').lower()
    text = (getattr(resp, 'text', '') or '')[:10000].lower()
    return ('/login' in url or
            ('__viewstate' in text and ('btnlogin' in text or 'تسجيل الدخول' in text)))

def response_is_server_error(resp) -> bool:
    text = (getattr(resp, 'text', '') or '')[:20000].lower()
    return (getattr(resp, 'status_code', 200) >= 500 or
            'server error' in text or 'runtime error' in text)

def hidden_form_fields(soup: BeautifulSoup) -> dict:
    """التقاط كل حقول ASP.NET المخفية بدلاً من افتراض ثلاثة حقول فقط."""
    return {
        element.get('name'): element.get('value', '')
        for element in soup.find_all('input', {'type': 'hidden'})
        if element.get('name')
    }

# ══════════════════════════════════════════════════════════════
#  تسجيل الدخول - يعمل تلقائياً ويجدد الجلسة عند انتهائها
# ══════════════════════════════════════════════════════════════

def do_login(username: str, password: str) -> tuple:
    """
    يسجل الدخول ويُرجع (cookie_value, session)
    يحتفظ بالـ session كاملة للاستخدام لاحقاً
    """
    s = requests.Session()
    s.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

    # الخطوة 1: GET صفحة Login للحصول على ViewState
    for attempt in range(MAX_RETRIES):
        try:
            get_resp = s.get(LOGIN_URL, timeout=30)
            get_resp.raise_for_status()
            get_resp.encoding = 'utf-8'
            break
        except Exception as e:
            print(f"⚠️ خطأ GET Login (محاولة {attempt+1}): {e}")
            wait_for_internet()
    else:
        raise RuntimeError("❌ فشل الاتصال بصفحة تسجيل الدخول")

    get_soup = BeautifulSoup(get_resp.text, 'html.parser')
    vs  = get_soup.find('input', {'name': '__VIEWSTATE'})
    ev  = get_soup.find('input', {'name': '__EVENTVALIDATION'})
    vsg = get_soup.find('input', {'name': '__VIEWSTATEGENERATOR'})

    # الخطوة 2: POST بيانات الدخول. نكتشف أسماء الحقول إن أمكن مع fallback
    # للأسماء الحالية في المنصة.
    user_input = get_soup.find('input', {'type': re.compile(r'^(text|email)$', re.I)})
    pass_input = get_soup.find('input', {'type': re.compile(r'^password$', re.I)})
    submit_input = get_soup.find('input', {'type': re.compile(r'^submit$', re.I)})
    user_key = user_input.get('name') if user_input and user_input.get('name') else 'ctl12$ctl03'
    pass_key = pass_input.get('name') if pass_input and pass_input.get('name') else 'ctl12$ctl07'
    submit_key = submit_input.get('name') if submit_input and submit_input.get('name') else 'ctl12$btnLogin'
    submit_value = submit_input.get('value', 'تسجيل الدخول') if submit_input else 'تسجيل الدخول'

    login_data = hidden_form_fields(get_soup)
    login_data.update({
        '__EVENTTARGET':        '',
        '__EVENTARGUMENT':      '',
        '__LASTFOCUS':          '',
        '__VIEWSTATE':          vs['value']  if vs  else '',
        '__VIEWSTATEGENERATOR': vsg['value'] if vsg else '',
        '__EVENTVALIDATION':    ev['value']  if ev  else '',
        user_key:               username,
        pass_key:               password,
        submit_key:             submit_value,
    })

    for attempt in range(MAX_RETRIES):
        try:
            post_resp = s.post(LOGIN_URL, data=login_data,
                               timeout=30, allow_redirects=True)
            post_resp.raise_for_status()
            break
        except Exception as e:
            print(f"⚠️ خطأ POST Login (محاولة {attempt+1}): {e}")
            wait_for_internet()
    else:
        raise RuntimeError("❌ فشل إرسال بيانات تسجيل الدخول")

    # التحقق من نجاح الدخول
    if response_is_login(post_resp):
        raise RuntimeError("❌ فشل تسجيل الدخول - تحقق من اسم المستخدم وكلمة المرور")

    # استخراج الكوكي
    cookie_val = s.cookies.get('Heritage_Survey_Access_Token')
    if not cookie_val:
        for c in s.cookies:
            if 'token' in c.name.lower() or 'access' in c.name.lower():
                cookie_val = c.value
                break

    if not cookie_val:
        raise RuntimeError("❌ تسجيل الدخول نجح لكن لم يُعثر على الكوكي")

    return cookie_val, s


def make_session() -> requests.Session:
    """ينشئ session جديدة بالكوكي الحالي"""
    s = requests.Session()
    if AUTH_COOKIES:
        s.cookies.update(AUTH_COOKIES)
    else:
        s.cookies.set('Heritage_Survey_Access_Token', SESSION_COOKIE,
                      domain='heritage2.itqan-consultant.com')
    s.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    return s


def refresh_session():
    """يجدد الجلسة تلقائياً عند انتهائها"""
    global SESSION_COOKIE, AUTH_COOKIES
    print("🔄 الجلسة انتهت - جاري تجديدها تلقائياً...")
    SESSION_COOKIE, login_session = do_login(USERNAME, PASSWORD)
    AUTH_COOKIES = requests.utils.dict_from_cookiejar(login_session.cookies)
    print("✅ تم تجديد الجلسة")



# ===== CELL 3 =====
# ============================================================
#  الخلية 4 - جمع المباني (GET أولاً ثم POST)
# ============================================================
from urllib.parse import urljoin

def get_all_buildings(decision: str, progress: dict, start_from_page: int = 1) -> list:

    if progress.get('buildings_complete'):
        print(f"📋 القائمة مكتملة مسبقاً: {len(progress['buildings']):,} مبنى")
        return progress['buildings']

    session        = make_session()
    # إزالة أي تكرار سابق مع الحفاظ على ترتيب أول ظهور.
    known = {}
    for item in progress.get('buildings', []):
        if isinstance(item, dict) and item.get('id'):
            known[str(item['id'])] = item
    all_buildings  = list(known.values())
    last_page_done = progress.get('last_page_collected', 0)
    actual_start   = max(start_from_page, last_page_done + 1)

    if start_from_page > 1 and last_page_done < start_from_page - 1:
        raise ValueError(
            "لا يمكن البدء من صفحة متقدمة من دون وجود الصفحات السابقة في ملف التقدم؛ "
            "استخدم --start-page 1 لضمان عدم فقد أي موقع."
        )

    if last_page_done > 0 and start_from_page == 1:
        print(f"📂 استئناف من الصفحة {actual_start} ({len(all_buildings):,} مبنى محفوظ)")
    elif start_from_page > 1:
        print(f"▶️  بدء يدوي من الصفحة {start_from_page}")
        actual_start = start_from_page

    print(f"🔍 قرار التسجيل: {decision} | البدء من صفحة: {actual_start}")

    # ── GET أولاً للحصول على ViewState ──
    print("⏳ جاري جلب ViewState...")
    get_resp = None
    for attempt in range(MAX_RETRIES):
        try:
            get_resp = session.get(DATAVIEW_URL, timeout=30)
            get_resp.encoding = 'utf-8'
            if response_is_login(get_resp):
                print("🔄 الجلسة منتهية - جاري التجديد...")
                refresh_session()
                session = make_session()
                continue
            get_resp.raise_for_status()
            if response_is_server_error(get_resp):
                raise RuntimeError("صفحة الخادم أعادت خطأ")
            break
        except Exception as e:
            print(f"⚠️ خطأ (محاولة {attempt+1}/{MAX_RETRIES}): {e}")
            get_resp = None
            wait_for_internet()

    if get_resp is None or response_is_login(get_resp):
        raise RuntimeError("❌ فشل الاتصال بالخادم")

    def get_vs(s):
        fields = hidden_form_fields(s)
        for required in ('__VIEWSTATE', '__EVENTVALIDATION'):
            fields.setdefault(required, '')
        return fields

    get_soup = BeautifulSoup(get_resp.text, 'html.parser')
    init_vs  = get_vs(get_soup)
    print("✅ ViewState جاهز")

    # ── POST مع البحث ──
    first_post = {
        **init_vs,
        '__EVENTTARGET':        '',
        '__EVENTARGUMENT':      '',
        '__LASTFOCUS':          '',
        'ctl12$ctl04': '', 'ctl12$ctl08': '', 'ctl12$ctl12': '',
        'ctl12$ctl16': '', 'ctl12$ctl20': '', 'ctl12$ctl24': '',
        'ctl12$ctl28': decision,
        'ctl12$ctl32': '',
        'ctl12$btnSearch': 'بحــــث',
    }

    resp = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = session.post(DATAVIEW_URL, data=first_post, timeout=30)
            resp.encoding = 'utf-8'
            if response_is_login(resp):
                refresh_session()
                session = make_session()
                resp = None
                continue
            resp.raise_for_status()
            if response_is_server_error(resp):
                raise RuntimeError("صفحة البحث أعادت خطأ من الخادم")
            break
        except Exception as e:
            print(f"⚠️ خطأ POST (محاولة {attempt+1}/{MAX_RETRIES}): {e}")
            resp = None
            wait_for_internet()

    if resp is None:
        raise RuntimeError("❌ فشل طلب البحث")

    soup = BeautifulSoup(resp.text, 'html.parser')

    result_text = soup.get_text(" ", strip=True)
    total_str = re.search(r'نتيجة\s*البحث[^\d]*([\d,٬]+)\s*موقع', result_text)
    total     = int(total_str.group(1).replace(',', '').replace('٬', '')) if total_str else 0
    if total <= 0:
        raise RuntimeError(
            "لم يمكن قراءة عدد نتائج القرار، أو أن القرار لا يحتوي مواقع. "
            "لم يبدأ التصدير حتى لا تُسجّل نتيجة ناقصة كأنها مكتملة."
        )
    def extract_page(s: BeautifulSoup) -> list:
        buildings = []
        grid = s.find('table', {'id': 'ctl12_GridView1'})
        if not grid:
            return buildings
        for row in grid.find_all('tr')[1:]:
            cells = row.find_all('td')
            if len(cells) < 12:
                continue
            bid = cells[0].get_text(strip=True)
            if not bid.isdigit():
                continue
            name = cells[1].get_text(strip=True)
            def href(cell):
                a = cell.find('a')
                return a['href'] if a else None
            def full(h):
                if not h: return None
                h = h.strip()
                if h.startswith('http'): return h
                if h.lower().startswith('javascript'): return None   # تجاهل postback
                return urljoin(DATAVIEW_URL, h)   # ✅ يحلّ المسار النسبي نسبةً لصفحة Pools
            buildings.append({
                'id':   bid,
                'name': name,
                'nomination_url':   full(href(cells[9])),
                'summary_url':      full(href(cells[10])),
                'registration_url': full(href(cells[11])),
                'photos_url': f"{BASE_URL}/Web/App/Pools/DataEdit/{bid}/1",
            })
        return buildings

    first_page_data = extract_page(soup)
    if not first_page_data:
        raise RuntimeError("تعذر قراءة مباني الصفحة الأولى رغم وجود نتائج")
    detected_page_size = len(first_page_data) if total > len(first_page_data) else total
    page_size = detected_page_size or ROWS_PER_PAGE
    pages = max(1, -(-total // page_size))
    print(
        f"📊 إجمالي المواقع: {total:,} | حجم الصفحة: {page_size} | "
        f"الصفحات: {pages} | من: {actual_start} إلى {pages}"
    )

    # عدد الصفوف المقروءة فعلياً (قد يفوق عدد المواقع الفريدة لأن الموقع الواحد
    # قد يتكرّر في القائمة عند إسناده لأكثر من شخص في سير العمل).
    rows_seen = progress.get('rows_seen', 0)

    if actual_start == 1:
        page_data = first_page_data
        known.update({b['id']: b for b in page_data})
        all_buildings = list(known.values())
        rows_seen = len(page_data)
        progress['last_page_collected'] = 1
        progress['buildings']           = all_buildings
        progress['rows_seen']           = rows_seen
        save_progress(PROGRESS_FILE, progress)
        print(f"   ✅ صفحة   1/{pages}: {len(page_data):>3} صف | مواقع فريدة: {len(all_buildings):,}")
    # نعيد المرور تسلسلياً من الصفحة 2 حتى عند الاستئناف لأن GridView يعتمد
    # ViewState الصفحة السابقة. الصفحات المحفوظة تُمرّر فقط ولا تُضاف ثانية.
    range_start = 2

    for pg in range(range_start, pages + 1):
        time.sleep(DELAY_REQUESTS)

        vs_data   = get_vs(soup)
        post_data = {
            **vs_data,
            '__EVENTTARGET':        'ctl12$GridView1',
            '__EVENTARGUMENT':      f'Page${pg}',
            '__LASTFOCUS':          '',
            'ctl12$ctl04': '', 'ctl12$ctl08': '', 'ctl12$ctl12': '',
            'ctl12$ctl16': '', 'ctl12$ctl20': '', 'ctl12$ctl24': '',
            'ctl12$ctl28': decision,
            'ctl12$ctl32': '',
        }

        resp = None
        for attempt in range(MAX_RETRIES):
            try:
                resp = session.post(DATAVIEW_URL, data=post_data, timeout=30)
                resp.encoding = 'utf-8'
                if response_is_login(resp):
                    refresh_session()
                    session = make_session()
                    resp = None
                    continue
                resp.raise_for_status()
                if response_is_server_error(resp):
                    raise RuntimeError("الخادم أعاد صفحة خطأ")
                break
            except Exception as e:
                print(f"\n⚠️ خطأ صفحة {pg} (محاولة {attempt+1}/{MAX_RETRIES}): {e}")
                resp = None
                wait_for_internet()

        if resp is None:
            raise RuntimeError(
                f"❌ فشل تحميل الصفحة {pg} بعد {MAX_RETRIES} محاولات. "
                "تم حفظ التقدم؛ شغّل البرنامج مجدداً للاستكمال دون فقد صفحات."
            )

        soup      = BeautifulSoup(resp.text, 'html.parser')
        if pg < actual_start:
            continue
        page_data = extract_page(soup)
        if not page_data:
            raise RuntimeError(
                f"الصفحة {pg} لم تُرجع مباني رغم أنها ضمن صفحات النتائج؛ "
                "أوقف البرنامج لمنع تصدير ناقص."
            )
        known.update({b['id']: b for b in page_data})
        all_buildings = list(known.values())
        rows_seen += len(page_data)

        progress['last_page_collected'] = pg
        progress['buildings']           = all_buildings
        progress['rows_seen']           = rows_seen
        save_progress(PROGRESS_FILE, progress)

        if pg % 25 == 0 or pg == pages:
            print(f"   ✅ صفحة {pg:>4}/{pages}: مواقع فريدة {len(all_buildings):,} (صفوف مقروءة {rows_seen:,})")

    # اكتمال الجمع مضمون بوصولنا لآخر صفحة (كل صفحة فاشلة/فارغة تُوقف التشغيل أعلاه).
    # لذلك لا نُفشل عند اختلاف العدد: الموقع قد يتكرّر مشروعاً في القائمة (مُسنَد
    # لأكثر من شخص)، فعدد المواقع الفريدة أقل من عدد صفوف المنصة — وهذا طبيعي.
    if len(all_buildings) != total:
        print(f"ℹ️ المنصة تعرض {total:,} صف (منها مكرّر — الموقع عند أكثر من شخص). "
              f"صفوف مقروءة: {rows_seen:,} → سيُصدَّر {len(all_buildings):,} موقع فريد.")
    if rows_seen < total:
        print(f"⚠️ تنبيه: الصفوف المقروءة ({rows_seen:,}) أقل من عدد المنصة ({total:,}). "
              "جرى المرور على كل الصفحات؛ إن نقص موقع أعد التشغيل للاستكمال الآمن.")

    progress['buildings_complete'] = True
    progress['buildings']          = all_buildings
    save_progress(PROGRESS_FILE, progress)
    print(f"\n✅ اكتمل جمع القائمة: {len(all_buildings):,} مبنى")
    return all_buildings

# ===== CELL 4 =====
# ============================================================
#  تحميل الصور5
# ============================================================
def download_photos(building: dict, photos_root: str, session: requests.Session) -> dict:
    """تحميل صور مبنى واحد - يرجع إحصائيات"""
    bname  = artifact_stem(building)
    folder = os.path.join(photos_root, bname)
    os.makedirs(folder, exist_ok=True)

    result = {'downloaded': 0, 'skipped': 0, 'failed': 0, 'found': 0}

    for attempt in range(MAX_RETRIES):
        try:
            resp = session.get(building['photos_url'], timeout=20)
            resp.encoding = 'utf-8'
            if resp.status_code in (401, 403) or response_is_login(resp):
                raise RuntimeError("SESSION_EXPIRED")
            resp.raise_for_status()
            if response_is_server_error(resp):
                raise RuntimeError("SERVER_ERROR")
            break
        except RuntimeError:
            raise
        except requests.RequestException as exc:
            if attempt == MAX_RETRIES - 1:
                result['failed'] = -1
                return result
            print(f"\n⚠️ تعذر فتح صفحة الصور ({attempt+1}/{MAX_RETRIES}): {exc}")
            time.sleep(min(RETRY_WAIT, 5))
    else:
        result['failed'] = -1
        return result

    soup = BeautifulSoup(resp.text, 'html.parser')

    # استخراج روابط الصور - من الـ anchor tags أو img src
    photo_urls = []
    # الطريقة الأولى: anchor links مباشرة للصورة الكاملة
    for a in soup.find_all('a', href=re.compile(r'/[Ss]3/|\.(?:jpe?g|png|webp|gif)', re.I)):
        url = a['href']
        if not url.startswith('http'):
            url = urljoin(resp.url, url)
        if url not in photo_urls:
            photo_urls.append(url)
    # الطريقة الثانية (احتياطية): img src
    if not photo_urls:
        for img in soup.find_all('img'):
            src = img.get('src', '')
            if '/s3/' in src.lower() or '/heritage/' in src.lower():
                if not src.startswith('http'):
                    src = urljoin(resp.url, src)
                if src not in photo_urls:
                    photo_urls.append(src)

    # احتياط في حال كانت الروابط مخزنة داخل JavaScript أو data attributes.
    if not photo_urls:
        raw_urls = re.findall(
            r'''["']([^"']*(?:/[sS]3/|/heritage/)[^"']+)["']''',
            resp.text,
            flags=re.I,
        )
        for raw_url in raw_urls:
            url = urljoin(resp.url, raw_url)
            if url not in photo_urls:
                photo_urls.append(url)

    if not photo_urls:
        page_text = re.sub(r'\s+', ' ', soup.get_text(' ', strip=True))
        no_photos_markers = ('لا توجد صور', 'لا يوجد صور', 'لا توجد صورة', 'لا يوجد صورة')
        if not any(marker in page_text for marker in no_photos_markers):
            # لا نعد الصفحة نجاحاً بصمت؛ قد يكون تصميم المنصة قد تغير.
            result['failed'] = -2
            return result

    # استبعاد الروابط الزخرفية الواضحة (ليست صوراً للمبنى): أيقونة/شعار/مصغّرة/خريطة…
    # هذه عناصر قالبية ثابتة في كل صفحة وكثيراً ما تعيد 403 أو محتوى غير صورة،
    # فكانت تُحسب خطأً «صورة فاشلة» وتُفشِل المبنى كاملاً رغم نزول كل صوره الحقيقية.
    _JUNK = ('thumb', 'thumbnail', 'icon', 'logo', 'placeholder', 'no-image',
             'noimage', 'no_image', 'avatar', 'spacer', 'blank', 'loading',
             'default-', 'map-marker', 'marker.', 'sprite', 'favicon')
    filtered = [u for u in photo_urls if not any(j in u.lower() for j in _JUNK)]
    if filtered:  # لا نفرّغ القائمة لو كانت كلها «مشبوهة» الأسماء
        photo_urls = filtered

    result['found'] = len(photo_urls)

    for idx, url in enumerate(photo_urls, start=1):
        existing = next(
            (os.path.join(folder, f) for f in os.listdir(folder)
             if re.fullmatch(rf"{idx}\.(jpg|jpeg|png|webp|gif)", f, re.I)
             and os.path.getsize(os.path.join(folder, f)) > 1000),
            None,
        )
        if existing:
            result['skipped'] += 1
            continue
        for attempt in range(MAX_RETRIES):
            try:
                r = session.get(url, timeout=30, headers={'Referer': resp.url})
                if r.status_code == 200:
                    content_type = (r.headers.get('Content-Type') or '').lower()
                    signatures = (
                        r.content.startswith(b'\xff\xd8\xff'),
                        r.content.startswith(b'\x89PNG\r\n\x1a\n'),
                        r.content.startswith((b'GIF87a', b'GIF89a')),
                        r.content.startswith(b'RIFF') and r.content[8:12] == b'WEBP',
                    )
                    if not r.content or 'text/html' in content_type or not any(signatures):
                        if attempt == MAX_RETRIES - 1:
                            result['failed'] += 1
                        else:
                            time.sleep(2)
                        continue
                    ext = ('png' if signatures[1] else 'gif' if signatures[2]
                           else 'webp' if signatures[3] else 'jpg')
                    save_path = os.path.join(folder, f"{idx}.{ext}")
                    part_path = save_path + '.part'
                    with open(part_path, 'wb') as f:
                        f.write(r.content)
                    os.replace(part_path, save_path)
                    result['downloaded'] += 1
                    break
                elif r.status_code in (401, 403) and response_is_login(r):
                    raise RuntimeError("SESSION_EXPIRED")
                elif r.status_code == 403:
                    # بعض روابط S3 المحمية تعيد 403 رغم أن جلسة المنصة ما زالت سليمة.
                    # لا نعتبر ذلك انتهاءً للجلسة؛ نسجل الصورة كفشل ونكمل باقي الموقع.
                    result['failed'] += 1
                    break
                elif r.status_code in (404, 410):
                    result['failed'] += 1
                    break
                else:
                    if attempt == MAX_RETRIES - 1:
                        result['failed'] += 1
                    else:
                        time.sleep(min(RETRY_WAIT, 5))
            except RuntimeError:
                raise
            except requests.RequestException:
                if attempt == MAX_RETRIES - 1 or not wait_for_internet():
                    result['failed'] += 1
                    break
        time.sleep(0.3)

    return result


# ===== CELL 5 =====
# ============================================================
#  الخلية 5.5 - إنشاء/تحديث ملف Excel
# ============================================================

def update_excel(dirs: dict, decision: str, completed_buildings: list):
    excel_path = os.path.join(dirs['root'], f"دليل_المباني_{decision}.xlsx")

    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"قرار {decision}"
        ws.sheet_view.rightToLeft = True

        headers      = ['#', 'رقم الموقع', 'اسم المبنى', 'مجلد الصور', 'عدد الصور', 'الحالة']
        header_fill  = PatternFill("solid", fgColor="2E4057")
        header_font  = Font(name='Arial', bold=True, color="FFFFFF", size=11)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for col, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = header_fill; cell.font = header_font; cell.alignment = center_align

        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 45
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.row_dimensions[1].height = 25
        ws.freeze_panes = 'A2'

    existing_ids = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 1 and row[1]:
            existing_ids.add(str(row[1]))

    row_fill_even = PatternFill("solid", fgColor="F0F4F8")
    row_fill_odd  = PatternFill("solid", fgColor="FFFFFF")
    link_font     = Font(name='Arial', color="1155CC", underline='single', size=10)
    normal_font   = Font(name='Arial', size=10)
    center_align  = Alignment(horizontal='center', vertical='center')
    right_align   = Alignment(horizontal='right',  vertical='center')

    added = 0
    for idx, building in enumerate(completed_buildings, start=1):
        bid = str(building['id'])
        bname = sanitize(building['name'])
        if bid in existing_ids:
            continue

        photo_folder = os.path.join(dirs['photos'], artifact_stem(building))
        photo_count  = 0
        if os.path.exists(photo_folder):
            photo_count = len([f for f in os.listdir(photo_folder)
                               if f.lower().endswith(('.jpg','.jpeg','.png'))])

        new_row = ws.max_row + 1
        fill    = row_fill_even if new_row % 2 == 0 else row_fill_odd

        c = ws.cell(row=new_row, column=1, value=new_row - 1)
        c.font = normal_font; c.alignment = center_align; c.fill = fill

        c = ws.cell(row=new_row, column=2, value=bid)
        c.font = normal_font; c.alignment = center_align; c.fill = fill

        c = ws.cell(row=new_row, column=3, value=bname)
        c.font = normal_font; c.alignment = right_align; c.fill = fill

        c = ws.cell(row=new_row, column=4)
        if os.path.exists(photo_folder) and photo_count > 0:
            folder_url  = "file:///" + photo_folder.replace("\\", "/").replace(" ", "%20")
            c.value     = "📁 فتح الصور"
            c.hyperlink = folder_url
            c.font      = link_font
        else:
            c.value = "لا توجد صور"; c.font = normal_font
        c.alignment = center_align; c.fill = fill

        c = ws.cell(row=new_row, column=5, value=photo_count)
        c.font = normal_font; c.alignment = center_align; c.fill = fill

        c = ws.cell(row=new_row, column=6, value="✅ مكتمل")
        c.font = Font(name='Arial', color="1E8449", size=10)
        c.alignment = center_align; c.fill = fill

        existing_ids.add(bid); added += 1

    temp_excel = excel_path + '.tmp.xlsx'
    try:
        wb.save(temp_excel)
        os.replace(temp_excel, excel_path)
    finally:
        if os.path.exists(temp_excel):
            os.remove(temp_excel)

    return excel_path, added

# ===== CELL 6 =====
# ============================================================
#  الخلية 6 - الدالة الرئيسية للأتمتة (نسخة محسّنة)
# ============================================================

def valid_pdf(path: str) -> bool:
    try:
        if os.path.getsize(path) <= 5000:
            return False
        with open(path, 'rb') as f:
            return f.read(5) == b'%PDF-'
    except OSError:
        return False

def playwright_cookies() -> list:
    cookies = AUTH_COOKIES or {'Heritage_Survey_Access_Token': SESSION_COOKIE}
    return [{
        'name': str(name), 'value': str(value),
        'domain': 'heritage2.itqan-consultant.com',
        'path': '/', 'secure': True,
    } for name, value in cookies.items()]

def local_artifacts_complete(building: dict, dirs: dict) -> bool:
    """لا نثق بعلامة completed وحدها؛ نتحقق من الملفات الأساسية على القرص."""
    stem = artifact_stem(building)
    expected_pdfs = (
        ('nomination_url', 'nominations'),
        ('registration_url', 'registration'),
        ('summary_url', 'summary'),
    )
    for url_key, dir_key in expected_pdfs:
        if not building.get(url_key):
            return False
        if not valid_pdf(os.path.join(dirs[dir_key], f"{stem}.pdf")):
            return False
    return os.path.isdir(os.path.join(dirs['photos'], stem))

async def save_pdf(page, url: str, path: str) -> bool:
    """تحميل صفحة كـ PDF مع التحقق من صحة المحتوى"""
    if valid_pdf(path):
        return True

    temp_path = path + '.part'
    if os.path.exists(temp_path):
        os.remove(temp_path)

    # نفس عدد المحاولات (لا نُطيل)، لكن بانتظار أذكى للـ500 ليلتقط النقص المؤقت
    pdf_retries = MAX_RETRIES
    for attempt in range(pdf_retries):
        try:
            response = await page.goto(url, wait_until='domcontentloaded', timeout=60000)
            await page.wait_for_timeout(int(DELAY_PDF * 1000))

            if response is None or response.status >= 500:
                raise RuntimeError(f"SERVER_ERROR_{getattr(response, 'status', 'NO_RESPONSE')}")
            if response.status in (401, 403):
                raise RuntimeError("SESSION_EXPIRED")

            # تحقق من انتهاء الجلسة
            body_text = await page.evaluate("document.body ? document.body.innerText : ''")
            lower_text = body_text.lower()
            if ('/login' in page.url.lower() or
                    ('تسجيل الدخول' in body_text and '__viewstate' in (await page.content()).lower())):
                print(f"\n⚠️ redirect للـ Login (محاولة {attempt+1}) - جاري تجديد الجلسة...")
                raise Exception("SESSION_EXPIRED")

            # تحقق من أن الصفحة تحتوي على محتوى فعلي
            if ('server error' in lower_text or 'runtime error' in lower_text):
                raise RuntimeError("SERVER_ERROR_PAGE")
            if len(body_text.strip()) < 80:
                print(f"\n⚠️ صفحة فارغة أو خطأ (محاولة {attempt+1})")
                await asyncio.sleep(5)
                continue

            # تطبيق @media print لإخفاء الـ navigation والـ sidebar
            await page.emulate_media(media='print')

            await page.pdf(
                path=temp_path, format='A4', print_background=True,
                margin={'top':'1.5cm','bottom':'1.5cm','left':'1cm','right':'1cm'}
            )

            # إرجاع الوضع لـ screen بعد الحفظ
            await page.emulate_media(media='screen')

            if valid_pdf(temp_path):
                os.replace(temp_path, path)
                return True
            else:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                print(f"\n⚠️ الملف المحفوظ صغير جداً (محاولة {attempt+1})")
                await asyncio.sleep(5)

        except Exception as e:
            err = str(e)
            if 'SESSION_EXPIRED' in err:
                refresh_session()
                await page.context.add_cookies(playwright_cookies())
                await asyncio.sleep(3)
            elif 'net::ERR' in err or 'ERR_INTERNET' in err:
                print(f"\n⚠️ انقطاع إنترنت ({attempt+1}/{MAX_RETRIES})")
                wait_for_internet()
            elif 'Timeout' in err:
                print(f"\n⚠️ Timeout ({attempt+1}/{pdf_retries})")
                await asyncio.sleep(RETRY_WAIT)
            elif 'SERVER_ERROR' in err:
                # خطأ خادم (500) غالباً مؤقت تحت الضغط — ننتظر بتصاعد ونعطي السيرفر وقتاً
                wait = min(10 + attempt * 10, 45)   # 10, 20, 30, 40, 45…
                print(f"\n⚠️ خطأ خادم المنصة 500 ({attempt+1}/{pdf_retries}) — انتظار {wait}ث ثم إعادة المحاولة")
                await asyncio.sleep(wait)
            else:
                print(f"\n⚠️ خطأ ({attempt+1}): {err[:80]}")
                await asyncio.sleep(5)
        finally:
            try:
                await page.emulate_media(media='screen')
            except Exception:
                pass
    if os.path.exists(temp_path):
        os.remove(temp_path)
    return False


async def run_automation(buildings: list, dirs: dict, progress: dict):
    done_ids = {str(x) for x in progress.get('completed', [])}
    verified_done = {
        str(b['id']) for b in buildings
        if str(b['id']) in done_ids and local_artifacts_complete(b, dirs)
    }
    if verified_done != done_ids.intersection({str(b['id']) for b in buildings}):
        selected_ids = {str(b['id']) for b in buildings}
        progress['completed'] = [
            str(x) for x in progress.get('completed', [])
            if str(x) not in selected_ids or str(x) in verified_done
        ]
        save_progress(PROGRESS_FILE, progress)
        done_ids = {str(x) for x in progress['completed']}
    pending  = [b for b in buildings if str(b['id']) not in done_ids]

    # تهيئة الداش بورد في مركز التحكم: الكل = كل المباني المحددة، المكتمل = ما تحقق مسبقاً.
    _R.total(len(buildings))
    _R.progress(done=len(done_ids), failed=len(progress.get('failed', [])), current="")

    print(f"\n{'='*55}")
    print("  🚀 بدء الأتمتة")
    print(f"  📦 إجمالي: {len(buildings):,} | ✅ مكتمل: {len(done_ids):,} | ⏳ متبقي: {len(pending):,}")
    print(f"{'='*55}\n")

    if not pending:
        completed_list = [b for b in buildings if str(b['id']) in done_ids]
        try:
            excel_path, _ = update_excel(dirs, DECISION, completed_list)
            print(f"🎉 كل المباني مكتملة! تم التحقق من الملفات وتحديث Excel: {excel_path}")
        except PermissionError:
            print("🎉 كل ملفات المباني مكتملة، لكن Excel مفتوح. أغلقه ثم أعد التشغيل لتحديثه.")
        return

    photo_session = make_session()

    async with async_playwright() as p:
        browser = None
        launch_errors = []
        # جرّب متصفح Playwright أولاً، ثم Edge/Chrome المثبتين عادةً على Windows.
        for launch_options in (
            {'headless': True},
            {'headless': True, 'channel': 'msedge'},
            {'headless': True, 'channel': 'chrome'},
        ):
            try:
                browser = await p.chromium.launch(**launch_options)
                break
            except Exception as exc:
                launch_errors.append(str(exc))
        if browser is None:
            raise RuntimeError(
                "تعذر تشغيل Chromium أو Edge أو Chrome. نفّذ مرة واحدة: "
                "py -m playwright install chromium\n"
                f"آخر خطأ: {launch_errors[-1][:200]}"
            )
        context = await browser.new_context(
            viewport={'width': 1280, 'height': 900},
            extra_http_headers={'Accept-Language': 'ar,en;q=0.9'}
        )

        async def inject_cookie():
            """حقن الكوكي الحالي في Playwright"""
            await context.add_cookies(playwright_cookies())

        await inject_cookie()
        page = await context.new_page()
        pbar = tqdm(pending, desc="المباني", unit="مبنى")
        successful_since_excel = 0

        stopped = False
        for building in pbar:
            # نقطة آمنة: نحترم الإيقاف المؤقت/الآمن قبل البدء بأي موقع (ليس وسط حفظ).
            if not _R.wait_if_paused() or _R.should_stop():
                _R.log("⏹ إيقاف آمن بطلب المستخدم — حُفظ التقدّم.")
                print("\n⏹ تم الإيقاف الآمن بطلب المستخدم؛ التقدّم محفوظ ويمكن الاستئناف لاحقاً.")
                stopped = True
                break
            bid   = str(building['id'])
            bname = artifact_stem(building)
            pbar.set_postfix_str(f"{bname[:35]}")
            _R.progress(current=f"{bname[:40]}")
            errors = []
            warnings = []

            # ── 1. قرار الترشيح ──
            if building.get('nomination_url'):
                path = os.path.join(dirs['nominations'], f"{bname}.pdf")
                ok = await save_pdf(page, building['nomination_url'], path)
                if not ok: errors.append('nomination')
            else:
                errors.append('nomination_url_missing')

            # ── 2. استمارة التسجيل ──
            if building.get('registration_url'):
                path = os.path.join(dirs['registration'], f"{bname}.pdf")
                ok = await save_pdf(page, building['registration_url'], path)
                if not ok: errors.append('registration')
            else:
                errors.append('registration_url_missing')

            # ── 3. ملخص الاستمارة ──
            if building.get('summary_url'):
                path = os.path.join(dirs['summary'], f"{bname}.pdf")
                ok = await save_pdf(page, building['summary_url'], path)
                if not ok: errors.append('summary')
            else:
                errors.append('summary_url_missing')

            # ── 4. الصور ──
            def _classify_photos(pr):
                """يُفشِل المبنى فقط عند ضياع صور حقيقية؛ فشل رابط/رابطين مع نجاح
                باقي الصور يُسجَّل كتنبيه ولا يُعدّ فشلاً."""
                failed = pr.get('failed', 0)
                got    = pr.get('downloaded', 0) + pr.get('skipped', 0)
                found  = pr.get('found', 0)
                if failed in (-1, -2):
                    errors.append('photos_page_error')          # تعذّر فتح صفحة الصور
                elif found > 0 and got == 0:
                    errors.append('photos_none')                # يوجد صور لكن لم ينزل أيٌّ منها
                elif failed and failed > 0:
                    warnings.append(f"photos_partial_{failed}/{found}")  # تنبيه فقط

            try:
                photo_result = download_photos(building, dirs['photos'], photo_session)
                if photo_result.get('failed', 0) == -1:
                    refresh_session()
                    photo_session = make_session()
                    photo_result  = download_photos(building, dirs['photos'], photo_session)
                _classify_photos(photo_result)
            except RuntimeError as exc:
                if 'SESSION_EXPIRED' in str(exc):
                    refresh_session()
                    photo_session = make_session()
                    try:
                        photo_result = download_photos(building, dirs['photos'], photo_session)
                        _classify_photos(photo_result)
                    except Exception as retry_exc:
                        errors.append(f'photos_error: {str(retry_exc)[:40]}')
                else:
                    errors.append(f'photos_error: {str(exc)[:40]}')
            except Exception as e:
                errors.append(f'photos_error: {str(e)[:40]}')

            # ── حفظ التقدم ──
            if errors:
                existing = next((f for f in progress['failed'] if f['id'] == bid), None)
                entry = {'name': sanitize(building['name']), 'errors': errors}
                if warnings:
                    entry['warnings'] = warnings
                if existing:
                    existing.update(entry)
                else:
                    progress['failed'].append({'id': bid, **entry})
            else:
                # نجاح — حتى لو كان هناك تنبيه صور جزئي (كل الصور الحقيقية نزلت)
                if bid not in progress['completed']:
                    progress['completed'].append(bid)
                progress['failed'] = [f for f in progress['failed'] if f['id'] != bid]
                if warnings:
                    progress.setdefault('warned', {})[bid] = {'name': bname, 'warnings': warnings}
                else:
                    progress.get('warned', {}).pop(bid, None)
                successful_since_excel += 1

            save_progress(PROGRESS_FILE, progress)

            # تحديث الداش بورد الحيّ في مركز التحكم
            _R.progress(done=len(progress['completed']),
                        failed=len(progress['failed']),
                        skipped=len(progress.get('warned', {})),
                        current=f"{bname[:40]}")

            # ── تحديث Excel ──
            if not errors and successful_since_excel >= 25:
                completed_list = [b for b in buildings if b['id'] in set(progress['completed'])]
                try:
                    update_excel(dirs, DECISION, completed_list)
                    successful_since_excel = 0
                except PermissionError:
                    print("\n⚠️ ملف Excel مفتوح. سيستمر التصدير وسيُعاد تحديثه في النهاية أو عند التشغيل القادم.")

            await asyncio.sleep(DELAY_REQUESTS)

        await browser.close()

    # ── Excel نهائي ──
    completed_list = [b for b in buildings if b['id'] in set(progress['completed'])]
    try:
        excel_path, _ = update_excel(dirs, DECISION, completed_list)
    except PermissionError:
        excel_path = "تعذر تحديثه لأن الملف مفتوح؛ أغلق Excel وأعد تشغيل البرنامج"

    warned = progress.get('warned', {})
    _R.done(ok=len(progress['completed']), failed=len(progress['failed']), skipped=len(warned))
    print(f"\n{'='*55}")
    print("  🎉 اكتملت العملية!")
    print(f"  ✅ مكتمل : {len(progress['completed']):,}")
    print(f"  ❌ فاشل  : {len(progress['failed']):,}")
    if warned:
        print(f"  ⚠️ مكتمل مع تنبيه (رابط غير-صورة فشل، وكل الصور الحقيقية نزلت): {len(warned):,}")
    print(f"  📊 Excel : {excel_path}")
    if progress['failed']:
        print("\n  المباني الفاشلة:")
        for f in progress['failed'][:10]:
            print(f"    • [{f['id']}] {f['name']}: {f['errors']}")
        if len(progress['failed']) > 10:
            print(f"    ... و {len(progress['failed'])-10} أخرى")
    print(f"{'='*55}")

def ask_decision_popup() -> str:
    """اطلب رقم قرار التسجيل من نافذة منبثقة، مع fallback للطرفية عند تعذر Tkinter."""
    if tk is None or simpledialog is None:
        while True:
            value = input("أدخل رقم قرار التسجيل المطلوب تصديره: ").strip()
            if value.isdigit() and int(value) > 0:
                return str(int(value))
            print("❌ الرجاء إدخال رقم قرار صحيح أكبر من صفر.")

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()

        while True:
            value = simpledialog.askstring(
                "تصدير مواقع التراث العمراني",
                "أدخل رقم قرار التسجيل المطلوب تصديره:",
                parent=root,
            )
            if value is None:
                raise SystemExit("تم إلغاء التشغيل بواسطة المستخدم.")

            value = value.strip()
            if value.isdigit() and int(value) > 0:
                return str(int(value))

            messagebox.showerror(
                "رقم غير صحيح",
                "الرجاء إدخال رقم قرار صحيح أكبر من صفر، مثال: 18",
                parent=root,
            )
    except SystemExit:
        raise
    except Exception as exc:
        print(f"⚠️ تعذر فتح نافذة رقم القرار ({exc}). سيتم استخدام الطرفية بدلاً منها.")
        while True:
            value = input("أدخل رقم قرار التسجيل المطلوب تصديره: ").strip()
            if value.isdigit() and int(value) > 0:
                return str(int(value))
            print("❌ الرجاء إدخال رقم قرار صحيح أكبر من صفر.")
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass

def main():
    global USERNAME, PASSWORD, OUTPUT_DIR, PROGRESS_FILE, DECISION, AUTH_COOKIES, SESSION_COOKIE
    parser = argparse.ArgumentParser(description="تصدير ملفات مواقع مشروع التراث العمراني")
    parser.add_argument("--decision", help="رقم قرار التسجيل")
    parser.add_argument("--output", default="HeritageExports", help="مجلد الحفظ")
    parser.add_argument("--start-page", type=int, default=1)
    parser.add_argument("--building-id", default="")
    parser.add_argument("--range-start", type=int)
    parser.add_argument("--range-end", type=int)
    args = parser.parse_args()

    if args.start_page < 1:
        parser.error("--start-page يجب أن يكون 1 أو أكبر")
    if args.range_start is not None and args.range_start < 0:
        parser.error("--range-start لا يمكن أن يكون سالباً")
    if args.range_end is not None and args.range_end < 0:
        parser.error("--range-end لا يمكن أن يكون سالباً")
    if (args.range_start is not None and args.range_end is not None
            and args.range_end <= args.range_start):
        parser.error("--range-end يجب أن يكون أكبر من --range-start")

    DECISION = (args.decision.strip() if args.decision else ask_decision_popup())
    if not DECISION.isdigit() or int(DECISION) <= 0:
        parser.error("--decision يجب أن يكون رقماً صحيحاً أكبر من صفر")
    DECISION = str(int(DECISION))

    # بيانات الدخول: من البيئة (يمرّرها المجمّع) ثم القيم الافتراضية ثم السؤال.
    USERNAME = USERNAME or os.environ.get("HERITAGE_USERNAME") or "H.qurah"
    PASSWORD = PASSWORD or os.environ.get("HERITAGE_PASSWORD") or "313146"
    while not USERNAME:
        USERNAME = input("أدخل اسم المستخدم: ").strip()
        if not USERNAME:
            print("❌ اسم المستخدم لا يمكن أن يكون فارغاً.")
    while not PASSWORD:
        PASSWORD = getpass.getpass("أدخل كلمة المرور: ")
        if not PASSWORD:
            print("❌ كلمة المرور لا يمكن أن تكون فارغة.")
    OUTPUT_DIR = os.path.abspath(args.output)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    PROGRESS_FILE = os.path.join(OUTPUT_DIR, f"progress_{DECISION}.json")

    global START_FROM_PAGE, START_FROM_BUILDING_ID, SPECIFIC_BUILDING_ID, RANGE_START, RANGE_END
    START_FROM_PAGE = args.start_page
    START_FROM_BUILDING_ID = ""
    SPECIFIC_BUILDING_ID = args.building_id
    RANGE_START, RANGE_END = args.range_start, args.range_end

    print("🔐 جاري تسجيل الدخول...")
    SESSION_COOKIE, login_session = do_login(USERNAME, PASSWORD)
    AUTH_COOKIES = requests.utils.dict_from_cookiejar(login_session.cookies)
    print("✅ تم تسجيل الدخول بنجاح")

    # ══════════════════════════════════════════════════════════════
    #  خيارات التشغيل - عدّل هنا حسب حاجتك
    # ══════════════════════════════════════════════════════════════


    # تحديد مبنى واحد بالـ ID (اتركه فارغاً لتصدير الكل)

    # تحديد range من المباني بالـ index (0-based) - مثال: من 0 إلى 50
    # اتركهما None لتصدير الكل

    # ══════════════════════════════════════════════════════════════

    dirs     = make_dirs(OUTPUT_DIR, DECISION)
    progress = load_progress(PROGRESS_FILE)

    # ── جمع قائمة المباني ──
    buildings = get_all_buildings(DECISION, progress, start_from_page=START_FROM_PAGE)

    if not buildings:
        print("❌ لا توجد مباني!")
    else:
        print(f"📋 إجمالي المباني: {len(buildings):,}")

        # ── تطبيق فلتر المبنى المحدد ──
        if SPECIFIC_BUILDING_ID:
            buildings_to_run = [b for b in buildings if b['id'] == SPECIFIC_BUILDING_ID]
            if buildings_to_run:
                print(f"🎯 تصدير مبنى واحد: ID={SPECIFIC_BUILDING_ID} | {buildings_to_run[0]['name']}")
            else:
                print(f"❌ لم يُعثر على ID={SPECIFIC_BUILDING_ID}")
                buildings_to_run = []

        # ── تطبيق فلتر الـ range ──
        elif RANGE_START is not None or RANGE_END is not None:
            s = RANGE_START if RANGE_START is not None else 0
            e = RANGE_END   if RANGE_END   is not None else len(buildings)
            buildings_to_run = buildings[s:e]
            print(f"📌 تصدير من index {s} إلى {e} | ({len(buildings_to_run):,} مبنى)")

        # ── تطبيق فلتر START_FROM_BUILDING_ID ──
        elif START_FROM_BUILDING_ID:
            ids_list = [b['id'] for b in buildings]
            if START_FROM_BUILDING_ID in ids_list:
                start_idx    = ids_list.index(START_FROM_BUILDING_ID)
                already_done = set(progress.get('completed', []))
                for bid in ids_list[:start_idx]:
                    if bid not in already_done:
                        progress['completed'].append(bid)
                buildings_to_run = buildings
                print(f"▶️  البدء من المبنى ID={START_FROM_BUILDING_ID} (تخطي {start_idx:,} مبنى)")
            else:
                print("⚠️  ID غير موجود - سيبدأ تلقائياً من الأول")
                buildings_to_run = buildings

        # ── تصدير الكل ──
        else:
            buildings_to_run = buildings

        # ── تشغيل الأتمتة ──
        if buildings_to_run:
            asyncio.run(run_automation(buildings_to_run, dirs, progress))
        else:
            print("⚠️ لا يوجد شيء للتصدير!")

if __name__ == "__main__":
    main()
