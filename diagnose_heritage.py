# -*- coding: utf-8 -*-
"""
🔎 سكربت تشخيص منصة التراث
يطلّع أسماء كل الحقول والأزرار في صفحة البحث وصفحة تعديل المبنى،
ويحفظها في ملف heritage_fields.txt — أرسل هذا الملف للمطور.

الاستخدام:
  1) شغّله: شغّل upload_croquis_v5_autologin.py مرة واحدة أولاً حتى
     يسجّل الدخول ويحفظ الجلسة (أو سجّل دخولك يدوياً في نافذة الأتمتة).
  2) ثم شغّل هذا الملف. سيفتح نفس جلسة Chrome ويطلّع التقرير.
"""

import subprocess, sys, time, os, importlib, urllib.request

for mod, pkg in {'selenium': 'selenium', 'webdriver_manager': 'webdriver-manager'}.items():
    try:
        importlib.import_module(mod)
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '-q'])

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

DATA_VIEW_URL = 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataView'

DEBUG_PORT = 9333
DEBUG_ADDRESS = f'127.0.0.1:{DEBUG_PORT}'
DEBUG_PROFILE = r'C:\heritage_automation_profile'

chrome_paths = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]
chrome_exe = next((p for p in chrome_paths if os.path.exists(p)), None)


def debugger_is_ready():
    try:
        with urllib.request.urlopen(f'http://{DEBUG_ADDRESS}/json/version', timeout=1) as r:
            return r.status == 200
    except Exception:
        return False


if not debugger_is_ready():
    if not chrome_exe:
        print('❌ Chrome غير موجود')
        sys.exit(1)
    subprocess.Popen([
        chrome_exe, f'--remote-debugging-port={DEBUG_PORT}',
        f'--user-data-dir={DEBUG_PROFILE}',
        '--no-first-run', '--no-default-browser-check', '--start-maximized'
    ])
    for _ in range(20):
        if debugger_is_ready():
            break
        time.sleep(0.5)

options = Options()
options.add_experimental_option('debuggerAddress', DEBUG_ADDRESS)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

OUT = os.path.join(os.getcwd(), 'heritage_fields.txt')
lines = []


def dump_page(title):
    lines.append('=' * 70)
    lines.append(title)
    lines.append(f'URL: {driver.current_url}')
    lines.append(f'TITLE: {driver.title}')
    lines.append('=' * 70)

    info = driver.execute_script("""
        function grab(sel) {
            return Array.from(document.querySelectorAll(sel)).map(function(e){
                return {
                    tag: e.tagName,
                    type: e.getAttribute('type') || '',
                    name: e.getAttribute('name') || '',
                    id: e.id || '',
                    placeholder: e.getAttribute('placeholder') || '',
                    value: (e.getAttribute('value') || '').slice(0,30),
                    text: (e.innerText || e.textContent || '').trim().slice(0,40),
                    visible: !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length)
                };
            });
        }
        return {
            inputs: grab('input'),
            buttons: grab('button'),
            selects: grab('select'),
            textareas: grab('textarea'),
            iframes: Array.from(document.querySelectorAll('iframe')).map(function(f){
                return {name: f.getAttribute('name')||'', id: f.id||'', src: f.getAttribute('src')||''};
            })
        };
    """)

    for group in ('inputs', 'buttons', 'selects', 'textareas'):
        items = info.get(group, [])
        lines.append(f'\n--- {group.upper()} ({len(items)}) ---')
        for it in items:
            lines.append(
                f"  tag={it['tag']} type={it['type']!r} name={it['name']!r} "
                f"id={it['id']!r} placeholder={it['placeholder']!r} "
                f"value={it['value']!r} text={it['text']!r} visible={it['visible']}"
            )
    iframes = info.get('iframes', [])
    if iframes:
        lines.append(f'\n--- IFRAMES ({len(iframes)}) ---')
        for fr in iframes:
            lines.append(f"  name={fr['name']!r} id={fr['id']!r} src={fr['src']!r}")
    lines.append('')


print('📍 فتح صفحة البحث (DataView)...')
driver.get(DATA_VIEW_URL)
time.sleep(4)
if 'Login' in driver.current_url or driver.title == 'Runtime Error':
    print('⚠️ يبدو أنك غير مسجّل دخول. سجّل دخولك في نافذة Chrome ثم اضغط Enter هنا...')
    input('')
    driver.get(DATA_VIEW_URL)
    time.sleep(4)

dump_page('صفحة البحث — DataView')
print('✅ تم تصوير صفحة البحث')

print('\n👉 الآن (اختياري لكن مفيد جداً): في نافذة Chrome، ابحث عن أي مبنى يدوياً')
print('   وافتح صفحة "تعديل / استكمال بيانات الميدان" لأي مبنى،')
print('   ثم ارجع هنا واضغط Enter لتصوير صفحة التعديل أيضاً.')
print('   أو اضغط Enter مباشرة لتخطي هذه الخطوة.')
input('')
dump_page('صفحة تعديل المبنى — DataEdit (إن فُتحت)')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f'\n💾 تم حفظ التقرير في: {OUT}')
print('📤 أرسل هذا الملف للمطور.')
