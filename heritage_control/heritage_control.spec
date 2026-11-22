# -*- mode: python ; coding: utf-8 -*-
# ملف بناء EXE لمركز التحكم.  البناء:  pyinstaller heritage_control.spec
# الناتج:  dist/HeritageControl/HeritageControl.exe
import os
block_cipher = None
ROOT = os.path.abspath('.')

# نضمّ السكربتات والـplugins كملفات فعلية (تُشغّل بالمسار عبر runpy).
datas = [('scripts', 'scripts'), ('plugins', 'plugins')]

# المكتبات التي تستوردها الأتمتات أثناء التشغيل (لا يراها التحليل الساكن
# لأنها تُشغّل عبر runpy) + وحدات الـplugins.
hiddenimports = [
    'plugins.p_historical', 'plugins.p_croqui', 'plugins.p_complementary',
    'plugins.p_region_update', 'plugins.p_export_files', 'plugins.p_pdf_table',
    'plugins.p_failed_tracker',
    'selenium', 'webdriver_manager', 'openpyxl', 'pandas', 'bs4', 'requests',
    'tqdm', 'rich', 'arabic_reshaper', 'bidi', 'bidi.algorithm', 'playwright',
    'playwright.sync_api', 'playwright.async_api', 'tkinter',
]

a = Analysis(['app.py'], pathex=[ROOT], binaries=[], datas=datas,
             hiddenimports=hiddenimports, hookspath=[], hooksconfig={},
             runtime_hooks=[], excludes=[], cipher=block_cipher, noarchive=False)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='HeritageControl',
          debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
          console=False, disable_windowed_traceback=False, icon=None)
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True,
               upx_exclude=[], name='HeritageControl')
