# -*- coding: utf-8 -*-
"""
نظام الـ Plugins: يكتشف كل وحدات الأتمتة تلقائياً.
كل ملف في هذا المجلد يبدأ بـ p_ ويعرّف متغيّر PLUGIN (قاموس) يظهر تلقائياً
في الواجهة — بدون تعديل الواجهة الرئيسية.

مواصفة الـ PLUGIN:
  id, title, desc, help, order
  script           : اسم ملف الأتمتة داخل scripts/
  argv_prefix      : وسائط ثابتة قبل الإعدادات (مثل ["run"])
  settings         : قائمة حقول الإعداد (choice/text/textarea/bool/int/savefile/folder)
  supports_control : True لو يدعم إيقاف/استئناف حقيقي (فيه hooks)
  emits_progress   : True لو يبعث تقدّماً منظّماً (شريط إنجاز)؛ وإلا سجل حيّ فقط
  self_dialogs     : True لو السكربت يفتح نوافذه الخاصة (إعداداته داخلية)
  cwd_scripts      : True لو يجب التشغيل من مجلد scripts (استيراد وحدات مجاورة)
  pass_creds_positional / env on fields ... (مثل باني الأوامر)
"""
import os
import importlib
import pkgutil


# قائمة احتياطية تضمن اكتشاف الـplugins حتى في نسخة EXE المجمّعة (حيث قد لا
# يرى pkgutil الملفات على القرص). أي plugin جديد يُضاف هنا تلقائياً عبر المسح،
# ويُسجّل اسمه هنا أيضاً ليعمل بعد التجميد.
_KNOWN = [
    "p_historical", "p_croqui", "p_complementary", "p_region_update",
    "p_export_files", "p_pdf_table", "p_failed_tracker",
]


def discover():
    names = []
    try:
        pkg_dir = os.path.dirname(os.path.abspath(__file__))
        names = [m.name for m in pkgutil.iter_modules([pkg_dir]) if m.name.startswith("p_")]
    except Exception:
        names = []
    for n in _KNOWN:                       # دمج القائمة الاحتياطية بدون تكرار
        if n not in names:
            names.append(n)

    plugins, seen = [], set()
    for name in names:
        try:
            m = importlib.import_module(f"plugins.{name}")
            pid = getattr(m, "PLUGIN", {}).get("id")
            if pid and pid not in seen:
                seen.add(pid)
                plugins.append(m.PLUGIN)
        except Exception as exc:
            print(f"[plugins] تعذّر تحميل {name}: {exc}")
    plugins.sort(key=lambda p: p.get("order", 999))
    return plugins
