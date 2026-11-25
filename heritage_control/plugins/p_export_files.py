# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "export_files", "order": 5,
    "title": "📸 تصدير ملفات المواقع (3 PDF + الصور)",
    "desc": "لكل موقع: قرار الترشيح + استمارة التسجيل + ملخص الاستمارة (PDF) + كل الصور + فهرس Excel. قابل للاستئناف.",
    "help": "اكتب رقم القرار واختر مجلد الحفظ (إجباري). يحتاج متصفح Playwright (chromium). قابل "
            "للاستئناف ويعيد الفاشل.",
    "script": "export_files.py",
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "arg": "--decision", "required": True},
        {"name": "output", "label": "مجلد الحفظ", "type": "folder", "arg": "--output", "required": True},
        {"name": "EXPORT_NOMINATION",   "label": "☑ تصدير: قرار الترشيح",   "type": "bool", "env": True, "default": True},
        {"name": "EXPORT_REGISTRATION", "label": "☑ تصدير: استمارة التسجيل", "type": "bool", "env": True, "default": True},
        {"name": "EXPORT_SUMMARY",      "label": "☑ تصدير: ملخص الاستمارة",  "type": "bool", "env": True, "default": True},
        {"name": "EXPORT_PHOTOS",       "label": "☑ تصدير: الصور",           "type": "bool", "env": True, "default": True},
    ],
}
