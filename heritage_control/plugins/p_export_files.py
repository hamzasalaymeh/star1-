# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "export_files", "order": 5,
    "title": "📸 تصدير ملفات المواقع (3 PDF + الصور)",
    "desc": "لكل موقع: قرار الترشيح + استمارة التسجيل + ملخص الاستمارة (PDF) + كل الصور + فهرس Excel. قابل للاستئناف.",
    "help": "اكتب رقم القرار واختر مجلد الحفظ. يحتاج متصفح Playwright (chromium). قابل للاستئناف ويعيد الفاشل.",
    "script": "export_files.py",
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "arg": "--decision", "required": True},
        {"name": "output", "label": "مجلد الحفظ", "type": "folder", "arg": "--output"},
    ],
}
