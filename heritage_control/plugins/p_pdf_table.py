# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "pdf_table", "order": 6,
    "title": "📄 تصدير جداول معايير التصنيف (PDF)",
    "desc": "PDF واحد لكل موقع يحتوي جدول «معايير التصنيف» + فهرس Excel (بدون صور).",
    "help": "اكتب رقم القرار واختر مجلد الحفظ (إجباري). يحتاج متصفح Playwright. مرحلة جمع الروابط "
            "قد تأخذ دقائق للقرارات الكبيرة (يظهر بالسجل). قابل للاستئناف ويعرض التقدّم حيّاً.",
    "script": "pdf_export.py", "cwd_scripts": True,
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "positional": True, "required": True},
        {"name": "HERITAGE_OUTPUT", "label": "مجلد الحفظ", "type": "folder", "env": True, "required": True},
        {"name": "limit", "label": "عدد الاستمارات (فاضي = الكل)", "type": "text", "positional": True},
        {"name": "HERITAGE_ONLY_IDS", "label": "أرقام مواقع محددة فقط (اختياري، تفصلها فاصلة)",
         "type": "textarea", "env": True},
    ],
}
