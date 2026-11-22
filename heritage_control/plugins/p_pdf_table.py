# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "pdf_table", "order": 6,
    "title": "📄 تصدير جداول معايير التصنيف (PDF)",
    "desc": "PDF واحد لكل موقع يحتوي جدول «معايير التصنيف» + فهرس Excel (بدون صور).",
    "help": "اكتب رقم القرار، ويمكن تحديد عدد الاستمارات. يحتاج متصفح Playwright. قابل للاستئناف.",
    "script": "pdf_export.py", "cwd_scripts": True, "pass_creds_positional": True,
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "positional": True, "required": True},
        {"name": "limit", "label": "عدد الاستمارات (فاضي = الكل)", "type": "text", "positional": True},
    ],
}
