# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "coords_update", "order": 3,
    "title": "📍 إدخال الإحداثيات من Excel (بدون كروكي)",
    "desc": "ترفع ملف Excel فيه (Site ID + خط الطول X + خط العرض Y + المساحة اختياري)، "
            "فيدخل لكل موقع ويكتب إحداثياته ويحفظ — بدون رسم كروكي وبدون مطابقة بالاسم.",
    "help": "اختر ملف الإحداثيات (Excel) ومجلد الحفظ (للاستئناف). يفتح Chrome وتسجّل الدخول "
            "مرة، ثم يكمل تلقائياً. ⚠️ يكتب فعلياً على المنصة — جرّب على عدد محدود أولاً عبر «الحد». "
            "قابل للاستئناف ويتخطّى المُدخل مسبقاً. يحتاج Chrome مثبّتاً.",
    "script": "coords_update.py", "cwd_scripts": True,
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "HERITAGE_COORDS_EXCEL", "label": "ملف الإحداثيات (Excel)", "type": "openfile",
         "env": True, "required": True},
        {"name": "HERITAGE_OUTPUT", "label": "مجلد الحفظ (للاستئناف)", "type": "folder",
         "env": True, "required": True},
        {"name": "HERITAGE_COORDS_LIMIT", "label": "الحد: عدد المواقع (فاضي = الكل) — للتجربة",
         "type": "text", "env": True},
    ],
}
