# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "coords_update", "order": 3,
    "title": "📍 إدخال الإحداثيات من Excel (بدون كروكي)",
    "desc": "ترفع ملف Excel فيه الإحداثيات (X و Y) + «كود الموقع» (رقم المبنى على المنصة) أو «اسم الموقع»، "
            "فيدخل لكل موقع ويكتب إحداثياته ويحفظ — بدون رسم كروكي. الأدق: عمود كود الموقع.",
    "help": "اختر ملف الإحداثيات (Excel) ومجلد الحفظ. لو الملف فيه عمود «كود الموقع» (رقم المبنى على "
            "المنصة) → إدخال مباشر بلا رقم قرار (الأدق). لو فيه «اسم الموقع» فقط → أدخل رقم القرار "
            "وتتم المطابقة بالاسم. يفتح Chrome وتسجّل الدخول مرة ثم يكمل تلقائياً، ويُنتج تقرير. "
            "⚠️ يكتب فعلياً على المنصة — جرّب على عدد محدود أولاً عبر «الحد». قابل للاستئناف. يحتاج Chrome.",
    "script": "coords_update.py", "cwd_scripts": True,
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "HERITAGE_COORDS_EXCEL", "label": "ملف الإحداثيات (Excel)", "type": "openfile",
         "env": True, "required": True},
        {"name": "HERITAGE_OUTPUT", "label": "مجلد الحفظ (تقرير + استئناف)", "type": "folder",
         "env": True, "required": True},
        {"name": "HERITAGE_DECISION", "label": "رقم القرار (فقط لو الملف بأسماء بدون كود)", "type": "text",
         "env": True},
        {"name": "HERITAGE_COORDS_LIMIT", "label": "الحد: عدد المواقع (فاضي = الكل) — للتجربة",
         "type": "text", "env": True},
    ],
}
