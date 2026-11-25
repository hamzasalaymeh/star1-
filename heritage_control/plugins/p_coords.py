# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "coords_update", "order": 3,
    "title": "📍 إدخال الإحداثيات من Excel (بدون كروكي)",
    "desc": "ترفع ملف Excel فيه (اسم الموقع + خط الطول X + خط العرض Y + المساحة اختياري) ورقم القرار، "
            "فيطابق كل اسم بأقرب مبنى على المنصة ويكتب إحداثياته ويحفظ — بدون رسم كروكي.",
    "help": "اختر ملف الإحداثيات (Excel)، أدخل رقم القرار (لتحديد نطاق المطابقة)، واختر مجلد الحفظ. "
            "يفتح Chrome وتسجّل الدخول مرة ثم يكمل تلقائياً. المطابقة بالاسم (تشابه)، ويُنتج تقرير "
            "بنسب التطابق. ⚠️ يكتب فعلياً على المنصة — جرّب على عدد محدود أولاً عبر «الحد» وراجع "
            "التقرير قبل التوسّع. قابل للاستئناف. يحتاج Chrome مثبّتاً.",
    "script": "coords_update.py", "cwd_scripts": True,
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "HERITAGE_COORDS_EXCEL", "label": "ملف الإحداثيات (Excel)", "type": "openfile",
         "env": True, "required": True},
        {"name": "HERITAGE_DECISION", "label": "رقم قرار التسجيل (لنطاق المطابقة)", "type": "text",
         "env": True, "required": True},
        {"name": "HERITAGE_OUTPUT", "label": "مجلد الحفظ (تقرير + استئناف)", "type": "folder",
         "env": True, "required": True},
        {"name": "HERITAGE_COORDS_LIMIT", "label": "الحد: عدد المواقع (فاضي = الكل) — للتجربة",
         "type": "text", "env": True},
    ],
}
