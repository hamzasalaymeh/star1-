# -*- coding: utf-8 -*-
from core.constants import REGIONS
PLUGIN = {
    "id": "historical", "order": 1,
    "title": "📜 الأوصاف التاريخية",
    "desc": "كتابة فقرة الوصف التاريخي لكل مبنى — باختيار منطقة/قرارات أو قرى أو أكواد محددة.",
    "help": "اختر المنطقة ثم نطاق العمل (قرارات/قرى/أكواد). «مرن» يشمل الآبار والمساجد بلا أدوار. "
            "«معاينة» يعبّي بدون حفظ. يدعم الاستئناف والإيقاف الحقيقي.",
    "script": "historical.py", "argv_prefix": ["run"],
    "supports_control": True, "emits_progress": True,
    "settings": [
        {"name": "region", "label": "المنطقة", "type": "choice", "choices": REGIONS, "default": "عسير", "arg": "--region"},
        {"name": "scope", "label": "نطاق العمل", "type": "choice", "choices": ["قرارات", "قرى محددة", "أكواد محددة"], "default": "قرارات"},
        {"name": "decisions", "label": "أرقام القرارات (مثال: 25 26)", "type": "text", "arg": "--decisions", "nargs": True, "when": {"field": "scope", "equals": ["قرارات"]}},
        {"name": "villages", "label": "أسماء القرى (سطر لكل قرية)", "type": "textarea", "arg": "--villages", "nargs": True, "when": {"field": "scope", "equals": ["قرى محددة"]}},
        {"name": "only_codes", "label": "أكواد المباني (مثال: 40125 40159)", "type": "text", "arg": "--only-codes", "nargs": True, "when": {"field": "scope", "equals": ["أكواد محددة"]}},
        {"name": "loose", "label": "وضع مرن (آبار/مساجد/حصون بلا أدوار)", "type": "bool", "arg": "--loose", "default": True},
        {"name": "preview", "label": "معاينة فقط (بدون حفظ فعلي)", "type": "bool", "arg": "--preview", "default": False},
        {"name": "excel", "label": "ملف نتائج Excel (اختياري)", "type": "savefile", "arg": "--excel"},
    ],
}
