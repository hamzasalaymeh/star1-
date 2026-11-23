# -*- coding: utf-8 -*-
from core.constants import REGIONS
PLUGIN = {
    "id": "region_update", "order": 4,
    "title": "🏛️ تعديل الطراز المعماري (المنطقة)",
    "desc": "يحدّث حقل الطراز المعماري لكل مواقع منطقة/خطوة مختارة — معمّم لأي منطقة وأي طراز.",
    "help": "اختر المنطقة الجغرافية والطراز المستهدف ومجلد الحفظ. «عدد مواقع للتجربة» للتجربة على "
            "عيّنة. يحفظ تقدّمه بقاعدة بيانات ويعيد الفاشل تلقائياً.",
    "script": "region_update.py",
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "HERITAGE_REGION_NAME", "label": "المنطقة الجغرافية", "type": "choice", "choices": REGIONS, "default": "الباحة", "env": True},
        {"name": "HERITAGE_TARGET_STYLE", "label": "الطراز المعماري المستهدف", "type": "text", "default": "المنطقة الجنوبية", "env": True},
        {"name": "HERITAGE_OUTPUT", "label": "مجلد الحفظ", "type": "folder", "env": True, "required": True},
        {"name": "limit", "label": "عدد مواقع للتجربة (0 = الكل)", "type": "int", "default": 0, "arg": "--limit"},
        {"name": "retry_failed", "label": "إعادة محاولة الفاشلة فقط", "type": "bool", "arg": "--retry-failed"},
    ],
}
