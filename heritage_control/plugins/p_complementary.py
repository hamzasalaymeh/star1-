# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "complementary", "order": 3,
    "title": "📐 الأتمتة التكميلية (إحداثيات المضلّع + ارتفاع/مطار)",
    "desc": "يدخل نقاط المضلّع والحقول الثابتة لمباني القرى داخل المجلد المختار.",
    "help": "اختر مجلد القرى (فيه ملف Excel لكل قرية). سيسألك البرنامج عن الارتفاع/المطار/المسافة "
            "لكل قرية في نافذة منفصلة (قيم لكل قرية). يدعم الاستئناف. السجل الحيّ يظهر هنا.",
    "script": "complementary.py", "self_dialogs": True,
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "HERITAGE_FOLDER", "label": "مجلد القرى", "type": "folder", "env": True, "required": True},
    ],
}
