# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "croqui", "order": 2,
    "title": "🗺️ الكروكي + الإحداثيات (X/Y) + المساحة",
    "desc": "يرفع صورة الكروكي ويكتب الإحداثيات والمساحة لمباني القرى داخل المجلد المختار.",
    "help": "اختر المجلد الأب (جوّاه مجلد لكل قرية فيه Excel + صور)، ورقم القرار اختياري. "
            "يعالج كل المباني، ويدعم الاستئناف (يتخطّى المكتمل ويعيد الفاشل).",
    "script": "croqui.py",
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "HERITAGE_FOLDER", "label": "مجلد القرى (الأب)", "type": "folder", "env": True, "required": True},
        {"name": "HERITAGE_DECISION", "label": "رقم قرار التسجيل (اختياري)", "type": "text", "env": True},
    ],
}
