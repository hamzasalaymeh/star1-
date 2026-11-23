# -*- coding: utf-8 -*-
PLUGIN = {
    "id": "failed_tracker", "order": 9,
    "title": "🧹 متابعة الفاشلين",
    "desc": "يمسح مجلد التقارير ويجمع كل المواقع الفاشلة/المتخطّاة في قائمة أكواد جاهزة + ملف Excel. بدون منصة.",
    "help": "اختر مجلد تقارير Excel (إجباري). الأداة تجمع كل اللي حالته فشل/تخطّي/معلّق وتطلّع "
            "قائمة أكواد + ملف Excel مجمّع في نفس المجلد.",
    "script": "failed_tracker.py",
    "supports_control": False, "emits_progress": False,
    "settings": [
        {"name": "folder", "label": "مجلد التقارير", "type": "folder", "arg": "--folder", "required": True},
    ],
}
