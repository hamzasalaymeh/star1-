# -*- coding: utf-8 -*-
"""حفظ تفضيلات الواجهة: آخر إعدادات لكل أتمتة + ملفات تعريف + الثيم.
ملاحظة أمان: لا نحفظ كلمة المرور على القرص إطلاقاً (تبقى في النافذة فقط)."""
import os
import json

STATE_FILE = os.path.join(os.path.expanduser("~"), ".heritage_control_state.json")


def load():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d.setdefault("last", {})
    d.setdefault("presets", {})
    d.setdefault("theme", "light")
    d.setdefault("username", "H.qurah")
    return d


def save(state):
    try:
        clean = {k: state[k] for k in ("last", "presets", "theme", "username") if k in state}
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(clean, f, ensure_ascii=False, indent=1)
    except Exception:
        pass
