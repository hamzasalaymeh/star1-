# -*- coding: utf-8 -*-
"""
اختبارات تلقائية للمجمّع. تشتغل بدون منصة (headless عبر xvfb على لينكس،
ومباشرة على ويندوز). تتحقق من:
  • بناء النافذة وكل المهام بدون خطأ Tk.
  • بناء أمر التشغيل الصحيح لكل مهمة (وتمرير بيانات الدخول).
  • حفظ/استرجاع آخر الإعدادات والـ Presets.
  • منطق الاستئناف (يكمل من التوقّف + يعيد الفاشل فقط).

التشغيل:
  python -m pytest tests/            (لو عندك pytest)
  أو:  python tests/test_suite.py    (يشغّل الفحوص ويطبع النتيجة)
"""
import os
import sys
import json
import tempfile
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def _load_hub(state_file):
    # نوجّه ملف الحالة لمكان مؤقت حتى لا نلمس إعدادات المستخدم
    spec = importlib.util.spec_from_file_location("hub", os.path.join(ROOT, "hub.py"))
    hub = importlib.util.module_from_spec(spec)
    sys.modules["hub"] = hub
    spec.loader.exec_module(hub)
    hub.STATE_FILE = state_file
    hub.messagebox.showinfo = lambda *a, **k: None
    hub.messagebox.showerror = lambda *a, **k: None
    return hub


def test_hub_builds_and_commands():
    with tempfile.TemporaryDirectory() as d:
        hub = _load_hub(os.path.join(d, "state.json"))
        cap = []
        hub.launch = lambda cmd, cwd=None, env_extra=None: cap.append((cmd, env_extra))
        app = hub.Hub(); app.update()
        assert app.creds() == ("H.qurah", "313146"), "بيانات الدخول الافتراضية"

        for task in hub.TASKS:
            cap.clear()
            if task.get("launch") == "direct":
                app.open_task(task)
            else:
                f = hub.TaskForm(app, task, app.creds, app.state); f.update()
                if task["id"] == "historical":
                    f.vars["scope"].set("أكواد محددة")
                    f.vars["only_codes"].set("40125 40159")
                    f._refresh_conditional()
                f._run()
            assert cap, f"{task['id']} لم يبنِ أمراً"
            cmd, env = cap[0]
            # بيانات الدخول تمر (env أو positional)
            ok_creds = (env and env.get("HERITAGE_USERNAME")) or task.get("pass_creds_positional")
            assert ok_creds, f"{task['id']} لم يمرّر بيانات الدخول"
        app.destroy()
    print("✅ test_hub_builds_and_commands")


def test_remember_and_presets():
    with tempfile.TemporaryDirectory() as d:
        sf = os.path.join(d, "state.json")
        hub = _load_hub(sf)
        hub.launch = lambda *a, **k: None
        app = hub.Hub(); app.update()
        hist = next(t for t in hub.TASKS if t["id"] == "historical")

        f = hub.TaskForm(app, hist, app.creds, app.state); f.update()
        f.vars["region"].set("الباحة"); f.vars["scope"].set("قرارات")
        f.vars["decisions"].set("24"); f._refresh_conditional()
        # حفظ كـ preset
        app.state.setdefault("presets", {}).setdefault("historical", {})["ب24"] = f._collect_values()
        hub.save_state(app.state)
        f._run()  # يحفظ آخر القيم أيضاً
        f.destroy()

        # نموذج جديد: لازم يتذكّر آخر قيم (الباحة/24)
        f2 = hub.TaskForm(app, hist, app.creds, app.state); f2.update()
        assert f2._value("region") == "الباحة", "لم يتذكّر المنطقة"
        assert f2._value("decisions") == "24", "لم يتذكّر القرارات"
        # تحميل preset
        f2.vars["region"].set("عسير")
        f2._preset_var.set("ب24"); f2._load_preset()
        assert f2._value("region") == "الباحة", "الـ preset لم يُحمّل"
        f2.destroy(); app.destroy()

        # الحالة محفوظة على القرص
        saved = json.load(open(sf, encoding="utf-8"))
        assert saved["last"]["historical"]["region"] == "الباحة"
        assert "ب24" in saved["presets"]["historical"]
    print("✅ test_remember_and_presets")


def test_resume_retry_logic():
    """يحاكي نفس منطق الاستئناف المضاف للأكواد."""
    with tempfile.TemporaryDirectory() as d:
        ck = os.path.join(d, ".checkpoint.json")

        def load():
            try:
                return set(json.load(open(ck, encoding="utf-8")).get("done", []))
            except Exception:
                return set()

        def save(s):
            t = ck + ".tmp"
            json.dump({"done": sorted(s)}, open(t, "w", encoding="utf-8"), ensure_ascii=False)
            os.replace(t, ck)

        items = ["101", "102", "103", "104", "105"]
        # تشغيل أول: 103 يفشل، توقّف بعد 104
        done = load()
        for bid in items[:4]:
            if bid in done:
                continue
            if bid != "103":
                done.add(bid); save(done)
        assert sorted(load()) == ["101", "102", "104"]
        # إعادة: يعالج 103 (الفاشل) و105 (الناقص) فقط
        done = load(); processed = []
        for bid in items:
            if bid in done:
                continue
            processed.append(bid); done.add(bid); save(done)
        assert processed == ["103", "105"], processed
        assert sorted(load()) == items
    print("✅ test_resume_retry_logic")


def _run_all():
    test_resume_retry_logic()  # لا يحتاج Tk
    try:
        test_hub_builds_and_commands()
        test_remember_and_presets()
    except Exception as exc:
        import tkinter  # noqa
        raise
    print("\n🎉 كل الاختبارات نجحت")


if __name__ == "__main__":
    _run_all()
