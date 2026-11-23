# -*- coding: utf-8 -*-
"""اختبار الواجهة فعلياً (headless): بناء كل الشاشات بدون خطأ، وقيادة شاشة
تشغيل بأتمتة وهمية عبر المحرّك: ابدأ/إيقاف مؤقت/استئناف/إيقاف — والتأكد أن
العدّادات بالواجهة تتحرّك فعلاً. التشغيل:  xvfb-run python3.12 tests/test_ui.py"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from ui import main_window as mw

SCRIPTS = os.path.join(ROOT, "scripts")
MOCK = os.path.join(SCRIPTS, "_mock_automation.py")


def pump(app, seconds):
    end = time.time() + seconds
    while time.time() < end:
        app.update()
        time.sleep(0.03)


def main():
    app = mw.App()
    app.update()
    print(f"✅ النافذة الرئيسية بُنيت — {len(app.plugins)} أتمتة")

    # 1) بناء شاشة كل أتمتة بدون خطأ Tk
    for p in app.plugins:
        app.open_plugin(p); app.update()
        app.show_gallery(); app.update()
    print("✅ كل شاشات الأتمتات تُبنى بدون خطأ")

    # 2) قيادة شاشة تشغيل بأتمتة وهمية عبر الأزرار الحقيقية
    hist = next(p for p in app.plugins if p["id"] == "historical")
    rv = mw.RunView(app.body, app, hist)
    rv.pack(fill="both", expand=True)
    app._cmd_override = ([sys.executable, MOCK, "60"], {})
    rv._start()
    pump(app, 1.0)
    s = rv.runner.stats
    assert s["total"] == 60 and s["done"] > 0, f"لا تقدّم: {s}"
    assert rv.tiles["done"].cget("text") == str(s["done"]), "بطاقة «نجح» لم تتحدّث"
    assert rv.tiles["total"].cget("text") == "60", "بطاقة «الكل» لم تتحدّث"
    print(f"✅ داش بورد حيّ: الكل={rv.tiles['total'].cget('text')} نجح={rv.tiles['done'].cget('text')} "
          f"فشل={rv.tiles['failed'].cget('text')} متبقّي={rv.tiles['remaining'].cget('text')} {rv.lbl_pct.cget('text')}")

    # إيقاف مؤقت
    rv._pause(); pump(app, 1.0)
    before = rv.runner.stats["done"]; pump(app, 1.0); after = rv.runner.stats["done"]
    assert after == before, f"ما توقّف: {before}->{after}"
    assert rv.btn_resume["state"] == "normal", "زر الاستئناف غير مفعّل"
    print(f"✅ زر الإيقاف المؤقت أوقف العدّاد عند {after}")

    # استئناف
    rv._resume(); pump(app, 1.0)
    assert rv.runner.stats["done"] > after, "ما استأنف"
    print(f"✅ زر الاستئناف كمّل إلى {rv.runner.stats['done']}")

    # إيقاف آمن
    rv._stop()
    end = time.time() + 4
    while rv.runner.running and time.time() < end:
        pump(app, 0.2)
    assert not rv.runner.running, "ما توقّف"
    assert rv.runner.stats["state"] == "stopped", rv.runner.stats["state"]
    assert rv.btn_start["state"] == "normal", "زر ابدأ لم يرجع"
    print(f"✅ زر الإيقاف الآمن أوقف نظيفاً عند {rv.runner.stats['done']}/{rv.runner.stats['total']}")

    app.destroy()
    print("\n🎉 الواجهة متصلة بالمحرّك فعلياً — كل أزرار التحكم تعمل")


if __name__ == "__main__":
    main()
