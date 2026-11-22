# -*- coding: utf-8 -*-
"""اختبار المحرّك فعلياً: ابدأ → تقدّم حيّ → إيقاف مؤقت → استئناف → إيقاف آمن.
لا يحتاج واجهة ولا منصة. التشغيل:  python tests/test_engine.py"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from core.engine import AutomationRunner

SCRIPTS = os.path.join(ROOT, "scripts")
MOCK = os.path.join(SCRIPTS, "_mock_automation.py")


def drain(r, seconds, logs):
    """يسحب الأحداث لمدة معيّنة ويرجّع آخر إحصائيات."""
    end = time.time() + seconds
    last = dict(r.stats)
    while time.time() < end:
        for ev in r.drain():
            if ev["type"] == "log":
                logs.append(ev["msg"])
            elif ev["type"] in ("progress", "state"):
                last = dict(r.stats)
            elif ev["type"] == "exit":
                last = dict(r.stats)
                return last, True
        time.sleep(0.05)
    return dict(r.stats), False


def main():
    r = AutomationRunner()
    logs = []
    r.start([sys.executable, MOCK, "60"], cwd=SCRIPTS)
    assert r.running, "لم يبدأ"

    # 1) تقدّم حيّ
    stats, _ = drain(r, 0.8, logs)
    assert stats["total"] == 60, f"total={stats['total']}"
    assert stats["done"] > 0, "لا تقدّم"
    print(f"✅ تقدّم حيّ: done={stats['done']} total={stats['total']} %={stats['percent']:.0f}")

    # 2) إيقاف مؤقت → بعد استقراره يتجمّد العدّاد
    r.pause()
    drain(r, 1.0, logs)              # نسحب باستمرار حتى يستقر الإيقاف ويُفرّغ ما كان في الطريق
    before = dict(r.stats)
    drain(r, 1.0, logs)             # الآن يجب ألا يزيد العدّاد
    after = dict(r.stats)
    assert after["done"] == before["done"], f"ما توقّف: {before['done']}→{after['done']}"
    print(f"✅ إيقاف مؤقت: العدّاد ثابت عند done={after['done']}")

    # 3) استئناف → يكمل
    r.resume()
    stats, _ = drain(r, 0.8, logs)
    assert stats["done"] > after["done"], "ما استأنف"
    print(f"✅ استئناف: done تقدّم إلى {stats['done']}")

    # 4) إيقاف آمن → خروج نظيف قبل إكمال الكل
    r.stop()
    stats, ended = drain(r, 3.0, logs)
    assert ended, "ما خرج بعد الإيقاف"
    assert stats["state"] == "stopped", f"state={stats['state']}"
    assert stats["done"] < stats["total"], "المفروض يوقف قبل النهاية"
    assert any("الإيقاف الآمن" in m for m in logs), "ما سجّل الإيقاف الآمن"
    print(f"✅ إيقاف آمن: خرج نظيفاً عند done={stats['done']}/{stats['total']} (state={stats['state']})")

    r.cleanup()
    assert not os.path.exists(r.control_file), "ملف التحكم ما انحذف"
    print("\n🎉 المحرّك يعمل: ابدأ/تقدّم/إيقاف مؤقت/استئناف/إيقاف آمن — كلها مثبتة")


if __name__ == "__main__":
    main()
