# -*- coding: utf-8 -*-
"""أتمتة وهمية للاختبار فقط — تحاكي معالجة مواقع وتحترم الإيقاف/الاستئناف."""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from agent_runtime import R

N = int(sys.argv[1]) if len(sys.argv) > 1 else 10
R.total(N)
R.log(f"بدء معالجة {N} موقع (محاكاة)", "info")

done = failed = 0
for i in range(1, N + 1):
    # نقطة آمنة قبل كل موقع: إيقاف مؤقت + فحص الإيقاف الآمن
    if not R.wait_if_paused():
        R.log("تم طلب الإيقاف الآمن — خروج نظيف بعد آخر موقع مكتمل", "warn")
        break
    if R.should_stop():
        R.log("تم طلب الإيقاف الآمن", "warn")
        break

    R.progress(current=f"موقع رقم {i}")
    time.sleep(0.05)                    # محاكاة زمن المعالجة
    ok = (i % 7 != 0)                   # كل موقع سابع يفشل (محاكاة)
    if ok:
        done += 1
    else:
        failed += 1
    R.progress(done=done, failed=failed, current=f"موقع رقم {i}")
    R.log(("✅ نجح" if ok else "❌ فشل") + f" الموقع {i}", "ok" if ok else "error")

R.done(ok=done, failed=failed)
print(f"انتهى المحاكي — نجح {done} فشل {failed}")
