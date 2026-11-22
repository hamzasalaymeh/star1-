# -*- coding: utf-8 -*-
"""
وحدة اتصال خفيفة بين أي أتمتة ومركز التحكم (مستقلة، بدون أي اعتماديات).
- الأتمتة تستوردها:  from agent_runtime import R
- تبعث تقدّمها:       R.total(n) / R.progress(done=..,failed=..,current=..) / R.log("..")
- تفحص التحكم بأمان:  R.wait_if_paused()  و  R.should_stop()  عند نقاط آمنة فقط
                      (قبل كل موقع — أبداً وسط عملية حفظ).

إذا شُغّلت الأتمتة وحدها (خارج مركز التحكم) تتحوّل كل الدوال إلى صامتة/ممرّرة،
فيبقى السكربت الأصلي يعمل كما هو بدون أي تغيير بالسلوك.
"""
import os
import sys
import json
import time

EVT = "@@EVT@@"                                   # بادئة سطر الحدث المنظّم
_CONTROL = os.environ.get("HERITAGE_CONTROL")     # مسار ملف التحكم (يضبطه المحرّك)
_ENABLED = bool(_CONTROL)                          # مُفعّل فقط عند التشغيل من المركز


def _emit(kind, **data):
    if not _ENABLED:
        return
    try:
        payload = {"kind": kind}
        payload.update({k: v for k, v in data.items() if v is not None})
        sys.stdout.write(EVT + json.dumps(payload, ensure_ascii=False) + "\n")
        sys.stdout.flush()
    except Exception:
        pass


def _control():
    if not _CONTROL:
        return {}
    try:
        with open(_CONTROL, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


class _Reporter:
    # -- أحداث التقدّم --
    def total(self, n):
        _emit("total", total=int(n))

    def progress(self, done=None, total=None, failed=None, skipped=None, current=None):
        _emit("progress", done=done, total=total, failed=failed, skipped=skipped, current=current)

    def log(self, msg, level="info"):
        _emit("log", msg=str(msg), level=level)

    def done(self, ok=0, failed=0, skipped=0):
        _emit("done", ok=int(ok), failed=int(failed), skipped=int(skipped))

    # -- التحكم (تُستدعى عند نقاط آمنة فقط) --
    def should_stop(self):
        return bool(_control().get("stop"))

    def wait_if_paused(self):
        """يحبس التنفيذ طالما الإيقاف المؤقت مفعّل. يرجّع False إذا طُلب إيقاف آمن."""
        notified = False
        while _control().get("pause") and not _control().get("stop"):
            if not notified:
                _emit("state", state="paused")
                notified = True
            time.sleep(0.4)
        if notified:
            _emit("state", state="running")
        return not _control().get("stop")

    @property
    def enabled(self):
        return _ENABLED


R = _Reporter()
