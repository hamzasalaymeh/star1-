# -*- coding: utf-8 -*-
"""مركز تحكم أتمتة مشروع التراث العمراني — نقطة التشغيل.

وضعان:
  • عادي:  python app.py              → يفتح الواجهة.
  • داخلي: ...app.py --hc-run <سكربت> [وسائط]  → يشغّل سكربت أتمتة في نفس العملية.
    يستخدمه المحرّك لتشغيل الأتمتات، فيعمل سواء كان Python عادياً أو ملف EXE
    مجمّعاً (حيث لا يوجد python منفصل).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# --- الوضع الداخلي: تشغيل سكربت أتمتة في نفس العملية (قبل استيراد الواجهة) ---
if len(sys.argv) >= 3 and sys.argv[1] == "--hc-run":
    import runpy
    target = sys.argv[2]
    sys.argv = [target] + sys.argv[3:]
    # مجلد السكربت على المسار حتى يجد agent_runtime والوحدات المجاورة
    sys.path.insert(0, os.path.dirname(os.path.abspath(target)))
    runpy.run_path(target, run_name="__main__")
    sys.exit(0)


def main():
    from ui.main_window import App
    App().mainloop()


if __name__ == "__main__":
    main()
