# -*- coding: utf-8 -*-
"""
🏛️ مجمّع أتمتة مشروع التراث العمراني — نافذة واحدة لكل الأتمتات
المؤلف: ENG. Hamza ALsalaymeh

شغّل هذا الملف فقط:  python hub.py
تفتح نافذة فيها أزرار لكل أتمتة. تضغط اللي بدك، يسألك عن (المنطقة/القرار/القرى/
المجلد…) — ولا شيء محفور، إنت تختار كل مرة — ثم يشغّل الأتمتة بنافذة طرفية
تشوف فيها التقدّم.

لا تُحفظ أي كلمة مرور على القرص؛ تُدخَل مرة واحدة في النافذة وتُمرَّر للعملية فقط.
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "مجمّع أتمتة مشروع التراث العمراني"
AUTHOR = "ENG. Hamza ALsalaymeh"
HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "scripts")

# مناطق المملكة (أسماء كما تظهر بقائمة المنطقة في المنصة).
REGIONS = [
    "عسير", "مكة المكرمة", "المدينة المنورة", "الرياض", "الباحة", "جازان",
    "نجران", "تبوك", "حائل", "القصيم", "الحدود الشمالية", "الجوف", "الشرقية",
]

# ================================================================= مواصفات المهام
# كل حقل إدخال: name, label, type, وخصائص حسب النوع.
#   type: choice | text | textarea | bool | int | savefile | folder
#   arg:  اسم المعامل بسطر الأوامر (مثل --region). لو غاب، يُمرَّر عبر متغير بيئة env.
#   when: {"field": "<اسم حقل آخر>", "equals": [قيم]} — يُظهر الحقل شرطياً.
TASKS = [
    {
        "id": "historical",
        "title": "📜 أتمتة الأوصاف التاريخية",
        "desc": "كتابة فقرة الوصف التاريخي لكل مبنى — باختيار منطقة/قرارات أو قرى أو أكواد محددة.",
        "script": "historical.py",
        "argv_prefix": ["run"],
        "params": [
            {"name": "region", "label": "المنطقة", "type": "choice",
             "choices": REGIONS, "default": "عسير", "arg": "--region"},
            {"name": "scope", "label": "نطاق العمل", "type": "choice",
             "choices": ["قرارات", "قرى محددة", "أكواد محددة"], "default": "قرارات"},
            {"name": "decisions", "label": "أرقام القرارات (مثال: 25 26)", "type": "text",
             "arg": "--decisions", "nargs": True,
             "when": {"field": "scope", "equals": ["قرارات"]}},
            {"name": "villages", "label": "أسماء القرى (سطر لكل قرية)", "type": "textarea",
             "arg": "--villages", "nargs": True,
             "when": {"field": "scope", "equals": ["قرى محددة"]}},
            {"name": "only_codes", "label": "أكواد المباني (مثال: 40125 40159)", "type": "text",
             "arg": "--only-codes", "nargs": True,
             "when": {"field": "scope", "equals": ["أكواد محددة"]}},
            {"name": "loose", "label": "وضع مرن (آبار/مساجد/حصون بلا أدوار)", "type": "bool",
             "arg": "--loose", "default": True},
            {"name": "preview", "label": "معاينة فقط (بدون حفظ فعلي)", "type": "bool",
             "arg": "--preview", "default": False},
            {"name": "excel", "label": "ملف نتائج Excel (اختياري)", "type": "savefile",
             "arg": "--excel", "filetypes": [("Excel", "*.xlsx")]},
        ],
    },
    {
        "id": "croqui",
        "title": "🗺️ رفع الكروكي + الإحداثيات (X/Y) + المساحة",
        "desc": "يرفع صورة الكروكي ويكتب الإحداثيات والمساحة. يسألك عن مجلد القرى ورقم القرار بنفسه.",
        "script": "croqui.py",
        "launch": "direct",  # السكربت فيه نوافذه الخاصة لاختيار المجلدات/القرار
    },
    {
        "id": "complementary",
        "title": "📐 الأتمتة التكميلية (إحداثيات المضلّع + ارتفاع/مطار)",
        "desc": "يدخل نقاط المضلّع والحقول الثابتة (ارتفاع/اسم مطار/مسافة). يسألك عن المجلد والبيانات بنفسه.",
        "script": "complementary.py",
        "launch": "direct",
    },
    {
        "id": "region_update",
        "title": "🏛️ تعديل الطراز المعماري (المنطقة)",
        "desc": "يحدّث حقل الطراز المعماري لكل مواقع منطقة/خطوة مختارة — معمّم لأي منطقة وأي طراز.",
        "script": "region_update.py",
        "params": [
            {"name": "HERITAGE_REGION_NAME", "label": "المنطقة الجغرافية (فلتر البحث)", "type": "choice",
             "choices": REGIONS, "default": "الباحة", "env": True},
            {"name": "HERITAGE_TARGET_STYLE", "label": "الطراز المعماري المستهدف", "type": "text",
             "default": "المنطقة الجنوبية", "env": True},
            {"name": "limit", "label": "عدد مواقع للتجربة (اتركه 0 للكل)", "type": "int",
             "default": 0, "arg": "--limit"},
            {"name": "retry_failed", "label": "إعادة محاولة الفاشلة فقط", "type": "bool",
             "arg": "--retry-failed", "default": False},
        ],
    },
    {
        "id": "export_files",
        "title": "📸 تصدير ملفات المواقع (3 PDF + الصور)",
        "desc": "لكل موقع في القرار: قرار الترشيح + استمارة التسجيل + ملخص الاستمارة (PDF) "
                "+ تحميل كل الصور + فهرس Excel. قابل للاستئناف.",
        "script": "export_files.py",
        "params": [
            {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "arg": "--decision"},
            {"name": "output", "label": "مجلد الحفظ", "type": "folder", "arg": "--output"},
        ],
    },
    {
        "id": "pdf_table",
        "title": "📄 تصدير جداول معايير التصنيف (PDF)",
        "desc": "PDF واحد لكل موقع يحتوي جدول «معايير التصنيف» + فهرس Excel (بدون صور).",
        "script": "pdf_export.py",
        "cwd": SCRIPTS,  # ليجد الوحدة المجاورة table_pdf_exporter_advanced
        "params": [
            {"name": "decision", "label": "رقم قرار التسجيل", "type": "text", "positional": True},
            {"name": "limit", "label": "عدد الاستمارات (اتركه فاضي للكل)", "type": "text",
             "positional": True, "optional": True},
        ],
        "pass_creds_positional": True,  # pdf_export يأخذ user/pass كوسيطين 3و4
    },
]


# ================================================================= أدوات مساعدة
def python_exe():
    return sys.executable or "python"


def launch(cmd, cwd=None, env_extra=None):
    """يشغّل أمراً في نافذة طرفية جديدة (على ويندوز) ليشوف المستخدم التقدّم.
    نلفّه بـ (cmd /k) حتى تبقى النافذة مفتوحة بعد انتهاء/تعطّل السكربت، فيظهر
    أي خطأ بدل ما تختفي النافذة فوراً."""
    env = os.environ.copy()
    if env_extra:
        env.update({k: str(v) for k, v in env_extra.items() if v is not None})
    kwargs = {"cwd": cwd or HERE, "env": env}
    if os.name == "nt":
        kwargs["creationflags"] = 0x00000010  # CREATE_NEW_CONSOLE
        # نافذة تبقى مفتوحة بعد الانتهاء (نقرأ الرسائل/الأخطاء)
        cmd = "cmd /k " + subprocess.list2cmdline(cmd)
    subprocess.Popen(cmd, **kwargs)


# ================================================================= نموذج إدخال المهمة
class TaskForm(tk.Toplevel):
    def __init__(self, master, task, creds):
        super().__init__(master)
        self.task = task
        self.creds = creds
        self.vars = {}
        self.rows = {}  # name -> (label_widget, input_widget) للإظهار الشرطي
        self.title(task["title"])
        self.configure(bg="#f4f6f8")
        self.resizable(False, False)
        self.attributes("-topmost", True)

        tk.Label(self, text=task["title"], font=("Arial", 14, "bold"),
                 bg="#0d47a1", fg="white", anchor="e", padx=12, pady=8).pack(fill="x")
        tk.Label(self, text=task["desc"], font=("Arial", 9), bg="#f4f6f8",
                 fg="#444", anchor="e", justify="right", wraplength=460,
                 padx=12, pady=6).pack(fill="x")

        body = tk.Frame(self, bg="#f4f6f8", padx=14, pady=6)
        body.pack(fill="both", expand=True)

        for spec in task.get("params", []):
            self._build_row(body, spec)

        self._refresh_conditional()

        btns = tk.Frame(self, bg="#f4f6f8", pady=10)
        btns.pack(fill="x")
        tk.Button(btns, text="▶ تشغيل", command=self._run, bg="#1565c0", fg="white",
                  font=("Arial", 12, "bold"), width=14).pack(side="right", padx=14)
        tk.Button(btns, text="إلغاء", command=self.destroy, width=10).pack(side="right")

    def _build_row(self, parent, spec):
        name, typ = spec["name"], spec["type"]
        row = tk.Frame(parent, bg="#f4f6f8")
        row.pack(fill="x", pady=4)
        lbl = tk.Label(row, text=spec["label"], font=("Arial", 10), bg="#f4f6f8",
                       anchor="e", width=34, justify="right")
        lbl.pack(side="right")

        if typ == "bool":
            var = tk.BooleanVar(value=spec.get("default", False))
            widget = tk.Checkbutton(row, variable=var, bg="#f4f6f8")
        elif typ == "choice":
            var = tk.StringVar(value=spec.get("default", spec["choices"][0]))
            widget = ttk.Combobox(row, textvariable=var, values=spec["choices"],
                                  state="readonly", width=28)
            if "when" not in spec:  # الحقول المتحكّمة تحدّث الإظهار الشرطي
                pass
            var.trace_add("write", lambda *a: self._refresh_conditional())
        elif typ == "textarea":
            var = tk.Text(row, height=4, width=30, font=("Arial", 10))
            widget = var
        elif typ in ("savefile", "folder"):
            var = tk.StringVar(value="")
            inner = tk.Frame(row, bg="#f4f6f8")
            ent = tk.Entry(inner, textvariable=var, width=22, justify="right")
            ent.pack(side="right")
            tk.Button(inner, text="…", width=3,
                      command=lambda s=spec, v=var: self._pick(s, v)).pack(side="right", padx=3)
            widget = inner
        else:  # text, int
            var = tk.StringVar(value=str(spec.get("default", "")))
            widget = tk.Entry(row, textvariable=var, width=30, justify="right")

        widget.pack(side="right", padx=6)
        self.vars[name] = var
        self.rows[name] = (lbl, row)

    def _pick(self, spec, var):
        if spec["type"] == "folder":
            val = filedialog.askdirectory(parent=self)
        else:
            val = filedialog.asksaveasfilename(
                parent=self, defaultextension=".xlsx",
                filetypes=spec.get("filetypes", [("All", "*.*")]))
        if val:
            var.set(val)

    def _value(self, name):
        v = self.vars[name]
        if isinstance(v, tk.Text):
            return v.get("1.0", "end").strip()
        return v.get()

    def _refresh_conditional(self):
        for spec in self.task.get("params", []):
            cond = spec.get("when")
            if not cond:
                continue
            show = self._value(cond["field"]) in cond["equals"]
            _, row = self.rows[spec["name"]]
            if show:
                row.pack(fill="x", pady=4)
            else:
                row.pack_forget()

    def _run(self):
        task = self.task
        script_path = os.path.join(SCRIPTS, task["script"])
        if not os.path.exists(script_path):
            messagebox.showerror("خطأ", f"لم يُعثر على السكربت:\n{script_path}", parent=self)
            return

        cmd = [python_exe(), script_path] + list(task.get("argv_prefix", []))
        env_extra = {}
        positional = []
        # بيانات الدخول المشتركة تُمرّر عبر البيئة لكل السكربتات التي تقرأها.
        user, pw = self.creds()
        if user:
            env_extra["HERITAGE_USERNAME"] = user
        if pw:
            env_extra["HERITAGE_PASSWORD"] = pw

        for spec in task.get("params", []):
            cond = spec.get("when")
            if cond and self._value(cond["field"]) not in cond["equals"]:
                continue
            name = spec["name"]
            val = self._value(name)
            typ = spec["type"]

            if spec.get("env"):
                env_extra[name] = val
                continue
            if spec.get("positional"):
                if val.strip() or not spec.get("optional"):
                    if val.strip():
                        positional.append(val.strip())
                continue
            if typ == "bool":
                if val and spec.get("arg"):
                    cmd.append(spec["arg"])
                continue
            if typ == "int":
                if str(val).strip() and str(val).strip() != "0" and spec.get("arg"):
                    cmd += [spec["arg"], str(val).strip()]
                continue
            # text / textarea / choice / savefile
            if not str(val).strip():
                continue
            if spec.get("arg"):
                if spec.get("nargs"):
                    parts = [p for p in str(val).replace("،", " ").replace("\n", " ").split(" ") if p.strip()] \
                        if typ != "textarea" else [ln.strip() for ln in val.splitlines() if ln.strip()]
                    cmd += [spec["arg"]] + parts
                else:
                    cmd += [spec["arg"], str(val).strip()]

        cmd += positional
        if task.get("pass_creds_positional") and user and pw:
            cmd += [user, pw]

        try:
            launch(cmd, cwd=task.get("cwd"), env_extra=env_extra)
            messagebox.showinfo("انطلقت الأتمتة",
                                "فُتحت نافذة طرفية فيها تقدّم العمل.\n"
                                "تأكد أن ملف الإكسل مسكّر قبل التشغيل.", parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("فشل التشغيل", str(exc), parent=self)


# ================================================================= النافذة الرئيسية
class Hub(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.configure(bg="#f4f6f8")
        self.geometry("560x640")
        self.resizable(False, True)

        tk.Label(self, text="🏛️  " + APP_TITLE, font=("Arial", 16, "bold"),
                 bg="#0d47a1", fg="white", pady=12).pack(fill="x")
        tk.Label(self, text=AUTHOR, font=("Arial", 10), bg="#0d47a1", fg="#bbdefb").pack(fill="x")

        # بيانات الدخول المشتركة (تُدخل مرة، لا تُحفظ على القرص)
        cred = tk.LabelFrame(self, text="بيانات الدخول (اختياري — مرة واحدة)",
                             bg="#f4f6f8", padx=10, pady=8, font=("Arial", 9, "bold"))
        cred.pack(fill="x", padx=14, pady=10)
        self.user_var = tk.StringVar(value="H.qurah")
        self.pw_var = tk.StringVar(value="313146")
        r1 = tk.Frame(cred, bg="#f4f6f8"); r1.pack(fill="x", pady=2)
        tk.Label(r1, text="اسم المستخدم", bg="#f4f6f8", width=12, anchor="e").pack(side="right")
        tk.Entry(r1, textvariable=self.user_var, justify="right").pack(side="right", fill="x", expand=True, padx=6)
        r2 = tk.Frame(cred, bg="#f4f6f8"); r2.pack(fill="x", pady=2)
        tk.Label(r2, text="كلمة المرور", bg="#f4f6f8", width=12, anchor="e").pack(side="right")
        tk.Entry(r2, textvariable=self.pw_var, show="•", justify="right").pack(side="right", fill="x", expand=True, padx=6)
        tk.Label(cred, text="بعض الأتمتات تسجّل الدخول بنفسها (لا تحتاج تعبئة هنا).",
                 bg="#f4f6f8", fg="#777", font=("Arial", 8), anchor="e").pack(fill="x")

        tk.Label(self, text="اختر الأتمتة:", font=("Arial", 11, "bold"),
                 bg="#f4f6f8", anchor="e").pack(fill="x", padx=16, pady=(4, 0))

        wrap = tk.Frame(self, bg="#f4f6f8")
        wrap.pack(fill="both", expand=True, padx=14, pady=6)
        for task in TASKS:
            card = tk.Frame(wrap, bg="white", bd=1, relief="solid")
            card.pack(fill="x", pady=5)
            tk.Button(card, text=task["title"], font=("Arial", 12, "bold"),
                      bg="white", fg="#0d47a1", anchor="e", relief="flat",
                      command=lambda t=task: self.open_task(t)).pack(fill="x", padx=8, pady=(6, 0))
            tk.Label(card, text=task["desc"], font=("Arial", 9), bg="white",
                     fg="#666", anchor="e", justify="right", wraplength=500,
                     padx=10).pack(fill="x", pady=(0, 6))

    def creds(self):
        return self.user_var.get().strip(), self.pw_var.get().strip()

    def open_task(self, task):
        if task.get("launch") == "direct":
            # السكربت مستقل بنوافذه — نشغّله مباشرة
            path = os.path.join(SCRIPTS, task["script"])
            user, pw = self.creds()
            env = {}
            if user: env["HERITAGE_USERNAME"] = user
            if pw: env["HERITAGE_PASSWORD"] = pw
            try:
                launch([python_exe(), path], env_extra=env)
                messagebox.showinfo("انطلقت الأتمتة",
                                    "فُتحت نافذة الأتمتة — تابع تعليماتها.", parent=self)
            except Exception as exc:
                messagebox.showerror("فشل التشغيل", str(exc), parent=self)
            return
        TaskForm(self, task, self.creds)


if __name__ == "__main__":
    Hub().mainloop()
