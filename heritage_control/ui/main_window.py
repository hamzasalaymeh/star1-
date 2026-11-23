# -*- coding: utf-8 -*-
"""مركز تحكم أتمتة التراث — الواجهة (Tkinter). تتكلّم مع المحرّك المعزول،
فلا تتجمّد أثناء التشغيل (تسحب الأحداث عبر مؤقّت)، وأزرارها متصلة فعلاً بالمحرّك."""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from core.engine import AutomationRunner
from core.launcher import build_command
from core import state as state_mod
from plugins.registry import discover

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(HERE, "scripts")

THEMES = {
    "light": {"bg": "#eef2f7", "card": "#ffffff", "fg": "#1a2733", "muted": "#5b6b7a",
              "primary": "#1565c0", "head": "#0d47a1", "headfg": "#ffffff",
              "ok": "#2e7d32", "err": "#c62828", "warn": "#e67e22",
              "logbg": "#0f1b2a", "logfg": "#d6e4f0", "bar": "#1565c0", "entry": "#ffffff"},
    "dark": {"bg": "#121a24", "card": "#1b2735", "fg": "#e6eef6", "muted": "#9fb2c4",
             "primary": "#4f9be0", "head": "#0b2238", "headfg": "#eaf3ff",
             "ok": "#66bb6a", "err": "#ef5350", "warn": "#ffa726",
             "logbg": "#0a1420", "logfg": "#cfe2f2", "bar": "#4f9be0", "entry": "#223345"},
}

LEVEL_COLORS = {"ok": "ok", "error": "err", "fail": "err", "warn": "warn",
                "skip": "muted", "preview": "primary", "info": "logfg"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مركز تحكم أتمتة مشروع التراث العمراني")
        self.geometry("920x720")
        self.minsize(820, 600)
        self.state_data = state_mod.load()
        self.theme_name = self.state_data.get("theme", "light")
        self.plugins = discover()
        self.user_var = tk.StringVar(value=self.state_data.get("username", "H.qurah"))
        self.pw_var = tk.StringVar(value="313146")  # معبّأة، لا تُحفظ على القرص
        self._cmd_override = None  # للاختبار فقط
        self._build()
        self.apply_theme()
        self.show_gallery()

    # ---------------------------------------------------------------- الثيم
    @property
    def T(self):
        return THEMES[self.theme_name]

    def toggle_theme(self):
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self.state_data["theme"] = self.theme_name
        state_mod.save(self.state_data)
        self.apply_theme()
        # إعادة بناء الشاشة الحالية لتطبيق الألوان
        if getattr(self, "_current", "gallery") == "gallery":
            self.show_gallery()

    def apply_theme(self):
        t = self.T
        self.configure(bg=t["bg"])
        self.header.configure(bg=t["head"])
        self.title_lbl.configure(bg=t["head"], fg=t["headfg"])
        self.theme_btn.configure(text="🌙 داكن" if self.theme_name == "light" else "☀️ فاتح")
        for w in self.cred_widgets:
            try:
                w.configure(bg=t["head"], fg=t["headfg"])
            except tk.TclError:
                pass
        self.body.configure(bg=t["bg"])

    # ---------------------------------------------------------------- البناء الأساسي
    def _build(self):
        t = self.T
        self.header = tk.Frame(self, height=58)
        self.header.pack(fill="x")
        self.title_lbl = tk.Label(self.header, text="🏛️  مركز تحكم أتمتة التراث العمراني",
                                  font=("Arial", 15, "bold"))
        self.title_lbl.pack(side="right", padx=14, pady=10)
        self.theme_btn = tk.Button(self.header, text="🌙", command=self.toggle_theme,
                                   relief="flat", font=("Arial", 10))
        self.theme_btn.pack(side="left", padx=10)

        # بيانات الدخول (تُدخل مرة، كلمة المرور لا تُحفظ)
        self.cred_widgets = []
        cf = tk.Frame(self.header)
        cf.pack(side="left", padx=6)
        self.cred_widgets.append(cf)
        for label, var, show in (("المستخدم", self.user_var, ""), ("المرور", self.pw_var, "•")):
            lb = tk.Label(cf, text=label, font=("Arial", 8)); lb.pack(side="left")
            en = tk.Entry(cf, textvariable=var, width=9, show=show, justify="center")
            en.pack(side="left", padx=2)
            self.cred_widgets.append(lb)

        self.body = tk.Frame(self)
        self.body.pack(fill="both", expand=True)

    def creds(self):
        return self.user_var.get().strip(), self.pw_var.get().strip()

    def _clear_body(self):
        for w in self.body.winfo_children():
            w.destroy()

    # ---------------------------------------------------------------- شاشة البطاقات
    def show_gallery(self):
        self._current = "gallery"
        self._clear_body()
        t = self.T
        tk.Label(self.body, text="اختر الأتمتة:", bg=t["bg"], fg=t["fg"],
                 font=("Arial", 12, "bold"), anchor="e").pack(fill="x", padx=18, pady=(10, 2))

        canvas = tk.Canvas(self.body, bg=t["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(self.body, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=t["bg"])
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=880)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=12)
        sb.pack(side="right", fill="y")

        for p in self.plugins:
            card = tk.Frame(inner, bg=t["card"], bd=0, highlightthickness=1,
                            highlightbackground="#c9d6e3")
            card.pack(fill="x", pady=6, padx=6)
            top = tk.Frame(card, bg=t["card"]); top.pack(fill="x", padx=10, pady=(8, 0))
            tk.Button(top, text="؟", width=2, font=("Arial", 9, "bold"), relief="groove",
                      command=lambda pl=p: messagebox.showinfo(pl["title"], pl.get("help") or pl["desc"])
                      ).pack(side="left")
            badge = "● تحكم كامل" if p.get("supports_control") else "○ سجل حيّ"
            tk.Label(top, text=badge, bg=t["card"], fg=t["muted"], font=("Arial", 8)).pack(side="left", padx=8)
            tk.Button(top, text=p["title"], font=("Arial", 13, "bold"), bg=t["card"], fg=t["head"],
                      relief="flat", anchor="e", cursor="hand2",
                      command=lambda pl=p: self.open_plugin(pl)).pack(side="right", fill="x", expand=True)
            tk.Label(card, text=p["desc"], bg=t["card"], fg=t["muted"], font=("Arial", 9),
                     anchor="e", justify="right", wraplength=820).pack(fill="x", padx=12, pady=(0, 9))

    # ---------------------------------------------------------------- شاشة الأتمتة
    def open_plugin(self, plugin):
        self._current = plugin["id"]
        self._clear_body()
        RunView(self.body, self, plugin).pack(fill="both", expand=True)


class RunView(tk.Frame):
    def __init__(self, parent, app, plugin):
        super().__init__(parent, bg=app.T["bg"])
        self.app = app
        self.plugin = plugin
        self.t = app.T
        self.runner = AutomationRunner()
        self.vars = {}
        self.rows = {}
        self._polling = False
        self._build()

    def _build(self):
        t = self.t
        bar = tk.Frame(self, bg=t["bg"]); bar.pack(fill="x", padx=12, pady=6)
        tk.Button(bar, text="⬅ رجوع", command=self._back, relief="groove").pack(side="left")
        tk.Label(bar, text=self.plugin["title"], bg=t["bg"], fg=t["head"],
                 font=("Arial", 13, "bold"), anchor="e").pack(side="right", fill="x", expand=True)

        # ---- الإعدادات ----
        box = tk.LabelFrame(self, text="الإعدادات", bg=t["card"], fg=t["fg"],
                            font=("Arial", 10, "bold"), padx=10, pady=6)
        box.pack(fill="x", padx=12, pady=4)
        if self.plugin.get("self_dialogs"):
            tk.Label(box, text="هذه الأتمتة تسأل عن إعداداتها بنوافذها الخاصة عند التشغيل.",
                     bg=t["card"], fg=t["muted"], anchor="e").pack(fill="x")
        for spec in self.plugin.get("settings", []):
            self._build_row(box, spec)
        self._apply_values(self.app.state_data.get("last", {}).get(self.plugin["id"], {}))
        self._refresh_conditional()
        if self.plugin.get("settings"):
            self._presets_bar(box)

        # ---- أزرار التحكم ----
        ctl = tk.Frame(self, bg=t["bg"]); ctl.pack(fill="x", padx=12, pady=6)
        self.btn_start = tk.Button(ctl, text="▶ ابدأ", bg=t["ok"], fg="white",
                                   font=("Arial", 11, "bold"), width=10, command=self._start)
        self.btn_pause = tk.Button(ctl, text="⏸ إيقاف مؤقت", width=12, state="disabled", command=self._pause)
        self.btn_resume = tk.Button(ctl, text="▶ استئناف", width=10, state="disabled", command=self._resume)
        self.btn_stop = tk.Button(ctl, text="⏹ إيقاف آمن", bg=t["err"], fg="white",
                                  width=10, state="disabled", command=self._stop)
        for b in (self.btn_start, self.btn_pause, self.btn_resume, self.btn_stop):
            b.pack(side="right", padx=4)
        if not self.plugin.get("supports_control"):
            self.btn_pause.configure(text="⏸ (غير مدعوم)")

        # ---- لوحة المراقبة (داش بورد): بطاقات مؤشرات + تقدّم + أزمنة ----
        mon = tk.Frame(self, bg=t["bg"]); mon.pack(fill="x", padx=12, pady=(2, 4))
        tiles = tk.Frame(mon, bg=t["bg"]); tiles.pack(fill="x")
        self.tiles = {}
        for key, label, color in (("total", "الكل", t["fg"]), ("done", "نجح", t["ok"]),
                                  ("failed", "فشل", t["err"]), ("skipped", "متخطّى", t["warn"]),
                                  ("remaining", "متبقّي", t["primary"])):
            cell = tk.Frame(tiles, bg=t["card"], highlightthickness=1, highlightbackground="#c9d6e3")
            cell.pack(side="right", expand=True, fill="both", padx=3)
            val = tk.Label(cell, text="0", bg=t["card"], fg=color, font=("Arial", 20, "bold"))
            val.pack(pady=(8, 0))
            tk.Label(cell, text=label, bg=t["card"], fg=t["muted"], font=("Arial", 9)).pack(pady=(0, 8))
            self.tiles[key] = val

        prog = tk.Frame(mon, bg=t["card"], highlightthickness=1, highlightbackground="#c9d6e3")
        prog.pack(fill="x", pady=6)
        head = tk.Frame(prog, bg=t["card"]); head.pack(fill="x", padx=10, pady=(8, 0))
        self.lbl_state = tk.Label(head, text="جاهز", bg=t["card"], fg=t["muted"], font=("Arial", 11, "bold"))
        self.lbl_state.pack(side="right")
        self.lbl_pct = tk.Label(head, text="0%", bg=t["card"], fg=t["primary"], font=("Arial", 15, "bold"))
        self.lbl_pct.pack(side="left")
        self.pbar = ttk.Progressbar(prog, maximum=100, value=0)
        self.pbar.pack(fill="x", padx=10, pady=(3, 4))
        self.lbl_extra = tk.Label(prog, text="", bg=t["card"], fg=t["muted"], font=("Arial", 9), anchor="e")
        self.lbl_extra.pack(fill="x", padx=10, pady=(0, 8))

        # ---- السجل الحيّ ----
        logf = tk.Frame(self, bg=t["bg"]); logf.pack(fill="both", expand=True, padx=12, pady=(2, 10))
        self.log = tk.Text(logf, bg=t["logbg"], fg=t["logfg"], font=("Consolas", 9),
                           wrap="word", height=12, state="disabled")
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); self.log.pack(side="right", fill="both", expand=True)
        for lvl, key in LEVEL_COLORS.items():
            self.log.tag_configure(lvl, foreground=self.t.get(key, self.t["logfg"]))

    # ---- بناء الحقول (مثل باني الإعدادات) ----
    def _build_row(self, parent, spec):
        t = self.t
        name, typ = spec["name"], spec["type"]
        row = tk.Frame(parent, bg=t["card"]); row.pack(fill="x", pady=3)
        lbl = tk.Label(row, text=spec["label"], bg=t["card"], fg=t["fg"], width=32,
                       anchor="e", justify="right")
        lbl.pack(side="right")
        if typ == "bool":
            var = tk.BooleanVar(value=spec.get("default", False))
            w = tk.Checkbutton(row, variable=var, bg=t["card"])
        elif typ == "choice":
            var = tk.StringVar(value=spec.get("default", spec["choices"][0]))
            w = ttk.Combobox(row, textvariable=var, values=spec["choices"], state="readonly", width=26)
            var.trace_add("write", lambda *a: self._refresh_conditional())
        elif typ == "textarea":
            var = tk.Text(row, height=4, width=28); w = var
        elif typ in ("folder", "savefile"):
            var = tk.StringVar()
            inner = tk.Frame(row, bg=t["card"])
            tk.Entry(inner, textvariable=var, width=20, justify="right").pack(side="right")
            tk.Button(inner, text="…", width=3, command=lambda s=spec, v=var: self._pick(s, v)).pack(side="right", padx=3)
            w = inner
        else:
            var = tk.StringVar(value=str(spec.get("default", "")))
            w = tk.Entry(row, textvariable=var, width=28, justify="right")
        w.pack(side="right", padx=6)
        self.vars[name] = var
        self.rows[name] = row

    def _pick(self, spec, var):
        if spec["type"] == "folder":
            v = filedialog.askdirectory(parent=self)
        else:
            v = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")])
        if v:
            var.set(v)

    def _value(self, name):
        v = self.vars[name]
        return v.get("1.0", "end").strip() if isinstance(v, tk.Text) else v.get()

    def _collect(self):
        return {n: self._value(n) for n in self.vars}

    def _apply_values(self, values):
        for n, val in (values or {}).items():
            var = self.vars.get(n)
            if var is None:
                continue
            try:
                if isinstance(var, tk.Text):
                    var.delete("1.0", "end"); var.insert("1.0", val)
                elif isinstance(var, tk.BooleanVar):
                    var.set(bool(val))
                else:
                    var.set(val)
            except Exception:
                pass

    def _refresh_conditional(self):
        for spec in self.plugin.get("settings", []):
            cond = spec.get("when")
            if not cond:
                continue
            show = self._value(cond["field"]) in cond["equals"]
            row = self.rows[spec["name"]]
            row.pack(fill="x", pady=3) if show else row.pack_forget()

    def _presets_bar(self, parent):
        t = self.t
        bar = tk.Frame(parent, bg=t["card"]); bar.pack(fill="x", pady=(4, 0))
        tk.Label(bar, text="إعداد محفوظ:", bg=t["card"], fg=t["muted"], anchor="e").pack(side="right")
        self._preset_var = tk.StringVar()
        names = sorted(self.app.state_data.get("presets", {}).get(self.plugin["id"], {}).keys())
        self._preset_cb = ttk.Combobox(bar, textvariable=self._preset_var, values=names,
                                       state="readonly", width=16)
        self._preset_cb.pack(side="right", padx=4)
        tk.Button(bar, text="حمّل", width=5, command=self._load_preset).pack(side="right", padx=2)
        tk.Button(bar, text="احفظ", width=6, command=self._save_preset).pack(side="right", padx=2)

    def _load_preset(self):
        p = self.app.state_data.get("presets", {}).get(self.plugin["id"], {}).get(self._preset_var.get())
        if p:
            self._apply_values(p); self._refresh_conditional()

    def _save_preset(self):
        name = simpledialog.askstring("حفظ إعداد", "اسم الإعداد:", parent=self)
        if not name:
            return
        self.app.state_data.setdefault("presets", {}).setdefault(self.plugin["id"], {})[name] = self._collect()
        state_mod.save(self.app.state_data)
        self._preset_cb["values"] = sorted(self.app.state_data["presets"][self.plugin["id"]].keys())
        self._preset_var.set(name)

    # ---- السجل ----
    def _append_log(self, msg, level="info"):
        self.log.configure(state="normal")
        tag = level if level in LEVEL_COLORS else "info"
        self.log.insert("end", msg + "\n", tag)
        # نحدّ طول السجل
        if int(self.log.index("end-1c").split(".")[0]) > 1500:
            self.log.delete("1.0", "500.0")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---- التشغيل والتحكم ----
    def _start(self):
        # نحفظ آخر القيم
        self.app.state_data.setdefault("last", {})[self.plugin["id"]] = self._collect()
        self.app.state_data["username"] = self.app.creds()[0]
        state_mod.save(self.app.state_data)
        try:
            if self.app._cmd_override:
                cmd, env = self.app._cmd_override
            else:
                cmd, env = build_command(self.plugin, self._collect(), self.app.creds(), SCRIPTS)
        except ValueError as exc:
            messagebox.showerror("إعداد ناقص", str(exc), parent=self)
            return
        cwd = SCRIPTS if self.plugin.get("cwd_scripts") else HERE
        self._append_log("▶ بدء التشغيل…", "info")
        self.runner.start(cmd, cwd=cwd, env_extra=env)
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        if self.plugin.get("supports_control"):
            self.btn_pause.configure(state="normal")
        self._polling = True
        self._poll()

    def _pause(self):
        self.runner.pause()
        self.btn_pause.configure(state="disabled")
        self.btn_resume.configure(state="normal")

    def _resume(self):
        self.runner.resume()
        self.btn_resume.configure(state="disabled")
        self.btn_pause.configure(state="normal")

    def _stop(self):
        if not self.plugin.get("supports_control"):
            if not messagebox.askyesno("إيقاف", "هذه الأتمتة لا تدعم الإيقاف الآمن — سيُغلق التشغيل. متابعة؟", parent=self):
                return
            self.runner.kill()
        else:
            self.runner.stop()
        self.btn_stop.configure(state="disabled")
        self.lbl_state.configure(text="جاري الإيقاف الآمن…")

    def _poll(self):
        if not self._polling:
            return
        for ev in self.runner.drain():
            if ev["type"] == "log":
                self._append_log(ev["msg"], ev.get("level", "info"))
            elif ev["type"] == "exit":
                self._on_finish(ev)
            elif ev["type"] == "state":
                pass
        self._update_monitor()
        if self.runner.running:
            self.after(150, self._poll)
        else:
            self._polling = False

    def _update_monitor(self):
        s = self.runner.stats
        rem = self.runner.remaining()
        self.pbar.configure(value=s["percent"])
        self.lbl_pct.configure(text=f"{s['percent']:.0f}%")
        vals = {"total": s["total"], "done": s["done"], "failed": s["failed"],
                "skipped": s["skipped"], "remaining": rem}
        for k, lbl in self.tiles.items():
            lbl.configure(text=str(vals[k]))

        def fmt(x):
            x = int(x)
            return f"{x // 60}:{x % 60:02d}"

        el = s["elapsed"]
        eta = int(s["avg"] * rem) if (s["avg"] and rem) else 0
        self.lbl_extra.configure(
            text=f"الموقع الحالي: {s['current'] or '—'}    |    متوسط الموقع: {s['avg']:.1f}ث"
                 f"    |    مضى: {fmt(el)}    |    متبقّي ≈ {fmt(eta)}")
        names = {"running": "▶ شغّال", "paused": "⏸ متوقّف مؤقتاً", "stopping": "⏹ جاري الإيقاف",
                 "stopped": "⏹ أُوقف", "finished": "✅ انتهى", "idle": "جاهز"}
        self.lbl_state.configure(text=names.get(s["state"], s["state"]))

    def _on_finish(self, ev):
        self._update_monitor()
        self.btn_start.configure(state="normal")
        for b in (self.btn_pause, self.btn_resume, self.btn_stop):
            b.configure(state="disabled")
        msg = "✅ انتهى التشغيل" if ev["state"] == "finished" else "⏹ تم الإيقاف"
        self._append_log(msg, "ok" if ev["state"] == "finished" else "warn")
        self.runner.cleanup()

    def _back(self):
        if self.runner.running:
            if not messagebox.askyesno("تشغيل جارٍ", "هناك تشغيل جارٍ. إيقافه والرجوع؟", parent=self):
                return
            self.runner.kill()
        self._polling = False
        self.app.show_gallery()
