# -*- coding: utf-8 -*-
"""مركز تحكم أتمتة التراث العمراني — الواجهة (Tkinter، تصميم احترافي عصري).

تتكلّم مع المحرّك المعزول عبر مؤقّت (after) فلا تتجمّد أثناء التشغيل، وأزرارها
متصلة فعلاً بالمحرّك. تدعم ثيم فاتح/داكن، RTL عربي، وداش بورد حيّ ببطاقات مؤشرات.
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from core.engine import AutomationRunner
from core.launcher import build_command
from core import state as state_mod
from plugins.registry import discover

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(HERE, "scripts")

# ── الهوية / حقوق العمل والنشر (تظهر دائماً أسفل البرنامج) ──
APP_NAME   = "مركز تحكم أتمتة التراث العمراني"
APP_VER    = "2.0"
AUTHOR_AR  = "المهندس حمزة السلايمة"
AUTHOR_EN  = "ENG. Hamza ALsalaymeh"
COPYRIGHT  = f"© 2026  تطوير: {AUTHOR_AR}  ({AUTHOR_EN})  ·  جميع الحقوق محفوظة"

UIFONT  = "Segoe UI"      # يسقط تلقائياً لخط افتراضي إن لم يتوفّر
MONO    = "Consolas"

THEMES = {
    "light": {
        "bg": "#eef1f8", "card": "#ffffff", "fg": "#141b2d", "muted": "#64748b",
        "primary": "#4f46e5", "primaryd": "#4338ca",
        "head": "#15223b", "head2": "#1e3a8a", "headfg": "#ffffff", "headmut": "#9db4d8",
        "ok": "#059669", "err": "#e11d48", "warn": "#d97706", "accent": "#0891b2",
        "ring": "#6366f1", "ringtrack": "#e2e8f0", "panel": "#f7f9fd",
        "border": "#dbe3ef", "logbg": "#0b1220", "logfg": "#cbd5e1",
        "footer": "#e3e8f2", "footerfg": "#5a6b85", "pill": "#eef4ff",
    },
    "dark": {
        "bg": "#0f172a", "card": "#1b263b", "fg": "#e6edf6", "muted": "#94a3b8",
        "primary": "#6366f1", "primaryd": "#4f46e5",
        "head": "#0a1120", "head2": "#14213d", "headfg": "#eaf2ff", "headmut": "#7f93b4",
        "ok": "#10b981", "err": "#fb7185", "warn": "#fbbf24", "accent": "#22d3ee",
        "ring": "#818cf8", "ringtrack": "#334155", "panel": "#111c30",
        "border": "#334155", "logbg": "#060b16", "logfg": "#cbd5e1",
        "footer": "#0a1120", "footerfg": "#7c8aa5", "pill": "#1e2b44",
    },
}

LEVEL_COLORS = {"ok": "ok", "error": "err", "fail": "err", "warn": "warn",
                "skip": "muted", "preview": "primary", "info": "logfg"}


def _shade(hexc, f):
    """تفتيح/تغميق لون سداسي (f<1 أغمق، f>1 أفتح)."""
    hexc = hexc.lstrip("#")
    if len(hexc) == 3:
        hexc = "".join(c * 2 for c in hexc)
    if len(hexc) != 6:
        hexc = "808080"
    r, g, b = int(hexc[0:2], 16), int(hexc[2:4], 16), int(hexc[4:6], 16)
    r, g, b = (max(0, min(255, int(c * f))) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def _rgb(hexc):
    hexc = hexc.lstrip("#")
    if len(hexc) == 3:
        hexc = "".join(c * 2 for c in hexc)
    return int(hexc[0:2], 16), int(hexc[2:4], 16), int(hexc[4:6], 16)


def _mix(c1, c2, t):
    """امزج c1→c2 بنسبة t (0..1). يُستخدم للخلفيات الخفيفة الملوّنة."""
    a, b = _rgb(c1), _rgb(c2)
    r, g, bl = (int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return f"#{r:02x}{g:02x}{bl:02x}"


def enable_clipboard(widget):
    """يجعل النسخ/اللصق/القص يعمل على حقول الإدخال مهما كانت لغة لوحة المفاتيح
    (عربي/إنجليزي) + قائمة كليك-يمين. يُطبّق على الحقل وكل أبنائه."""
    def is_text(w):
        return isinstance(w, tk.Text)

    def do(w, ev):
        try:
            w.event_generate(ev)
        except Exception:
            pass

    def select_all(w):
        try:
            if is_text(w):
                w.tag_add("sel", "1.0", "end-1c")
            else:
                w.select_range(0, "end")
        except Exception:
            pass

    def attach(w):
        if not isinstance(w, (tk.Entry, tk.Text)):
            return
        menu = tk.Menu(w, tearoff=0)
        menu.add_command(label="قص", command=lambda: do(w, "<<Cut>>"))
        menu.add_command(label="نسخ", command=lambda: do(w, "<<Copy>>"))
        menu.add_command(label="لصق", command=lambda: do(w, "<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="تحديد الكل", command=lambda: select_all(w))

        def popup(e):
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()
            return "break"

        def on_ctrl(e):
            # نعتمد على keycode (VK على ويندوز) فيعمل مهما كانت لغة لوحة المفاتيح
            kc = getattr(e, "keycode", None)
            if kc == 86:   # V
                do(w, "<<Paste>>"); return "break"
            if kc == 67:   # C
                do(w, "<<Copy>>"); return "break"
            if kc == 88:   # X
                do(w, "<<Cut>>"); return "break"
            if kc == 65:   # A
                select_all(w); return "break"

        w.bind("<Button-3>", popup, add="+")
        w.bind("<Control-KeyPress>", on_ctrl, add="+")

    attach(widget)
    for child in widget.winfo_children():
        enable_clipboard(child)


def flat_button(parent, text, command, bg, fg, hover=None, font=(UIFONT, 10, "bold"),
                padx=16, pady=8, border=None, width=None):
    """زر مسطّح عصري مع تأثير تمرير (hover)."""
    hover = hover or _shade(bg, 0.90)
    kw = dict(text=text, command=command, relief="flat", bd=0, cursor="hand2",
              font=font, fg=fg, bg=bg, activebackground=hover, activeforeground=fg,
              padx=padx, pady=pady, highlightthickness=0)
    if width:
        kw["width"] = width
    if border:
        kw.update(highlightthickness=1, highlightbackground=border, highlightcolor=border)
    b = tk.Button(parent, **kw)
    b._bg, b._hover = bg, hover
    b.bind("<Enter>", lambda e: e.widget["state"] == "normal" and e.widget.config(bg=e.widget._hover))
    b.bind("<Leave>", lambda e: e.widget["state"] == "normal" and e.widget.config(bg=e.widget._bg))
    return b


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}")
        self.geometry("980x760")
        self.minsize(860, 640)
        self.state_data = state_mod.load()
        self.theme_name = self.state_data.get("theme", "light")
        self.plugins = discover()
        self.user_var = tk.StringVar(value=self.state_data.get("username", "H.qurah"))
        self.pw_var = tk.StringVar(value="313146")   # معبّأة، لا تُحفظ على القرص
        self._cmd_override = None                      # للاختبار فقط
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
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
        if getattr(self, "_current", "gallery") == "gallery":
            self.show_gallery()

    def _style_ttk(self):
        t = self.T
        s = self.style
        s.configure("HC.Horizontal.TProgressbar", troughcolor=_shade(t["bg"], 0.94),
                    background=t["primary"], bordercolor=t["bg"], lightcolor=t["primary"],
                    darkcolor=t["primary"], thickness=20, borderwidth=0)
        s.configure("TCombobox", fieldbackground=t["card"], background=t["card"],
                    foreground=t["fg"], arrowcolor=t["fg"], bordercolor=t["border"],
                    lightcolor=t["border"], darkcolor=t["border"])
        s.map("TCombobox", fieldbackground=[("readonly", t["card"])],
              foreground=[("readonly", t["fg"])])
        s.configure("Vertical.TScrollbar", background=t["border"], troughcolor=t["bg"],
                    bordercolor=t["bg"], arrowcolor=t["muted"])

    def apply_theme(self):
        t = self.T
        self._style_ttk()
        self.theme_btn.configure(text=("🌙  داكن" if self.theme_name == "light" else "☀️  فاتح"),
                                 bg=_shade(t["head"], 1.35), fg=t["headfg"],
                                 activebackground=_shade(t["head"], 1.5))
        self.theme_btn._bg = _shade(t["head"], 1.35)
        self.theme_btn._hover = _shade(t["head"], 1.5)
        self.configure(bg=t["bg"])
        self.header.configure(bg=t["head"])
        self.accent_strip.configure(bg=t["primary"])
        self.title_lbl.configure(bg=t["head"], fg=t["headfg"])
        self.sub_lbl.configure(bg=t["head"], fg=t["headmut"])
        for w in self.cred_widgets:
            try:
                w.configure(bg=t["head"], fg=t["headfg"])
            except tk.TclError:
                pass
        for e in self.cred_entries:
            e.configure(bg=t["pill"], fg=t["fg"], insertbackground=t["fg"],
                        highlightbackground=_shade(t["head"], 1.4), highlightcolor=t["primary"])
        self.body.configure(bg=t["bg"])
        self.footer.configure(bg=t["footer"])
        self.footer_lbl.configure(bg=t["footer"], fg=t["footerfg"])
        self.footer_brand.configure(bg=t["footer"], fg=t["primary"])

    # ---------------------------------------------------------------- البناء الأساسي
    def _build(self):
        # ===== الهيدر =====
        self.header = tk.Frame(self, height=70)
        self.header.pack(fill="x", side="top")
        self.accent_strip = tk.Frame(self, height=3)
        self.accent_strip.pack(fill="x", side="top")

        titles = tk.Frame(self.header, bg=self.T["head"])
        titles.pack(side="right", padx=16, pady=8)
        self.title_lbl = tk.Label(titles, text="🏛️  مركز تحكم أتمتة التراث العمراني",
                                  font=(UIFONT, 16, "bold"), anchor="e")
        self.title_lbl.pack(side="top", anchor="e")
        self.sub_lbl = tk.Label(titles, text="منصّة موحّدة لتشغيل ومراقبة كل الأتمتات",
                                font=(UIFONT, 9), anchor="e")
        self.sub_lbl.pack(side="top", anchor="e")

        self.theme_btn = flat_button(self.header, "🌙", self.toggle_theme,
                                     bg="#334155", fg="#fff", padx=12, pady=7)
        self.theme_btn.pack(side="left", padx=12)

        # بيانات الدخول (تُدخل مرة؛ كلمة المرور لا تُحفظ على القرص)
        self.cred_widgets = []
        self.cred_entries = []
        cf = tk.Frame(self.header)
        cf.pack(side="left", padx=4)
        self.cred_widgets.append(cf)
        for label, var, show in (("المستخدم", self.user_var, ""), ("المرور", self.pw_var, "•")):
            lb = tk.Label(cf, text=label, font=(UIFONT, 8))
            lb.pack(side="left", padx=(6, 2))
            en = tk.Entry(cf, textvariable=var, width=10, show=show, justify="center",
                          relief="flat", bd=0, highlightthickness=1, font=(UIFONT, 10))
            en.pack(side="left", padx=2, ipady=3)
            self.cred_widgets.append(lb)
            self.cred_entries.append(en)

        # ===== الفوتر (حقوق دائمة على كل الشاشات) =====
        self.footer = tk.Frame(self, height=30)
        self.footer.pack(fill="x", side="bottom")
        self.footer_brand = tk.Label(self.footer, text=f"  ●  v{APP_VER}",
                                     font=(UIFONT, 8, "bold"))
        self.footer_brand.pack(side="left", padx=8)
        self.footer_lbl = tk.Label(self.footer, text=COPYRIGHT, font=(UIFONT, 9))
        self.footer_lbl.pack(side="right", padx=14, pady=5)

        # ===== جسم المحتوى =====
        self.body = tk.Frame(self)
        self.body.pack(fill="both", expand=True, side="top")

        enable_clipboard(self.header)   # نسخ/لصق يعمل على حقلي المستخدم والمرور

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

        head = tk.Frame(self.body, bg=t["bg"])
        head.pack(fill="x", padx=22, pady=(16, 4))
        tk.Label(head, text="اختر الأتمتة المطلوبة", bg=t["bg"], fg=t["fg"],
                 font=(UIFONT, 15, "bold"), anchor="e").pack(side="right")
        tk.Label(head, text=f"{len(self.plugins)} أتمتة متاحة", bg=t["bg"], fg=t["muted"],
                 font=(UIFONT, 10), anchor="w").pack(side="left")

        wrap = tk.Frame(self.body, bg=t["bg"])
        wrap.pack(fill="both", expand=True, padx=12)
        canvas = tk.Canvas(wrap, bg=t["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=t["bg"])
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 4))
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        for p in self.plugins:
            self._gallery_card(inner, p)

    def _gallery_card(self, parent, p):
        t = self.T
        full = p.get("supports_control")
        accent = t["ok"] if full else t["accent"]

        card = tk.Frame(parent, bg=t["card"], highlightthickness=1, highlightbackground=t["border"])
        card.pack(fill="x", pady=7, padx=8)
        # شريط لوني جانبي (دلالة نوع التحكم)
        tk.Frame(card, bg=accent, width=5).pack(side="right", fill="y")

        inner = tk.Frame(card, bg=t["card"])
        inner.pack(side="right", fill="both", expand=True, padx=14, pady=12)

        top = tk.Frame(inner, bg=t["card"])
        top.pack(fill="x")
        title = tk.Label(top, text=p["title"], font=(UIFONT, 14, "bold"), bg=t["card"],
                         fg=t["fg"], anchor="e", cursor="hand2")
        title.pack(side="right", fill="x", expand=True)
        title.bind("<Button-1>", lambda e, pl=p: self.open_plugin(pl))

        badge_txt = "● تحكم كامل" if full else "○ سجل حيّ"
        badge = tk.Label(top, text=badge_txt, bg=t["pill"], fg=accent,
                         font=(UIFONT, 8, "bold"), padx=8, pady=2)
        badge.pack(side="left")
        tk.Label(top, text="؟", bg=t["pill"], fg=t["muted"], font=(UIFONT, 8, "bold"),
                 padx=7, pady=2, cursor="hand2").pack(side="left", padx=6)

        tk.Label(inner, text=p["desc"], bg=t["card"], fg=t["muted"], font=(UIFONT, 10),
                 anchor="e", justify="right", wraplength=760).pack(fill="x", pady=(6, 10))

        openb = flat_button(inner, "فتح  ▸", lambda pl=p: self.open_plugin(pl),
                            bg=t["primary"], fg="#fff", hover=t["primaryd"], padx=20, pady=7)
        openb.pack(side="left")
        helpb = flat_button(inner, "؟ مساعدة", lambda pl=p: messagebox.showinfo(
            pl["title"], pl.get("help") or pl["desc"]),
            bg=t["card"], fg=t["muted"], hover=t["pill"], border=t["border"], padx=12, pady=7)
        helpb.pack(side="left", padx=6)

        def enter(_): card.configure(highlightbackground=accent)
        def leave(_): card.configure(highlightbackground=t["border"])
        for w in (card, inner, top, title):
            w.bind("<Enter>", enter)
            w.bind("<Leave>", leave)

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

        # ---- شريط العنوان ----
        bar = tk.Frame(self, bg=t["bg"])
        bar.pack(fill="x", padx=16, pady=(12, 6))
        flat_button(bar, "⬅  رجوع", self._back, bg=t["card"], fg=t["fg"],
                    hover=t["pill"], border=t["border"], padx=14, pady=6).pack(side="left")
        tk.Label(bar, text=self.plugin["title"], bg=t["bg"], fg=t["fg"],
                 font=(UIFONT, 15, "bold"), anchor="e").pack(side="right", fill="x", expand=True)

        # ---- الإعدادات ----
        box = self._card(self, "⚙  الإعدادات")
        if self.plugin.get("self_dialogs"):
            tk.Label(box, text="هذه الأتمتة تسأل عن إعداداتها بنوافذها الخاصة عند التشغيل.",
                     bg=t["card"], fg=t["muted"], anchor="e", font=(UIFONT, 10)).pack(fill="x")
        for spec in self.plugin.get("settings", []):
            self._build_row(box, spec)
        self._apply_values(self.app.state_data.get("last", {}).get(self.plugin["id"], {}))
        self._refresh_conditional()
        if self.plugin.get("settings"):
            self._presets_bar(box)

        # ---- أزرار التحكم ----
        ctl = tk.Frame(self, bg=t["bg"])
        ctl.pack(fill="x", padx=16, pady=(2, 6))
        self.btn_start = flat_button(ctl, "▶  ابدأ", self._start, bg=t["ok"], fg="#fff",
                                     hover=_shade(t["ok"], 0.88), padx=22, pady=9,
                                     font=(UIFONT, 11, "bold"))
        self.btn_pause = flat_button(ctl, "⏸  إيقاف مؤقت", self._pause, bg=t["card"], fg=t["fg"],
                                     hover=t["pill"], border=t["border"], padx=14, pady=9)
        self.btn_resume = flat_button(ctl, "▶  استئناف", self._resume, bg=t["card"], fg=t["fg"],
                                      hover=t["pill"], border=t["border"], padx=14, pady=9)
        self.btn_stop = flat_button(ctl, "⏹  إيقاف آمن", self._stop, bg=t["err"], fg="#fff",
                                    hover=_shade(t["err"], 0.88), padx=16, pady=9)
        self.btn_pause.configure(state="disabled")
        self.btn_resume.configure(state="disabled")
        self.btn_stop.configure(state="disabled")
        for b in (self.btn_start, self.btn_pause, self.btn_resume, self.btn_stop):
            b.pack(side="right", padx=5)
        if not self.plugin.get("supports_control"):
            self.btn_pause.configure(text="⏸  (غير مدعوم)")

        # ============ الداش بورد الاحترافي التفاعلي ============
        dash = tk.Frame(self, bg=t["card"], highlightthickness=1, highlightbackground=t["border"])
        dash.pack(fill="x", padx=16, pady=(4, 6))

        # شريط علوي: العنوان + نقطة الحالة النابضة + نص الحالة
        topbar = tk.Frame(dash, bg=t["card"])
        topbar.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(topbar, text="لوحة المتابعة الحيّة", bg=t["card"], fg=t["fg"],
                 font=(UIFONT, 12, "bold"), anchor="e").pack(side="right")
        self.state_dot = tk.Canvas(topbar, width=14, height=14, bg=t["card"],
                                   highlightthickness=0)
        self.state_dot.pack(side="left")
        self._dot_id = self.state_dot.create_oval(2, 2, 12, 12, fill=t["muted"], outline="")
        self.lbl_state = tk.Label(topbar, text="جاهز", bg=t["card"], fg=t["muted"],
                                  font=(UIFONT, 11, "bold"))
        self.lbl_state.pack(side="left", padx=6)

        body = tk.Frame(dash, bg=t["card"])
        body.pack(fill="x", padx=16, pady=(0, 6))

        # يمين: حلقة التقدّم الدائرية (Canvas) مع النسبة في المنتصف
        ringwrap = tk.Frame(body, bg=t["card"])
        ringwrap.pack(side="right", padx=(10, 0))
        self.RING = 150
        self.ring = tk.Canvas(ringwrap, width=self.RING, height=self.RING,
                              bg=t["card"], highlightthickness=0)
        self.ring.pack()
        self.lbl_pct = tk.Label(ringwrap, text="0%", bg=t["card"], fg=t["ring"],
                                font=(UIFONT, 30, "bold"))
        self.lbl_pct.place(in_=self.ring, relx=0.5, rely=0.44, anchor="center")
        self.lbl_ring_sub = tk.Label(ringwrap, text="جاهز", bg=t["card"], fg=t["muted"],
                                     font=(UIFONT, 9))
        self.lbl_ring_sub.place(in_=self.ring, relx=0.5, rely=0.66, anchor="center")
        self._ring_angle = 0

        # يسار: بطاقات المؤشّرات (شبكة 2×2 + الإجمالي عريض) بخلفيات ملوّنة خفيفة
        grid = tk.Frame(body, bg=t["card"])
        grid.pack(side="right", fill="both", expand=True, padx=(0, 14))
        self.tiles = {}
        specs = [("done", "نجح", t["ok"], "✓"), ("failed", "فشل", t["err"], "✕"),
                 ("remaining", "المتبقّي", t["accent"], "⏳"), ("skipped", "تنبيه/تخطّي", t["warn"], "!"),
                 ("total", "الإجمالي", t["ring"], "∑")]
        posn = [(0, 1), (0, 0), (1, 1), (1, 0), (2, 0, 2)]
        for (spec, pos) in zip(specs, posn):
            key, label, color, icon = spec
            tint = _mix(t["card"], color, 0.12)
            cell = tk.Frame(grid, bg=tint, highlightthickness=1, highlightbackground=_mix(t["card"], color, 0.30))
            if len(pos) == 3:
                cell.grid(row=pos[0], column=pos[1], columnspan=pos[2], sticky="nsew", padx=4, pady=4)
            else:
                cell.grid(row=pos[0], column=pos[1], sticky="nsew", padx=4, pady=4)
            line = tk.Frame(cell, bg=tint)
            line.pack(fill="x", padx=12, pady=(10, 0))
            tk.Label(line, text=label, bg=tint, fg=t["muted"], font=(UIFONT, 10)).pack(side="right")
            tk.Label(line, text=icon, bg=tint, fg=color, font=(UIFONT, 12, "bold")).pack(side="left")
            val = tk.Label(cell, text="0", bg=tint, fg=color, font=(UIFONT, 24, "bold"), anchor="e")
            val.pack(fill="x", padx=12, pady=(0, 10))
            self.tiles[key] = val
        for c in (0, 1):
            grid.grid_columnconfigure(c, weight=1, uniform="k")

        # شريط تقدّم أفقي (يتحوّل «غير محدّد» أثناء جمع الروابط) + سطر المعلومات
        self.pbar = ttk.Progressbar(dash, maximum=100, value=0,
                                    style="HC.Horizontal.TProgressbar")
        self.pbar.pack(fill="x", padx=16, pady=(6, 4))
        self._indeterminate = False
        self.lbl_extra = tk.Label(dash, text="— بانتظار بدء التشغيل —", bg=t["card"],
                                  fg=t["muted"], font=(UIFONT, 10), anchor="e")
        self.lbl_extra.pack(fill="x", padx=16, pady=(0, 12))

        self._draw_ring(0)
        self._anim()   # حلقة تحريك ناعمة (نقطة نابضة + دوران أثناء التحضير)

        # ---- السجل الحيّ ----
        logcard = tk.Frame(self, bg=t["card"], highlightthickness=1, highlightbackground=t["border"])
        logcard.pack(fill="both", expand=True, padx=16, pady=(2, 12))
        tk.Label(logcard, text="📜  السجل الحيّ", bg=t["card"], fg=t["fg"],
                 font=(UIFONT, 10, "bold"), anchor="e").pack(fill="x", padx=12, pady=(8, 2))
        logf = tk.Frame(logcard, bg=t["card"])
        logf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log = tk.Text(logf, bg=t["logbg"], fg=t["logfg"], font=(MONO, 10),
                           wrap="word", height=10, state="disabled", relief="flat", bd=0,
                           padx=10, pady=8, insertbackground=t["logfg"])
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="right", fill="both", expand=True)
        for lvl, key in LEVEL_COLORS.items():
            self.log.tag_configure(lvl, foreground=self.t.get(key, self.t["logfg"]))

        enable_clipboard(self)   # نسخ/لصق (كليك يمين + Ctrl مهما كانت لغة الكيبورد) لكل الحقول

    # ---- بطاقة بعنوان ----
    def _card(self, parent, title):
        t = self.t
        outer = tk.Frame(parent, bg=t["card"], highlightthickness=1, highlightbackground=t["border"])
        outer.pack(fill="x", padx=16, pady=6)
        tk.Label(outer, text=title, bg=t["card"], fg=t["fg"],
                 font=(UIFONT, 11, "bold"), anchor="e").pack(fill="x", padx=12, pady=(10, 2))
        inner = tk.Frame(outer, bg=t["card"])
        inner.pack(fill="x", padx=12, pady=(0, 10))
        return inner

    # ---- بناء الحقول ----
    def _build_row(self, parent, spec):
        t = self.t
        name, typ = spec["name"], spec["type"]
        row = tk.Frame(parent, bg=t["card"])
        row.pack(fill="x", pady=4)
        lbl = tk.Label(row, text=spec["label"], bg=t["card"], fg=t["fg"], width=30,
                       anchor="e", justify="right", font=(UIFONT, 10))
        lbl.pack(side="right")
        if typ == "bool":
            var = tk.BooleanVar(value=spec.get("default", False))
            w = tk.Checkbutton(row, variable=var, bg=t["card"], activebackground=t["card"])
        elif typ == "choice":
            var = tk.StringVar(value=spec.get("default", spec["choices"][0]))
            w = ttk.Combobox(row, textvariable=var, values=spec["choices"], state="readonly", width=26)
            var.trace_add("write", lambda *a: self._refresh_conditional())
        elif typ == "textarea":
            var = tk.Text(row, height=4, width=28, relief="flat", bd=0,
                          highlightthickness=1, highlightbackground=t["border"])
            w = var
        elif typ in ("folder", "savefile", "openfile"):
            var = tk.StringVar()
            inner = tk.Frame(row, bg=t["card"])
            tk.Entry(inner, textvariable=var, width=22, justify="right", relief="flat", bd=0,
                     highlightthickness=1, highlightbackground=t["border"],
                     font=(UIFONT, 10)).pack(side="right", ipady=3)
            flat_button(inner, "…", lambda s=spec, v=var: self._pick(s, v),
                        bg=t["primary"], fg="#fff", hover=t["primaryd"],
                        padx=10, pady=4).pack(side="right", padx=4)
            w = inner
        else:
            var = tk.StringVar(value=str(spec.get("default", "")))
            w = tk.Entry(row, textvariable=var, width=28, justify="right", relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=t["border"], font=(UIFONT, 10))
        w.pack(side="right", padx=8, ipady=2)
        self.vars[name] = var
        self.rows[name] = row

    def _pick(self, spec, var):
        if spec["type"] == "folder":
            v = filedialog.askdirectory(parent=self)
        elif spec["type"] == "openfile":
            v = filedialog.askopenfilename(parent=self,
                                           filetypes=[("Excel", "*.xlsx *.xls"), ("الكل", "*.*")])
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
            row.pack(fill="x", pady=4) if show else row.pack_forget()

    def _presets_bar(self, parent):
        t = self.t
        bar = tk.Frame(parent, bg=t["card"])
        bar.pack(fill="x", pady=(8, 0))
        tk.Label(bar, text="إعداد محفوظ:", bg=t["card"], fg=t["muted"],
                 anchor="e", font=(UIFONT, 10)).pack(side="right")
        self._preset_var = tk.StringVar()
        names = sorted(self.app.state_data.get("presets", {}).get(self.plugin["id"], {}).keys())
        self._preset_cb = ttk.Combobox(bar, textvariable=self._preset_var, values=names,
                                       state="readonly", width=16)
        self._preset_cb.pack(side="right", padx=4)
        flat_button(bar, "حمّل", self._load_preset, bg=t["card"], fg=t["fg"],
                    hover=t["pill"], border=t["border"], padx=10, pady=4).pack(side="right", padx=2)
        flat_button(bar, "احفظ", self._save_preset, bg=t["card"], fg=t["fg"],
                    hover=t["pill"], border=t["border"], padx=10, pady=4).pack(side="right", padx=2)

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
        if int(self.log.index("end-1c").split(".")[0]) > 1500:
            self.log.delete("1.0", "500.0")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---- التشغيل والتحكم ----
    def _start(self):
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
        self.lbl_state.configure(text="⏹ جاري الإيقاف الآمن…")

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

    # ---------- رسم الحلقة الدائرية ----------
    def _draw_ring(self, pct, spin_angle=None):
        try:
            c = self.ring
        except Exception:
            return
        c.delete("all")
        t = self.t
        W = 15
        m = W / 2 + 3
        R = self.RING
        box = (m, m, R - m, R - m)
        # المسار (خلفية الحلقة)
        c.create_arc(*box, start=0, extent=359.9, style="arc", width=W, outline=t["ringtrack"])
        if spin_angle is not None:
            # وضع «جاري التحضير»: قوس قصير يدور
            c.create_arc(*box, start=spin_angle, extent=-90, style="arc", width=W, outline=t["ring"])
        else:
            ext = -(max(0.0, min(100.0, pct)) / 100.0 * 360.0)
            if ext:
                c.create_arc(*box, start=90, extent=ext, style="arc", width=W, outline=t["ring"])

    # ---------- حلقة التحريك الناعمة ----------
    def _anim(self):
        if not self.winfo_exists():
            return
        self._ring_angle = (self._ring_angle + 11) % 360
        running = self.runner.running
        collecting = running and self.runner.stats.get("total", 0) == 0
        # نقطة الحالة النابضة
        try:
            import math
            self._pulse = getattr(self, "_pulse", 0) + 1
            base = getattr(self, "_state_color", self.t["muted"])
            if running:
                k = 0.78 + 0.22 * (0.5 + 0.5 * math.sin(self._pulse / 3.0))
                self.state_dot.itemconfigure(self._dot_id, fill=_shade(base, k))
            else:
                self.state_dot.itemconfigure(self._dot_id, fill=base)
        except Exception:
            pass
        if collecting:
            self._draw_ring(0, spin_angle=self._ring_angle)
        self.after(90, self._anim)

    def _update_monitor(self):
        t = self.t
        s = self.runner.stats
        rem = self.runner.remaining()
        running = self.runner.running
        collecting = running and s["total"] == 0 and s["state"] == "running"

        # الشريط الأفقي: غير-محدّد أثناء التحضير، ومحدّد بعد معرفة الإجمالي
        if collecting and not self._indeterminate:
            self.pbar.configure(mode="indeterminate"); self.pbar.start(16); self._indeterminate = True
        elif not collecting and self._indeterminate:
            self.pbar.stop(); self.pbar.configure(mode="determinate"); self._indeterminate = False
        if not collecting:
            self.pbar.configure(value=s["percent"])
            self.lbl_pct.configure(text=f"{s['percent']:.0f}%")
            self._draw_ring(s["percent"])
            self.lbl_ring_sub.configure(text=f"{s['done']+s['failed']+s['skipped']:,} / {s['total']:,}"
                                        if s["total"] else "—")
        else:
            self.lbl_pct.configure(text="…")
            self.lbl_ring_sub.configure(text="تحضير")

        vals = {"total": s["total"], "done": s["done"], "failed": s["failed"],
                "skipped": s["skipped"], "remaining": rem}
        for k, lbl in self.tiles.items():
            lbl.configure(text=f"{vals[k]:,}")

        def fmt(x):
            x = int(x)
            h, rem2 = divmod(x, 3600)
            m, sec = divmod(rem2, 60)
            return f"{h}:{m:02d}:{sec:02d}" if h else f"{m}:{sec:02d}"

        if collecting:
            self.lbl_extra.configure(
                text="⏳ جاري تسجيل الدخول وجمع روابط المواقع… اللوحة تبدأ بالتحرّك بعد انتهاء الجمع")
        else:
            el = s["elapsed"]
            eta = int(s["avg"] * rem) if (s["avg"] and rem) else 0
            self.lbl_extra.configure(
                text=f"📍 {s['current'] or '—'}      ⏱ متوسط: {s['avg']:.1f}ث      "
                     f"⏲ مضى: {fmt(el)}      ⏳ متبقٍّ ≈ {fmt(eta)}")

        states = {"running": ("شغّال", t["ok"]), "paused": ("متوقّف مؤقتاً", t["warn"]),
                  "stopping": ("جاري الإيقاف", t["warn"]), "stopped": ("أُوقف", t["err"]),
                  "finished": ("انتهى ✅", t["ok"]), "idle": ("جاهز", t["muted"])}
        if collecting:
            txt, col = "جاري التحضير…", t["accent"]
        else:
            txt, col = states.get(s["state"], (s["state"], t["muted"]))
        self._state_color = col
        self.lbl_state.configure(text=txt, fg=col)
        self.lbl_ring_sub.configure(fg=col if collecting else t["muted"])

    def _on_finish(self, ev):
        self._update_monitor()
        self.btn_start.configure(state="normal")
        for b in (self.btn_pause, self.btn_resume, self.btn_stop):
            b.configure(state="disabled")
        msg = "✅ انتهى التشغيل بنجاح" if ev["state"] == "finished" else "⏹ تم الإيقاف — التقدّم محفوظ ويمكن الاستئناف"
        self._append_log(msg, "ok" if ev["state"] == "finished" else "warn")
        self.runner.cleanup()

    def _back(self):
        if self.runner.running:
            if not messagebox.askyesno("تشغيل جارٍ", "هناك تشغيل جارٍ. إيقافه والرجوع؟", parent=self):
                return
            self.runner.kill()
        self._polling = False
        self.app.show_gallery()
