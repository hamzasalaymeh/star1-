#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Heritage historical village description automation — PRO edition.

نسخة احترافية مطوّرة من السكربت الأصلي. المميزات:
  • تشغيل ذاتي كامل: ينشئ متصفحه بنفسه ويسجّل الدخول تلقائياً من بيانات الاعتماد.
  • يعالج أكثر من قرار في تشغيل واحد (الافتراضي: قرار 25 ثم 26 لمنطقة عسير).
  • استئناف ذكي: لا يبدأ من الصفر لو توقّف — يتخطّى ما أُنجز (checkpoint + Excel) والمطابق مسبقاً.
  • صمود: يعيد تسجيل الدخول لو انتهت الجلسة، ويعيد تشغيل المتصفح لو انهار، ويكمل من مكانه.
  • حارس جودة: لا يكتب وصفاً فيه نص نائب (XXXX) أو قرية/اسم خاطئ — يتخطّى المبنى بدلاً من ذلك.
  • بانر حيّ أثناء العمل باسم: ENG. Hamza ALsalaymeh + شريط تقدّم وعدّادات لحظية.
  • إكسل منسّق: رأس مجمّد، تلوين شرطي للحالة، التفاف نص، فلتر تلقائي، ورقتا RUN_LOG و SUMMARY.

المؤلف: ENG. Hamza ALsalaymeh
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import random
import re
import subprocess
import sys
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime
from getpass import getpass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ------------------------------------------------------------------ Selenium
try:
    from selenium import webdriver
    from selenium.common.exceptions import (
        StaleElementReferenceException,
        TimeoutException,
        WebDriverException,
    )
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.selenium_manager import SeleniumManager
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import Select, WebDriverWait
except ImportError:  # pragma: no cover
    webdriver = None
    StaleElementReferenceException = TimeoutException = WebDriverException = Exception
    ChromeOptions = ChromeService = EdgeOptions = EdgeService = By = EC = Select = None
    SeleniumManager = WebDriverWait = None

# ------------------------------------------------------------------ Rich (optional, graceful fallback)
try:
    from rich.console import Console, Group
    from rich.live import Live
    from rich.panel import Panel
    from rich.text import Text

    _RICH = True
except Exception:  # pragma: no cover
    _RICH = False
    Console = Group = Live = Panel = Text = None

# ------------------------------------------------------------------ Arabic display shaping
# الطرفيات على ويندوز لا تُرتّب العربية ثنائية الاتجاه؛ نعيد تشكيلها وترتيبها للعرض فقط.
# هذا يخص الطباعة على الشاشة حصرًا — النص المكتوب على المنصة وفي الإكسل يبقى بترتيبه المنطقي.
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    _ARABIC_SHAPING = True
except Exception:
    _ARABIC_SHAPING = False

import builtins as _builtins

_ARABIC_CHARS = re.compile(r"[؀-ۿ]")
_ORIG_PRINT = _builtins.print


def shape(text):
    """يعيد تشكيل وترتيب نص عربي لعرضه صحيحًا في طرفية لا تدعم bidi. آمن لأي نص آخر."""
    if not _ARABIC_SHAPING or not isinstance(text, str) or not _ARABIC_CHARS.search(text):
        return text
    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


def print(*args, **kwargs):  # noqa: A001 - يُظلّل print المدمج داخل هذا الملف فقط
    _ORIG_PRINT(*[shape(a) if isinstance(a, str) else a for a in args], **kwargs)


# =================================================================== Constants
BASE_URL = "https://heritage2.itqan-consultant.com/Web/"
BASE_VIEW = "https://heritage2.itqan-consultant.com/Web/App/Pools/DataView"
BASE_EDIT = "https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/"
LOGIN_URL = "https://heritage2.itqan-consultant.com/Web/App/Home/Login"

DEBUG_PORT = 9222
DEFAULT_AUTOMATION_PROFILE = r"C:\heritage_chrome_profile"
STEP_VALUE = "410"
DELAY_BETWEEN = 1.2
IMPLICIT_WAIT = 8

CONFIG_FILE = Path("excel_path_config.json")
CREDENTIALS_FILE = Path(".env")

# بيانات الدخول الافتراضية حتى يشتغل السكربت لوحده بدون أي إعداد.
# الأفضل أمنياً وضعها في ملف .env (عبر الأمر write-env)، وعدم مشاركة هذا الملف أو رفعه على GitHub.
DEFAULT_USERNAME = "H.qurah"
DEFAULT_PASSWORD = "313146"

LOGIN_USERNAME_FIELD_NAME = "ctl12$ctl03"
LOGIN_PASSWORD_FIELD_NAME = "ctl12$ctl07"
LOGIN_BUTTON_ID = "ctl12_btnLogin"
AUTH_MARKER_ID = "TopBar_lblName"

AUTHOR = "ENG. Hamza ALsalaymeh"
DEFAULT_REGION = "عسير"
DEFAULT_DECISIONS = [25, 26]
DEFAULT_EXCEL_NAME = "نتائج_الوصف_التاريخي.xlsx"

FLOOR_PHRASES = {
    1: "دور واحد",
    2: "دورين",
    3: "ثلاثة أدوار",
    4: "أربعة أدوار",
    5: "خمسة أدوار",
    6: "ستة أدوار",
    7: "سبعة أدوار",
    8: "ثمانية أدوار",
    9: "تسعة أدوار",
    10: "عشرة أدوار",
}
ALL_FLOOR_PATTERNS = sorted(FLOOR_PHRASES.values(), key=len, reverse=True)

EXCEL_COLUMNS = [
    "الكود",
    "اسم المبنى",
    "المنطقة",
    "المحافظة",
    "القرية (بحث)",
    "رقم القرار",
    "عدد الأدوار",
    "الحالة",
    "السبب",
    "الوصف المكتوب",
    "النص السابق",
    "وقت المعالجة",
]
STATUS_COL = EXCEL_COLUMNS.index("الحالة") + 1  # 1-based

RUN_LOG_COLUMNS = [
    "Run ID",
    "وقت الحدث",
    "المرحلة",
    "المنطقة",
    "رقم القرار",
    "الوضع",
    "صفحة النتائج",
    "ترتيب المبنى",
    "إجمالي المباني",
    "الكود",
    "الحالة",
    "التفاصيل",
]

SUMMARY_COLUMNS = ["المؤشر", "القيمة"]

TEMPLATES = {
    "عسير": """تمتاز منطقة عسير بقراها التاريخية وتراثها الفريد، حيث ما زالت هذه القرى قائمة منذ مئات السنين، وتتفاوت في أحجامها ونمطها العمراني، وتتفق على ما تعكسه من إرث ثقافي وإنساني، وما اشتملت عليه من موروثات فنية وثقافية اجتماعية حملت في طيّاتها دلالات حضارة هذه المنطقة على مر العصور.
ومن ضمن هذه المواقع مبنى رقم (16) في قرية القارية، حيث يتسم ببساطته ومواءمته للبيئة المحلية والعمارة التقليدية، ويتألف من دورين، وتكاد تخلو واجهاته الرئيسة من الزخارف، مما يعكس الطابع المحلي والاجتماعي السائد في أساليب البناء خلال تلك الفترة، حيث شُيِّد المبنى باستخدام مواد طبيعية محلية، ويُصنَّف المبنى حاليًا ضمن المباني ذات الحالة الإنشائية المتوسطة، مما يستلزم جهودًا للحفاظ عليه كأحد الشواهد الباقية على التراث المعماري لمحافظة أبها.""",
}


# =================================================================== Logging
LOGGER = logging.getLogger("heritage")


def setup_logging(log_path: Path) -> None:
    LOGGER.setLevel(logging.INFO)
    LOGGER.handlers.clear()
    try:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        LOGGER.addHandler(handler)
    except Exception:
        pass


# =================================================================== Live UI
class UI:
    """بانر حيّ أثناء العمل. يعمل بـ rich إن توفّر، ويرجع تلقائياً لنص عادي إن لم يتوفّر."""

    def __init__(self, enabled: bool = True):
        self.rich = bool(_RICH and enabled)
        self.console = Console() if self.rich else None
        self.live = None
        self.events: deque = deque(maxlen=7)
        self.state: Dict[str, object] = {
            "decisions": [],
            "current_decision": "",
            "mode": "",
            "total": 0,
            "index": 0,
            "code": "",
            "ok": 0,
            "fail": 0,
            "skip": 0,
            "preview": 0,
            "start": time.time(),
        }

    # -- lifecycle -------------------------------------------------
    def start(self) -> None:
        if self.rich:
            self.live = Live(
                _BannerRenderable(self),
                console=self.console,
                refresh_per_second=4,
                redirect_stdout=True,
                redirect_stderr=True,
                screen=False,
            )
            self.live.start()
        else:
            self._print_plain_header()

    def stop(self) -> None:
        if self.rich and self.live is not None:
            try:
                self.live.update(_BannerRenderable(self))
                self.live.stop()
            except Exception:
                pass

    # -- state updates ---------------------------------------------
    def set_run(self, decisions: List[int], mode: str) -> None:
        self.state["decisions"] = list(decisions)
        self.state["mode"] = mode

    def set_decision(self, decision: object, total: int) -> None:
        self.state["current_decision"] = str(decision)
        self.state["total"] = total
        self.state["index"] = 0
        self.state["code"] = ""

    def set_progress(self, index: int, code: str) -> None:
        self.state["index"] = index
        self.state["code"] = code
        self.tick()

    def incr(self, status: str) -> None:
        if status == "نجح":
            self.state["ok"] = int(self.state["ok"]) + 1
        elif status == "فشل":
            self.state["fail"] = int(self.state["fail"]) + 1
        elif status == "معاينة":
            self.state["preview"] = int(self.state["preview"]) + 1
        else:
            self.state["skip"] = int(self.state["skip"]) + 1

    # -- logging ---------------------------------------------------
    _STYLE = {
        "ok": ("✓", "bold green"),
        "fail": ("✗", "bold red"),
        "skip": ("⏭", "yellow"),
        "warn": ("⚠", "bold yellow"),
        "info": ("•", "white"),
        "preview": ("👁", "cyan"),
    }

    def log(self, message: str, level: str = "info") -> None:
        symbol, style = self._STYLE.get(level, ("•", "white"))
        line = f"{symbol} {message}"
        LOGGER.info("[%s] %s", level, message)
        if self.rich:
            self.events.append((line, style))
        else:
            print(line, flush=True)

    def tick(self) -> None:
        if self.rich:
            return
        s = self.state
        print(
            f"══ {AUTHOR} ══ {DEFAULT_REGION} ══ قرار {s['current_decision']} ══ "
            f"[{s['index']}/{s['total']}] ✓{s['ok']} ✗{s['fail']} ⏭{s['skip']} ══",
            flush=True,
        )

    def _print_plain_header(self) -> None:
        bar = "═" * 64
        print(bar, flush=True)
        print(f"   {AUTHOR}   |   أتمتة أوصاف القرى التاريخية — {DEFAULT_REGION}", flush=True)
        print(bar, flush=True)

    # -- rich rendering --------------------------------------------
    def render_group(self):
        s = self.state
        elapsed = int(time.time() - float(s["start"]))

        name = Text(f"🏛  {AUTHOR}", style="bold cyan", justify="center")
        subtitle = Text(shape("أتمتة أوصاف القرى التاريخية — منطقة عسير"), style="white", justify="center")
        info = Text(
            shape(
                f"القرار الحالي: {s['current_decision'] or '—'}    |    "
                f"القرارات: {', '.join(map(str, s['decisions'])) or '—'}    |    "
                f"الوضع: {s['mode'] or '—'}"
            ),
            style="yellow",
            justify="center",
        )
        banner = Panel(
            Group(name, subtitle, info),
            border_style="cyan",
            title=shape("⚙  العمل جارٍ"),
            title_align="right",
        )

        total = int(s["total"])
        index = int(s["index"])
        pct = (index / total * 100.0) if total else 0.0
        bar_len = 32
        filled = int(bar_len * pct / 100.0)
        bar = "█" * filled + "░" * (bar_len - filled)

        prog = Text()
        prog.append(f" [{index}/{total}] ", style="bold white")
        prog.append(bar + " ", style="green")
        prog.append(f"{pct:5.1f}%   ", style="bold green")
        prog.append(f"✓ {s['ok']}  ", style="bold green")
        prog.append(f"✗ {s['fail']}  ", style="bold red")
        prog.append(f"⏭ {s['skip']}  ", style="bold yellow")
        if int(s["preview"]):
            prog.append(f"👁 {s['preview']}  ", style="bold cyan")
        prog.append(f"  ⏱ {elapsed // 60}:{elapsed % 60:02d}   ", style="white")
        prog.append(shape(f"الكود: {s['code'] or '—'}"), style="magenta")

        if self.events:
            body = Group(*[Text(shape(txt), style=style) for txt, style in self.events])
        else:
            body = Text(shape("بانتظار بدء المعالجة…"), style="grey58")
        events = Panel(body, border_style="grey37", title=shape("آخر الأحداث"), title_align="right")

        return Group(banner, prog, events)


if _RICH:

    class _BannerRenderable:
        def __init__(self, ui: "UI"):
            self.ui = ui

        def __rich__(self):
            return self.ui.render_group()
else:  # pragma: no cover
    class _BannerRenderable:  # type: ignore
        def __init__(self, ui: "UI"):
            self.ui = ui


# =================================================================== Data classes
@dataclass
class ProcessResult:
    code: str
    name: str = ""
    gov: str = ""
    floors: object = ""
    old: str = ""
    status: str = "فشل"
    reason: str = ""
    new: str = ""


@dataclass
class CollectionResult:
    buildings: List[Tuple[str, str]]
    last_page: int = 1
    pages_collected: int = 1
    stopped_reason: str = ""


# =================================================================== Text builders
def floors_to_arabic(n: int) -> str:
    return FLOOR_PHRASES.get(int(n), f"{int(n)} أدوار")


def replace_building_name(text: str, new_name: str) -> str:
    pattern = r"(?:ال)?مبن[ىي]\s+رقم\s*\(?\s*\d+\s*\)?\s*(?:\"[^\"]*\"\s*)?في\s+قرية\s+[^،.\n]+"
    if re.search(pattern, text):
        return re.sub(pattern, lambda _m: new_name.strip(), text, count=1)
    return text


def replace_floors(text: str, floors_count: int) -> str:
    correct = floors_to_arabic(floors_count)
    variants = sorted(set(ALL_FLOOR_PATTERNS + ["دوار واحد", "دوار واحدة"]), key=len, reverse=True)
    for phrase in variants:
        if phrase in text:
            return text.replace(phrase, correct, 1)
    return text


def replace_governorate(text: str, governorate: str) -> str:
    return re.sub(r"محافظة\s+[^،.\n]+", f"محافظة {governorate.strip()}", text)


def bracket_building_number(name: str) -> str:
    def repl(match: "re.Match[str]") -> str:
        return f"{match.group(1)} ({match.group(2)})"

    # يقبل «مبنى/مبني» و«ميني» (خطأ إملائي شائع) و«موقع» (آبار/مساجد/حصون).
    return re.sub(r"((?:مبن[ىي]|ميني|موقع)\s+رقم)\s*\(?\s*(\d+)\s*\)?", repl, name, count=1)


def normalize_site_name(name: str) -> str:
    """يصحّح «ميني رقم N» إلى «مبنى رقم N» (خطأ إملائي على المنصة). «موقع» يبقى كما هو."""
    return re.sub(r"ميني(?=\s*رقم)", "مبنى", name or "")


def remove_floors_clause(text: str) -> str:
    """يحذف جملة الأدوار («ويتألف من ...») بالكامل للمواقع التي لا أدوار لها (مثل الآبار)."""
    new = re.sub(r"\s*[،,]\s*و?يتألف من[^،,.]*", "", text, count=1)
    return new if new != text else text


def build_description(template: str, building_name: str, floors_count: int, governorate: str) -> str:
    text = template.strip()
    text = replace_building_name(text, bracket_building_number(building_name.strip()))
    text = replace_floors(text, floors_count)
    text = replace_governorate(text, governorate)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{2,}", "\n", text).strip()


def build_description_safe(
    template: str, building_name: str, floors_count, governorate: str, loose: bool = False
) -> Tuple[str, bool, str]:
    """يبني الوصف مع ضمان عدم كتابة نص خاطئ أو غير قياسي.

    يرجّع (النص, نجح؟, السبب). إذا كان اسم المبنى غير قياسي (لا يحتوي "مبنى" ورقمًا)
    أو بقي نص نائب، يُرجع نجح=False ويُترك المبنى للمراجعة اليدوية بدل كتابة وصف غير دقيق.

    الوضع المرن (loose=True): يقبل أيضًا الأسماء بصيغة «موقع رقم N» و«ميني رقم N»
    (آبار/مساجد/حصون)، ويكتب الوصف بدون جملة الأدوار إذا كانت خانة الأدوار فاضية
    أو غير رقمية — لأن هذه المواقع أصلًا بلا أدوار.
    """
    base = template.strip()
    name_clean = (building_name or "").strip()
    if not name_clean:
        return "", False, "اسم المبنى فارغ على المنصة — تُرك للمراجعة اليدوية"

    # الوضع القياسي: يحتوي كلمة "مبنى/مبني" ورقمًا. الوضع المرن: يقبل «موقع/ميني رقم N».
    standard = bool(re.search(r"مبن[ىي]", name_clean) and re.search(r"\d", name_clean))
    loose_ok = bool(
        loose and re.search(r"(?:مبن[ىي]|ميني|موقع)\s*رقم", name_clean) and re.search(r"\d", name_clean)
    )
    if not (standard or loose_ok):
        return (
            "",
            False,
            "اسم المبنى غير قياسي (لا يطابق صيغة «مبنى رقم N») — تُرك للمراجعة اليدوية",
        )

    if not governorate.strip():
        return "", False, "المحافظة فارغة على المنصة — تُرك للمراجعة اليدوية"

    bracketed = bracket_building_number(normalize_site_name(name_clean))
    text = replace_building_name(base, bracketed)
    if isinstance(floors_count, int):
        text = replace_floors(text, floors_count)
    elif loose:
        # موقع بلا أدوار (بئر/…): نكتب الوصف بدون ذكر عدد الأدوار بدل الفشل.
        text = remove_floors_clause(text)
    else:
        return "", False, "خانة الأدوار غير رقمية — تُرك للمراجعة اليدوية"
    text = replace_governorate(text, governorate)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text).strip()

    if "XXXX" in text or "xxxx" in text.lower():
        return "", False, "بقي نص نائب (XXXX) في الوصف — تُرك للمراجعة اليدوية"
    return text, True, ""


def validate_template(
    template: str,
    sample_name: str = "مبنى رقم 99 في قرية الاختبار",
    sample_floors: int = 3,
    sample_gov: str = "الاختبار",
    verbose: bool = True,
) -> bool:
    checks = [
        ("اسم المبنى", replace_building_name(template, sample_name) != template),
        ("الأدوار", replace_floors(template, sample_floors) != template),
        ("المحافظة", replace_governorate(template, sample_gov) != template),
    ]
    if verbose:
        print("=" * 70)
        print(">> فحص القالب:")
        for label, ok in checks:
            print(f"   {label}: {'مسك' if ok else 'ما مسك'}")
        print("=" * 70)
    return all(ok for _, ok in checks)


# =================================================================== Browser
def launch_chrome(
    chrome_path: Optional[str] = None,
    profile: str = DEFAULT_AUTOMATION_PROFILE,
    port: int = DEBUG_PORT,
) -> None:
    profile_path = Path(profile)
    profile_path.mkdir(parents=True, exist_ok=True)
    candidates = [chrome_path] if chrome_path else [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "chrome",
        "google-chrome",
    ]
    cmd = None
    for candidate in candidates:
        if candidate in {"chrome", "google-chrome"} or Path(candidate).exists():
            cmd = [
                candidate,
                f"--remote-debugging-port={port}",
                f"--user-data-dir={profile_path}",
                "--start-maximized",
                BASE_URL,
            ]
            break
    if not cmd:
        raise RuntimeError("لم أجد Chrome. مرر المسار عبر --chrome-path.")
    subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print(">> فتحت نافذة Chrome للأتمتة.")


def ensure_selenium() -> None:
    if webdriver is None or ChromeOptions is None or EdgeOptions is None:
        raise RuntimeError("Selenium غير مثبت. شغّل: pip install -r requirements_historical_villages.txt")


def load_env_file(path: Path = CREDENTIALS_FILE) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def write_env_file(username: str, password: str, path: Path = CREDENTIALS_FILE) -> None:
    path.write_text(
        f"HERITAGE_USERNAME={username.strip()}\nHERITAGE_PASSWORD={password.strip()}\n",
        encoding="utf-8",
    )
    print(f">> تم حفظ بيانات الدخول محليًا في {path}. لا ترفع هذا الملف ولا تشاركه.")


def get_credentials(args: argparse.Namespace, interactive: bool = False) -> Tuple[str, str]:
    load_env_file()
    username = getattr(args, "username", None) or os.environ.get("HERITAGE_USERNAME") or DEFAULT_USERNAME
    password = getattr(args, "password", None) or os.environ.get("HERITAGE_PASSWORD") or DEFAULT_PASSWORD
    if interactive and (not username or not password):
        username = username or input("اسم المستخدم: ").strip()
        password = password or getpass("كلمة المرور: ").strip()
    if not username or not password:
        raise RuntimeError("بيانات الدخول ناقصة.")
    return username, password


def managed_browser_service(browser: str):
    try:
        paths = SeleniumManager().binary_paths(["--browser", browser])
        driver_path = paths.get("driver_path")
        if driver_path:
            service_type = EdgeService if browser == "edge" else ChromeService
            return service_type(executable_path=driver_path)
    except Exception:
        pass
    return EdgeService() if browser == "edge" else ChromeService()


def create_background_driver(headless: bool = True, profile: Optional[str] = None, browser: str = "edge"):
    ensure_selenium()
    browser = browser.lower()
    if browser not in {"edge", "chrome"}:
        raise ValueError("المتصفح المدعوم هو edge أو chrome فقط.")
    options = EdgeOptions() if browser == "edge" else ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.page_load_strategy = "eager"
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-component-update")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option("useAutomationExtension", False)
    if profile:
        profile_path = Path(profile)
        profile_path.mkdir(parents=True, exist_ok=True)
        options.add_argument(f"--user-data-dir={profile_path}")

    last_error = None
    for attempt in range(1, 3):
        driver = None
        try:
            builder = webdriver.Edge if browser == "edge" else webdriver.Chrome
            driver = builder(service=managed_browser_service(browser), options=options)
            driver.implicitly_wait(IMPLICIT_WAIT)
            driver.set_page_load_timeout(45)
            driver.execute_script("return navigator.userAgent")
            return driver
        except Exception as exc:
            last_error = exc
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            if attempt == 1:
                print(f">> تعذر بدء {browser.capitalize()} Driver، سأعيد المحاولة بدرايفر متوافق...")
                time.sleep(3)
    raise RuntimeError(
        f"تعذر تشغيل {browser.capitalize()} Driver بشكل مستقر. حدّث Selenium ثم أعد التشغيل. "
        f"التفصيل: {type(last_error).__name__}: {last_error}"
    )


def get_attached_driver(port: int = DEBUG_PORT):
    ensure_selenium()
    options = ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    return webdriver.Chrome(options=options)


# =================================================================== Login / session
def quick_present(driver, by, value) -> bool:
    """فحص سريع لوجود عنصر بدون انتظار ضمني طويل."""
    try:
        driver.implicitly_wait(0)
    except Exception:
        pass
    try:
        return len(driver.find_elements(by, value)) > 0
    except WebDriverException:
        return False
    finally:
        try:
            driver.implicitly_wait(IMPLICIT_WAIT)
        except Exception:
            pass


def is_logged_in(driver) -> bool:
    return quick_present(driver, By.ID, AUTH_MARKER_ID)


def login(driver, username: str, password: str) -> bool:
    driver.get(LOGIN_URL)
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, AUTH_MARKER_ID)))
        return True
    except TimeoutException:
        pass

    user_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.NAME, LOGIN_USERNAME_FIELD_NAME))
    )
    password_input = driver.find_element(By.NAME, LOGIN_PASSWORD_FIELD_NAME)
    user_input.clear()
    user_input.send_keys(username)
    password_input.clear()
    password_input.send_keys(password)
    driver.find_element(By.ID, LOGIN_BUTTON_ID).click()
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, AUTH_MARKER_ID)))
    return True


# =================================================================== Search / collect
def wait_postback(driver, element, timeout: int = 10) -> None:
    try:
        WebDriverWait(driver, timeout).until(EC.staleness_of(element))
    except TimeoutException:
        pass
    time.sleep(0.8)


def select_by_text(driver, name: str, text: str):
    element = driver.find_element(By.NAME, name)
    select = Select(element)
    target = text.strip()
    for option in select.options:
        if (option.text or "").strip() == target:
            select.select_by_value(option.get_attribute("value"))
            return element
    raise ValueError(f"ما لقيت '{target}' بالقائمة {name}")


def do_search(driver, region: str, registration_decision: str) -> None:
    # نبدأ دائمًا من صفحة بحث نظيفة حتى لا تتأثر نتائج قرار بما قبله (ترقيم صفحات قديم).
    print(">> فتح صفحة بحث نظيفة DataView...")
    driver.get(BASE_VIEW)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "ctl12$ctl08")))
    time.sleep(1)

    print(f">> اختيار المنطقة: {region}")
    region_select = select_by_text(driver, "ctl12$ctl08", region)
    wait_postback(driver, region_select)

    print(">> اختيار الخطوة: مراجعة بيانات الميدان 2")
    step_select = driver.find_element(By.NAME, "ctl12$ctl32")
    Select(step_select).select_by_value(STEP_VALUE)
    wait_postback(driver, step_select)

    print(">> تفريغ حقل القرية/الكود")
    village_input = driver.find_element(By.NAME, "ctl12$ctl04")
    village_input.clear()

    if registration_decision.strip():
        print(f">> إدخال رقم القرار: {registration_decision.strip()}")
        decision_input = driver.find_element(By.NAME, "ctl12$ctl28")
        decision_input.clear()
        decision_input.send_keys(registration_decision.strip())

    print(">> الضغط على بحث...")
    search_button = driver.find_element(By.ID, "ctl12_btnSearch")
    search_button.click()
    try:
        WebDriverWait(driver, 15).until(EC.staleness_of(search_button))
    except TimeoutException:
        pass
    time.sleep(1.2)
    print(">> انتهى البحث، جاري جمع النتائج...")


def page_buildings(driver) -> List[Tuple[str, str]]:
    output: List[Tuple[str, str]] = []
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        table = max(tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")), default=None)
    except Exception:
        return output
    if not table:
        return output
    try:
        rows = table.find_elements(By.TAG_NAME, "tr")
    except Exception:
        return output
    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) < 2:
                continue
            code = cells[0].text.strip()
            name = cells[1].text.strip()
            if code.isdigit() and len(code) >= 4:
                output.append((code, name))
        except Exception:
            continue
    return output


def page_signature(driver) -> Tuple[str, ...]:
    codes: List[str] = []
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        table = None
        best = -1
        for candidate in tables:
            try:
                rows_count = len(candidate.find_elements(By.TAG_NAME, "tr"))
            except StaleElementReferenceException:
                continue
            if rows_count > best:
                best = rows_count
                table = candidate
        if table is None:
            return tuple()
        for row in table.find_elements(By.TAG_NAME, "tr"):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                code = cells[0].text.strip() if cells else ""
                if code.isdigit() and len(code) >= 4:
                    codes.append(code)
                    if len(codes) >= 5:
                        break
            except StaleElementReferenceException:
                continue
    except Exception:
        return tuple()
    return tuple(codes)


def postback_target_for_pager(driver) -> Optional[str]:
    try:
        anchors = driver.find_elements(By.TAG_NAME, "a")
    except Exception:
        return None
    for anchor in anchors:
        try:
            href = anchor.get_attribute("href") or ""
        except Exception:
            continue
        match = re.search(r"__doPostBack\('([^']+)','Page\$\d+'\)", href)
        if match:
            return match.group(1)
    return None


def goto_page_once(driver, page_number: int) -> bool:
    old_signature = page_signature(driver)
    if not old_signature:
        return False
    target_link = None
    event_target = postback_target_for_pager(driver)
    try:
        anchors = driver.find_elements(By.TAG_NAME, "a")
    except Exception:
        return False
    for anchor in anchors:
        try:
            href = anchor.get_attribute("href") or ""
        except Exception:
            continue
        if re.search(rf"Page\${page_number}\b", href):
            target_link = anchor
            break

    def wait_for_change() -> bool:
        for _ in range(35):
            time.sleep(0.6)
            current_signature = page_signature(driver)
            if current_signature and current_signature != old_signature:
                return True
        return False

    if target_link is not None:
        try:
            driver.execute_script("arguments[0].click();", target_link)
            if wait_for_change():
                return True
        except Exception:
            pass

    if not event_target:
        return False
    try:
        print(f"   (استخدام PostBack مباشر للصفحة {page_number})")
        driver.execute_script("__doPostBack(arguments[0], arguments[1]);", event_target, f"Page${page_number}")
    except Exception:
        return False
    return wait_for_change()


def goto_page(driver, page_number: int, attempts: int = 3) -> bool:
    for attempt in range(1, attempts + 1):
        if goto_page_once(driver, page_number):
            return True
        if attempt < attempts:
            print(f"   !! محاولة انتقال {attempt} فشلت للصفحة {page_number}، إعادة المحاولة...")
            time.sleep(2.0)
    return False


def collect_all_buildings(driver, max_pages: int = 50, progress_callback=None) -> CollectionResult:
    seen: Set[str] = set()
    result: List[Tuple[str, str]] = []
    last_page = 1
    pages_collected = 0
    stopped_reason = ""

    def add_current_page(page_number: int) -> Tuple[int, int]:
        nonlocal last_page, pages_collected
        before = len(result)
        page_rows = page_buildings(driver)
        for code, name in page_rows:
            if code not in seen:
                seen.add(code)
                result.append((code, name))
        page_count = len(page_rows)
        added = len(result) - before
        last_page = page_number
        pages_collected += 1
        print(f">> صفحة النتائج {page_number}: ظاهر {page_count} | جديد {added} | الإجمالي {len(result)}")
        if progress_callback:
            progress_callback(page_number, added, len(result), f"تم جمع الصفحة | ظاهر: {page_count}")
        return page_count, added

    page_count, _ = add_current_page(1)
    if page_count < 25:
        print(">> هذه آخر صفحة نتائج. اكتمل جمع النتائج.")
        return CollectionResult(buildings=result, last_page=last_page, pages_collected=pages_collected)
    next_page = 2
    while next_page <= max_pages:
        print(f">> الانتقال إلى صفحة النتائج {next_page}...")
        if not goto_page(driver, next_page):
            stopped_reason = f"فشل الانتقال إلى صفحة النتائج {next_page}"
            print(f"!! {stopped_reason}. توقف الجمع عند إجمالي {len(result)}.")
            if progress_callback:
                progress_callback(next_page, 0, len(result), stopped_reason)
            break
        page_count, _ = add_current_page(next_page)
        if page_count < 25:
            print(">> هذه آخر صفحة نتائج. اكتمل جمع النتائج.")
            break
        next_page += 1

    if next_page > max_pages:
        stopped_reason = f"وصل إلى حد الصفحات --max-pages={max_pages}"
        print(f"!! {stopped_reason}. إذا النتائج أكثر، ارفع الرقم.")
        if progress_callback:
            progress_callback(last_page, 0, len(result), stopped_reason)

    return CollectionResult(
        buildings=result, last_page=last_page, pages_collected=pages_collected, stopped_reason=stopped_reason
    )


# =================================================================== Per-building
def read_floors(driver):
    row = driver.find_element(By.XPATH, "//*[contains(text(),'عدد أدوار المبنى')]/ancestor::tr[1]")
    for radio in row.find_elements(By.CSS_SELECTOR, "input[type='radio']"):
        if radio.is_selected():
            radio_id = radio.get_attribute("id")
            labels = driver.find_elements(By.CSS_SELECTOR, f"label[for='{radio_id}']") if radio_id else []
            text = labels[0].text.strip() if labels else ""
            return int(text) if text.isdigit() else (text or None)
    return None


def get_description_textarea(driver):
    row = driver.find_element(By.XPATH, "//*[contains(text(),'وصف المبنى')]/ancestor::tr[1]")
    return row.find_element(By.TAG_NAME, "textarea")


def process_one(driver, code: str, template: str, do_save: bool, loose: bool = False) -> ProcessResult:
    driver.get(f"{BASE_EDIT}{code}/1")
    time.sleep(2.3)

    name = driver.find_element(By.NAME, "ctl12$RequestBasics$ctl19").get_attribute("value").strip()
    gov = Select(driver.find_element(By.NAME, "ctl12$RequestBasics$ctl15")).first_selected_option.text.strip()
    floors = read_floors(driver)
    base = ProcessResult(code=code, name=name, gov=gov, floors=floors)

    # في الوضع العادي نرفض أي موقع بلا رقم أدوار. في الوضع المرن (--loose) نكمل
    # ونكتب وصفًا بدون ذكر الأدوار (مناسب للآبار والمواقع التي لا أدوار لها).
    if not isinstance(floors, int) and not loose:
        reason = "خانة الأدوار فاضية" if floors is None else f"الأدوار='{floors}' (مش رقم)"
        base.status = "فشل"
        base.reason = f"{reason} — لم يتم الحفظ حتى لا يُكتب وصف غير دقيق"
        return base

    driver.get(f"{BASE_EDIT}{code}/3")
    time.sleep(2.3)
    textarea = get_description_textarea(driver)
    existing = (textarea.get_attribute("value") or "").strip()
    base.old = existing

    desc, ok, build_reason = build_description_safe(template, name, floors, gov, loose=loose)
    if not ok:
        base.status = "متخطّى"
        base.reason = build_reason
        base.new = ""
        return base
    base.new = desc

    # تجاهل إعادة الكتابة لو الوصف المحفوظ مطابق أصلاً (idempotent)
    if existing == desc.strip():
        base.status = "نجح"
        base.reason = "الوصف مطابق مسبقًا (لم تتم إعادة الحفظ)"
        return base

    driver.execute_script(
        """
        const t = arguments[0];
        t.focus();
        t.value = arguments[1];
        t.dispatchEvent(new Event('input', {bubbles: true}));
        t.dispatchEvent(new Event('change', {bubbles: true}));
        t.dispatchEvent(new Event('blur', {bubbles: true}));
        """,
        textarea,
        desc,
    )
    if (textarea.get_attribute("value") or "").strip() != desc.strip():
        base.status = "فشل"
        base.reason = "فشل تعبئة الخانة"
        return base

    reason = "استُبدل الوصف القديم بوصف جديد" if existing else "كُتب وصف جديد (كانت فاضية)"
    if not do_save:
        base.status = "معاينة"
        base.reason = reason + " (بدون حفظ)"
        return base

    buttons = driver.find_elements(By.ID, "ctl12_btnSave")
    if not buttons or not buttons[0].is_displayed():
        base.status = "فشل"
        base.reason = "ما لقيت زر الحفظ"
        return base

    driver.execute_script("arguments[0].click();", buttons[0])
    saved_ok = False
    try:
        WebDriverWait(driver, 15).until(EC.text_to_be_present_in_element((By.ID, "lblAlertMsg"), "تم الحفظ"))
        saved_ok = True
    except TimeoutException:
        pass

    time.sleep(1.2)
    driver.get(f"{BASE_EDIT}{code}/3")
    time.sleep(2.3)
    saved = (get_description_textarea(driver).get_attribute("value") or "").strip()
    if saved == desc.strip():
        base.status = "نجح"
        base.reason = reason
        return base

    base.status = "فشل"
    base.reason = "حُفظ بس النص غير مطابق" if saved_ok else "ما ظهرت رسالة الحفظ"
    return base


# =================================================================== Context (resilient driver)
@dataclass
class Context:
    args: argparse.Namespace
    ui: UI
    username: str
    password: str
    driver: object = None

    def make_driver(self):
        if self.args.attach:
            return get_attached_driver(self.args.port)
        return create_background_driver(
            headless=self.args.headless,
            profile=self.args.profile,
            browser=self.args.browser,
        )

    def restart_driver(self) -> None:
        try:
            if self.driver is not None:
                self.driver.quit()
        except Exception:
            pass
        self.driver = self.make_driver()
        login(self.driver, self.username, self.password)

    def ensure_session(self) -> None:
        try:
            if is_logged_in(self.driver):
                return
        except WebDriverException:
            self.restart_driver()
            return
        self.ui.log("الجلسة غير مسجّلة — إعادة تسجيل الدخول…", "warn")
        login(self.driver, self.username, self.password)


def safe_process(ctx: Context, code: str, template: str, do_save: bool, retries: int, loose: bool = False) -> ProcessResult:
    last_error = None
    for attempt in range(retries + 1):
        try:
            ctx.ensure_session()
            return process_one(ctx.driver, code, template, do_save, loose=loose)
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            dead = False
            try:
                _ = ctx.driver.current_url
            except Exception:
                dead = True
            if dead:
                ctx.ui.log("المتصفح توقّف — إعادة تشغيله ومتابعة من نفس المكان…", "warn")
                try:
                    ctx.restart_driver()
                except Exception as exc2:
                    last_error = f"فشل إعادة تشغيل المتصفح: {exc2}"
            if attempt < retries:
                time.sleep(3)
    return ProcessResult(code=code, floors="?", status="فشل", reason=f"خطأ متكرر: {last_error}")


# =================================================================== Checkpoint
def checkpoint_path(excel_path: Path, decision: object) -> Path:
    safe = re.sub(r"[^0-9A-Za-z]", "_", str(decision)) or "بدون_قرار"
    return excel_path.parent / f".checkpoint_قرار_{safe}.json"


def load_checkpoint(path: Path) -> Dict[str, Dict[str, str]]:
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and isinstance(data.get("done"), dict):
                return data
        except Exception:
            pass
    return {"done": {}}


def save_checkpoint(path: Path, data: Dict[str, Dict[str, str]]) -> None:
    try:
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")
        os.replace(tmp, path)
    except Exception as exc:
        LOGGER.warning("تعذّر حفظ checkpoint: %s", exc)


def read_done_from_excel(excel_path: Path, decision: object) -> Set[str]:
    done: Set[str] = set()
    if not excel_path.exists():
        return done
    try:
        workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        sheet_name = sheet_name_for(str(decision))
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            return done
        worksheet = workbook[sheet_name]
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            code = str(row[0]).strip() if row[0] is not None else ""
            status = str(row[STATUS_COL - 1]).strip() if len(row) >= STATUS_COL and row[STATUS_COL - 1] else ""
            if code and status == "نجح":
                done.add(code)
        workbook.close()
    except Exception as exc:
        LOGGER.warning("تعذّر قراءة الإنجاز من الإكسل: %s", exc)
    return done


# =================================================================== Excel
def sheet_name_for(registration_decision: str) -> str:
    name = f"قرار {str(registration_decision).strip() or 'بدون_قرار'}"
    for char in r"[]:*?/\\":
        name = name.replace(char, "_")
    return name[:31]


def prepare_sheet(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    worksheet = None
    if len(workbook.sheetnames) == 1:
        only = workbook[workbook.sheetnames[0]]
        if only.max_row <= 1 and (only.cell(row=1, column=1).value in (None, "")):
            worksheet = only
            worksheet.title = sheet_name
    if worksheet is None:
        worksheet = workbook.create_sheet(title=sheet_name)
    # الكتابة الصريحة على الصف الأول تتجنّب علة append بعد لمس الخلايا (عناوين فارغة).
    for col_index, col in enumerate(EXCEL_COLUMNS, 1):
        worksheet.cell(row=1, column=col_index, value=str(col))
    worksheet.sheet_view.rightToLeft = True
    return worksheet


def rewrite_table(worksheet, sheet_name: str) -> None:
    for table_name in list(worksheet.tables.keys()):
        del worksheet.tables[table_name]
    ref = f"A1:{get_column_letter(len(EXCEL_COLUMNS))}{max(worksheet.max_row, 1)}"
    safe_name = "tbl_" + "".join(char if char.isalnum() else "_" for char in sheet_name)
    table = Table(displayName=safe_name[:31], ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False
    )
    worksheet.add_table(table)
    widths = [10, 32, 10, 14, 16, 10, 10, 22, 30, 60, 40, 20]
    for index, width in enumerate(widths, 1):
        worksheet.column_dimensions[get_column_letter(index)].width = width


def style_sheet(worksheet) -> None:
    """تنسيق جمالي: تجميد الرأس + التفاف الوصف + تلوين شرطي للحالة."""
    worksheet.freeze_panes = "A2"
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="right")
    desc_col = EXCEL_COLUMNS.index("الوصف المكتوب") + 1
    prev_col = EXCEL_COLUMNS.index("النص السابق") + 1
    for row_index in range(2, worksheet.max_row + 1):
        worksheet.cell(row=row_index, column=desc_col).alignment = wrap
        worksheet.cell(row=row_index, column=prev_col).alignment = wrap

    worksheet.conditional_formatting = ConditionalFormattingList()
    col = get_column_letter(STATUS_COL)
    rng = f"{col}2:{col}{max(worksheet.max_row, 2)}"
    palette = {
        "نجح": "C6EFCE",
        "فشل": "FFC7CE",
        "معاينة": "FFEB9C",
        "متخطّى": "D9D9D9",
        "متروك": "D9D9D9",
    }
    for value, color in palette.items():
        worksheet.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[f'"{value}"'], fill=PatternFill("solid", fgColor=color)),
        )


def upsert_excel_row(path: Path, record: Dict[str, object]) -> bool:
    sheet_name = sheet_name_for(str(record.get("رقم القرار", "")))
    for attempt in range(6):
        try:
            workbook = openpyxl.load_workbook(path) if path.exists() else openpyxl.Workbook()
            worksheet = prepare_sheet(workbook, sheet_name)
            code = str(record.get("الكود", "")).strip()
            target_row = None
            for row_index in range(2, worksheet.max_row + 1):
                if str(worksheet.cell(row=row_index, column=1).value).strip() == code:
                    target_row = row_index
                    break
            row_values = [record.get(col, "") if record.get(col, "") is not None else "" for col in EXCEL_COLUMNS]
            if target_row is None:
                target_row = max(worksheet.max_row + 1, 2)
            for col_index, value in enumerate(row_values, 1):
                worksheet.cell(row=target_row, column=col_index, value=value)
            rewrite_table(worksheet, sheet_name)
            style_sheet(worksheet)
            workbook.save(path)
            return True
        except PermissionError:
            if attempt < 5:
                print("   (الإكسل مفتوح؟ إعادة محاولة خلال 3 ثواني...)")
                time.sleep(3)
            else:
                print("   !! ما قدرت أحفظ بالإكسل — سكّر الملف؛ سيُكمِل السكربت لاحقًا بدون فقدان التقدّم.")
                return False
        except Exception as exc:
            print(f"   !! خطأ إكسل: {type(exc).__name__}: {exc}")
            return False
    return False


def append_run_log(path: Path, record: Dict[str, object]) -> bool:
    for attempt in range(6):
        try:
            workbook = openpyxl.load_workbook(path) if path.exists() else openpyxl.Workbook()
            sheet_name = "RUN_LOG"
            worksheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook.create_sheet(sheet_name)
            if worksheet.max_row <= 1 and worksheet.cell(row=1, column=1).value in (None, ""):
                for col_index, col in enumerate(RUN_LOG_COLUMNS, 1):
                    worksheet.cell(row=1, column=col_index, value=col)
                worksheet.sheet_view.rightToLeft = True
            row_values = [record.get(col, "") if record.get(col, "") is not None else "" for col in RUN_LOG_COLUMNS]
            target_row = max(worksheet.max_row + 1, 2)
            for col_index, value in enumerate(row_values, 1):
                worksheet.cell(row=target_row, column=col_index, value=value)
            for index, width in enumerate([24, 20, 18, 14, 12, 12, 12, 12, 14, 12, 14, 80], 1):
                worksheet.column_dimensions[get_column_letter(index)].width = width
            workbook.save(path)
            return True
        except PermissionError:
            if attempt < 5:
                time.sleep(3)
            else:
                return False
        except Exception as exc:
            LOGGER.warning("خطأ RUN_LOG: %s", exc)
            return False
    return False


def write_summary(path: Path, summary: Dict[str, object]) -> bool:
    for attempt in range(6):
        try:
            workbook = openpyxl.load_workbook(path) if path.exists() else openpyxl.Workbook()
            sheet_name = "SUMMARY"
            if sheet_name in workbook.sheetnames:
                del workbook[sheet_name]
            worksheet = workbook.create_sheet(sheet_name, 0)
            worksheet.sheet_view.rightToLeft = True
            worksheet.append(SUMMARY_COLUMNS)
            for col in range(1, len(SUMMARY_COLUMNS) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E5F")
            for key, value in summary.items():
                worksheet.append([key, value])
            worksheet.column_dimensions["A"].width = 35
            worksheet.column_dimensions["B"].width = 80
            workbook.save(path)
            return True
        except PermissionError:
            if attempt < 5:
                time.sleep(3)
            else:
                return False
        except Exception as exc:
            LOGGER.warning("خطأ SUMMARY: %s", exc)
            return False
    return False


# =================================================================== Decision runner
def run_decision(ctx: Context, decision: object, template: str, grand: Dict[str, int]) -> None:
    args = ctx.args
    ui = ctx.ui
    excel_path = Path(args.excel)
    do_save = not args.preview
    mode = "حفظ فعلي" if do_save else "معاينة"
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    start_time = datetime.now()
    cp_path = checkpoint_path(excel_path, decision)
    checkpoint = load_checkpoint(cp_path)
    done_codes: Set[str] = set(checkpoint.get("done", {}).keys()) | read_done_from_excel(excel_path, decision)

    state = {"last_page": 0, "last_index": 0, "last_code": "", "total": 0}

    def log_event(stage, details="", result_page="", building_index="", total_buildings="", code="", status=""):
        append_run_log(
            excel_path,
            {
                "Run ID": run_id,
                "وقت الحدث": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "المرحلة": stage,
                "المنطقة": args.region,
                "رقم القرار": decision,
                "الوضع": mode,
                "صفحة النتائج": result_page,
                "ترتيب المبنى": building_index,
                "إجمالي المباني": total_buildings,
                "الكود": code,
                "الحالة": status,
                "التفاصيل": details,
            },
        )

    def update_summary(final_status, extra=""):
        write_summary(
            excel_path,
            {
                "المؤلف": AUTHOR,
                "Run ID": run_id,
                "حالة التشغيل": final_status,
                "وقت البداية": start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "آخر تحديث": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "المنطقة": args.region,
                "رقم القرار": decision,
                "الوضع": mode,
                "آخر صفحة نتائج": state["last_page"],
                "آخر ترتيب مبنى": state["last_index"],
                "إجمالي المباني": state["total"],
                "آخر كود": state["last_code"],
                "نجح": grand.get("نجح", 0),
                "معاينة": grand.get("معاينة", 0),
                "متخطّى": grand.get("متخطّى", 0),
                "متروك": grand.get("متروك", 0),
                "فشل": grand.get("فشل", 0),
                "ملاحظات": extra,
            },
        )

    ui.set_decision(decision, 0)
    ui.log(f"بدء قرار {decision} — الوضع: {mode}", "info")
    log_event("START", f"بدأ التشغيل لقرار {decision}")
    update_summary("STARTED")

    # ----- collection (with crash safety) -----
    collection: Optional[CollectionResult] = None

    only_codes = [str(c).strip() for c in (getattr(args, "only_codes", None) or []) if str(c).strip()]
    if only_codes:
        # وضع الأكواد المحددة: لا بحث ولا جمع — نعالج الأكواد المطلوبة مباشرةً.
        collection = CollectionResult(
            buildings=[(c, "") for c in only_codes], last_page=0, pages_collected=0
        )
        ui.log(f"وضع الأكواد المحددة: {len(only_codes)} كود فقط — بدون بحث/جمع لكل المواقع", "info")
        log_event("COLLECTED", f"أكواد محددة ({len(only_codes)}) بدون جمع", total_buildings=len(only_codes))
    else:
        for collect_attempt in range(2):
            try:
                ctx.ensure_session()
                do_search(ctx.driver, args.region, str(decision))

                def page_progress(page_number, added, running_total, details):
                    state["last_page"] = page_number
                    log_event("PAGE", f"{details} | جديد: {added} | الإجمالي: {running_total}",
                              result_page=page_number, total_buildings=running_total)

                collection = collect_all_buildings(ctx.driver, max_pages=args.max_pages, progress_callback=page_progress)
                break
            except Exception as exc:
                ui.log(f"خطأ أثناء البحث/الجمع: {type(exc).__name__}: {exc}", "warn")
                try:
                    _ = ctx.driver.current_url
                    dead = False
                except Exception:
                    dead = True
                if dead and collect_attempt == 0:
                    ui.log("إعادة تشغيل المتصفح وإعادة الجمع…", "warn")
                    try:
                        ctx.restart_driver()
                        continue
                    except Exception:
                        pass
                log_event("ERROR", f"فشل الجمع: {exc}")
                update_summary("ERROR", str(exc))
                return

    if collection is None:
        return

    buildings = collection.buildings
    if getattr(args, "limit", 0) and args.limit > 0:
        buildings = buildings[: args.limit]
        ui.log(f"وضع التجربة: سيُعالَج أول {len(buildings)} مبنى فقط (--limit)", "warn")
    state["last_page"] = collection.last_page
    state["total"] = len(buildings)
    ui.set_decision(decision, len(buildings))

    # حماية الدفعة الناقصة تُطبّق فقط عند الحفظ الفعلي؛ المعاينة تكمل دائمًا.
    if collection.stopped_reason and do_save and not args.allow_partial_collection:
        msg = (
            f"توقف جمع النتائج قبل الاكتمال: {collection.stopped_reason}. لم تبدأ المعالجة لتفادي دفعة ناقصة. "
            "استخدم --allow-partial-collection لمعالجة المجموع فقط."
        )
        ui.log(msg, "warn")
        log_event("END", msg, result_page=state["last_page"], total_buildings=len(buildings), status="جمع غير مكتمل")
        update_summary("COLLECTION_INCOMPLETE", msg)
        return

    if not buildings:
        msg = "لا توجد مبانٍ مطابقة. راجع المنطقة/رقم القرار/الخطوة."
        ui.log(msg, "warn")
        log_event("END", msg, status="لا نتائج")
        update_summary("NO_RESULTS", msg)
        return

    already = len([c for c, _ in buildings if c in done_codes])
    ui.log(f"تم جمع {len(buildings)} مبنى | مُنجز مسبقًا: {already} | متبقّي: {len(buildings) - already}", "info")
    log_event("COLLECTED", f"جُمع {len(buildings)} مبنى من {collection.pages_collected} صفحة",
              result_page=state["last_page"], total_buildings=len(buildings))
    update_summary("PROCESSING", collection.stopped_reason)

    # ----- per-building processing -----
    for item_index, (code, _) in enumerate(buildings, 1):
        state["last_index"] = item_index
        state["last_code"] = code
        ui.set_progress(item_index, code)

        # في وضع الأكواد المحددة نعيد المعالجة دائمًا حتى لو سبق نجاحها.
        if code in done_codes and not getattr(args, "only_codes", None):
            ui.incr("متخطّى")
            ui.log(f"[{item_index}/{len(buildings)}] {code} — مُنجز مسبقًا، تخطّي", "skip")
            continue

        result = safe_process(ctx, code, template, do_save, retries=args.retries,
                              loose=getattr(args, "loose", False))
        status = result.status
        grand[status] = grand.get(status, 0) + 1
        ui.incr(status)
        level = {"نجح": "ok", "فشل": "fail", "معاينة": "preview"}.get(status, "skip")
        ui.log(f"[{item_index}/{len(buildings)}] {code} — {status}: {result.reason}", level)

        log_event("BUILDING", result.reason, result_page=state["last_page"], building_index=item_index,
                  total_buildings=len(buildings), code=code, status=status)

        upsert_excel_row(
            excel_path,
            {
                "الكود": code,
                "اسم المبنى": result.name,
                "المنطقة": args.region,
                "المحافظة": result.gov,
                "القرية (بحث)": "بدون فلترة قرية",
                "رقم القرار": decision,
                "عدد الأدوار": result.floors,
                "الحالة": status,
                "السبب": result.reason,
                "الوصف المكتوب": result.new,
                "النص السابق": result.old,
                "وقت المعالجة": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
        )

        if status == "نجح":
            done_codes.add(code)
            checkpoint.setdefault("done", {})[code] = status
            save_checkpoint(cp_path, checkpoint)

        if item_index % 10 == 0:
            update_summary("PROCESSING")

        time.sleep(args.delay + random.uniform(0.0, 0.6))

    log_event("END", "انتهى التشغيل", result_page=state["last_page"], building_index=state["last_index"],
              total_buildings=len(buildings), code=state["last_code"], status="انتهى")
    update_summary("DONE")
    ui.log(f"انتهى قرار {decision} ✓", "ok")


# =================================================================== Auto run
def run_auto(args: argparse.Namespace) -> int:
    # إعدادات افتراضية للتشغيل الذاتي
    if not args.region:
        args.region = DEFAULT_REGION
    if not args.decisions:
        args.decisions = list(DEFAULT_DECISIONS)
    if not args.excel:
        args.excel = Path.cwd() / DEFAULT_EXCEL_NAME
    args.excel = Path(args.excel)
    do_save = not args.preview
    mode = "حفظ فعلي" if do_save else "معاينة"

    setup_logging(args.excel.parent / "heritage_log.txt")
    LOGGER.info("بدء تشغيل | المؤلف=%s | المنطقة=%s | القرارات=%s | الوضع=%s",
                AUTHOR, args.region, args.decisions, mode)

    if args.region not in TEMPLATES or not TEMPLATES.get(args.region, "").strip():
        print(f"!! ما في قالب للمنطقة '{args.region}'. أضفه داخل TEMPLATES.")
        return 2
    if not validate_template(TEMPLATES[args.region], verbose=True):
        print("!! القالب فيه تعديل ما مسك. صحّحه قبل التشغيل.")
        return 2
    template = TEMPLATES[args.region]

    username, password = get_credentials(args, interactive=True)

    ui = UI(enabled=not args.no_color)
    ui.set_run(args.decisions, mode)

    ctx = Context(args=args, ui=ui, username=username, password=password)
    try:
        print(f">> تشغيل المتصفح ({'خلفي' if args.headless else 'ظاهر'} | {args.browser}) وتسجيل الدخول تلقائيًا…")
        ctx.driver = ctx.make_driver()
        login(ctx.driver, username, password)
        if not is_logged_in(ctx.driver):
            print("!! تعذّر تأكيد تسجيل الدخول. تحقّق من بيانات الاعتماد.")
            return 2
        print(f">> تم تسجيل الدخول بنجاح | المنطقة={args.region} | القرارات={args.decisions} | الوضع={mode}")
    except Exception as exc:
        print(f"!! فشل الإقلاع/تسجيل الدخول: {type(exc).__name__}: {exc}")
        LOGGER.error("فشل الإقلاع: %s", exc)
        return 2

    grand: Dict[str, int] = {}
    # وضع الأكواد المحددة: صفحة الإكسل تُسمّى حسب القرار، فنمرّ مرة واحدة فقط
    # (على أول قرار) حتى لا تتكرر معالجة نفس الكود تحت قرارين.
    decisions_to_run = args.decisions
    if getattr(args, "only_codes", None):
        decisions_to_run = args.decisions[:1] or list(DEFAULT_DECISIONS)[:1]
        print(f">> وضع الأكواد المحددة: {len(args.only_codes)} كود | ستُكتب في ورقة «قرار {decisions_to_run[0]}»"
              + (" | مرن (يشمل الآبار والمواقع بلا أدوار)" if getattr(args, "loose", False) else ""))
    ui.start()
    try:
        for decision in decisions_to_run:
            run_decision(ctx, decision, template, grand)
    except KeyboardInterrupt:
        ui.log("إيقاف يدوي (Ctrl+C) — التقدّم محفوظ، أعد التشغيل للاستئناف.", "warn")
    finally:
        ui.stop()
        if not args.attach and ctx.driver is not None:
            try:
                ctx.driver.quit()
            except Exception:
                pass

    print(f"\n{'=' * 64}")
    print(f">> الملخّص الإجمالي | {AUTHOR}")
    for key in ("نجح", "معاينة", "متخطّى", "متروك", "فشل"):
        if grand.get(key):
            print(f"   {key}: {grand[key]}")
    print(f">> التفاصيل الكاملة بالإكسل: {args.excel}")
    print(f">> سجل التشغيل: {args.excel.parent / 'heritage_log.txt'}")
    print("=" * 64)
    return 0


# =================================================================== Diagnose
def diagnose_search(args: argparse.Namespace) -> int:
    """يفتح صفحة البحث ويجمع أسماء كل الحقول مع عناوينها، ويختبر فلتر رقم القرار.

    الهدف: معرفة الحقل الصحيح لـ «رقم القرار» عندما لا يعمل الفلتر الافتراضي.
    """
    setup_console()
    username, password = get_credentials(args, interactive=True)
    out_path = Path.cwd() / "diagnose_form.txt"
    test_decision = str(args.decision or "26")
    lines: List[str] = [f"تقرير تشخيص البحث | المنطقة={args.region} | قرار اختبار={test_decision}"]

    driver = create_background_driver(headless=False, profile=None, browser=args.browser)

    def dump_fields(tag: str) -> None:
        lines.append(f"\n===== {tag} =====")
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, "input, select, textarea")
        except Exception as exc:
            lines.append(f"(تعذّر قراءة الحقول: {exc})")
            return
        for el in elements:
            try:
                if not el.is_displayed():
                    continue
                tag_name = el.tag_name
                etype = el.get_attribute("type") or tag_name
                if etype in {"hidden"}:
                    continue
                name = el.get_attribute("name") or ""
                eid = el.get_attribute("id") or ""
                value = (el.get_attribute("value") or "")[:30]
                selected = ""
                if tag_name == "select":
                    try:
                        selected = Select(el).first_selected_option.text.strip()[:30]
                    except Exception:
                        selected = ""
                try:
                    label = el.find_element(By.XPATH, "./ancestor::tr[1]").text.strip().replace("\n", " ")[:55]
                except Exception:
                    label = ""
                star = "  <== فيه (قرار)" if "قرار" in label else ""
                lines.append(
                    f"[{tag_name}/{etype}] name={name} id={eid} value={value!r} selected={selected!r} | label={label}{star}"
                )
            except Exception:
                continue

    try:
        print(">> تسجيل الدخول…")
        login(driver, username, password)
        print(">> فتح صفحة البحث…")
        driver.get(BASE_VIEW)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "ctl12$ctl08")))
        time.sleep(1)
        try:
            region_select = select_by_text(driver, "ctl12$ctl08", args.region)
            wait_postback(driver, region_select)
        except Exception as exc:
            lines.append(f"خطأ اختيار المنطقة: {exc}")
        try:
            step_select = driver.find_element(By.NAME, "ctl12$ctl32")
            Select(step_select).select_by_value(STEP_VALUE)
            wait_postback(driver, step_select)
        except Exception as exc:
            lines.append(f"خطأ اختيار الخطوة: {exc}")

        dump_fields("كل حقول البحث بعد اختيار المنطقة والخطوة")

        # اختبار: البحث بدون أي رقم قرار (كم مبنى؟)
        try:
            driver.find_element(By.NAME, "ctl12$ctl04").clear()
            btn = driver.find_element(By.ID, "ctl12_btnSearch")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(3)
            lines.append(f"\n>> نتائج البحث بدون رقم قرار (أول صفحة): {len(page_buildings(driver))} صف")
        except Exception as exc:
            lines.append(f"خطأ البحث بدون قرار: {exc}")

        # اختبار: البحث برقم القرار عبر الحقل الحالي ctl12$ctl28
        try:
            driver.get(BASE_VIEW)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "ctl12$ctl08")))
            time.sleep(1)
            region_select = select_by_text(driver, "ctl12$ctl08", args.region)
            wait_postback(driver, region_select)
            step_select = driver.find_element(By.NAME, "ctl12$ctl32")
            Select(step_select).select_by_value(STEP_VALUE)
            wait_postback(driver, step_select)
            driver.find_element(By.NAME, "ctl12$ctl04").clear()
            dec = driver.find_element(By.NAME, "ctl12$ctl28")
            dec.clear()
            dec.send_keys(test_decision)
            btn = driver.find_element(By.ID, "ctl12_btnSearch")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(3)
            lines.append(
                f">> نتائج البحث بإدخال القرار {test_decision} في الحقل ctl12$ctl28 (أول صفحة): "
                f"{len(page_buildings(driver))} صف"
            )
        except Exception as exc:
            lines.append(f"خطأ البحث برقم القرار: {exc}")

        out_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n>> تم حفظ تقرير التشخيص في: {out_path}")
        print(">> ابعت هذا الملف لأحدّد حقل رقم القرار الصحيح.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    return 0


# =================================================================== CLI
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=f"أتمتة أوصاف القرى التاريخية (PRO) — {AUTHOR}"
    )
    subparsers = parser.add_subparsers(dest="command")

    launch = subparsers.add_parser("launch", help="فتح نافذة Chrome يدوية مع منفذ تصحيح")
    launch.add_argument("--chrome-path", default=None)
    launch.add_argument("--profile", default=DEFAULT_AUTOMATION_PROFILE)
    launch.add_argument("--port", type=int, default=DEBUG_PORT)

    env_cmd = subparsers.add_parser("write-env", help="حفظ بيانات الدخول محليًا في .env")
    env_cmd.add_argument("--username", required=True)
    env_cmd.add_argument("--password", required=True)

    run = subparsers.add_parser("run", help="تشغيل ذاتي كامل (الافتراضي)")
    run.add_argument("--region", default=DEFAULT_REGION)
    run.add_argument("--decisions", nargs="+", type=int, default=list(DEFAULT_DECISIONS),
                     help="أرقام القرارات (الافتراضي: 25 26)")
    run.add_argument("--excel", type=Path, default=None)
    run.add_argument("--preview", action="store_true", help="تعبئة ومعاينة بدون حفظ")
    run.add_argument("--username", default=None)
    run.add_argument("--password", default=None)
    run.add_argument("--profile", default=None)
    run.add_argument("--browser", choices=("edge", "chrome"), default="chrome")
    run.add_argument("--headless", action="store_true", help="تشغيل بالخلفية دون إظهار المتصفح")
    run.add_argument("--attach", action="store_true", help="الالتصاق بجلسة Chrome تصحيح على --port")
    run.add_argument("--port", type=int, default=DEBUG_PORT)
    run.add_argument("--max-pages", type=int, default=200)
    run.add_argument("--retries", type=int, default=2)
    run.add_argument("--limit", type=int, default=0, help="معالجة أول N مبنى فقط (0 = الكل) — مفيد للتجربة")
    run.add_argument("--only-codes", nargs="+", default=None, dest="only_codes",
                     help="معالجة هذه الأكواد فقط مباشرةً بدون بحث/جمع لكل المواقع (مثال: --only-codes 40125 40159)")
    run.add_argument("--loose", action="store_true",
                     help="وضع مرن: يقبل «موقع/ميني رقم N» (آبار/مساجد/حصون) ويكتب الوصف بدون ذكر الأدوار إن لم تكن رقمية")
    run.add_argument("--delay", type=float, default=DELAY_BETWEEN)
    run.add_argument("--no-color", action="store_true", help="تعطيل البانر الملوّن (وضع نصي بسيط)")
    run.add_argument("--allow-partial-collection", action="store_true")

    diag = subparsers.add_parser("diagnose", help="تشخيص حقول البحث (لمعرفة حقل رقم القرار الصحيح)")
    diag.add_argument("--region", default=DEFAULT_REGION)
    diag.add_argument("--decision", default="26", help="رقم قرار للاختبار")
    diag.add_argument("--browser", choices=("edge", "chrome"), default="chrome")
    diag.add_argument("--username", default=None)
    diag.add_argument("--password", default=None)

    subparsers.add_parser("validate-templates", help="فحص القوالب")
    return parser


def setup_console() -> None:
    """تحسين عرض العربية في الطرفية (UTF-8). ملاحظة: نافذة cmd القديمة قد تعكس
    الأحرف العربية بسبب قيودها؛ استخدم Windows Terminal لعرض صحيح."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    if os.name == "nt":
        try:
            os.system("chcp 65001 >nul")
        except Exception:
            pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "launch":
        launch_chrome(args.chrome_path, args.profile, args.port)
        return 0
    if args.command == "write-env":
        write_env_file(args.username, args.password)
        return 0
    if args.command == "validate-templates":
        ok = True
        print(">> المناطق المتوفّرة:", list(TEMPLATES.keys()))
        for region, template in TEMPLATES.items():
            print(f"\n########## فحص قالب: {region} ##########")
            ok = validate_template(template, verbose=True) and ok
        return 0 if ok else 2
    if args.command == "diagnose":
        return diagnose_search(args)

    # لا أمر → تشغيل ذاتي بالإعدادات الافتراضية
    if args.command is None:
        args = parser.parse_args(["run"])
    return run_auto(args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n>> أُوقف يدويًا.")
        raise SystemExit(130)
