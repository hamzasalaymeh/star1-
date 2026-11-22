# -*- coding: utf-8 -*-
"""
🧹 أداة متابعة الفاشلين
تمسح مجلداً وتجمع من كل تقارير Excel المواقع/المباني التي فشلت أو تُخطّيت أو
بقيت معلّقة، وتطلّع:
  • ملف Excel واحد مجمّع «متابعة_الفاشلين_<وقت>.xlsx»
  • قائمة أكواد جاهزة للصق في خانة «أكواد المباني» بالمجمّع لإعادة تشغيلها.

يشتغل مستقلاً: python failed_tracker.py            (يفتح نافذة اختيار مجلد)
            أو: python failed_tracker.py --folder "C:\\...\\المجلد"
"""

import os
import re
import sys
import argparse
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# كلمات تدل على أن الصف يحتاج متابعة (فشل/تخطّي/معلّق).
NEEDS = ("فشل", "فاشل", "خطأ", "متخطّى", "متخطّي", "معلق", "معلّق",
         "failed", "pending", "error", "skip", "❌", "✗")
# رؤوس أعمدة الكود المحتملة عبر كل التقارير.
CODE_HEADERS = ("الكود", "رقم المبنى", "رقم الموقع", "كود", "code", "site_id", "رقم الموقع ")
STATUS_HEADERS = ("الحالة", "status")


def ask_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        val = filedialog.askdirectory(title="اختر المجلد اللي فيه تقارير Excel")
        root.destroy()
        return val
    except Exception:
        return input("مسار المجلد: ").strip()


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def status_needs_followup(value):
    v = norm(value)
    return any(k in v for k in NEEDS) and "لا يوجد" not in v and "نجح" not in v


def scan_workbook(path):
    """يرجّع قائمة صفوف تحتاج متابعة من كل أوراق الملف."""
    rows_out = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        print(f"   ⚠️ تعذّر فتح {os.path.basename(path)}: {exc}")
        return rows_out

    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            continue
        cols = {norm(h): i for i, h in enumerate(header) if h is not None}
        status_i = next((cols[h] for h in cols if h in STATUS_HEADERS), None)
        code_i = next((cols[h] for h in cols if h in CODE_HEADERS), None)
        if status_i is None:
            continue
        detail_i = next((cols[h] for h in cols if h in ("السبب", "التفاصيل", "الخطأ (إن وجد)")), None)
        for r in rows:
            if not r or status_i >= len(r):
                continue
            if status_needs_followup(r[status_i]):
                code = norm(r[code_i]) if code_i is not None and code_i < len(r) else ""
                detail = norm(r[detail_i]) if detail_i is not None and detail_i < len(r) else ""
                rows_out.append({
                    "الكود": code,
                    "الحالة": norm(r[status_i]),
                    "المصدر": f"{os.path.basename(path)} › {ws.title}",
                    "التفاصيل": detail,
                })
    try:
        wb.close()
    except Exception:
        pass
    return rows_out


def collect(folder):
    found = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                # نتجاهل ملف المتابعة المجمّع نفسه حتى لا ندور بحلقة.
                if f.startswith("متابعة_الفاشلين"):
                    continue
                found.extend(scan_workbook(os.path.join(root, f)))
    return found


def write_report(folder, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "متابعة"
    ws.sheet_view.rightToLeft = True
    headers = ["الكود", "الحالة", "المصدر", "التفاصيل"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0D47A1")
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append([r["الكود"], r["الحالة"], r["المصدر"], r["التفاصيل"]])
    for col, w in zip("ABCD", (14, 22, 45, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    out = os.path.join(folder, f"متابعة_الفاشلين_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(out)
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description="تجميع المواقع الفاشلة/المتخطّاة من تقارير Excel")
    parser.add_argument("--folder", default=None)
    args = parser.parse_args(argv)

    folder = args.folder or ask_folder()
    if not folder or not os.path.isdir(folder):
        print("❌ لم يتم اختيار مجلد صالح.")
        return 2

    print(f"🔍 جاري مسح: {folder}")
    rows = collect(folder)
    if not rows:
        print("🎉 ما في مواقع فاشلة/متخطّاة في التقارير — كله تمام!")
        return 0

    # أكواد فريدة (رقمية) للصق المباشر في خانة «أكواد المباني»
    codes = sorted({r["الكود"] for r in rows if r["الكود"].isdigit()}, key=int)

    print(f"\n📊 عدد الصفوف المحتاجة متابعة: {len(rows)}")
    print(f"🔢 أكواد فريدة: {len(codes)}")
    if codes:
        print("\n— الصق هذه الأكواد في خانة «أكواد المباني» (نطاق=أكواد محددة) —")
        print(" ".join(codes))

    out = write_report(folder, rows)
    print(f"\n💾 تقرير مجمّع: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
