# -*- coding: utf-8 -*-
"""باني الأوامر: يحوّل (plugin + قيم الإعدادات + بيانات الدخول) إلى أمر تشغيل
جاهز للمحرّك. مستقل عن الواجهة، فيسهل اختباره."""
import os
import sys


def python_exe():
    return sys.executable or "python"


def _value(values, name):
    return values.get(name, "")


def visible(spec, values):
    cond = spec.get("when")
    if not cond:
        return True
    return _value(values, cond["field"]) in cond["equals"]


def _run_base():
    """الأساس الذي يشغّل سكربت أتمتة — يعمل كـ python عادي أو EXE مجمّع.
    نعيد استدعاء البرنامج نفسه بوضع --hc-run (نمط قياسي لتطبيقات PyInstaller)."""
    if getattr(sys, "frozen", False):          # EXE مجمّع: البرنامج نفسه هو المشغّل
        return [sys.executable, "--hc-run"]
    app_py = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")
    return [python_exe(), app_py, "--hc-run"]


def build_command(plugin, values, creds, scripts_dir):
    """يرجّع (cmd:list, env:dict). يرفع ValueError لو نقص حقل مطلوب."""
    script_path = os.path.join(scripts_dir, plugin["script"])
    if not os.path.exists(script_path):
        raise ValueError(f"لم يُعثر على السكربت: {script_path}")

    cmd = _run_base() + [script_path] + list(plugin.get("argv_prefix", []))
    env = {}
    positional = []

    user, pw = creds
    if user:
        env["HERITAGE_USERNAME"] = user
    if pw:
        env["HERITAGE_PASSWORD"] = pw

    for spec in plugin.get("settings", []):
        if not visible(spec, values):
            continue
        name = spec["name"]
        val = _value(values, name)
        typ = spec["type"]

        if spec.get("required") and typ in ("text", "int", "choice") and not str(val).strip():
            raise ValueError(f"الحقل «{spec['label']}» مطلوب")

        if spec.get("env"):
            if str(val).strip():
                env[name] = val
            continue
        if spec.get("positional"):
            if str(val).strip():
                positional.append(str(val).strip())
            continue
        if typ == "bool":
            if val and spec.get("arg"):
                cmd.append(spec["arg"])
            continue
        if typ == "int":
            s = str(val).strip()
            if s and s != "0" and spec.get("arg"):
                cmd += [spec["arg"], s]
            continue
        if not str(val).strip():
            continue
        if spec.get("arg"):
            if spec.get("nargs"):
                if typ == "textarea":
                    parts = [ln.strip() for ln in str(val).splitlines() if ln.strip()]
                else:
                    parts = [p for p in str(val).replace("،", " ").split(" ") if p.strip()]
                cmd += [spec["arg"]] + parts
            else:
                cmd += [spec["arg"], str(val).strip()]

    cmd += positional
    if plugin.get("pass_creds_positional") and user and pw:
        cmd += [user, pw]
    return cmd, env
