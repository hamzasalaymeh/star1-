# -*- coding: utf-8 -*-
"""
محرّك التشغيل — مستقل تماماً عن الواجهة (Tkinter/PySide لا فرق).
يشغّل أي أتمتة كعملية فرعية، يقرأ أحداث التقدّم منها سطراً سطراً في خيط منفصل،
ويتحكم بها (إيقاف مؤقت / استئناف / إيقاف آمن) عبر ملف تحكم.

الاستخدام من الواجهة:
    runner = AutomationRunner()
    runner.start(cmd, cwd, env_extra)
    ... كل ~150ms:  for ev in runner.drain(): handle(ev)
    runner.pause() / runner.resume() / runner.stop()
    runner.stats  → إحصائيات حيّة
"""
import os
import json
import time
import queue
import signal
import tempfile
import threading
import subprocess

EVT = "@@EVT@@"


class AutomationRunner:
    def __init__(self):
        self.proc = None
        self.control_file = None
        self._q = queue.Queue()
        self._reader = None
        self.running = False
        self.paused = False
        self.stopping = False
        self.stats = self._fresh_stats()
        self._start_ts = None
        self._last_done = 0
        self._last_done_ts = None
        self._site_times = []

    @staticmethod
    def _fresh_stats():
        return {"total": 0, "done": 0, "failed": 0, "skipped": 0,
                "current": "", "percent": 0.0, "avg": 0.0, "elapsed": 0,
                "state": "idle"}

    # ----------------------------------------------------- تشغيل
    def start(self, cmd, cwd=None, env_extra=None):
        if self.running:
            raise RuntimeError("هناك تشغيل جارٍ بالفعل")
        self.stats = self._fresh_stats()
        self.stats["state"] = "running"
        self._start_ts = time.time()
        self._last_done = 0
        self._last_done_ts = self._start_ts
        self._site_times = []
        self.paused = False
        self.stopping = False

        fd, self.control_file = tempfile.mkstemp(prefix="hc_ctrl_", suffix=".json")
        os.close(fd)
        self._write_control({})

        env = os.environ.copy()
        env["HERITAGE_CONTROL"] = self.control_file
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        if env_extra:
            env.update({k: str(v) for k, v in env_extra.items() if v is not None})

        self.proc = subprocess.Popen(
            cmd, cwd=cwd, env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
        )
        self.running = True
        self._reader = threading.Thread(target=self._read_loop, daemon=True)
        self._reader.start()

    def _read_loop(self):
        try:
            for line in self.proc.stdout:
                self._q.put(("line", line.rstrip("\n")))
        except Exception as exc:
            self._q.put(("line", f"[خطأ قراءة]: {exc}"))
        finally:
            rc = self.proc.wait()
            self._q.put(("exit", rc))

    # ----------------------------------------------------- تحكم
    def _write_control(self, data):
        try:
            tmp = self.control_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f)
            os.replace(tmp, self.control_file)
        except Exception:
            pass

    def _merge_control(self, **kw):
        cur = {}
        try:
            with open(self.control_file, encoding="utf-8") as f:
                cur = json.load(f)
        except Exception:
            pass
        cur.update(kw)
        self._write_control(cur)
        return cur

    def pause(self):
        if self.running:
            self.paused = True
            self._merge_control(pause=True)

    def resume(self):
        if self.running:
            self.paused = False
            self._merge_control(pause=False)

    def stop(self):
        """إيقاف آمن: نطلب التوقّف، والأتمتة تخرج بعد إنهاء الموقع الحالي."""
        if self.running:
            self.stopping = True
            self.stats["state"] = "stopping"
            self._merge_control(stop=True, pause=False)

    def kill(self):
        """إيقاف قسري (آخر حل) — قد يوقف وسط عملية، فنتجنّبه إلا عند التعليق."""
        try:
            if self.proc and self.proc.poll() is None:
                self.proc.terminate()
        except Exception:
            pass

    # ----------------------------------------------------- سحب الأحداث (تستدعيها الواجهة دورياً)
    def drain(self):
        """يرجّع قائمة أحداث جاهزة للعرض: نوعها 'log' أو 'progress' أو 'state' أو 'exit'."""
        out = []
        while True:
            try:
                kind, val = self._q.get_nowait()
            except queue.Empty:
                break
            if kind == "exit":
                self.running = False
                self.stats["state"] = "stopped" if self.stopping else "finished"
                self._update_elapsed()
                out.append({"type": "exit", "code": val, "state": self.stats["state"]})
                continue
            # kind == "line"
            ev = self._parse_event(val)
            if ev is None:
                out.append({"type": "log", "msg": val, "level": "info"})
            else:
                out.append(self._apply_event(ev))
        self._update_elapsed()
        return out

    @staticmethod
    def _parse_event(line):
        if line.startswith(EVT):
            try:
                return json.loads(line[len(EVT):])
            except Exception:
                return None
        return None

    def _apply_event(self, ev):
        kind = ev.get("kind")
        if kind == "total":
            self.stats["total"] = int(ev.get("total", 0))
            return {"type": "progress", "stats": dict(self.stats)}
        if kind == "progress":
            for k in ("done", "failed", "skipped", "total"):
                if ev.get(k) is not None:
                    self.stats[k] = int(ev[k])
            if ev.get("current") is not None:
                self.stats["current"] = str(ev["current"])
            self._recompute()
            return {"type": "progress", "stats": dict(self.stats)}
        if kind == "state":
            self.stats["state"] = ev.get("state", self.stats["state"])
            return {"type": "state", "state": self.stats["state"]}
        if kind == "done":
            self.stats["done"] = int(ev.get("ok", self.stats["done"]))
            self.stats["failed"] = int(ev.get("failed", self.stats["failed"]))
            self.stats["skipped"] = int(ev.get("skipped", self.stats["skipped"]))
            self._recompute()
            return {"type": "progress", "stats": dict(self.stats)}
        if kind == "log":
            return {"type": "log", "msg": ev.get("msg", ""), "level": ev.get("level", "info")}
        return {"type": "log", "msg": str(ev), "level": "info"}

    def _recompute(self):
        t = self.stats["total"]
        d = self.stats["done"] + self.stats["failed"] + self.stats["skipped"]
        self.stats["percent"] = (d / t * 100.0) if t else 0.0
        # متوسط زمن الموقع (بناءً على المواقع المكتملة منذ البداية)
        now = time.time()
        if self.stats["done"] > self._last_done:
            inc = self.stats["done"] - self._last_done
            dt = now - (self._last_done_ts or now)
            if inc > 0 and dt > 0:
                self._site_times.append(dt / inc)
                self._site_times = self._site_times[-50:]
            self._last_done = self.stats["done"]
            self._last_done_ts = now
        if self._site_times:
            self.stats["avg"] = sum(self._site_times) / len(self._site_times)

    def _update_elapsed(self):
        if self._start_ts:
            self.stats["elapsed"] = int(time.time() - self._start_ts)

    def remaining(self):
        done = self.stats["done"] + self.stats["failed"] + self.stats["skipped"]
        return max(0, self.stats["total"] - done)

    def cleanup(self):
        try:
            if self.control_file and os.path.exists(self.control_file):
                os.remove(self.control_file)
        except Exception:
            pass
