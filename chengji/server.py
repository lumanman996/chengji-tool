"""界面背后的本地服务：只在本机（127.0.0.1）上开一个端口，界面通过它调用计算、导出、设置。
不联网，外面的电脑访问不到；每次启动生成一个随机口令，界面带着口令才能调用。
"""
from __future__ import annotations

import json
import os
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import yaml

from . import __version__, license as lic
from . import service as sv
from .config import load_school, load_teachers
from .paths import ROOT, ensure_layout

UI_FILE = Path(__file__).resolve().parent / "ui" / "index.html"


def reveal(path: Path):
    """在资源管理器 / 访达里打开文件夹（是文件就把它选中）。"""
    if os.environ.get("CHENGJI_NO_OPEN"):
        return
    path = Path(path)
    try:
        if os.name == "nt":
            if path.is_file():
                subprocess.Popen(["explorer", "/select,", str(path)])
            else:
                os.startfile(str(path))                        # noqa
        elif sys.platform == "darwin":
            subprocess.run(["open", "-R", str(path)] if path.is_file() else ["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path if path.is_dir() else path.parent)], check=False)
    except Exception:
        pass


def open_file(path: Path):
    """用系统默认的程序打开文件（Excel / WPS）。"""
    if os.environ.get("CHENGJI_NO_OPEN"):
        return
    try:
        if os.name == "nt":
            os.startfile(str(path))                            # noqa
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


class App:
    """一次运行期间的状态：当前导入的登分表、当前的核算结果。"""

    def __init__(self, root: Path = ROOT):
        self.root = Path(root)
        self.config = self.root / "config" / "学校设置.yaml"
        self.local = self.root / "data" / "本校设置.yaml"
        self.teacher = self.root / "data" / "任课总表.xlsx"
        self.output = self.root / "output"
        self.templates = self.root / "templates"
        self.tmp = Path(tempfile.mkdtemp(prefix="chengji_"))
        self.src: Path | None = None      # 当前登分表
        self.opts: dict | None = None     # 当前核算的设置
        self.result = None                # (R, 结论, 提示)
        self.lock = threading.Lock()
        self.last_ping = time.time()

    def school(self):
        return load_school(self.config, self.local)

    # ---------------- 总体状态
    def api_state(self, _=None):
        s = self.school()
        return {"version": __version__, "school": s.name, "license": lic.status(), "recent": sv.list_runs(self.output),
                "templates": list(s.grade_subjects), "pdfDefaults": s.pdf_sections, "hasTeacherTable": self.teacher.is_file()}

    def api_ping(self, _=None):
        self.last_ping = time.time()
        return {}

    # ---------------- 导入、核算
    def upload(self, name: str, blob: bytes):
        name = Path(unquote(name)).name or "登分表.xlsx"
        if not name.lower().endswith((".xlsx", ".xlsm")):
            raise ValueError("请选择 Excel 文件（.xlsx）。老式的 .xls 文件请先用 Excel 另存为 .xlsx。")
        dst = self.tmp / name
        dst.write_bytes(blob)
        info = sv.inspect(dst, self.school(), self.teacher)
        self.src, self.opts, self.result = dst, None, None
        info["nameTaken"] = [r["name"] for r in sv.list_runs(self.output, 500)]
        return info

    def api_compute(self, opts):
        if not self.src:
            raise ValueError("请先导入登分表。")
        R, concl, msgs = sv.compute(self.src, self.school(), opts, self.teacher)
        view = sv.to_view(R, concl, msgs)
        sv.save_run(self.output, self.src, opts, view)
        self.opts, self.result = {**opts, "exam": view["exam"]}, (R, concl, msgs)
        self.src = self.output / view["exam"] / sv.RUN_COPY
        return view

    def api_open_run(self, body):
        meta = sv.load_run(self.output, body["name"])
        self.src = self.output / meta["view"]["exam"] / sv.RUN_COPY
        self.opts, self.result = meta["opts"], None
        return meta["view"]

    def _ensure_result(self):
        if self.result is None:
            if not (self.src and self.opts and Path(self.src).is_file()):
                raise ValueError("请先核算，或从“最近的考试”里打开一次结果。")
            self.result = sv.compute(self.src, self.school(), self.opts, self.teacher)
        return self.result

    # ---------------- 导出
    def api_export_excel(self, _=None):
        lic.require_export()
        R, concl, _m = self._ensure_result()
        path = sv.export_excel(R, concl, self.output)
        reveal(path)
        return {"path": str(path), "name": path.name}

    def api_export_pdf(self, body):
        lic.require_export()
        R, concl, _m = self._ensure_result()
        path = sv.export_pdf(R, concl, body.get("sections") or [], self.output)
        reveal(path)
        return {"path": str(path), "name": path.name}

    def api_open_output(self, body=None):
        name = sv.clean_name((body or {}).get("name") or "")
        p = self.output / name if name and (self.output / name).is_dir() else self.output
        p.mkdir(parents=True, exist_ok=True)
        reveal(p)
        return {}

    # ---------------- 模板、任课总表
    def api_template(self, body):
        key = sv.clean_name(body.get("key") or "")
        f = self.templates / (f"登分表_{key}.xlsx" if key != "任课总表" else "任课总表_模板.xlsx")
        if not f.is_file():
            raise ValueError("没有找到这个模板。")
        reveal(f)
        return {"name": f.name}

    def _teacher_info(self):
        if not self.teacher.is_file():
            return {"exists": False, "text": "还没有任课总表"}
        s = self.school()
        n = {g: len(load_teachers(self.teacher, s, g)) for g in s.grades}
        return {"exists": True, "text": "已导入 · " + " · ".join(f"{g} {k} 位" for g, k in n.items() if k) if any(n.values())
                else "已导入，但还没有填任课安排"}

    def api_teacher_open(self, _=None):
        if not self.teacher.is_file():
            self.teacher.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.templates / "任课总表_模板.xlsx", self.teacher)
        open_file(self.teacher)
        return self._teacher_info()

    def teacher_upload(self, name: str, blob: bytes):
        tmp = self.tmp / "任课总表_待导入.xlsx"
        tmp.write_bytes(blob)
        s = self.school()
        n = sum(len(load_teachers(tmp, s, g)) for g in s.grades)      # 表头不对会抛出说明
        if not n:
            raise ValueError("这张表里没有读到任课安排，请检查是否填写。")
        self.teacher.parent.mkdir(parents=True, exist_ok=True)
        if self.teacher.is_file():
            shutil.copy2(self.teacher, self.teacher.with_name("任课总表_上一份.xlsx"))
        shutil.copy2(tmp, self.teacher)
        return self._teacher_info()

    # ---------------- 设置
    def api_settings(self, _=None):
        s = self.school()
        return {"school": s.name, "org": s.org_template, "full": s.full,
                "schemes": {k: {"weights": {i: v for i, v in s.structures.get(k, {}).items() if v}} for k in s.schemes},
                "gradeSubjects": s.grade_subjects, "teacher": self._teacher_info(), "license": lic.status(), "version": __version__}

    def api_settings_save(self, body):
        cur = {}
        if self.local.is_file():
            cur = yaml.safe_load(self.local.read_text(encoding="utf-8")) or {}
        s = self.school()
        if "school" in body:
            cur["学校"] = str(body["school"]).strip() or s.name
        if "org" in body:
            cur["落款"] = str(body["org"]).strip() or s.org_template
        if "full" in body:
            full = dict(s.full)
            for k, v in body["full"].items():
                if k in full:
                    try:
                        full[k] = max(float(v or 0), 0)
                    except (TypeError, ValueError):
                        raise ValueError(f"{k} 的满分要填数字。") from None
            cur["科目满分"] = {k: (int(v) if float(v).is_integer() else v) for k, v in full.items()}
        self.local.parent.mkdir(parents=True, exist_ok=True)
        self.local.write_text("# 只在这台电脑上生效的设置（在界面的“设置”里修改）\n"
                              + yaml.safe_dump(cur, allow_unicode=True, sort_keys=False), encoding="utf-8")
        return self.api_settings()

    # ---------------- 许可
    def api_activate(self, body):
        return lic.activate(str(body.get("code") or ""))


def make_handler(app: App, token: str):
    class H(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, *a):          # 不往控制台刷日志
            pass

        def _send(self, code, body: bytes, ctype="application/json; charset=utf-8"):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _json(self, code, obj):
            self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

        def do_GET(self):
            if urlparse(self.path).path in ("/", "/index.html"):
                html = UI_FILE.read_text(encoding="utf-8").replace("__TOKEN__", token)
                self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            else:
                self._send(404, b"{}")

        def do_POST(self):
            u = urlparse(self.path)
            if self.headers.get("X-Token") != token or not u.path.startswith("/api/"):
                return self._json(403, {"ok": False, "error": "拒绝访问"})
            name = u.path[5:]
            n = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(n) if n else b""
            try:
                with app.lock:
                    if name in ("upload", "teacher_upload"):
                        fn = (parse_qs(u.query).get("name") or ["登分表.xlsx"])[0]
                        data = getattr(app, name)(fn, raw)
                    else:
                        fnc = getattr(app, "api_" + name, None)
                        if fnc is None:
                            return self._json(404, {"ok": False, "error": "没有这个功能"})
                        data = fnc(json.loads(raw.decode("utf-8")) if raw else {})
                self._json(200, {"ok": True, "data": data})
            except lic.LicenseError as e:
                self._json(200, {"ok": False, "error": str(e), "license": True})
            except PermissionError as e:
                self._json(200, {"ok": False, "error": f"文件写不进去：{Path(e.filename or '').name}。多半是它正被 Excel 或 WPS 打开着，请先关掉再试。"})
            except (ValueError, RuntimeError) as e:
                self._json(200, {"ok": False, "error": str(e)})
            except Exception as e:                       # 没料到的错误：留下记录
                try:
                    (app.root / "出错记录.txt").write_text(traceback.format_exc(), encoding="utf-8")
                except Exception:
                    pass
                self._json(200, {"ok": False, "error": f"出错了：{type(e).__name__}: {e}（详细信息已存到程序文件夹里的“出错记录.txt”）"})
    return H


def start(app: App | None = None, port: int = 0):
    """启动本地服务，返回 (服务, 网址, App)。"""
    ensure_layout()
    app = app or App()
    token = secrets.token_urlsafe(18)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), make_handler(app, token))
    httpd.daemon_threads = True
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd, f"http://127.0.0.1:{httpd.server_address[1]}/", app
