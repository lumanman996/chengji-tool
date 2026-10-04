"""界面背后的本地服务：只在本机（127.0.0.1）上开一个端口，界面通过它调用计算、导出、设置。
外面的电脑访问不到；每次启动生成一个随机口令，界面带着口令才能调用。
唯一会上网的是检查和下载更新（updater.py，只访问 GitHub 上本项目的发布页）。
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

from . import CONTACT, __version__, license as lic, updater
from . import service as sv
from . import teachers as tt
from .config import load_school, load_teachers
from .paths import ROOT, ensure_layout

UI_FILE = Path(__file__).resolve().parent / "ui" / "index.html"


def newer_version(a: str, b: str) -> bool:
    """版本号 a 是不是比 b 新（2.1.0 > 2.0.9）。"""
    import re
    ta, tb = [int(x) for x in re.findall(r"\d+", a)], [int(x) for x in re.findall(r"\d+", b)]
    return bool(ta) and ta > tb


def auto_update_check() -> bool:
    """打开程序时要不要查一次新版本：打包好的程序查；源码运行（开发、测试）不查，除非特意打开。"""
    if os.environ.get("CHENGJI_NO_UPDATE_CHECK"):
        return False
    return updater.FROZEN or os.environ.get("CHENGJI_UPDATE_CHECK") == "1"


def friendly_path(path: Path, home: Path | None = None) -> str:
    """把文件夹位置说成人话：个人文件夹里的写成“文稿 → 分寸成绩核算”，不出现用户名；其他位置（如 D 盘）原样显示。"""
    home = (home or Path.home())
    try:
        parts = list(Path(path).resolve().relative_to(home.resolve()).parts)
    except (ValueError, OSError):
        return str(path)
    known = {"Documents": "文稿" if sys.platform == "darwin" else "文档", "Desktop": "桌面", "Downloads": "下载"}
    if not parts:
        return "个人文件夹"
    head = known.get(parts[0])
    return " → ".join(([head] if head else ["个人文件夹", parts[0]]) + parts[1:])


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


ALGO_KEYS = ["算法方案", "结构分方案"] + [y for _a, y, _p in sv.LINES.values()]      # 本机改过的算法、分数线


SAMPLE = ("示例登分表_九年级.xlsx", "示例任课总表.xlsx", "示例：九年级第一次月考")      # “用示例数据试一试”


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
        dl = Path(os.environ.get("CHENGJI_DOWNLOADS") or Path.home() / "Downloads")       # 模板“下载”到这里
        self.downloads = dl if dl.is_dir() else (Path.home() / "Desktop" if (Path.home() / "Desktop").is_dir() else Path.home())
        self.src: Path | None = None      # 当前登分表
        self.opts: dict | None = None     # 当前核算的设置
        self.result = None                # (R, 结论, 提示)
        self.lock = threading.Lock()
        self.last_ping = time.time()
        self.window = None                # 程序窗口（有它才能弹出系统的“存储”对话框）；浏览器方式下是 None
        self.updater = updater.Updater()  # 一键更新：下载、替换
        self.update_info = None

    def school(self):
        return load_school(self.config, self.local)

    # ---------------- 总体状态
    def api_state(self, _=None):
        s = self.school()
        loc = self._local()
        return {"version": __version__, "school": s.name, "nativeDialog": self.window is not None,
                "firstRun": not loc.get("已引导") and not loc.get("学校") and not sv.list_runs(self.output, 1),
                "contact": CONTACT, "autoUpdate": auto_update_check(), "fontSize": loc.get("界面字号") or "标准", "root": friendly_path(self.root), "hasSample": (self.root / "samples" / SAMPLE[0]).is_file(), "license": lic.status(), "recent": sv.list_runs(self.output),
                "templates": list(s.grade_subjects), "pdfDefaults": s.pdf_sections, "hasTeacherTable": self.teacher.is_file()}

    def api_welcome(self, body):
        """第一次打开的引导：记下学校名称（可以跳过），以后不再出现。"""
        cur = self._local()
        name = str(body.get("school") or "").strip()
        if name:
            cur["学校"] = name
            cur["落款"] = name + "{年级}组"
        cur["已引导"] = True
        self._write_local(cur)
        return self.api_state()

    def api_ui_save(self, body):
        cur = self._local()
        if body.get("fontSize") in ("标准", "大", "特大"):
            cur["界面字号"] = body["fontSize"]
        self._write_local(cur)
        return {"fontSize": cur.get("界面字号") or "标准"}

    def api_try_sample(self, _=None):
        """用自带的虚构示例数据直接算一遍，让新用户先看到效果。用示例任课总表，不存进“最近的考试”。"""
        f, t = self.root / "samples" / SAMPLE[0], self.root / "samples" / SAMPLE[1]
        if not f.is_file():
            raise ValueError("没有找到示例数据（samples 文件夹）。")
        school = self.school()
        opts = {"scheme": next(iter(school.schemes)), "exam": SAMPLE[2], "date": "", "prev": {"九1": 3, "九2": 6, "九3": 1, "九4": 5, "九5": 4, "九6": 2}}
        R, concl, msgs = sv.compute(f, school, opts, t if t.is_file() else None)
        view = sv.to_view(R, concl, msgs)
        view["sample"] = True
        self.src, self.opts, self.result = f, opts, (R, concl, msgs)
        return view

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
        info["last"] = sv.last_run_for(self.output, info["grade"], info["defaultName"])     # 同年级上次考试：名次、应考人数供预填
        return info

    def api_compute(self, opts):
        if not self.src:
            raise ValueError("请先导入登分表。")
        school = self.school()
        opts = {k: v for k, v in opts.items() if k != "schemeDef"}            # 比例只认设置里的，不认界面传来的
        R, concl, msgs = sv.compute(self.src, school, opts, self.teacher)
        view = sv.to_view(R, concl, msgs)
        self._remember_subjects(view)
        opts = {**opts, "schemeDef": sv.scheme_snapshot(school, view["scheme"])}   # 把这次用的比例一起存下来
        sv.save_run(self.output, self.src, opts, view)
        self.opts, self.result = {**opts, "exam": view["exam"]}, (R, concl, msgs)
        self.src = self.output / view["exam"] / sv.RUN_COPY
        return view

    def _remember_subjects(self, view):
        """登分表里新加的科目，算过一次就记进本机设置（连同这次的满分），下次自动认得。"""
        known = self.school().full
        new = {s: view["full"][s] for s in view["subjects"] if s not in known}
        if new:
            self.api_settings_save({"full": new})

    def api_open_run(self, body):
        meta = sv.load_run(self.output, body["name"])
        self.src = self.output / meta["view"]["exam"] / sv.RUN_COPY
        self.opts, self.result = meta["opts"], None
        return meta["view"]

    def api_run_view(self, body):
        """只读地取一次算过的考试的结果（用来和当前的结果对比），不改变当前打开的考试。"""
        return sv.load_run(self.output, body["name"])["view"]

    def api_run_delete(self, body):
        sv.delete_run(self.output, body["name"])
        return {"recent": sv.list_runs(self.output)}

    def api_run_rename(self, body):
        new = sv.rename_run(self.output, body["name"], body.get("to") or "")
        if self.opts and self.opts.get("exam") == sv.clean_name(body["name"]):      # 正开着的就是它：跟着改
            self.opts["exam"], self.result = new, None
            self.src = self.output / new / sv.RUN_COPY
        return {"recent": sv.list_runs(self.output), "name": new}

    def api_open_path(self, body):
        """打开刚导出的文件，或它所在的文件夹。只允许打开结果文件夹里的东西。"""
        p = Path(str(body.get("path") or "")).resolve()
        if self.output.resolve() not in p.parents or not p.exists():
            raise ValueError("找不到这个文件，可能已经被移走了。")
        (reveal if body.get("folder") else open_file)(p)
        return {}

    def _ensure_result(self):
        if self.result is None:
            if not (self.src and self.opts and Path(self.src).is_file()):
                raise ValueError("请先核算，或从“最近的考试”里打开一次结果。")
            self.result = sv.compute(self.src, self.school(), self.opts, self.teacher)
        return self.result

    # ---------------- 导出
    def _compare(self, R, body):
        """导出时带上“和上次比”：body 里的 compare 是和哪一次考试比（空 = 不比）。"""
        name = str((body or {}).get("compare") or "").strip()
        return sv.compare_with(R, self.output, name) if name else None

    def api_export_excel(self, body=None):
        lic.require_export()
        R, concl, _m = self._ensure_result()
        path = sv.export_excel(R, concl, self.output, self._compare(R, body))
        return {"path": str(path), "name": path.name}

    def api_export_pdf(self, body):
        lic.require_export()
        R, concl, _m = self._ensure_result()
        sections = body.get("sections") or []
        cmp_ = self._compare(R, body) if "和上次比" in sections else None
        path = sv.export_pdf(R, concl, sections, self.output, cmp_)
        return {"path": str(path), "name": path.name}

    def api_open_output(self, body=None):
        name = sv.clean_name((body or {}).get("name") or "")
        p = self.output / name if name and (self.output / name).is_dir() else self.output
        p.mkdir(parents=True, exist_ok=True)
        reveal(p)
        return {}

    # ---------------- 模板、任课总表
    def _template_file(self, key: str) -> tuple[Path, str]:
        """(模板文件, 建议的文件名)"""
        key = sv.clean_name(key or "")
        src = self.templates / (f"登分表_{key}.xlsx" if key != "任课总表" else "任课总表_模板.xlsx")
        if not src.is_file():
            raise ValueError("没有找到这个模板。")
        return src, (f"登分表_{key}.xlsx" if key != "任课总表" else "任课总表.xlsx")

    def api_template(self, body):
        """下载空白模板：弹出系统的“存储”对话框，让使用者自己选位置和文件名。点取消就什么也不做。"""
        src, name = self._template_file(body.get("key"))
        if self.window is None:
            raise ValueError("当前是浏览器方式，请用浏览器的下载。")
        import webview
        kind = getattr(getattr(webview, "FileDialog", None), "SAVE", None) or webview.SAVE_DIALOG
        res = self.window.create_file_dialog(kind, directory=str(self.downloads), save_filename=name,
                                             file_types=("Excel 文件 (*.xlsx)",))
        if not res:
            return {"cancelled": True}
        dst = Path(res if isinstance(res, str) else res[0])
        if dst.suffix.lower() != ".xlsx":
            dst = dst.with_name(dst.name + ".xlsx")
        shutil.copy2(src, dst)
        return {"name": dst.name, "folder": dst.parent.name, "path": str(dst)}

    def template_bytes(self, key: str) -> tuple[bytes, str]:
        """浏览器方式下用：把模板内容交给浏览器，由浏览器自己的下载来保存。"""
        src, name = self._template_file(key)
        return src.read_bytes(), name

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

    def api_teacher_get(self, _=None):
        """任课总表的内容，给界面里的编辑器用。"""
        s = self.school()
        return {**tt.read_table(self.teacher, s), "grades": list(s.grades), "gradeChar": dict(s.grades), "subjects": s.subjects}

    def api_teacher_save(self, body):
        s = self.school()
        data = tt.check_table(body, s)                       # 有问题在这里就报出来，不会写坏文件
        if self.teacher.is_file():
            shutil.copy2(self.teacher, self.teacher.with_name("任课总表_上一份.xlsx"))
        tt.write_table(self.teacher, data, "任课总表" if s.name.startswith("××") else f"{s.name} 任课总表")
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
                "algo": sv.algo_view(s), "algoDefault": sv.algo_view(load_school(self.config)), "lines": sv.lines_view(s),
                "algoCustom": any(k in self._local() for k in ALGO_KEYS),
                "gradeSubjects": s.grade_subjects, "teacher": self._teacher_info(), "license": lic.status(), "version": __version__}

    def _local(self) -> dict:
        if self.local.is_file():
            return yaml.safe_load(self.local.read_text(encoding="utf-8")) or {}
        return {}

    def _write_local(self, cur: dict):
        self.local.parent.mkdir(parents=True, exist_ok=True)
        self.local.write_text("# 只在这台电脑上生效的设置（在界面的“设置”里修改）\n"
                              + yaml.safe_dump(cur, allow_unicode=True, sort_keys=False), encoding="utf-8")

    def api_algo_reset(self, _=None):
        """算法恢复成默认（去掉本机改过的比例）。"""
        cur = self._local()
        for k in ALGO_KEYS:
            cur.pop(k, None)
        self._write_local(cur)
        return self.api_settings()

    def api_settings_save(self, body):
        cur = self._local()
        s = self.school()
        if "algo" in body:                                   # 界面里改的比例：检查合计是不是 100%
            cur["算法方案"], cur["结构分方案"] = sv.algo_to_yaml(body["algo"], s)
        if "lines" in body:
            cur.update(sv.lines_to_yaml(body["lines"]))
        if "school" in body:
            cur["学校"] = str(body["school"]).strip() or s.name
        if "org" in body:
            cur["落款"] = str(body["org"]).strip() or s.org_template
        if "full" in body:
            from .loader import is_meta_column
            full = dict(s.full)
            for k, v in body["full"].items():
                k = sv.clean_name(k).replace(" ", "")
                if not k:
                    continue
                if k not in full and (is_meta_column(k) or k in ("姓名", "班级")):
                    raise ValueError(f"“{k}”不能当科目名称。")
                try:
                    full[k] = max(float(v or 0), 0)               # 已有的科目改满分；没有的科目就是新加
                except (TypeError, ValueError):
                    raise ValueError(f"{k} 的满分要填数字。") from None
            cur["科目满分"] = {k: (int(v) if float(v).is_integer() else v) for k, v in full.items()}
        self._write_local(cur)
        return self.api_settings()

    # ---------------- 帮助、关于
    def api_open_root(self, _=None):
        reveal(self.root)
        return {}

    def api_update_check(self, _=None):
        """查有没有新版本（只读 GitHub 上本项目的发布页，不上传任何东西）。打开程序时查一次，帮助页里也能手动查。"""
        self.update_info = updater.check()
        return {k: v for k, v in self.update_info.items() if k != "asset"}

    def api_update_start(self, _=None):
        info = self.update_info
        if not info or not info.get("newer") or not info.get("canAuto"):
            raise ValueError((info or {}).get("reason") or "没有可以一键安装的新版本。")
        self.updater.start(info["asset"])
        return self.updater.status()

    def api_update_status(self, _=None):
        return self.updater.status()

    def api_update_apply(self, _=None):
        """换上新版本并重新打开：先启动替换脚本，再退出本程序。"""
        self.updater.apply()
        threading.Timer(0.8, lambda: os._exit(0)).start()   # 留一点时间把这次回应发回界面
        return {}

    def api_open_repo(self, _=None):
        """用系统浏览器打开项目主页（只有使用者点了才会打开）。"""
        if not os.environ.get("CHENGJI_NO_OPEN"):
            import webbrowser
            webbrowser.open(CONTACT["github"])
        return {}

    def api_open_releases(self, _=None):
        if not os.environ.get("CHENGJI_NO_OPEN"):
            import webbrowser
            webbrowser.open(updater.RELEASES_PAGE)
        return {}

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
                if name == "template_file":                  # 返回文件本身，不是 JSON
                    blob, _fn = app.template_bytes(json.loads(raw.decode("utf-8") or "{}").get("key"))
                    return self._send(200, blob, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
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
