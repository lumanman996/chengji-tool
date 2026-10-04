"""总入口：没有参数就打开图形界面；带参数就走命令行（自动测试、批量处理用）。
--自检、--apply-update <安装包.zip | online> [--then-selftest]、--检查更新 是打包后自检用的。"""
from __future__ import annotations

import sys
import traceback


def _log_error():
    from .paths import ROOT
    try:
        ROOT.mkdir(parents=True, exist_ok=True)
        (ROOT / "出错记录.txt").write_text(traceback.format_exc(), encoding="utf-8")
    except Exception:
        pass


def _safe_output():
    """让输出永远不会因为编码报错：
    输出被别的程序接走时（管道、重定向）一律用 UTF-8；显示在终端里时沿用终端的编码，遇到显示不了的字用 ? 代替。
    带窗口的打包程序没有输出通道（None），跳过。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is None:
                continue
            if stream.isatty():
                stream.reconfigure(errors="replace")
            else:
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def launch(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    _safe_output()
    headless = bool(argv) and argv[0] != "--browser"          # 命令行 / 自检：没有人在看
    try:
        if argv and argv[0] == "--自检":
            from .gui import selftest
            sys.exit(selftest())
        if argv and argv[0] == "--apply-update":                # 打包自检：用安装包更新自己，再重新打开
            from .updater import apply_from_command
            sys.exit(apply_from_command(argv[1] if len(argv) > 1 else "", ["--自检"] if "--then-selftest" in argv else []))
        if argv and argv[0] == "--检查更新":                    # 打包自检：联网查一次版本（看证书等是否正常）
            import json
            from .updater import check
            print(json.dumps({k: v for k, v in check().items() if k != "notes"}, ensure_ascii=False), flush=True)
            sys.exit(0)
        if headless:
            from .cli import run
            return run(argv)
        from .gui import main
        main(browser="--browser" in argv)
    except SystemExit:
        raise
    except Exception:
        _log_error()
        if headless:                                          # 不弹对话框，直接以错误码退出
            sys.exit(1)
        raise                                                 # 窗口模式：让系统把错误显示给用户
