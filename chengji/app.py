"""总入口：没有参数就打开图形界面；带参数就走命令行（自动测试、批量处理用）。"""
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


def launch(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    headless = bool(argv) and argv[0] != "--browser"          # 命令行 / 自检：没有人在看
    try:
        if argv and argv[0] == "--自检":
            from .gui import selftest
            sys.exit(selftest())
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
