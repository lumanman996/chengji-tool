"""总入口：没有参数就打开图形界面；带参数就走命令行（自动测试、批量处理用）。"""
from __future__ import annotations

import sys
import traceback


def launch(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] == "--自检":
        from .gui import selftest
        sys.exit(selftest())
    if argv and argv[0] != "--browser":
        from .cli import run
        return run(argv)
    try:
        from .gui import main
        main(browser="--browser" in argv)
    except Exception:                                        # 窗口程序没有黑窗口可看，出错要留下记录
        from .paths import ROOT
        try:
            ROOT.mkdir(parents=True, exist_ok=True)
            (ROOT / "出错记录.txt").write_text(traceback.format_exc(), encoding="utf-8")
        except Exception:
            pass
        raise
