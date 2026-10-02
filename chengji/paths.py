"""程序的文件夹在哪里。

- 用源码运行（python -m chengji）：项目文件夹。
- 用打包好的程序运行：
    Windows：程序文件所在的文件夹（解压到哪儿，设置和结果就在哪儿）；
    Mac：“文稿/成绩核算”文件夹（Mac 的程序包里不能写东西）。
  第一次运行时，会把自带的学校设置、模板、示例数据放到这个文件夹里，方便修改和使用。
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))


def app_root() -> Path:
    if os.environ.get("CHENGJI_HOME"):                     # 指定数据文件夹（自动测试用）
        return Path(os.environ["CHENGJI_HOME"]).resolve()
    if FROZEN and sys.platform == "darwin":
        return Path.home() / "Documents" / "成绩核算"
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = app_root()


def ensure_layout() -> list[str]:
    """免安装程序第一次运行：把自带的 config / templates / samples 放到程序旁边，并建好 data、output 文件夹。
    已经有的文件不会被覆盖。返回本次新放出来的内容。"""
    made = []
    ROOT.mkdir(parents=True, exist_ok=True)
    for name in ("data", "output"):
        if not (ROOT / name).exists():
            (ROOT / name).mkdir(parents=True)
            made.append(name)
    if FROZEN:
        bundled = Path(getattr(sys, "_MEIPASS", ROOT))
        for folder in ("config", "templates", "samples"):
            src = bundled / folder
            if not src.is_dir():
                continue
            for f in src.iterdir():
                dst = ROOT / folder / f.name
                if f.is_file() and not dst.exists():
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(f, dst)
                    made.append(f"{folder}/{f.name}")
    return made
