"""程序的文件夹在哪里。

- 用源码运行（python -m chengji）：项目文件夹。
- 用打包好的免安装程序运行：程序文件所在的文件夹。第一次运行时，会把自带的
  学校设置、模板、示例数据放到程序旁边，方便修改和使用。
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))


def app_root() -> Path:
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = app_root()


def ensure_layout() -> list[str]:
    """免安装程序第一次运行：把自带的 config / templates / samples 放到程序旁边，并建好 data、output 文件夹。
    已经有的文件不会被覆盖。返回本次新放出来的内容。"""
    made = []
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
