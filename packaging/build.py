"""打包成免安装程序（在哪种电脑上运行，就打出哪种电脑用的程序）。

用法：python packaging/build.py
结果：dist/成绩核算工具/            程序 + config + templates + samples + data + 使用说明
      dist/chengji-tool-v版本-系统.zip   上面这个文件夹的压缩包（发布用）
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from chengji import __version__   # noqa: E402

NAME = "成绩核算工具"
DIST = ROOT / "dist"
APP = DIST / NAME


def main():
    for d in (ROOT / "build", DIST):
        shutil.rmtree(d, ignore_errors=True)
    sep = os.pathsep                                            # Windows 是 ;  其他系统是 :
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--console",
           "--name", NAME, "--distpath", str(APP), "--workpath", str(ROOT / "build"), "--specpath", str(ROOT / "build"),
           "--paths", str(ROOT)]
    for folder in ("config", "templates", "samples"):           # 程序里自带一份，文件夹丢了也能自己放出来
        cmd += ["--add-data", f"{ROOT / folder}{sep}{folder}"]
    for mod in ("playwright", "pytest", "matplotlib", "tkinter", "IPython", "PIL"):
        cmd += ["--exclude-module", mod]
    cmd.append(str(ROOT / "packaging" / "entry.py"))
    subprocess.run(cmd, check=True, cwd=ROOT)

    for folder in ("config", "templates", "samples"):
        shutil.copytree(ROOT / folder, APP / folder)
    (APP / "data").mkdir()
    (APP / "data" / "把任课总表放这里.txt").write_text(
        "把填好的任课总表存成“任课总表.xlsx”放在这个文件夹里。\n登分表也可以放在这里。\n", encoding="utf-8")
    shutil.copy2(ROOT / "packaging" / "使用说明.txt", APP / "使用说明.txt")

    system = {"Windows": "windows", "Darwin": "mac", "Linux": "linux"}.get(platform.system(), platform.system().lower())
    zpath = DIST / f"chengji-tool-v{__version__}-{system}.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(APP.rglob("*")):
            z.write(f, Path(NAME) / f.relative_to(APP))
    size = zpath.stat().st_size / 1024 / 1024
    print(f"\n打包完成：{APP}\n压缩包：{zpath}（{size:.1f} MB）")


if __name__ == "__main__":
    main()
