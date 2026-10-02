"""打包成双击就能用的程序（在哪种电脑上运行，就打出哪种电脑用的程序）。

用法：python packaging/build.py
结果：
  Windows  dist/分寸/                    分寸.exe + 程序文件 + config + templates + samples + data + 使用说明
           dist/chengji-tool-v版本-windows.zip
  Mac      dist/分寸.app                （设置和结果放在“文稿/分寸成绩核算”文件夹）
           dist/chengji-tool-v版本-mac.zip

有 chengji/_license_impl.py（试用与激活模块，不在公开仓库里）时会一起打进去；没有就是不限制的版本。
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

NAME = "分寸"
DIST = ROOT / "dist"
MAC = sys.platform == "darwin"
WIN = os.name == "nt"


def main():
    for d in (ROOT / "build", DIST):
        shutil.rmtree(d, ignore_errors=True)
    licensed = (ROOT / "chengji" / "_license_impl.py").is_file()
    sep = os.pathsep                                            # Windows 是 ;  其他系统是 :
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir", "--windowed",
           "--name", NAME, "--distpath", str(DIST), "--workpath", str(ROOT / "build"), "--specpath", str(ROOT / "build"),
           "--paths", str(ROOT), "--icon", str(ROOT / "packaging" / ("icon.icns" if MAC else "icon.ico")),
           "--add-data", f"{ROOT / 'chengji' / 'ui'}{sep}chengji/ui"]
    for folder in ("config", "templates", "samples"):           # 程序里自带一份，第一次运行时放到数据文件夹
        cmd += ["--add-data", f"{ROOT / folder}{sep}{folder}"]
    for mod in ("playwright", "pytest", "matplotlib", "tkinter", "IPython", "PIL"):
        cmd += ["--exclude-module", mod]
    if MAC:
        cmd += ["--osx-bundle-identifier", "cn.chengji.scorebook"]
    cmd.append(str(ROOT / "packaging" / "entry.py"))
    subprocess.run(cmd, check=True, cwd=ROOT)

    system = {"Windows": "windows", "Darwin": "mac", "Linux": "linux"}.get(platform.system(), platform.system().lower())
    zpath = DIST / f"chengji-tool-v{__version__}-{system}.zip"
    guide = (ROOT / "packaging" / "使用说明.txt").read_text(encoding="utf-8")
    if MAC:
        app = DIST / f"{NAME}.app"
        shutil.rmtree(DIST / NAME, ignore_errors=True)          # onedir 的中间文件夹，Mac 上只要 .app
        (DIST / "使用说明.txt").write_text(guide, encoding="utf-8")
        # 用 ditto 压缩：保留 .app 里的符号链接和权限（普通 zip 会把程序包弄坏）
        stage = DIST / "_stage" / NAME
        stage.mkdir(parents=True)
        subprocess.run(["ditto", str(app), str(stage / app.name)], check=True)
        shutil.copy2(DIST / "使用说明.txt", stage / "使用说明.txt")
        subprocess.run(["ditto", "-c", "-k", "--keepParent", str(stage), str(zpath)], check=True)
        shutil.rmtree(DIST / "_stage")
        target = app
    else:
        app = DIST / NAME
        for folder in ("config", "templates", "samples"):
            shutil.copytree(ROOT / folder, app / folder)
        (app / "data").mkdir()
        (app / "data" / "把任课总表放这里.txt").write_text(
            "任课总表会保存在这个文件夹里（在软件的“设置”里导入或修改）。\n", encoding="utf-8")
        (app / "使用说明.txt").write_text(guide, encoding="utf-8")
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(app.rglob("*")):
                z.write(f, Path(NAME) / f.relative_to(app))
        target = app
    size = zpath.stat().st_size / 1024 / 1024
    print(f"\n打包完成：{target}\n压缩包：{zpath}（{size:.1f} MB）\n试用与激活模块：{'已包含' if licensed else '没有（不限制的版本）'}")


if __name__ == "__main__":
    main()
