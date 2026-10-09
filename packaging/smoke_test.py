"""打包后的试跑：用打好的程序（不开窗口）自检，并算一遍虚构示例数据，检查结果是否齐全、数字是否正确；
再用刚打好的安装包走一遍一键更新（替换程序文件 → 自动重新打开），检查数据文件夹原样保留。

用法：python packaging/smoke_test.py      （先运行 packaging/build.py）
数据放在 dist/_试跑 里，不碰真实的数据文件夹。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from chengji import __version__   # noqa: E402

DIST = ROOT / "dist"
if sys.platform == "darwin":
    exe = DIST / "分寸.app" / "Contents" / "MacOS" / "分寸"
else:
    exe = DIST / "分寸" / ("分寸.exe" if os.name == "nt" else "分寸")
HOME = DIST / "_试跑"
ENV = {**os.environ, "CHENGJI_NO_OPEN": "1", "CHENGJI_HOME": str(HOME)}


def run(*args):
    print(">", " ".join(str(a) for a in args))
    r = subprocess.run([str(exe), *map(str, args)], cwd=HOME, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       timeout=600, env=ENV)
    if r.stdout.strip():
        print(r.stdout[-1200:])
    if r.returncode:
        print(r.stderr[-3000:])
        log = HOME / "出错记录.txt"
        if log.is_file():
            print(log.read_text(encoding="utf-8")[-3000:])
        sys.exit(f"程序退出码 {r.returncode}")


def main():
    assert exe.is_file(), f"没有找到 {exe}，请先运行 packaging/build.py"
    shutil.rmtree(HOME, ignore_errors=True)
    HOME.mkdir(parents=True)

    # 1) 自检：界面文件、本地服务、许可状态；自带的设置、模板、示例会放到数据文件夹
    run("--自检")
    info = json.loads((HOME / "自检结果.json").read_text(encoding="utf-8"))
    print("自检：", info)
    assert info["ok"] and info["license"] in ("open", "trial", "active", "expired"), info
    licensed = (ROOT / "chengji" / "_license_impl.py").is_file()
    if licensed:                                           # 正式版：验证模块必须是编译成机器码的，安装包里没有它的源码
        assert info.get("licenseModule") == "native", f"验证模块不是机器码：{info.get('licenseModule')}"
        tamper_check()
    expired = info["license"] == "expired"                 # 在本机试跑时，这台电脑的试用可能已经到期（GitHub 上每次都是新电脑）
    if expired:
        print("（提醒）这台电脑的试用已到期：跳过需要导出的两步试跑，直接检查一键更新。在 GitHub 上打包时每次都是完整试跑。")
    for f in ("config/学校设置.yaml", "templates/登分表_九年级.xlsx", "samples/示例登分表_九年级.xlsx"):
        assert (HOME / f).is_file(), f"没有放出 {f}"
    T = HOME / "samples" / "示例任课总表.xlsx"

    if not expired:
        try_samples(T)
    check_online()
    update_flow(info["license"])
    shutil.rmtree(HOME, ignore_errors=True)
    print("试跑全部通过。")


def try_samples(T):
    # 2) 九年级、平时方案、全部 PDF 内容
    name = "试跑_九年级"
    run(HOME / "samples" / "示例登分表_九年级.xlsx", "--任课", T, "--考试", name, "--上次名次", "九1=3 九2=6 九3=1 九4=5 九5=4 九6=2", "--不确认")
    out = HOME / "output" / name
    xlsx, pdf, note = out / f"{name}_各班综合统计.xlsx", out / f"{name}_成绩发布版.pdf", out / "核对说明.txt"
    assert xlsx.is_file() and xlsx.stat().st_size > 20000, "没有生成 Excel"
    text = note.read_text(encoding="utf-8")
    assert "九6 46.45（第1）" in text, "结构分结果不对：\n" + text[:600]
    if "--no-pdf" not in sys.argv:
        assert pdf.is_file() and pdf.stat().st_size > 50000, "没有生成 PDF（电脑上要有 Edge 或 Chrome）"

    # 3) 中考核算（10 科，其中生物、地理、体育要当场给满分）、不要 PDF
    name = "试跑_中考核算"
    run(HOME / "samples" / "示例登分表_中考.xlsx", "--任课", T, "--考试", name, "--方案", "中考", "--满分", "生物=50 地理=50 体育=50", "--PDF内容", "0", "--不确认")
    assert (HOME / "output" / name / f"{name}_各班综合统计.xlsx").is_file()


def tamper_check():
    """在程序的一份副本里删掉验证模块：正式版应当锁住导出，而不是变成不限制的版本。"""
    tmp = DIST / "_篡改测试"
    shutil.rmtree(tmp, ignore_errors=True)
    if sys.platform == "darwin":
        app = tmp / "分寸.app"
        subprocess.run(["ditto", str(DIST / "分寸.app"), str(app)], check=True)
        copy_exe = app / "Contents" / "MacOS" / "分寸"
        found = list(app.rglob("_license_impl*.so"))
    else:
        shutil.copytree(exe.parent, tmp / "分寸")
        copy_exe = tmp / "分寸" / exe.name
        found = list((tmp / "分寸").rglob("_license_impl*.pyd")) + list((tmp / "分寸").rglob("_license_impl*.so"))
    assert found, "安装包里没有找到编译后的验证模块"
    assert not list(tmp.rglob("_license_impl.py")), "安装包里不应该有验证模块的源码"
    for f in found:
        f.unlink()
    home = tmp / "数据"
    home.mkdir()
    subprocess.run([str(copy_exe), "--自检"], cwd=home, capture_output=True, timeout=300,
                   env={**os.environ, "CHENGJI_NO_OPEN": "1", "CHENGJI_HOME": str(home)})
    info = json.loads((home / "自检结果.json").read_text(encoding="utf-8"))
    print("删掉验证模块后：", info)
    assert info["licenseModule"] == "none" and not info["canExport"] and info["license"] == "expired", "删掉验证模块后导出没有锁住"
    shutil.rmtree(tmp, ignore_errors=True)
    print("防篡改检查通过：验证模块是机器码；删掉它，导出就锁住。")


def check_online():
    """联网查一次版本：主要看打包后的程序能不能正常建立 HTTPS 连接（证书）。GitHub 偶尔限流，那种情况只提醒。"""
    r = subprocess.run([str(exe), "--检查更新"], cwd=HOME, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120, env=ENV)
    line = (r.stdout.strip().splitlines() or ["{}"])[-1]
    print("检查更新：", line)
    res = json.loads(line)
    err = str(res.get("error") or "")
    if "CERTIFICATE" in err.upper() or "SSL" in err.upper():
        sys.exit("检查更新时 HTTPS 证书有问题：" + err)
    if not res.get("ok"):
        print("（提醒）这次没连上 GitHub，不算失败：", err)


def update_flow(license_before: str):
    """拿刚打好的安装包当“新版本”，让程序把自己换掉再重新打开。"""
    system = "mac" if sys.platform == "darwin" else "windows"
    zpath = DIST / f"chengji-tool-v{__version__}-{system}.zip"
    assert zpath.is_file(), f"没有找到安装包 {zpath}"
    result = HOME / "自检结果.json"
    if sys.platform == "darwin":
        app = DIST / "分寸.app"
        marker = app / "Contents" / "Resources" / "旧版本标记.txt"           # 整个 .app 被换掉后它应当消失
        marker.write_text("old", encoding="utf-8")
        kept = {HOME / "output" / "更新测试" / "结果.json": "旧的结果", HOME / "data" / "本校设置.yaml": "学校: 更新测试学校\n"}
    else:
        prog = exe.parent
        marker = prog / "使用说明.txt"                                      # 被新版覆盖后内容应当恢复
        marker.write_text("old", encoding="utf-8")
        cfg = prog / "config" / "学校设置.yaml"
        kept = {prog / "output" / "更新测试" / "结果.json": "旧的结果", prog / "data" / "本校设置.yaml": "学校: 更新测试学校\n",
                cfg: cfg.read_text(encoding="utf-8") + "\n# 本机改过的设置，更新后要还在\n"}
    for f, text in kept.items():
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(text, encoding="utf-8")
    result.unlink(missing_ok=True)
    log = HOME / "更新日志.txt"
    log.unlink(missing_ok=True)
    print(">", exe, "--apply-update", zpath.name, "--then-selftest")
    code = subprocess.run([str(exe), "--apply-update", str(zpath), "--then-selftest"], cwd=HOME, env=ENV, timeout=180, check=False,
                          stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode
    print("更新命令退出码:", code)
    deadline = time.time() + 240
    while time.time() < deadline and not result.exists():
        time.sleep(1)
    time.sleep(2)
    print("更新日志:\n" + (log.read_text(encoding="utf-8", errors="replace") if log.exists() else "（没有日志）"))
    if not result.exists():
        sys.exit("更新自检未通过：换上新版本后程序没有重新打开。")
    info = json.loads(result.read_text(encoding="utf-8"))
    replaced = (not marker.exists()) if sys.platform == "darwin" else (marker.read_text(encoding="utf-8") != "old")
    lost = [str(f) for f, text in kept.items() if not f.is_file() or f.read_text(encoding="utf-8") != text]
    print("更新后自检：", info, "| 程序文件已替换:", replaced, "| 数据丢失:", lost or "无")
    assert info["ok"] and exe.is_file() and replaced, "更新自检未通过"
    assert not lost, f"更新后这些数据文件变了或没了：{lost}"
    assert info["license"] == license_before, f"更新前后许可状态不一样：{license_before} → {info['license']}"
    if sys.platform != "darwin":                                            # 把测试放进程序文件夹的东西收拾掉
        shutil.rmtree(prog / "output" / "更新测试", ignore_errors=True)
        (prog / "data" / "本校设置.yaml").unlink(missing_ok=True)
        cfg.write_text(kept[cfg].replace("\n# 本机改过的设置，更新后要还在\n", ""), encoding="utf-8")
    print("一键更新自检通过。")


if __name__ == "__main__":
    main()
