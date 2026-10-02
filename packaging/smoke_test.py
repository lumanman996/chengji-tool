"""打包后的试跑：用打好的程序（不开窗口）自检，并算一遍虚构示例数据，检查结果是否齐全、数字是否正确。

用法：python packaging/smoke_test.py      （先运行 packaging/build.py）
数据放在 dist/_试跑 里，不碰真实的数据文件夹。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
if sys.platform == "darwin":
    exe = DIST / "成绩核算.app" / "Contents" / "MacOS" / "成绩核算"
else:
    exe = DIST / "成绩核算" / ("成绩核算.exe" if os.name == "nt" else "成绩核算")
HOME = DIST / "_试跑"


def run(*args):
    print(">", " ".join(str(a) for a in args))
    r = subprocess.run([str(exe), *map(str, args)], cwd=HOME, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       timeout=600, env={**os.environ, "CHENGJI_NO_OPEN": "1", "CHENGJI_HOME": str(HOME)})
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
    assert info["ok"] and info["license"] in ("open", "trial", "active"), info
    for f in ("config/学校设置.yaml", "templates/登分表_九年级.xlsx", "samples/示例登分表_九年级.xlsx"):
        assert (HOME / f).is_file(), f"没有放出 {f}"
    T = HOME / "samples" / "示例任课总表.xlsx"

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

    shutil.rmtree(HOME)
    print("试跑全部通过。")


if __name__ == "__main__":
    main()
