"""打包后的试跑：用打好的程序算一遍虚构示例数据，检查结果文件是否齐全、数字是否正确。

用法：python packaging/smoke_test.py      （先运行 packaging/build.py）
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
APP = ROOT / "dist" / "成绩核算工具"
exe = APP / ("成绩核算工具.exe" if os.name == "nt" else "成绩核算工具")


def run(args, **kw):
    print(">", " ".join(str(a) for a in args[1:]))
    r = subprocess.run([str(exe), *args[1:]], cwd=APP, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=300, env={**os.environ, "CHENGJI_NO_OPEN": "1"}, **kw)
    print(r.stdout[-1500:])
    if r.returncode:
        print(r.stderr[-3000:])
        sys.exit(f"程序退出码 {r.returncode}")
    return r


def main():
    assert exe.is_file(), f"没有找到 {exe}，请先运行 packaging/build.py"
    need_pdf = "--no-pdf" not in sys.argv

    # 1) 九年级、平时方案、全部 PDF 内容
    name = "试跑_九年级"
    run([exe, "samples/示例登分表_九年级.xlsx", "--任课", "samples/示例任课总表.xlsx", "--考试", name,
         "--上次名次", "九1=3 九2=6 九3=1 九4=5 九5=4 九6=2", "--不确认"])
    out = APP / "output" / name
    xlsx, pdf, note = out / f"{name}_各班综合统计.xlsx", out / f"{name}_成绩发布版.pdf", out / "核对说明.txt"
    assert xlsx.is_file() and xlsx.stat().st_size > 20000, "没有生成 Excel"
    text = note.read_text(encoding="utf-8")
    assert "九6 46.45（第1）" in text, "结构分结果不对：\n" + text[:600]
    if need_pdf:
        assert pdf.is_file() and pdf.stat().st_size > 50000, "没有生成 PDF（电脑上要有 Edge 或 Chrome）"

    # 2) 中考核算（10 科，其中生物、地理、体育要当场给满分）、不要 PDF
    name = "试跑_中考核算"
    run([exe, "samples/示例登分表_中考.xlsx", "--任课", "samples/示例任课总表.xlsx", "--考试", name, "--方案", "中考",
         "--满分", "生物=50 地理=50 体育=50", "--PDF内容", "0", "--不确认"])
    assert (APP / "output" / name / f"{name}_各班综合统计.xlsx").is_file()
    assert not (APP / "output" / name / f"{name}_成绩发布版.pdf").exists()

    # 3) 程序文件夹里的设置丢了，也能自己放出来
    import shutil
    shutil.rmtree(APP / "config")
    run([exe, "samples/示例登分表_七年级.xlsx", "--任课", "samples/示例任课总表.xlsx", "--考试", "试跑_七年级",
         "--满分", "生物=50 地理=50", "--PDF内容", "0", "--不确认"])
    assert (APP / "config" / "学校设置.yaml").is_file(), "没有自动放出学校设置"

    shutil.rmtree(APP / "output")                       # 试跑结果不留在发布包里
    (APP / "output").mkdir()
    print("试跑全部通过。")


if __name__ == "__main__":
    main()
