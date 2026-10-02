"""一键全面核对：用虚构示例数据把各年级、各方案都算一遍，
用 LibreOffice 重算导出的 Excel，再逐项对照程序的独立计算结果。必须 0 处不一致。

用法：python tests/verify_all.py        （需要先装好 LibreOffice）
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from verify_excel import verify   # noqa: E402

T = str(ROOT / "samples" / "示例任课总表.xlsx")
PREV = "九1=3 九2=6 九3=1 九4=5 九5=4 九6=2"
# (登分表, 方案, 升学比例, 上次名次, 满分)
FULL = "生物=50 地理=50 体育=50"      # 这三科在学校设置里满分还没定，示例数据按 50 分造的
CASES = [("七年级", "平时", None, "", FULL),
         ("八年级", "平时", None, "", FULL),
         ("九年级", "平时", None, PREV, ""),
         ("九年级", "中考", 0.8, "", ""),          # 只有 7 科的表用中考方案（会提示少了 3 科，照样要算对）
         ("中考", "中考", 0.8, "", FULL)]


def find_soffice() -> str:
    for c in ("soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        w = shutil.which(c) or (c if Path(c).is_file() else None)
        if w:
            return w
    sys.exit("没有找到 LibreOffice（soffice），无法重算 Excel 公式。")


def main():
    soffice = find_soffice()
    total = bad = 0
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        for i, (grade, scheme, ratio, prev, full) in enumerate(CASES):
            scores = str(ROOT / "samples" / f"示例登分表_{grade}.xlsx")
            name = f"核对{i}"
            cmd = [sys.executable, "-m", "chengji", scores, "--任课", T, "--考试", name, "--方案", scheme,
                   "--输出", str(tmp / "out"), "--不生成PDF", "--不确认"]
            cmd += ["--升学比例", str(ratio)] if ratio else []
            cmd += ["--上次名次", prev] if prev else []
            cmd += ["--满分", full] if full else []
            subprocess.run(cmd, check=True, cwd=ROOT, capture_output=True)
            xlsx = tmp / "out" / name / f"{name}_各班综合统计.xlsx"
            subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp / "re" / name), str(xlsx)],
                           check=True, capture_output=True, timeout=300)
            n, errs = verify(scores, tmp / "re" / name / xlsx.name, T, scheme, ratio, "", prev, full)
            total += n
            bad += len(errs)
            print(f"{grade} {scheme}方案：核对 {n} 项，不一致 {len(errs)} 项")
            for e in errs[:20]:
                print("   ", e)
    print(f"合计核对 {total} 项，不一致 {bad} 项")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
