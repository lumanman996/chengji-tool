"""程序入口。

- 双击免安装程序，或运行 python -m chengji：把登分表拖进窗口，按提示回答几个问题。
- 熟悉命令行的也可以：python -m chengji 登分表.xlsx --考试 名称 --方案 中考 …（见 --help）

读入登分表后会列出识别到的年级、科目和满分，确认后开始计算，生成：
  output/<考试名称>/<考试名称>_各班综合统计.xlsx
  output/<考试名称>/<考试名称>_成绩发布版.pdf
  output/<考试名称>/核对说明.txt
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import traceback
from pathlib import Path

from . import APP_TITLE, __version__, license as lic
from .analysis import analyze, draft_conclusions
from .config import load_enrolled, load_school, load_teachers, make_config, normalize_class
from .excel_report import build_excel
from .loader import check_data, load_scores
from .paths import FROZEN, ROOT, ensure_layout
from .service import clean_name, enrolment_notes, subject_note as _subject_note


def has_tty() -> bool:
    """有没有可以问答的终端。带窗口的打包程序（尤其 Windows）没有标准输入，sys.stdin 是 None。"""
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except Exception:
        return False


def clean_path(text: str) -> str:
    """把拖进窗口的文件路径整理干净：去掉引号、PowerShell 的 & 前缀、Mac 终端加的反斜杠。"""
    t = text.strip()
    if t.startswith("& "):
        t = t[2:].strip()
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "'\"":
        t = t[1:-1]
    if t and not Path(t).exists() and "\\" in t and os.name != "nt":
        t = re.sub(r"\\(.)", r"\1", t)
    return t


def parse_full_overrides(text: str) -> dict[str, float]:
    out = {}
    for part in re.split(r"[\s,，、;；]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(.+?)[=＝:：](\d+(?:\.\d+)?)", part)
        if not m:
            raise ValueError(f"看不懂“{part}”，请写成 科目=满分，例如 数学=120")
        out[m.group(1)] = float(m.group(2))
    return out


def parse_class_values(text: str, grade_char: str) -> dict[str, int]:
    """“九1=46 九2=45”“91=3、92=1” → {"九1": 46, "九2": 45}"""
    out = {}
    for part in re.split(r"[\s,，、;；]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(.+?)[=＝:：](\d+)", part)
        if not m:
            raise ValueError(f"看不懂“{part}”，请写成 班级=数字，例如 {grade_char}1=46")
        out[normalize_class(m.group(1), grade_char)] = int(m.group(2))
    return out


def ask_class_values(prompt: str, classes: list[str], grade_char: str, lo: int, hi: int | None,
                     allow_empty: bool) -> dict[str, int]:
    """逐次输入“班级=数字”，直到所有班级都有；allow_empty 时直接回车表示没有。"""
    got: dict[str, int] = {}
    while True:
        left = [c for c in classes if c not in got]
        if not left:
            return got
        ans = input(prompt.format(left="、".join(left))).strip()
        if not ans:
            if allow_empty and not got:
                return {}
            print(f"还缺：{'、'.join(left)}")
            continue
        try:
            for c, v in parse_class_values(ans, grade_char).items():
                if c not in classes:
                    print(f"本次没有“{c}”这个班，已忽略。")
                elif v < lo or (hi is not None and v > hi):
                    print(f"{c} 的数字 {v} 不对，应在 {lo}～{hi if hi is not None else '…'} 之间。")
                else:
                    got[c] = v
        except ValueError as e:
            print(e)


def parse_sections(text: str, avail: list[str], all_names: list[str]) -> list[str]:
    """“1 4 7”“成绩通报、班级结构分”“全部”“0” → 要导出的内容名称（按固定页面顺序）。"""
    t = text.strip()
    if t in ("全部", "all", "ALL"):
        return list(avail)
    if t in ("0", "无", "不要", "否"):
        return []
    want = set()
    for part in re.split(r"[\s,，、;；]+", t):
        if not part:
            continue
        if part.isdigit():
            if not 1 <= int(part) <= len(all_names):
                raise ValueError(f"没有第 {part} 项，序号是 1～{len(all_names)}")
            name = all_names[int(part) - 1]
        else:
            hit = [n for n in all_names if part == n or part in n]
            if len(hit) != 1:
                raise ValueError(f"看不懂“{part}”，请写序号或名称：{'、'.join(all_names)}")
            name = hit[0]
        if name not in avail:
            print(f"本次没有“{name}”，已跳过。")
        else:
            want.add(name)
    return [n for n in avail if n in want]


def parse_ratio(text) -> float:
    t = str(text).strip().replace("％", "%")
    v = float(t.rstrip("%"))
    if t.endswith("%") or v > 1:
        v /= 100
    if not 0 < v <= 1:
        raise ValueError(f"升学线比例“{text}”不对，请写成 85% 或 0.85")
    return v


def main(argv=None):
    ap = argparse.ArgumentParser(prog="分寸", description=f"{APP_TITLE} {__version__}")
    ap.add_argument("登分表", nargs="?", help="登分表 Excel 文件（不写则运行时询问，可以把文件拖进窗口）")
    ap.add_argument("--版本", action="version", version=__version__)
    ap.add_argument("--考试", help="考试名称，如：2026年秋季学期九年级第一次月考（不写则运行时询问）")
    ap.add_argument("--年级", help="七年级/八年级/九年级（不写则根据班级名称自动判断）")
    ap.add_argument("--日期", default=None, help="考试日期，如 2026年10月20日（不写则运行时询问，可以不填）")
    ap.add_argument("--满分", default="", help="本次满分与默认不同时填写，如：数学=120 英语=120")
    ap.add_argument("--不计", default="", help="本次不计入的科目（登分表里有这一列，但这次不算），如：--不计 体育")
    ap.add_argument("--方案", default=None, help="算法方案：平时 或 中考（不写则运行时询问，默认平时），方案内容见 学校设置.yaml")
    ap.add_argument("--升学比例", default="", help="升学线取全级前百分之几，如 80%%（不写则询问，默认见 学校设置.yaml）")
    ap.add_argument("--应考", "--学籍", dest="应考", default="", help="各班应考人数（不写则读任课总表的「班级信息」），如：九1=46 九2=45")
    ap.add_argument("--上次名次", default="", help="上次考试各班结构分综合名次（平时方案的增值评价用），如：九1=3 九2=1")
    ap.add_argument("--任课", default=str(ROOT / "data" / "任课总表.xlsx"), help="任课总表 Excel（默认 data/任课总表.xlsx）")
    ap.add_argument("--学校设置", default=str(ROOT / "config" / "学校设置.yaml"))
    ap.add_argument("--本校设置", default=str(ROOT / "data" / "本校设置.yaml"),
                    help="只在本机生效的设置（如学校名称），有这个文件就盖在 学校设置.yaml 上面")
    ap.add_argument("--输出", default=str(ROOT / "output"))
    ap.add_argument("--不确认", action="store_true", help="不询问，直接按当前满分计算")
    ap.add_argument("--不生成PDF", action="store_true")
    ap.add_argument("--PDF内容", default="", help="发布版 PDF 导出哪些内容：写序号或名称，如 “1 4 7”、“成绩通报 班级结构分”；"
                                                "“全部”= 都要；不写则按 学校设置.yaml 的“发布版内容”")
    a = ap.parse_args(argv)

    interactive = not a.不确认 and has_tty()
    school = load_school(a.学校设置, a.本校设置)

    # ---- 登分表：命令里没写就当场问（可以把文件拖进窗口）
    if not a.登分表 and not interactive:
        sys.exit("请指定登分表文件，例如：python -m chengji data/登分表.xlsx")
    while not a.登分表 or not Path(a.登分表).is_file():
        if a.登分表:
            print(f"找不到这个文件：{a.登分表}")
            if not interactive:
                sys.exit(1)
        a.登分表 = clean_path(input("请把登分表（Excel 文件）拖进这个窗口，然后按回车："))
    data = load_scores(a.登分表, school, a.年级, exclude=[x for x in re.split(r"[\s,，、]+", a.不计) if x])
    grade = data.grade

    # ---- 算法方案：命令里没写就当场问
    names = list(school.schemes)
    while a.方案 is None and interactive and len(names) > 1:
        ans = input("算法方案：" + "　".join(f"{i} {n}" for i, n in enumerate(names, 1))
                    + f"。直接回车 = {names[0]}；要用别的请输入序号：").strip()
        if not ans:
            break
        if ans in names:
            a.方案 = ans
        elif ans.isdigit() and 1 <= int(ans) <= len(names):
            a.方案 = names[int(ans) - 1]
        else:
            print(f"请输入 1～{len(names)} 的序号。")
    a.方案 = a.方案 or names[0]

    # ---- 对照这个年级的固定科目，少了、多了提醒一句（不影响计算）
    subject_note = _subject_note(school, grade, a.方案, data.subjects)

    # ---- 考试名称、日期：命令里没写就当场问。名称同时是结果文件夹和文件的名字
    default_name = clean_name(Path(a.登分表).stem)
    exam = clean_name(a.考试 or "")
    if not exam and interactive:
        print(f"\n已读入登分表：{grade}，{len(data.df)} 人。")
        exam = clean_name(input(f"请输入本次考试名称（如 2026年秋季学期{grade}第一次月考）；"
                                f"直接回车则用文件名“{default_name}”：")) or default_name
    exam = exam or default_name
    while interactive and any((Path(a.输出) / exam).glob("*.xlsx")):
        new = clean_name(input(f"“{exam}”以前算过，这次的结果会盖掉原来的。直接回车 = 覆盖；想两份都留，请输入一个新名称："))
        if not new or new == exam:
            break
        exam = new
    if a.日期 is None:
        a.日期 = input("考试日期（如 2026年10月20日，会印在发布版上；不填直接回车）：").strip() if interactive else ""

    # ---- 本次满分：默认值 → 命令行 --满分 → 交互确认
    full = {s: school.grade_full.get(grade, {}).get(s, school.full.get(s, 0)) for s in data.subjects}
    full.update({k: v for k, v in parse_full_overrides(a.满分).items() if k in full})
    print(f"\n年级：{grade}　考试：{exam}　实考人数：{len(data.df)}　未计入：{len(data.excluded)} 人")
    print(f"从登分表抓到 {len(data.subjects)} 门科目及满分："
          + "　".join(f"{s} {v:g}" if v else f"{s} 【待设置】" for s, v in full.items()))
    if data.empty_subjects:
        print(f"整列没有成绩、本次不计算的科目：{'、'.join(data.empty_subjects)}")
    if data.new_subjects:
        print(f"其中新加的科目：{'、'.join(data.new_subjects)}（科目总表里没有，按这一列是分数认出来的；不想算就加 --不计）")
    if data.ignored_cols:
        print("登分表里没当成科目、已忽略的列：" + "、".join(data.ignored_cols))

    # 有成绩、但学校设置里还没有满分的科目：导入时当场问
    unset = [s for s, v in full.items() if not v]
    if unset and not interactive:
        sys.exit(f"{'、'.join(unset)} 有成绩，但还没有满分。请用 --满分 填写，例如：--满分 {unset[0]}=50")
    for s in unset:
        while True:
            ans = input(f"{s} 有成绩，这次满分是多少？请输入数字：").strip()
            try:
                v = float(ans)
            except ValueError:
                v = 0
            if v > 0:
                full[s] = v
                break
            print("请输入大于 0 的数字，例如 50。")
    if unset:
        print("当前满分：" + "　".join(f"{s} {v:g}" for s, v in full.items()))

    while interactive:
        ans = input("满分正确请直接回车；需要修改请输入 科目=满分（如 数学=120 英语=120）：").strip()
        if not ans:
            break
        try:
            for k, v in parse_full_overrides(ans).items():
                if k not in full:
                    print(f"本次没有“{k}”这门科目，已忽略。")
                elif v <= 0:
                    print(f"{k} 的满分要大于 0，没有修改。")
                else:
                    full[k] = v
        except ValueError as e:
            print(e)
        print("当前满分：" + "　".join(f"{s} {v:g}" for s, v in full.items()))

    # ---- 升学线比例：默认值 → 命令行 → 交互确认
    ratio = parse_ratio(a.升学比例) if a.升学比例 else school.promote_ratio
    if interactive and not a.升学比例:
        while True:
            ans = input(f"升学线比例：全级前 {ratio:.0%}。正确请直接回车；需要修改请输入比例（如 80%）：").strip()
            if not ans:
                break
            try:
                ratio = parse_ratio(ans)
            except ValueError as e:
                print(e)
    if a.方案 not in school.schemes:
        sys.exit(f"没有名为“{a.方案}”的算法方案，可选：{'、'.join(school.schemes)}")
    print(f"算法方案：{a.方案}　升学线比例：前 {ratio:.0%}")

    school.grade_full.setdefault(grade, {}).update(full)
    teachers, enrolled = [], {}
    if a.任课 and Path(a.任课).exists():
        teachers = load_teachers(a.任课, school, grade)
        enrolled = load_enrolled(a.任课, school, grade)
    else:
        print(f"【提示】没有找到任课总表（{a.任课}），本次不出教师排名。"
              f"需要的话，把 templates 文件夹里的“任课总表_模板.xlsx”填好，存成 data 文件夹里的“任课总表.xlsx”。")

    # ---- 班级结构分要用的：应考人数、上次名次
    W = {k: v for k, v in school.structures.get(a.方案, {}).items() if v}
    gc, CL = school.grades[grade], data.classes
    prev = {}
    if "参考率" in W:
        enrolled.update(parse_class_values(a.应考, gc))
        miss = [c for c in CL if not enrolled.get(c)]
        while miss and interactive:                          # 不填也行：没填的班按实考人数算
            ans = input(f"还没有应考人数的班：{'、'.join(miss)}。请输入（如 {miss[0]}=46）；"
                        f"直接回车 = 按实考人数算（参考率 100%）：").strip()
            if not ans:
                break
            try:
                enrolled.update({c: v for c, v in parse_class_values(ans, gc).items() if c in CL and v > 0})
            except ValueError as e:
                print(e)
            miss = [c for c in CL if not enrolled.get(c)]
        enrolled = {c: v for c, v in enrolled.items() if v}
        print("应考人数：" + "　".join(f"{c} {enrolled.get(c) or '按实考'}" for c in CL))
    if "增值评价" in W:
        prev = parse_class_values(a.上次名次, gc)
        if interactive and not prev:
            prev = ask_class_values("上次考试各班结构分综合名次（如 " + f"{CL[0]}=3 {CL[-1]}=1" + "，还缺 {left}）；"
                                    "没有上次成绩直接回车：", CL, gc, 1, len(CL), allow_empty=True)
        miss = [c for c in CL if c not in prev]
        if prev and miss:
            sys.exit(f"上次名次缺少：{'、'.join(miss)}。要么全部填写，要么都不填。")
        print("上次名次：" + ("　".join(f"{c} 第{prev[c]}" for c in CL) if prev
                              else "未输入，增值评价只按本次名次给分，不算进步退步"))
    cfg = make_config(school, grade, data.subjects, exam, a.日期, teachers, scheme=a.方案, promote_ratio=ratio,
                      enrolled=enrolled, prev_rank=prev)

    msgs = check_data(data, cfg)
    if subject_note:
        msgs.insert(0, subject_note)
    msgs += enrolment_notes(cfg, data)
    for m in msgs:
        print(m)
    if any(m.startswith("【错误】") for m in msgs):
        sys.exit("发现错误，已停止。请修正满分或登分表后重新运行。")
    if teachers:
        by_subj: dict[str, list[str]] = {}
        for t in cfg.teachers:
            by_subj.setdefault(t.subject, []).append(t.name)
        print(f"从任课总表抓到 {grade} 任课教师 {sum(len(v) for v in by_subj.values())} 位："
              + "　".join(f"{s} {len(v)}人" for s, v in by_subj.items()))
        other = sorted({t.subject for t in teachers} - set(cfg.subjects))
        if other:
            print(f"（任课总表里 {grade} 还有 {'、'.join(other)}，本次考试没有，已跳过）")
        no_teacher = [s for s in cfg.subjects if s not in by_subj]
        if no_teacher:
            print(f"【提示】本次考了 {'、'.join(no_teacher)}，但任课总表里没有 {grade} 这些学科的教师，"
                  f"教师排名里不会出现。")
    taught = {(t.subject, c) for t in cfg.teachers for c in t.classes}
    has_teacher = {t.subject for t in cfg.teachers}
    missing = [f"{s}·{c}" for s in cfg.subjects for c in data.classes
               if s in has_teacher and (s, c) not in taught]   # 整科都缺的已在上面提示过
    if cfg.teachers and missing:
        print("【提示】以下学科班级在任课总表中没有找到任课教师：" + "、".join(missing))

    R = analyze(data, cfg)
    concl = draft_conclusions(R)
    out = Path(a.输出) / exam
    out.mkdir(parents=True, exist_ok=True)
    try:
        lic.require_export()                                 # 试用到期且未激活时，不能导出
    except lic.LicenseError as e:
        sys.exit(f"{e}（核算结果可以在图形界面里查看；激活后才能导出 Excel 和 PDF）")
    xlsx = build_excel(R, out / f"{exam}_各班综合统计.xlsx", concl)
    print(f"已生成：{xlsx}")
    pdf_note = "发布版 PDF：未生成"
    if not a.不生成PDF:
        from .pdf_report import SECTIONS, available_sections, build_html, html_to_pdf
        avail = available_sections(R)
        chosen = [k for k in avail if school.pdf_sections.get(k, True)]      # 学校设置里的默认选择
        if a.PDF内容:
            chosen = parse_sections(a.PDF内容, avail, SECTIONS)
        elif interactive:
            while True:
                print("发布版 PDF 可以导出下面这些内容，打 ✓ 的是现在要导出的。")
                print("　" + "　".join(f"{SECTIONS.index(k) + 1} {k}{' ✓' if k in chosen else ''}" for k in avail))
                ans = input("直接回车 = 按打 ✓ 的导出；要改就输入要导出的序号（如 1 4 7）；输入 0 = 这次不要 PDF：").strip()
                if not ans:
                    break
                try:
                    chosen = parse_sections(ans, avail, SECTIONS)
                except ValueError as e:
                    print(e)
        if chosen:
            try:
                pdf = html_to_pdf(build_html(R, concl, chosen), out / f"{exam}_成绩发布版.pdf")
                pdf_note = "发布版 PDF 内容：" + "、".join(chosen)
                print(f"已生成：{pdf}（{'、'.join(chosen)}）")
            except RuntimeError as e:                         # 没有浏览器：不影响 Excel
                pdf_note = "发布版 PDF：未生成（没有找到浏览器）"
                print(f"【提示】{e}")
        else:
            print("本次不生成 PDF。")
    note = [f"考试：{exam}", f"年级：{grade}", f"算法方案：{cfg.scheme}（单科得分 = {cfg.formula_text()}）", "满分：" + "、".join(f"{s}{v:g}" for s, v in cfg.full.items()),
            f"实考人数：{len(data.df)}", f"升学线：前{cfg.promote_ratio:.0%}，第{R.cut_rank}名 {R.cut:g}分，线上{R.above}人",
            *([f"班级结构分（{cfg.scheme}）：" + "　".join(f"{c} {R.structure[c]['结构分']:.2f}（第{R.structure[c]['名次']}）"
                                                     for c in sorted(R.classes, key=lambda c: R.structure[c]['名次']))]
              if R.structure else []),
            *([("上次名次：" + "、".join(f"{c}第{v}" for c, v in cfg.prev_rank.items())) if cfg.prev_rank
               else "上次名次：未输入，增值评价只按本次名次给分"] if "增值评价" in cfg.structure else []),
            pdf_note,
            "未计入：" + ("；".join(f"{e['班级']} {e['姓名']}（{e['原因']}）" for e in data.excluded) or "无"),
            *msgs, "", "自动起草的结论（请审核后再发布）：", *[f"{i + 1}. {c}" for i, c in enumerate(concl)]]
    (out / "核对说明.txt").write_text("\n".join(note), encoding="utf-8")
    print(f"已生成：{out / '核对说明.txt'}\n完成。结果都在这个文件夹里：{out}")
    return out


def _open_folder(path: Path):
    if os.environ.get("CHENGJI_NO_OPEN"):                 # 自动测试时不弹出文件夹窗口
        return
    try:
        if os.name == "nt":
            os.startfile(str(path))                       # noqa: 仅 Windows 有
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
    except Exception:
        pass


def run(argv=None):
    """带“引导”的入口：出错时说人话；双击启动的，算完自动打开结果文件夹，并等按回车再关窗口。"""
    argv = sys.argv[1:] if argv is None else argv
    for stream in (sys.stdout, sys.stderr):               # 个别环境的输出编码显示不了中文：宁可显示成 ?，也不要报错中断
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    tty = has_tty()
    guided = tty and (FROZEN or not any(x.startswith("--") for x in argv))     # 双击或只给了文件 = 引导模式
    code = 0
    try:
        new = ensure_layout()
        if guided:
            print(f"===== {APP_TITLE} {__version__} =====")
            if FROZEN and any(n.startswith("config/") for n in new):
                print("第一次使用：已在程序旁边放好 config（学校设置）、templates（空白模板）、samples（示例数据）文件夹。\n"
                      "　先打开 config 里的“学校设置.yaml”，把学校名称改成自己学校的；空白模板在 templates 里。")
        out = main(argv)
        if guided and out:
            _open_folder(out)
    except SystemExit as e:
        if isinstance(e.code, str):
            print("\n【没有算完】" + e.code)
            code = 1
        else:
            code = e.code or 0
    except KeyboardInterrupt:
        print("\n已取消。")
        code = 1
    except PermissionError as e:
        print(f"\n【没有算完】文件写不进去：{e.filename}\n多半是这个文件正被 Excel 或 WPS 打开着，请先关掉它，再重新算一次。")
        code = 1
    except ValueError as e:
        print(f"\n【没有算完】{e}")
        code = 1
    except Exception as e:                                   # 没料到的错误：留下记录，方便反馈
        log = ROOT / "出错记录.txt"
        try:
            log.write_text(traceback.format_exc(), encoding="utf-8")
            where = f"详细信息已存到：{log}（反馈问题时请附上这个文件）"
        except Exception:
            where = traceback.format_exc()
        print(f"\n【出错了】{type(e).__name__}: {e}\n{where}")
        code = 1
    if guided:
        try:
            input("\n按回车键关闭窗口…")
        except EOFError:
            pass
    sys.exit(code)


if __name__ == "__main__":
    run()
