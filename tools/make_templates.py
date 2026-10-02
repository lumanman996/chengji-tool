"""生成空白模板和虚构示例数据：

  templates/登分表_七年级.xlsx 等  各年级、中考的空白登分表（科目按 学校设置.yaml 的“年级科目”，带 2 行示例）
  templates/任课总表_模板.xlsx   空白任课总表（2 行示例）
  samples/示例登分表_*.xlsx      虚构学生成绩，供测试用
  samples/示例任课总表.xlsx      虚构教师，供测试用

科目和满分直接从 config/学校设置.yaml 读取，改了设置重新运行本脚本即可。
用法：python tools/make_templates.py
"""
import random
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from chengji.config import load_school   # noqa: E402

F = "微软雅黑"
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")     # 表头浅蓝
DEMO_FILL = PatternFill("solid", fgColor="FFF2CC")     # 示例行浅黄
NOTE_FONT = Font(name=F, size=10, color="C00000")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def _title(ws, text, span):
    ws["A1"] = text
    ws["A1"].font = Font(name=F, size=14, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws.row_dimensions[1].height = 24


def _header(ws, names):
    for j, h in enumerate(names, 1):
        c = ws.cell(2, j, h)
        c.font = Font(name=F, bold=True)
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "A3"


def _tips_sheet(wb, tips, width=86):
    s = wb.create_sheet("填写说明")
    for i, t in enumerate(tips, 1):
        c = s.cell(i, 1, t)
        c.font = Font(name=F, bold=(i == 1), size=13 if i == 1 else 11)
        c.alignment = Alignment(vertical="center", wrap_text=False)
    s.column_dimensions["A"].width = width
    return s


def _save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print("已生成", path.relative_to(ROOT))


# ---------------------------------------------------------------- 空白登分表
def blank_scorebook(path, school, subjects=None, title="＿＿年级　＿＿＿＿考试　登分表", demo_class="九1"):
    subjects = subjects or school.subjects
    head = ["序号", "班级", "姓名", "考号"] + subjects
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "登分表"
    _title(ws, title, len(head))
    _header(ws, head)

    demo = [(1, demo_class, "学生甲", "0101"), (2, demo_class, "学生乙", "0102")]
    for i, (no, cls, name, sid) in enumerate(demo):
        r = 3 + i
        vals = [no, cls, name, sid]
        for s in subjects:
            fm = school.full[s] or 50                 # 满分待确认的科目，示例分数按 50 分估
            vals.append(round(fm * (0.75 if i == 0 else 0.55)))
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font = Font(name=F)
            c.fill = DEMO_FILL
            c.alignment = Alignment(horizontal="center")
            c.border = THIN
    ws.cell(5, 1, "↑ 上面两行是示例，填写前请删除。哪一科这次没考，那一列空着不填就行。")
    ws.cell(5, 1).font = NOTE_FONT

    ws.column_dimensions["A"].width = 6
    for col, w in zip("BCD", [10, 12, 10]):
        ws.column_dimensions[col].width = w
    for j in range(5, len(head) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 8

    _tips_sheet(wb, [
        "登分表填写说明",
        "",
        "1. 第 2 行是表头，不要改动文字；从第 3 行开始，一名学生占一行。",
        "2. 第 3、4 行是示例（浅黄色），填写前请删除。",
        "3. 本次没考的科目，那一列空着不填就行（删掉也可以），程序自动不算它。",
        "   要加一门科目：在最后加一列，表头写科目名称（如 信息技术），下面填分数，程序会一起统计。",
        "4. 班级写法：年级用汉字，班号用数字，例如 七1、八3、九6。",
        "   同一张表只放一个年级。",
        "5. 考号可以不填（整列留空或删掉都行）。",
        "6. 缺考、请假、转学：那一格留空，或者写“缺考”“请假”“转学”。",
        "   这名学生不计入统计，会在核对说明里单独列出。",
        "7. 0 分是真实成绩，照常计入，不要写成空格。",
        "8. 不要合并单元格，不要在中间插小标题行，分数只填数字（不要带“分”字）。",
        "9. 表头上方可以有标题行，程序会自己找到表头。",
        "",
        "科目名称（可以写别名，括号里的写法也认）：",
        "　语文、数学、英语、物理、化学、生物（生物学）、地理、",
        "　道法（道德与法治 / 思想品德 / 政治）、历史、体育",
        "",
        "满分不用填在这张表里：导入后程序会列出本次有成绩的科目和满分，",
        "　让你确认或当场修改；还没有满分的科目，程序会逐个问你。"
    ])
    _save(wb, path)


# ---------------------------------------------------------------- 任课总表
def teacher_book(path, title, rows, note=None, tips=None, class_rows=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "任课总表"
    head = ["年级", "学科", "教师", "任教班级"]
    _title(ws, title, len(head))
    _header(ws, head)
    for i, r in enumerate(rows, start=3):
        for j, v in enumerate(r, 1):
            c = ws.cell(i, j, v)
            c.font = Font(name=F)
            c.border = THIN
            if note:
                c.fill = DEMO_FILL
            if j != 4:
                c.alignment = Alignment(horizontal="center")
    if note:
        ws.cell(3 + len(rows) + 1, 1, note).font = NOTE_FONT
    for col, w in zip("ABCD", [10, 10, 12, 26]):
        ws.column_dimensions[col].width = w

    # 班级信息：应考人数（算班级结构分的参考率用）
    ci = wb.create_sheet("班级信息")
    _title(ci, "班级信息（算参考率用：参考率 = 实考人数 ÷ 应考人数）", 3)
    _header(ci, ["年级", "班级", "应考人数"])
    for i, r in enumerate(class_rows or [], start=3):
        for j, v in enumerate(r, 1):
            c = ci.cell(i, j, v)
            c.font = Font(name=F)
            c.border = THIN
            c.alignment = Alignment(horizontal="center")
            if note:
                c.fill = DEMO_FILL
    if note and class_rows:
        ci.cell(3 + len(class_rows) + 1, 1, "↑ 上面两行是示例，填写前请删除。每个班一行，一学期填一次。").font = NOTE_FONT
    for col, w in zip("ABC", [10, 10, 12]):
        ci.column_dimensions[col].width = w

    _tips_sheet(wb, tips or [])
    _save(wb, path)


TEACHER_TIPS = [
    "任课总表填写说明",
    "",
    "1. 第 2 行是表头，不要改动文字；从第 3 行开始，一位教师教一个学科占一行。",
    "2. 第 3、4 行是示例（浅黄色），填写前请删除。",
    "3. 年级：写 七年级 / 八年级 / 九年级（只写 七 / 八 / 九 也行）。",
    "4. 学科：语文、数学、英语、物理、化学、生物、地理、道法、历史、体育。",
    "5. 任教班级：年级用汉字，班号用数字，多个班用顿号隔开，例如 九1、九2。",
    "   不要写成 9192、91 这样的纯数字。",
    "6. 一位教师教两个学科，就写两行；教两个年级，也写两行。",
    "7. 全校一张表就够了。算成绩时，程序自动挑出这次考试的年级和科目对应的教师，",
    "   表里多出来的年级、科目不影响计算。",
    "8. 教师姓名不要带空格；同名教师靠“学科 + 任教班级”区分。",
    "",
    "「班级信息」工作表（算班级结构分的参考率用）：",
    "9. 每个班一行：年级、班级、应考人数。班级写法同上，如 九1。",
    "10. 应考人数一学期填一次，有转入转出时改一下。",
    "11. 应考人数可以不填：没填的班按实考人数算（参考率按 100%）。",
]

# ---------------------------------------------------------------- 虚构示例成绩
def make_sample_scorebook(path, grade_char, subjects, full, n_classes=6, per_class=45, seed=1):
    """生成虚构学生成绩（只用于测试，不含任何真实信息）。"""
    rnd = random.Random(seed)
    fam = "王李张刘陈杨赵黄周吴徐孙马朱胡郭何高林罗"
    given = "砚秋书白墨言青禾听澜疏桐南乔北辰星野和舟"      # 特意选了不常用来起名的字，避免和真实学生重名
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "登分表"
    ws["A1"] = f"{grade_char}年级 示例登分表（虚构数据）"
    head = ["序号", "班级", "姓名", "考号"] + subjects
    for j, h in enumerate(head, 1):
        ws.cell(2, j, h)
    r, k = 3, 1
    for c in range(1, n_classes + 1):
        level = rnd.uniform(-0.06, 0.06)
        for i in range(per_class):
            ability = min(max(rnd.gauss(0.58 + level, 0.18), 0.05), 0.98)
            row = [k, f"{grade_char}{c}",
                   fam[rnd.randrange(len(fam))]
                   + given[rnd.randrange(0, len(given) - 1, 2)]
                   + given[rnd.randrange(1, len(given), 2)],
                   f"{c:02d}{i + 1:02d}"]
            for s in subjects:
                row.append(round(min(max(rnd.gauss(ability, 0.12), 0), 1) * full[s]))
            for j, v in enumerate(row, 1):
                ws.cell(r, j, v)
            r += 1
            k += 1
    ws.cell(r - 1, 4 + 2, "缺考")          # 一名缺考学生，演示“不计入”
    _save(wb, path)


if __name__ == "__main__":
    school = load_school(ROOT / "config" / "学校设置.yaml")

    # ---- 空白模板（给老师下载填写）
    old = ROOT / "templates" / "登分表_模板.xlsx"          # 以前是一张 10 科通用的，现在按年级各一张
    old.unlink(missing_ok=True)
    for key, subs in school.grade_subjects.items():
        ch = school.grades.get(key)                           # 七年级 → 七；“中考”用九年级的班
        title = f"{key}　＿＿＿＿考试　登分表" if ch else "九年级　＿＿＿＿年中考成绩　登分表"
        blank_scorebook(ROOT / "templates" / f"登分表_{key}.xlsx", school, subs, title, f"{ch or '九'}1")
    teacher_book(ROOT / "templates" / "任课总表_模板.xlsx",
                 "任课总表（模板）",
                 [("九年级", "语文", "张三", "九1、九2"), ("九年级", "数学", "李四", "九3")],
                 note="↑ 上面两行是示例，填写前请删除。表里只留实际的任课安排。",
                 tips=TEACHER_TIPS,
                 class_rows=[("九年级", "九1", 46), ("九年级", "九2", 45)])

    # ---- 虚构示例数据（测试用，可以上传）
    # 科目按“年级科目”。生物、地理、体育在学校设置里满分还没定，示例按 50 分造数据，
    # 运行示例时要给满分：--满分 "生物=50 地理=50 体育=50"
    FULL = {**{s: v or 50 for s, v in school.full.items()}}
    GS = school.grade_subjects
    make_sample_scorebook(ROOT / "samples" / "示例登分表_七年级.xlsx", "七", GS["七年级"], FULL, seed=7)
    make_sample_scorebook(ROOT / "samples" / "示例登分表_八年级.xlsx", "八", GS["八年级"], FULL, seed=8)
    make_sample_scorebook(ROOT / "samples" / "示例登分表_九年级.xlsx", "九", GS["九年级"], FULL, seed=9)
    make_sample_scorebook(ROOT / "samples" / "示例登分表_中考.xlsx", "九", GS["中考"], FULL, seed=10)
    rows = []
    names = iter([f"示例教师{i}" for i in range(1, 200)])
    for g, ch, subs in [("七年级", "七", GS["七年级"]), ("八年级", "八", GS["八年级"]), ("九年级", "九", GS["中考"])]:
        for s in subs:
            for a, b in (("1", "2"), ("3", "4"), ("5", "6")):
                rows.append((g, s, next(names), f"{ch}{a}、{ch}{b}"))
    enrolled = [(g, f"{ch}{k}", 45 + extra) for g, ch in (("七年级", "七"), ("八年级", "八"), ("九年级", "九"))
                for k, extra in zip(range(1, 7), (1, 2, 0, 3, 1, 2))]
    teacher_book(ROOT / "samples" / "示例任课总表.xlsx", "示例任课总表（虚构教师）", rows, tips=TEACHER_TIPS,
                 class_rows=enrolled)
