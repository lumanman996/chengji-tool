"""基本测试（只用 samples/ 里的虚构数据）。运行：python -m pytest -q"""
from pathlib import Path

import openpyxl
import pytest

from chengji.analysis import analyze, draft_conclusions
from chengji.config import load_enrolled, load_school, load_teachers, make_config, normalize_class, parse_classes
from chengji.excel_report import build_excel

ROOT = Path(__file__).resolve().parent.parent
SCHOOL = load_school(ROOT / "config" / "学校设置.yaml")


def test_class_names():
    assert normalize_class("九（1）", "九") == "九1"
    assert normalize_class("91", "九") == "九1"
    assert normalize_class("七2", "七") == "七2"
    assert normalize_class("八年级3班", "八") == "八3"
    assert parse_classes("9192", "九") == ["九1", "九2"]
    assert parse_classes("93949596", "九") == ["九3", "九4", "九5", "九6"]
    assert parse_classes("93", "九") == ["九3"]
    assert parse_classes("七1、七2", "七") == ["七1", "七2"]


def test_subject_aliases():
    assert SCHOOL.canonical("道德与法治") == "道法"
    assert SCHOOL.canonical("物 理") == "物理"
    assert SCHOOL.canonical("班级") is None


@pytest.mark.parametrize("fname,grade,nsub", [("示例登分表_七年级.xlsx", "七年级", 7), ("示例登分表_八年级.xlsx", "八年级", 8),
                                              ("示例登分表_九年级.xlsx", "九年级", 7), ("示例登分表_中考.xlsx", "九年级", 10)])
def test_pipeline(tmp_path, fname, grade, nsub):
    import copy
    from chengji.loader import load_scores
    SCHOOL = copy.deepcopy(globals()["SCHOOL"])
    data = load_scores(ROOT / "samples" / fname, SCHOOL)
    # 学校设置里满分待定的科目（生物、地理）：像运行时那样当场设定，示例数据按 50 分
    SCHOOL.grade_full.setdefault(grade, {}).update({s: 50 for s in data.subjects if not SCHOOL.full[s]})
    assert data.grade == grade and len(data.subjects) == nsub
    assert len(data.excluded) == 1                      # 示例数据里放了 1 名缺考学生
    T = load_teachers(ROOT / "samples" / "示例任课总表.xlsx", SCHOOL, grade)
    E = load_enrolled(ROOT / "samples" / "示例任课总表.xlsx", SCHOOL, grade)
    cfg = make_config(SCHOOL, grade, data.subjects, "测试", "", T, enrolled=E)
    R = analyze(data, cfg)
    n = len(data.df)
    assert R.cut_rank == round(n * 0.85)
    assert R.above >= R.cut_rank and R.above + R.below == n
    assert sum(b[4] for b in R.bands) == R.above        # 分数段合计 = 线上人数
    for s in cfg.subjects:                              # 每科教师覆盖全部学生
        assert sum(t["人数"] for t in R.teachers if t["学科"] == s) == n
    assert (R.near["距线"].abs() <= cfg.near_range).all()
    out = build_excel(R, tmp_path / "t.xlsx", draft_conclusions(R))
    wb = openpyxl.load_workbook(out)
    assert {"各班统计", "班级结构分", "教师排名", "临界生名单", "参数设置", "成绩明细", "任课安排"} <= set(wb.sheetnames)
    assert sorted(x["名次"] for x in R.structure.values()) == list(range(1, len(R.classes) + 1))


def test_empty_subject_column_ignored(tmp_path):
    """整列没有成绩的科目 = 本次没考：不计算，也不能把学生判成“成绩不全”。"""
    from chengji.loader import load_scores
    wb = openpyxl.load_workbook(ROOT / "samples" / "示例登分表_九年级.xlsx")
    ws = wb.active
    c = ws.max_column + 1
    ws.cell(2, c, "生物")                               # 只有表头，没有分数
    ws.cell(2, c + 1, "地理")
    ws.cell(5, c + 1, "缺考")                           # 只有文字、没有数字，也算没成绩
    wb.save(tmp_path / "t.xlsx")
    data = load_scores(tmp_path / "t.xlsx", SCHOOL)
    assert data.subjects == ["语文", "数学", "英语", "物理", "化学", "道法", "历史"]
    assert data.empty_subjects == ["生物", "地理"]
    assert len(data.excluded) == 1                      # 仍然只有原来那 1 名缺考学生


def test_missing_full_mark():
    with pytest.raises(ValueError, match="满分"):
        make_config(SCHOOL, "九年级", ["生物"], "测试")


def test_zhongkao_scheme():
    """中考方案：平均分（原始分）×60% + 合格率×40%，不看优秀率。"""
    from chengji.analysis import subject_stats
    import pandas as pd
    cfg = make_config(SCHOOL, "九年级", ["语文"], "测试", scheme="中考", promote_ratio=0.8)
    st = subject_stats(pd.Series([100, 72, 50, 30.0]), cfg, "语文")   # 满分120，合格线72
    assert abs(st["得分"] - (63 * 0.6 + 50 * 0.4)) < 1e-9
    assert cfg.promote_ratio == 0.8 and not cfg.normalize


def test_ratio_parse():
    from chengji.cli import parse_ratio
    assert parse_ratio("80%") == 0.8 and parse_ratio("0.75") == 0.75 and parse_ratio("90") == 0.9


# ---------------- 班级结构分：用一组能手算的小数据核对 ----------------
def _tiny(scheme, prev=None):
    """3 个班、每班 2 人，语文（满分120，合格72，优秀96）+ 数学（满分100，合格60，优秀80）。"""
    import pandas as pd
    from chengji.analysis import analyze
    from chengji.loader import ScoreData
    rows = [("九1", 110, 95), ("九1", 100, 90),      # 两人都全科优秀
            ("九2", 80, 70), ("九2", 60, 50),        # 一人全科合格
            ("九3", 50, 40), ("九3", 30, 20)]        # 都不合格
    df = pd.DataFrame([(c, f"生{i}", f"{i:02d}", a, b) for i, (c, a, b) in enumerate(rows)],
                      columns=["班级", "姓名", "考号", "语文", "数学"])
    data = ScoreData(df=df, grade="九年级", subjects=["语文", "数学"])
    cfg = make_config(SCHOOL, "九年级", ["语文", "数学"], "测试", scheme=scheme,
                      enrolled={"九1": 2, "九2": 3, "九3": 2}, prev_rank=prev)
    return analyze(data, cfg).structure


def test_structure_pingshi():
    """平时：平均成绩50 + 全科合格率20 + 全科优秀率20 + 参考率5 + 增值评价5。"""
    X = _tiny("平时", prev={"九1": 2, "九2": 1, "九3": 3})
    avg = {"九1": 197.5 / 220 * 50, "九2": 130 / 220 * 50, "九3": 70 / 220 * 50}   # 总分满分 220
    sub = {"九1": avg["九1"] + 20 + 20 + 5,
           "九2": avg["九2"] + 0.5 * 20 + 0 + 2 / 3 * 5,
           "九3": avg["九3"] + 0 + 0 + 5}
    # 小计名次 1、2、3 → 5、4.9、4.8；九1 上次第2进步1名 5.05 → 封顶 5；九2 上次第1退步1名 4.85；九3 不变 4.8
    add = {"九1": 5.0, "九2": 4.85, "九3": 4.8}
    for c in X:
        assert abs(X[c]["小计"] - sub[c]) < 1e-9
        assert abs(X[c]["增值评价分"] - add[c]) < 1e-9
        assert abs(X[c]["结构分"] - (sub[c] + add[c])) < 1e-9
    assert [X[c]["名次"] for c in ("九1", "九2", "九3")] == [1, 2, 3]


def test_structure_pingshi_no_previous():
    """没有输入上次名次：增值只按本次名次给分。"""
    X = _tiny("平时")
    assert [round(X[c]["增值评价分"], 9) for c in ("九1", "九2", "九3")] == [5.0, 4.9, 4.8]


def test_structure_zhongkao():
    """中考：平均成绩50 + 优秀率15 + 合格率15 + 参考率10 + 进线率10 + 前10名每人0.1分。"""
    X = _tiny("中考")
    # 共 6 人，前85% = 第5名 = 90分 → 进线：九1 2/2、九2 2/2、九3 1/2；不足10人，前10名 = 全部
    exp = {"九1": 197.5 / 220 * 50 + 15 + 15 + 10 + 10 + 0.2,
           "九2": 130 / 220 * 50 + 0 + 7.5 + 2 / 3 * 10 + 10 + 0.2,
           "九3": 70 / 220 * 50 + 0 + 0 + 10 + 5 + 0.2}
    for c in X:
        assert abs(X[c]["结构分"] - exp[c]) < 1e-9
        assert "增值评价" not in X[c]


def test_structure_needs_enrolment():
    import pandas as pd
    from chengji.analysis import analyze
    from chengji.loader import ScoreData
    df = pd.DataFrame([("九1", "甲", "01", 90.0)], columns=["班级", "姓名", "考号", "语文"])
    cfg = make_config(SCHOOL, "九年级", ["语文"], "测试")
    with pytest.raises(ValueError, match="应考人数"):
        analyze(ScoreData(df=df, grade="九年级", subjects=["语文"]), cfg)


def test_pdf_sections_parse():
    from chengji.cli import parse_sections
    from chengji.pdf_report import SECTIONS
    avail = [s for s in SECTIONS if s != "教师排名"]              # 假设本次没有任课教师
    assert parse_sections("1 4", avail, SECTIONS) == ["成绩通报", "班级结构分"]
    assert parse_sections("结构分、成绩通报", avail, SECTIONS) == ["成绩通报", "班级结构分"]   # 按页面顺序
    assert parse_sections("全部", avail, SECTIONS) == avail
    assert parse_sections("0", avail, SECTIONS) == []
    assert parse_sections("1 7", avail, SECTIONS) == ["成绩通报"]  # 本次没有的内容跳过
    with pytest.raises(ValueError):
        parse_sections("9", avail, SECTIONS)


def test_pdf_sections_pages():
    """只选两项时，发布版里只有这两项的页面。"""
    from chengji.loader import load_scores
    from chengji.pdf_report import build_html
    data = load_scores(ROOT / "samples" / "示例登分表_九年级.xlsx", SCHOOL)
    E = load_enrolled(ROOT / "samples" / "示例任课总表.xlsx", SCHOOL, "九年级")
    R = analyze(data, make_config(SCHOOL, "九年级", data.subjects, "测试", enrolled=E))
    html = build_html(R, draft_conclusions(R), ["成绩通报", "班级结构分"])
    assert html.count('<div class="page">') == 2
    assert "班级结构分评比" in html and "教师成绩排名" not in html and "临界生名单" not in html
    assert "第 2 页" in html and "第 3 页" not in html            # 页码连续


def test_class_rank():
    """班名次：本班内按总分排，同分同名次；与级名次互不影响。"""
    import pandas as pd
    from chengji.loader import ScoreData
    rows = [("九1", 100), ("九1", 90), ("九1", 90), ("九1", 50), ("九2", 95), ("九2", 60)]
    df = pd.DataFrame([(c, f"生{i}", f"{i:02d}", float(v)) for i, (c, v) in enumerate(rows)],
                      columns=["班级", "姓名", "考号", "语文"])
    cfg = make_config(SCHOOL, "九年级", ["语文"], "测试", enrolled={"九1": 4, "九2": 2})
    R = analyze(ScoreData(df=df, grade="九年级", subjects=["语文"]), cfg)
    assert list(R.df["班名次"]) == [1, 2, 2, 4, 1, 2]
    assert list(R.df["级名次"]) == [1, 3, 3, 6, 2, 5]


def test_grade_subjects():
    """各年级固定科目：七年级 7 科，八年级多一门物理，九年级 7 科，中考 10 科。示例登分表与之一致。"""
    from chengji.loader import load_scores
    assert SCHOOL.expected_subjects("七年级", "平时") == ["语文", "数学", "英语", "道法", "历史", "生物", "地理"]
    assert set(SCHOOL.expected_subjects("八年级", "平时")) == set(SCHOOL.expected_subjects("七年级", "平时")) | {"物理"}
    assert SCHOOL.expected_subjects("九年级", "平时") == ["语文", "数学", "英语", "物理", "化学", "道法", "历史"]
    assert SCHOOL.expected_subjects("九年级", "中考") == SCHOOL.subjects and len(SCHOOL.subjects) == 10
    for key, fname in [("七年级", "七年级"), ("八年级", "八年级"), ("九年级", "九年级"), ("中考", "中考")]:
        want = SCHOOL.grade_subjects[key]
        for folder, name in (("samples", f"示例登分表_{fname}.xlsx"), ("templates", f"登分表_{fname}.xlsx")):
            data = load_scores(ROOT / folder / name, SCHOOL)
            assert set(data.subjects) == set(want), (name, data.subjects)
