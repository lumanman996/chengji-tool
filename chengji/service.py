"""核算服务：图形界面和命令行共用的一层。

界面的每一步对应这里的一个函数：
  inspect()   读入登分表，告诉界面识别到了什么（年级、班级、科目、满分、任课、需要确认的事项）
  compute()   按界面上确认的设置计算，得到 Result
  to_view()   把 Result 整理成界面要显示的数据
  save_run() / list_runs() / load_run()   “最近的考试”
  export_excel() / export_pdf()   导出
"""
from __future__ import annotations

import json
import re
import shutil
import time
from pathlib import Path

from .analysis import RATE_ITEMS, Result, analyze, draft_conclusions
from .config import School, load_enrolled, load_teachers, make_config
from .loader import ScoreData, check_data, load_scores

RUN_FILE = "结果.json"          # 每次核算存一份，供“最近的考试”回看
RUN_COPY = "登分表副本.xlsx"    # 当时用的登分表，回看时导出要用它重算


def clean_name(text: str) -> str:
    """考试名称要当文件夹名用：去掉首尾空格和不能做文件名的符号。"""
    return re.sub(r'[\\/:*?"<>|\r\n\t]+', "", str(text or "")).strip().strip(".")


def subject_note(school: School, grade: str, scheme: str, subjects: list[str]) -> str:
    """对照这个年级的固定科目，少了、多了提醒一句（不影响计算）。没问题返回空字符串。"""
    expect = school.expected_subjects(grade, scheme)
    if not expect:
        return ""
    lack = [s for s in expect if s not in subjects]
    more = [s for s in subjects if s not in expect]
    if not (lack or more):
        return ""
    which = f"{scheme}核算" if scheme in school.grade_subjects else f"{grade}平时考试"
    return (f"【提示】{which}固定是 {len(expect)} 科（{'、'.join(expect)}），本次"
            + "，".join(x for x in (f"少了 {'、'.join(lack)}" if lack else "", f"多了 {'、'.join(more)}" if more else "") if x)
            + "。如果就是这样考的，不用管；否则请检查登分表。")


def _teacher_table(teacher_path, school: School, grade: str):
    if teacher_path and Path(teacher_path).is_file():
        return load_teachers(teacher_path, school, grade), load_enrolled(teacher_path, school, grade), True
    return [], {}, False


def inspect(path, school: School, teacher_path=None, grade: str | None = None) -> dict:
    """读入登分表，返回界面“确认设置”这一步要显示和预填的内容。"""
    data = load_scores(path, school, grade)
    g = data.grade
    teachers, enrolled, has_table = _teacher_table(teacher_path, school, g)
    counts = data.df["班级"].value_counts()
    schemes = []
    for name in school.schemes:
        w = {k: v for k, v in school.structures.get(name, {}).items() if v}
        schemes.append({"name": name, "note": subject_note(school, g, name, data.subjects).replace("【提示】", ""),
                        "needEnrolled": "参考率" in w, "needPrev": "增值评价" in w,
                        "formula": make_config(_with_full(school, g, data.subjects), g, data.subjects, "x", scheme=name).formula_text()})
    return {
        "file": Path(path).name, "grade": g, "classes": data.classes,
        "counts": {c: int(counts[c]) for c in data.classes}, "n": len(data.df),
        "subjects": data.subjects,
        "full": {s: school.grade_full.get(g, {}).get(s, school.full[s]) for s in data.subjects},
        "excluded": data.excluded, "emptySubjects": data.empty_subjects, "ignoredCols": data.ignored_cols,
        "hasTeacherTable": has_table,
        "teacherCount": len([t for t in teachers if t.subject in data.subjects]),
        "enrolled": {c: enrolled.get(c) for c in data.classes},
        "schemes": schemes, "ratio": school.promote_ratio,
        "defaultName": clean_name(Path(path).stem),
    }


def _with_full(school: School, grade: str, subjects: list[str], full: dict | None = None) -> School:
    """为了算公式说明等临时用：把没定满分的科目先按 100 填上（不影响真正的计算）。"""
    import copy
    s = copy.deepcopy(school)
    fm = s.grade_full.setdefault(grade, {})
    for sub in subjects:
        v = (full or {}).get(sub) or fm.get(sub) or s.full.get(sub) or 100
        fm[sub] = float(v)
    return s


def compute(path, school: School, opts: dict, teacher_path=None) -> tuple[Result, list[str], list[str]]:
    """按确认好的设置计算。opts：scheme、exam、date、full{科目:满分}、ratio、enrolled{班:人数}、prev{班:名次}、grade。
    返回 (Result, 自动起草的结论, 提示信息)。数据有硬伤（超过满分、应考人数不对等）时抛 ValueError。"""
    import copy
    school = copy.deepcopy(school)
    data = load_scores(path, school, opts.get("grade"))
    g = data.grade
    scheme = opts.get("scheme") or next(iter(school.schemes))
    if scheme not in school.schemes:
        raise ValueError(f"没有名为“{scheme}”的算法方案")

    full = {s: school.grade_full.get(g, {}).get(s, school.full[s]) for s in data.subjects}
    for s, v in (opts.get("full") or {}).items():
        if s in full and v not in (None, ""):
            full[s] = float(v)
    miss = [s for s, v in full.items() if not v or v <= 0]
    if miss:
        raise ValueError(f"{'、'.join(miss)} 还没有填满分。")
    school.grade_full.setdefault(g, {}).update(full)

    ratio = float(opts.get("ratio") or school.promote_ratio)
    if ratio > 1:
        ratio /= 100
    teachers, enrolled, has_table = _teacher_table(teacher_path, school, g)
    W = {k: v for k, v in school.structures.get(scheme, {}).items() if v}
    CL = data.classes
    prev = {}
    if "参考率" in W:
        for c, v in (opts.get("enrolled") or {}).items():
            if c in CL and v not in (None, ""):
                enrolled[c] = int(float(v))
        lack = [c for c in CL if not enrolled.get(c)]
        if lack:
            raise ValueError(f"还没有填应考人数：{'、'.join(lack)}。")
    if "增值评价" in W:
        prev = {c: int(float(v)) for c, v in (opts.get("prev") or {}).items() if c in CL and v not in (None, "")}
        if prev:
            lack = [c for c in CL if c not in prev]
            bad = [c for c, v in prev.items() if not 1 <= v <= len(CL)]
            if lack:
                raise ValueError(f"上次名次还缺：{'、'.join(lack)}。要么全部填写，要么都不填。")
            if bad:
                raise ValueError(f"上次名次应在 1～{len(CL)} 之间，请检查：{'、'.join(bad)}。")

    exam = clean_name(opts.get("exam") or "") or clean_name(Path(path).stem)
    cfg = make_config(school, g, data.subjects, exam, str(opts.get("date") or "").strip(), teachers, scheme=scheme,
                      promote_ratio=ratio, enrolled=enrolled, prev_rank=prev)
    msgs = check_data(data, cfg)
    note = subject_note(school, g, scheme, data.subjects)
    if note:
        msgs.insert(0, note)
    if "参考率" in cfg.structure:
        real = data.df["班级"].value_counts()
        for c in CL:
            if real[c] > cfg.enrolled[c]:
                msgs.append(f"【错误】{c} 实考人数 {real[c]} 比应考人数 {cfg.enrolled[c]} 还多，请检查应考人数")
    errors = [m.replace("【错误】", "") for m in msgs if m.startswith("【错误】")]
    if errors:
        raise ValueError("；".join(errors) + "。")
    if not has_table:
        msgs.append("【提示】没有任课总表，本次不出教师排名。")
    elif cfg.teachers:
        has = {t.subject for t in cfg.teachers}
        none = [s for s in cfg.subjects if s not in has]
        if none:
            msgs.append(f"【提示】任课总表里没有 {g} {'、'.join(none)} 的教师，这些学科不出教师排名。")
    else:
        msgs.append(f"【提示】任课总表里没有 {g} 的任课安排，本次不出教师排名。")
    R = analyze(data, cfg)
    return R, draft_conclusions(R), msgs


# ---------------------------------------------------------------- 给界面的数据
def _r2(v):
    return round(float(v), 2)


def _plain(v):
    """numpy 的数字换成普通数字，方便存成 JSON。"""
    if v is None or isinstance(v, (str, bool)):
        return v
    if isinstance(v, float):
        return round(float(v), 4)
    return int(v)


def to_view(R: Result, conclusions: list[str], msgs: list[str]) -> dict:
    """把计算结果整理成界面显示用的数据（都是普通的数字、文字、列表）。"""
    from .pdf_report import available_sections
    cfg = R.cfg
    S = cfg.subjects
    structure = []
    for c in R.classes:
        x = R.structure.get(c)
        if x:
            structure.append({"name": c, **{k: _plain(v) for k, v in x.items()}})
    return {
        "exam": cfg.exam, "date": cfg.date, "grade": cfg.grade, "school": cfg.school, "scheme": cfg.scheme,
        "formula": cfg.formula_text(),
        "subjects": S, "full": {s: cfg.full[s] for s in S}, "totalFull": cfg.total_full,
        "n": len(R.df), "excluded": R.data.excluded,
        "cut": R.cut, "cutRank": R.cut_rank, "above": R.above, "below": R.below, "ratio": cfg.promote_ratio,
        "nearRange": cfg.near_range, "highLine": cfg.high_line,
        "classes": [{"name": c, "n": R.online[c]["实考人数"], "composite": _r2(R.composite[c]), "rank": R.class_rank[c],
                     "online": R.online[c]["线上"], "onlineRate": _r2(R.online[c]["上线率"] * 100),
                     "subj": {s: {"avg": _r2(R.class_stats[c][s]["平均分"]), "pass": _r2(R.class_stats[c][s]["及格率"] * 100),
                                  "exc": _r2(R.class_stats[c][s]["优秀率"] * 100), "score": _r2(R.class_stats[c][s]["得分"])} for s in S}}
                    for c in R.classes],
        "structure": structure, "weights": cfg.structure, "rateItems": [k for k in RATE_ITEMS if k in cfg.structure],
        "hasPrev": bool(cfg.prev_rank),
        "teachers": [{"subject": t["学科"], "name": t["教师"], "classes": t["班级"], "n": int(t["人数"]), "avg": _r2(t["平均分"]),
                      "pass": _r2(t["及格率"] * 100), "exc": _r2(t["优秀率"] * 100), "score": _r2(t["得分"]), "rank": int(t["名次"])}
                     for t in R.teachers],
        "students": [[r.班级, r.姓名, r.考号, *[float(getattr(r, s)) for s in S], float(r.总分), int(r.级名次), int(r.班名次)]
                     for r in R.df.itertuples(index=False)],
        "bands": [{"label": b[0], "counts": b[3], "total": b[4]} for b in R.bands],
        "near": [{"cls": r["班级"], "name": r["姓名"], "total": float(r["总分"]), "pos": r["位置"], "gap": int(r["距线"]), "weak": r["薄弱学科"]}
                 for _, r in R.near.iterrows()],
        "conclusions": conclusions,
        "messages": [m.replace("【提示】", "") for m in msgs],
        "sections": available_sections(R),
    }


# ---------------------------------------------------------------- 最近的考试
def save_run(output_root, source_path, opts: dict, view: dict) -> Path:
    """把这次核算存下来（设置 + 界面数据 + 登分表副本），以后可以在“最近的考试”里回看、导出。"""
    out = Path(output_root) / view["exam"]
    out.mkdir(parents=True, exist_ok=True)
    src, dst = Path(source_path), out / RUN_COPY
    if src.resolve() != dst.resolve():
        shutil.copy2(src, dst)
    meta = {"savedAt": time.strftime("%Y-%m-%d %H:%M"), "opts": {**opts, "exam": view["exam"], "grade": view["grade"]}, "view": view}
    (out / RUN_FILE).write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    return out


def list_runs(output_root, limit: int = 30) -> list[dict]:
    runs = []
    root = Path(output_root)
    if not root.is_dir():
        return runs
    for f in root.glob(f"*/{RUN_FILE}"):
        try:
            m = json.loads(f.read_text(encoding="utf-8"))
            v = m["view"]
            runs.append({"name": v["exam"], "grade": v["grade"], "scheme": v["scheme"], "n": v["n"], "date": v.get("date") or "",
                         "savedAt": m.get("savedAt", ""), "_t": f.stat().st_mtime})
        except Exception:
            continue
    runs.sort(key=lambda r: -r["_t"])
    return [{k: v for k, v in r.items() if k != "_t"} for r in runs[:limit]]


def load_run(output_root, name: str) -> dict:
    f = Path(output_root) / clean_name(name) / RUN_FILE
    if not f.is_file():
        raise ValueError(f"找不到“{name}”的核算结果。")
    return json.loads(f.read_text(encoding="utf-8"))


# ---------------------------------------------------------------- 导出
def export_excel(R: Result, conclusions: list[str], output_root) -> Path:
    from .excel_report import build_excel
    out = Path(output_root) / R.cfg.exam
    out.mkdir(parents=True, exist_ok=True)
    return Path(build_excel(R, out / f"{R.cfg.exam}_各班综合统计.xlsx", conclusions))


def export_pdf(R: Result, conclusions: list[str], sections: list[str], output_root) -> Path:
    from .pdf_report import available_sections, build_html, html_to_pdf
    chosen = [k for k in available_sections(R) if k in set(sections)]
    if not chosen:
        raise ValueError("请至少勾选一项要导出的内容。")
    out = Path(output_root) / R.cfg.exam
    out.mkdir(parents=True, exist_ok=True)
    return Path(html_to_pdf(build_html(R, conclusions, chosen), out / f"{R.cfg.exam}_成绩发布版.pdf"))
