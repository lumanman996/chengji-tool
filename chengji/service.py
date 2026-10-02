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


def enrolment_notes(cfg, data: ScoreData) -> list[str]:
    """应考人数的检查和说明：填了的班，实考不能比应考多；没填的班，说明按实考人数算。"""
    if "参考率" not in cfg.structure:
        return []
    real, out = data.df["班级"].value_counts(), []
    for c in data.classes:
        e = cfg.enrolled.get(c)
        if e and real[c] > e:
            out.append(f"【错误】{c} 实考人数 {real[c]} 比应考人数 {e} 还多，请检查应考人数")
    miss = [c for c in data.classes if not cfg.enrolled.get(c)]
    if miss:
        who = "各班" if len(miss) == len(data.classes) else "、".join(miss)
        out.append(f"【提示】{who}没有填应考人数，按实考人数计算（参考率按 100% 算）。")
    return out


# ---------------------------------------------------------------- 算法（界面里可以改的比例）
STRUCT_PCT = ["平均成绩", "全科合格率", "全科优秀率", "参考率", "进线率", "增值评价"]     # 占结构分的百分比（满分 100）
STRUCT_PTS = {"增值名次差": 0.1, "增值进退步": 0.05, "前10名每人加分": 0}                 # 按“分”填的，右边是没写时的默认值


def _clean(v, nd=6):
    v = round(float(v), nd)
    return int(v) if v == int(v) else v


def algo_view(school: School) -> dict:
    """给界面的算法设置：单科得分的三个比例（百分数）、是否折算百分制；结构分各项占比（百分数）和几个按分计的数。"""
    out = {}
    for name, w in school.schemes.items():
        st = school.structures.get(name, {})
        out[name] = {
            "score": {"优秀率": _clean(w["优秀率"] * 100), "及格率": _clean(w["及格率"] * 100), "平均分": _clean(w["平均分"] * 100)},
            "normalize": bool(w["折算"]),
            "structure": {k: _clean(st.get(k, 0)) for k in STRUCT_PCT},
            "points": {k: _clean(st.get(k, d)) for k, d in STRUCT_PTS.items()},
        }
    return out


def algo_to_yaml(algo: dict, school: School) -> tuple[dict, dict]:
    """把界面交来的算法设置检查一遍，换成 学校设置.yaml 里“算法方案”“结构分方案”的写法。有问题抛 ValueError。"""
    schemes, structs = {}, {}
    for name in school.schemes:
        a = algo.get(name)
        if not a:
            raise ValueError(f"缺少“{name}”方案的设置。")
        label = {"平时": "平时考试", "中考": "中考核算"}.get(name, name)

        def num(group, key, lo=0.0, hi=100.0):
            try:
                v = float((a.get(group) or {}).get(key) or 0)
            except (TypeError, ValueError):
                raise ValueError(f"{label}的“{key}”要填数字。") from None
            if not lo <= v <= hi:
                raise ValueError(f"{label}的“{key}”应在 {lo:g}～{hi:g} 之间。")
            return v
        sc = {k: num("score", k) for k in ("优秀率", "及格率", "平均分")}
        if abs(sum(sc.values()) - 100) > 1e-6:
            raise ValueError(f"{label}的单科得分：三项加起来要等于 100%，现在是 {_clean(sum(sc.values()), 2)}%。")
        stc = {k: num("structure", k) for k in STRUCT_PCT}
        if abs(sum(stc.values()) - 100) > 1e-6:
            raise ValueError(f"{label}的班级结构分：各项加起来要等于 100%，现在是 {_clean(sum(stc.values()), 2)}%。")
        pts = {k: num("points", k, 0, 10) for k in STRUCT_PTS}
        schemes[name] = {"优秀率": _clean(sc["优秀率"] / 100), "及格率": _clean(sc["及格率"] / 100), "平均分": _clean(sc["平均分"] / 100),
                         "平均分折算百分制": "是" if a.get("normalize") else "否"}
        st = {k: _clean(v) for k, v in stc.items() if v}
        if st.get("增值评价"):
            st["增值名次差"], st["增值进退步"] = _clean(pts["增值名次差"]), _clean(pts["增值进退步"])
        if pts["前10名每人加分"]:
            st["前10名每人加分"] = _clean(pts["前10名每人加分"])
        structs[name] = st
    return schemes, structs


# 分数线：界面上的名字 → (School 的属性, 学校设置.yaml 里的写法, 是不是百分比)
LINES = {"及格线": ("pass_ratio", "及格线比例", True), "优秀线": ("excellent_ratio", "优秀线比例", True),
         "升学线": ("promote_ratio", "升学线比例", True), "高分线": ("high_ratio", "高分线比例", True),
         "临界范围": ("near_range", "临界范围", False)}


def lines_view(school: School) -> dict:
    """及格线、优秀线占满分的百分之几；升学线默认取前百分之几；高分线占总分的百分之几；临界生上下几分。"""
    return {k: _clean(getattr(school, attr) * (100 if pct else 1)) for k, (attr, _y, pct) in LINES.items()}


def lines_to_yaml(lines: dict) -> dict:
    v = {}
    for k, (_a, _y, pct) in LINES.items():
        try:
            v[k] = float(lines.get(k))
        except (TypeError, ValueError):
            raise ValueError(f"“{k}”要填数字。") from None
        if not 0 < v[k] <= 100:
            raise ValueError(f"“{k}”应在 1～100 之间。")
    if v["及格线"] >= v["优秀线"]:
        raise ValueError("优秀线要比及格线高。")
    return {y: _clean(v[k] / 100 if pct else v[k]) for k, (_a, y, pct) in LINES.items()}


def scheme_snapshot(school: School, scheme: str) -> dict:
    """这次核算用的比例和分数线，随结果一起存下来：以后改了设置，回看、导出旧考试仍按当时的算。"""
    return {"score": dict(school.schemes[scheme]), "structure": dict(school.structures.get(scheme, {})),
            "lines": {attr: getattr(school, attr) for attr, _y, _p in LINES.values() if attr != "promote_ratio"}}


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
    listed = {c: int(counts[c]) for c in data.classes}            # 登分表里这个班一共多少人（含缺考未计入的）
    for e in data.excluded:
        if e["班级"] in listed:
            listed[e["班级"]] += 1
    schemes = []
    for name in school.schemes:
        w = {k: v for k, v in school.structures.get(name, {}).items() if v}
        schemes.append({"name": name, "note": subject_note(school, g, name, data.subjects).replace("【提示】", ""),
                        "needEnrolled": "参考率" in w, "needPrev": "增值评价" in w,
                        "formula": make_config(_with_full(school, g, data.subjects), g, data.subjects, "x", scheme=name).formula_text()})
    return {
        "file": Path(path).name, "grade": g, "classes": data.classes,
        "counts": {c: int(counts[c]) for c in data.classes}, "listed": listed, "n": len(data.df),
        "subjects": data.subjects, "newSubjects": data.new_subjects,
        "full": {s: school.grade_full.get(g, {}).get(s, school.full.get(s, 0)) for s in data.subjects},
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
    """按确认好的设置计算。opts：scheme、exam、date、full{科目:满分}、ratio、enrolled{班:人数}、prev{班:名次}、grade、
    exclude[本次不计的科目]、schemeDef（当时的比例快照，回看旧考试时用）。
    返回 (Result, 自动起草的结论, 提示信息)。数据有硬伤（超过满分、应考人数不对等）时抛 ValueError。"""
    import copy
    school = copy.deepcopy(school)
    data = load_scores(path, school, opts.get("grade"), exclude=opts.get("exclude") or ())
    g = data.grade
    scheme = opts.get("scheme") or next(iter(school.schemes))
    snap = opts.get("schemeDef")                               # 回看旧考试：用当时存下来的比例
    if snap:
        school.schemes[scheme] = dict(snap["score"])
        school.structures[scheme] = dict(snap["structure"])
        for attr, val in (snap.get("lines") or {}).items():
            setattr(school, attr, val)
    if scheme not in school.schemes:
        raise ValueError(f"没有名为“{scheme}”的算法方案")

    full = {s: school.grade_full.get(g, {}).get(s, school.full.get(s, 0)) for s in data.subjects}
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
        enrolled = {c: v for c, v in enrolled.items() if v}       # 没填的班不强求：计算时按实考人数代替
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
    msgs += enrolment_notes(cfg, data)
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


def last_run_for(output_root, grade: str, not_name: str = "") -> dict | None:
    """同一个年级最近一次算过的考试：把它的结构分名次、应考人数带出来，供这次预填（使用者可以改）。"""
    root = Path(output_root)
    best = None
    for f in root.glob(f"*/{RUN_FILE}") if root.is_dir() else []:
        try:
            m = json.loads(f.read_text(encoding="utf-8"))
            v = m["view"]
        except Exception:
            continue
        if v.get("grade") != grade or v.get("exam") == not_name or not v.get("structure"):
            continue
        if best is None or f.stat().st_mtime > best[0]:
            best = (f.stat().st_mtime, v)
    if not best:
        return None
    v = best[1]
    return {"name": v["exam"], "date": v.get("date") or "",
            "ranks": {x["name"]: x["名次"] for x in v["structure"]},
            "enrolled": {x["name"]: x["应考人数"] for x in v["structure"] if not x.get("应考按实考")}}


DELETED = "已删除"          # 删掉的考试移到 output/已删除/ 里，需要时可以找回


def delete_run(output_root, name: str) -> Path:
    src = Path(output_root) / clean_name(name)
    if not (src / RUN_FILE).is_file():
        raise ValueError(f"找不到“{name}”。")
    bin_ = Path(output_root) / DELETED
    bin_.mkdir(exist_ok=True)
    dst = bin_ / f"{src.name}_{time.strftime('%Y%m%d_%H%M%S')}"
    shutil.move(str(src), str(dst))
    return dst


def rename_run(output_root, name: str, to: str) -> str:
    """给算过的考试改名：文件夹、里面带考试名的文件、存着的结果一起改。"""
    old, new = clean_name(name), clean_name(to)
    src, dst = Path(output_root) / old, Path(output_root) / new
    if not new:
        raise ValueError("请填写新的名称。")
    if new == old:
        return new
    if not (src / RUN_FILE).is_file():
        raise ValueError(f"找不到“{name}”。")
    if dst.exists():
        raise ValueError(f"已经有一个叫“{new}”的考试了，请换个名称。")
    src.rename(dst)
    for f in list(dst.iterdir()):
        if f.is_file() and f.name.startswith(old + "_"):
            f.rename(dst / (new + f.name[len(old):]))
    m = json.loads((dst / RUN_FILE).read_text(encoding="utf-8"))
    m["view"]["exam"] = new
    m.setdefault("opts", {})["exam"] = new
    (dst / RUN_FILE).write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
    return new


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
