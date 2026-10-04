"""两次考试的对比（“和上次比”），供导出 Excel、PDF 用。

本次用这次的计算结果（Result）；上次用当时存下来的结果（output/<考试>/结果.json 里的界面数据），
和界面上“和上次比”看到的是同一份数据。只读，不改动任何一次考试。

- 各班：结构分和名次、综合得分名次、上线率。
- 各科（全年级）：平均分、得分率（平均分 ÷ 满分，两次满分可能不同，所以比得分率）、及格率、优秀率。
- 学生：按“班级 + 姓名”对应；同班重名、换了班、缺考的不比。
- 上次的临界生这次怎么样了：线下的有没有上线，线上的有没有掉下去。
"""
from __future__ import annotations

from collections import Counter

from .analysis import Result


def _count(pct: float, n: int) -> int:
    """界面数据里的及格率、优秀率是保留两位小数的百分数；乘回人数取整就是准确的人数。"""
    return int(round(pct * n / 100))


def where(prev_pos: str, pos: str | None) -> str:
    """上次的临界生这次的去向。"""
    if pos is None:
        return "这次没有对上"
    if prev_pos == "线下":
        return "上线了" if pos == "线上" else "仍在线下"
    return "掉到线下" if pos == "线下" else "仍在线上"


def build(R: Result, P: dict) -> dict:
    """R：本次的计算结果；P：上次的界面数据（service.load_run(...)["view"]）。"""
    cfg = R.cfg
    if P.get("grade") != cfg.grade:
        raise ValueError(f"只能和同一个年级的考试比（上次是{P.get('grade')}，这次是{cfg.grade}）。")

    # ---- 各班
    has_s = bool(R.structure) and bool(P.get("structure"))
    pS = {x["name"]: x for x in P.get("structure") or []}
    pC = {c["name"]: c for c in P["classes"]}
    classes = []
    for c in R.classes:
        p = pC.get(c)
        row = {"name": c, "found": p is not None, "rank": R.class_rank[c], "onlineRate": float(R.online[c]["上线率"])}
        if has_s:
            row.update(struct=float(R.structure[c]["结构分"]), structRank=int(R.structure[c]["名次"]))
        if p:
            row.update(prevRank=int(p["rank"]), prevOnlineRate=p["online"] / p["n"] if p["n"] else 0.0)
            if has_s and c in pS:
                row.update(prevStruct=float(pS[c]["结构分"]), prevStructRank=int(pS[c]["名次"]))
        classes.append(row)
    classes.sort(key=lambda x: x["structRank"] if has_s else x["rank"])

    # ---- 各科（全年级）
    ps, pstud = P["subjects"], P["students"]
    n_prev = sum(c["n"] for c in P["classes"])
    subjects = []
    for s in [s for s in cfg.subjects if s in ps]:
        k = ps.index(s)
        avg_p = sum(r[3 + k] for r in pstud) / len(pstud) if pstud else 0.0
        g = R.class_stats["全年级"][s]
        subjects.append({
            "subject": s, "prevFull": float(P["full"][s]), "full": float(cfg.full[s]),
            "prevAvg": avg_p, "avg": float(g["平均分"]),
            "prevAvgRate": avg_p / P["full"][s], "avgRate": float(g["平均分"]) / cfg.full[s],
            "prevPass": sum(_count(c["subj"][s]["pass"], c["n"]) for c in P["classes"]) / n_prev, "pass": float(g["及格率"]),
            "prevExc": sum(_count(c["subj"][s]["exc"], c["n"]) for c in P["classes"]) / n_prev, "exc": float(g["优秀率"]),
        })

    # ---- 学生：两次都参加、按“班级 + 姓名”能唯一对上的
    nsp = len(ps)
    c0 = Counter((r[0], r[1]) for r in pstud)
    c1 = Counter(zip(R.df["班级"], R.df["姓名"]))
    old = {(r[0], r[1]): r for r in pstud if c0[(r[0], r[1])] == 1}
    students = []
    for r in R.df.itertuples(index=False):
        k = (r.班级, r.姓名)
        if c1[k] == 1 and k in old:
            o = old[k]
            students.append({"cls": r.班级, "name": r.姓名, "prevTotal": float(o[3 + nsp]), "total": float(r.总分),
                             "prevRank": int(o[4 + nsp]), "rank": int(r.级名次), "change": int(o[4 + nsp]) - int(r.级名次)})
    students.sort(key=lambda x: (-x["change"], x["rank"]))

    # ---- 上次的临界生
    now = {(x["cls"], x["name"]): x for x in students}
    near = []
    for x in P.get("near") or []:
        m = now.get((x["cls"], x["name"]))
        pos = None if m is None else ("线上" if m["total"] >= R.cut else "线下")
        near.append({"cls": x["cls"], "name": x["name"], "prevPos": x["pos"], "prevTotal": float(x["total"]),
                     "found": m is not None, "total": m["total"] if m else None, "pos": pos, "where": where(x["pos"], pos)})
    near.sort(key=lambda x: (x["prevPos"] != "线下", int("".join(ch for ch in x["cls"] if ch.isdigit()) or 0), -x["prevTotal"]))

    def tally(prev_pos, goal):
        L = [x for x in near if x["prevPos"] == prev_pos]
        return {"all": len(L), "found": sum(x["found"] for x in L), "moved": sum(x["where"] == goal for x in L)}

    return {
        "prev": {"exam": P["exam"], "date": P.get("date") or "", "scheme": P["scheme"], "n": P["n"], "cut": float(P["cut"])},
        "now": {"exam": cfg.exam, "scheme": cfg.scheme, "n": len(R.df), "cut": float(R.cut)},
        "hasStructure": has_s, "classes": classes, "subjects": subjects, "students": students, "near": near,
        "nearUp": tally("线下", "上线了"), "nearDown": tally("线上", "掉到线下"),
    }
