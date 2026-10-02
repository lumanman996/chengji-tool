"""用 pandas 独立计算全部统计结果。

这些结果用于：PDF 发布版、自动起草结论、以及核对 Excel 公式（两种方法结果必须一致）。
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

import pandas as pd

from .config import Config
from .loader import ScoreData


def subject_stats(scores: pd.Series, cfg: Config, s: str) -> dict:
    n = len(scores)
    avg = scores.mean() if n else 0.0
    pr = (scores >= cfg.pass_line(s)).mean() if n else 0.0
    er = (scores >= cfg.excellent_line(s)).mean() if n else 0.0
    score = er * 100 * cfg.w_excellent + pr * 100 * cfg.w_pass + cfg.avg_term(avg, s) * cfg.w_avg
    return {"人数": n, "总分": scores.sum(), "平均分": avg, "及格率": pr, "优秀率": er, "得分": score}


@dataclass
class Result:
    cfg: Config
    data: ScoreData
    df: pd.DataFrame
    classes: list[str]
    class_stats: dict = field(default_factory=dict)      # 班级/全年级 -> 科目 -> 指标
    composite: dict = field(default_factory=dict)        # 班级 -> 综合得分
    class_rank: dict = field(default_factory=dict)
    cut_rank: int = 0
    cut: float = 0
    above: int = 0
    below: int = 0
    bands: list = field(default_factory=list)           # (标签, 下限, 上限, {班级:人数}, 合计)
    below_counts: dict = field(default_factory=dict)
    top_n: list = field(default_factory=list)           # (N, 最低总分, {班级:人数}, 实际人数)
    online: dict = field(default_factory=dict)          # 班级/全级 -> 上线情况
    subj_compare: dict = field(default_factory=dict)    # 科目 -> 线上/线下对比
    near: pd.DataFrame | None = None
    teachers: list = field(default_factory=list)        # dict 列表
    structure: dict = field(default_factory=dict)       # 班级 -> 结构分各项（见 structure_scores）


def _weakest(row, cfg):
    return min(cfg.subjects, key=lambda s: row[s] / cfg.full[s])


def analyze(data: ScoreData, cfg: Config) -> Result:
    df = data.df.copy()
    df["总分"] = df[cfg.subjects].sum(axis=1)
    df["级名次"] = df["总分"].rank(ascending=False, method="min").astype(int)
    df["班名次"] = df.groupby("班级")["总分"].rank(ascending=False, method="min").astype(int)
    R = Result(cfg=cfg, data=data, df=df, classes=data.classes)

    groups = {c: df[df["班级"] == c] for c in R.classes}
    groups["全年级"] = df
    for c, g in groups.items():
        R.class_stats[c] = {s: subject_stats(g[s], cfg, s) for s in cfg.subjects}
        R.composite[c] = sum(v["得分"] for v in R.class_stats[c].values())
    ordered = sorted(R.classes, key=lambda c: -R.composite[c])
    R.class_rank = {c: ordered.index(c) + 1 for c in R.classes}

    # 升学线
    n = len(df)
    R.cut_rank = int(round(n * cfg.promote_ratio + 1e-9))
    tot_sorted = df["总分"].sort_values(ascending=False).reset_index(drop=True)
    R.cut = float(tot_sorted[R.cut_rank - 1])
    R.above = int((df["总分"] >= R.cut).sum())
    R.below = n - R.above

    # 分数段（线上）
    w = cfg.band_width
    # 最高一段从最高分所在段开始；人数少于 3 人时向下合并一段
    top_lo = df["总分"].max() // w * w
    while top_lo - w > R.cut and (df["总分"] >= top_lo).sum() < 3:
        top_lo -= w
    lo = top_lo
    edges = []
    while lo > R.cut:
        edges.append(lo)
        lo -= w
    for i, lo in enumerate(edges):
        hi = cfg.total_full + 1 if i == 0 else edges[i - 1]
        label = f"{int(lo)}分及以上" if i == 0 else f"{int(lo)}–{int(hi - 1)}"
        sub = df[(df["总分"] >= lo) & (df["总分"] < hi)]
        cnt = {c: int((sub["班级"] == c).sum()) for c in R.classes}
        R.bands.append((label, lo, hi, cnt, sum(cnt.values())))
    low_hi = edges[-1] if edges else cfg.total_full + 1
    sub = df[(df["总分"] >= R.cut) & (df["总分"] < low_hi)]
    cnt = {c: int((sub["班级"] == c).sum()) for c in R.classes}
    R.bands.append((f"{int(R.cut)}–{int(low_hi - 1)}（线上最低段）", R.cut, low_hi, cnt, sum(cnt.values())))
    R.below_counts = {c: int(((df["班级"] == c) & (df["总分"] < R.cut)).sum()) for c in R.classes}

    # 前N名
    for k in cfg.top_n + [R.cut_rank]:
        if k > n:
            continue
        line = float(tot_sorted[k - 1])
        sub = df[df["总分"] >= line]
        R.top_n.append((k, line, {c: int((sub["班级"] == c).sum()) for c in R.classes}, len(sub)))

    # 上线情况
    rng = cfg.near_range
    for c, g in groups.items():
        t = g["总分"]
        up = t >= R.cut
        R.online[c] = {
            "实考人数": len(g), "线上": int(up.sum()), "线下": int((~up).sum()),
            "上线率": up.mean() if len(g) else 0,
            "线上临界": int(((t >= R.cut) & (t < R.cut + rng)).sum()),
            "线下临界": int(((t < R.cut) & (t >= R.cut - rng)).sum()),
            "高分": int((t >= cfg.high_line).sum()),
            "线上平均总分": t[up].mean() if up.any() else 0,
        }
    up_df, dn_df = df[df["总分"] >= R.cut], df[df["总分"] < R.cut]
    for s in cfg.subjects:
        R.subj_compare[s] = {
            "线上平均分": up_df[s].mean() if len(up_df) else 0,
            "线下平均分": dn_df[s].mean() if len(dn_df) else 0,
            "线上及格率": (up_df[s] >= cfg.pass_line(s)).mean() if len(up_df) else 0,
            "线下及格率": (dn_df[s] >= cfg.pass_line(s)).mean() if len(dn_df) else 0,
        }

    # 临界生
    near = df[(df["总分"] >= R.cut - rng) & (df["总分"] < R.cut + rng)].copy()
    near["位置"] = near["总分"].map(lambda t: "线上" if t >= R.cut else "线下")
    near["距线"] = (near["总分"] - R.cut).astype(int)
    near["薄弱学科"] = near.apply(lambda r: _weakest(r, cfg), axis=1)
    R.near = near.sort_values("总分", ascending=False)

    # 教师
    for t in cfg.teachers:
        g = df[df["班级"].isin(t.classes)]
        st = subject_stats(g[t.subject], cfg, t.subject)
        R.teachers.append({"学科": t.subject, "教师": t.name, "班级": "、".join(t.classes), **st})
    for s in cfg.subjects:
        ts = sorted([x for x in R.teachers if x["学科"] == s], key=lambda x: -x["得分"])
        for i, x in enumerate(ts):
            x["名次"] = i + 1

    if cfg.structure:
        R.structure = structure_scores(R)
    return R


# ---------------- 班级结构分 ----------------
RATE_ITEMS = ["平均成绩", "全科合格率", "全科优秀率", "参考率", "进线率"]   # 比率 × 分值 的各项


def _rank_desc(vals: dict) -> dict:
    """从高到低排名，同分同名次。"""
    return {k: 1 + sum(1 for x in vals.values() if x > v) for k, v in vals.items()}


def structure_scores(R: Result) -> dict:
    """班级结构分。各项直接按比例折分（比率 × 分值），再加增值评价、前10名加分。"""
    cfg, df = R.cfg, R.df
    W = cfg.structure
    if "参考率" in W:
        miss = [c for c in R.classes if not cfg.enrolled.get(c)]
        if miss:
            raise ValueError(f"算参考率需要应考人数，以下班级没有：{'、'.join(miss)}。"
                             f"请填在任课总表的「班级信息」工作表里。")
    all_pass = pd.Series(True, index=df.index)
    all_exc = pd.Series(True, index=df.index)
    for s in cfg.subjects:
        all_pass &= df[s] >= cfg.pass_line(s)
        all_exc &= df[s] >= cfg.excellent_line(s)
    tot_sorted = df["总分"].sort_values(ascending=False).reset_index(drop=True)
    top10_line = float(tot_sorted[min(10, len(df)) - 1])

    out = {}
    for c in R.classes:
        m = df["班级"] == c
        n = int(m.sum())
        x = {"应考人数": cfg.enrolled.get(c, 0), "实考人数": n,
             "总分平均": df.loc[m, "总分"].mean(),
             "全科合格人数": int(all_pass[m].sum()), "全科优秀人数": int(all_exc[m].sum()),
             "进线人数": int((df.loc[m, "总分"] >= R.cut).sum())}
        rate = {"平均成绩": x["总分平均"] / cfg.total_full,
                "全科合格率": x["全科合格人数"] / n, "全科优秀率": x["全科优秀人数"] / n,
                "参考率": x["实考人数"] / x["应考人数"] if x["应考人数"] else 0,
                "进线率": x["进线人数"] / n}
        for k in RATE_ITEMS:
            if k in W:
                x[k] = rate[k]
                x[k + "分"] = rate[k] * W[k]
        x["小计"] = sum(x[k + "分"] for k in RATE_ITEMS if k in W)
        x["前10名人数"] = int((df.loc[m, "总分"] >= top10_line).sum())
        x["加分"] = x["前10名人数"] * W.get("前10名每人加分", 0)
        out[c] = x

    sub_rank = _rank_desc({c: out[c]["小计"] for c in R.classes})
    for c in R.classes:
        x = out[c]
        x["小计名次"] = sub_rank[c]
        x["上次名次"] = cfg.prev_rank.get(c)
        x["增值评价分"] = 0.0
        if "增值评价" in W:
            v = W["增值评价"] - W.get("增值名次差", 0.1) * (sub_rank[c] - 1)
            if x["上次名次"]:
                v += (x["上次名次"] - sub_rank[c]) * W.get("增值进退步", 0.05)   # 进步为正
            x["增值评价分"] = min(v, W["增值评价"])
        x["结构分"] = x["小计"] + x["增值评价分"] + x["加分"]
    final = _rank_desc({c: out[c]["结构分"] for c in R.classes})
    for c in R.classes:
        out[c]["名次"] = final[c]
    return out


# ---------------- 自动起草结论（需人工审核） ----------------
def pct(x, d=1):
    return f"{x * 100:.{d}f}%"


def draft_conclusions(R: Result) -> list[str]:
    cfg, C = R.cfg, R.classes
    out = []
    n = len(R.df)
    out.append(f"升学线：全级有效成绩{n}人，前{pct(cfg.promote_ratio, 0)}为第{R.cut_rank}名，对应{len(cfg.subjects)}科总分"
               f"{R.cut:g}分（满分{cfg.total_full:g}）；线上{R.above}人（{pct(R.above / n)}），线下{R.below}人。")
    big = max(R.bands[:-1] or R.bands, key=lambda b: b[4])
    out.append(f"分数段：{big[0]}分人数最多（{big[4]}人）；{R.bands[0][0]}{R.bands[0][4]}人，"
               f"其中{max(C, key=lambda c: R.bands[0][3][c])}最多。")
    ranked = sorted(C, key=lambda c: R.class_rank[c])
    best_sub = {}
    worst_sub = {}
    for s in cfg.subjects:
        vals = {c: R.class_stats[c][s]["得分"] for c in C}
        best_sub.setdefault(max(vals, key=vals.get), []).append(s)
        worst_sub.setdefault(min(vals, key=vals.get), []).append(s)
    first, last = ranked[0], ranked[-1]
    t = f"班级：{first}综合第一（{R.composite[first]:.2f}）"
    if best_sub.get(first):
        t += f"，{'、'.join(best_sub[first])}得分全级最高"
    t += f"；{last}综合最后（{R.composite[last]:.2f}）"
    if worst_sub.get(last):
        t += f"，{'、'.join(worst_sub[last])}得分全级最低"
    lowrate = min(C, key=lambda c: R.online[c]["上线率"])
    t += f"。上线率最低的是{lowrate}（{pct(R.online[lowrate]['上线率'])}），线下{R.online[lowrate]['线下']}人。"
    out.append(t)
    if R.structure:      # 结构分是另一张表，结论也单列一条
        X = R.structure
        sr = sorted(C, key=lambda c: X[c]["名次"])
        out.append("班级结构分：" + "、".join(f"{c}第{X[c]['名次']}（{X[c]['结构分']:.2f}）" for c in sr) + "。")
    g = R.class_stats["全年级"]
    weak = [s for s in cfg.subjects if g[s]["及格率"] < 0.5]
    strong = sorted(cfg.subjects, key=lambda s: -g[s]["及格率"])[:2]
    if weak:
        out.append(f"学科短板：{'、'.join(weak)}全级及格率不足50%（{'、'.join(pct(g[s]['及格率']) for s in weak)}）；"
                   f"及格率最高的是{'、'.join(strong)}（{'、'.join(pct(g[s]['及格率']) for s in strong)}）。")
    near = R.near
    wc = Counter(near["薄弱学科"])
    nd = int((near["位置"] == "线下").sum())
    out.append(f"临界生：升学线上下{cfg.near_range:g}分内共{len(near)}人，线下{nd}人（冲线重点）、线上{len(near) - nd}人（需稳住）。"
               f"薄弱学科：" + "、".join(f"{s}{k}人" for s, k in wc.most_common()) + "。")
    first10 = next((x for x in R.top_n if x[0] == 10), None)
    if first10:
        mx = max(C, key=lambda c: first10[2][c])
        zero = [c for c in C if first10[2][c] == 0]
        t = f"前N名：前10名{mx}占{first10[2][mx]}人"
        if zero:
            t += f"，{'、'.join(zero)}无人进入"
        out.append(t + "。")
    out.append("线下学生：" + f"{R.below}人各科平均分为" +
               "、".join(f"{s}{R.subj_compare[s]['线下平均分']:.1f}" for s in cfg.subjects) + "，建议以基础巩固为主。")
    return out
