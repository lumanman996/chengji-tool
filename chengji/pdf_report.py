"""生成成绩发布版（A4 横版 PDF）：先拼 HTML，再用浏览器内核（Chromium / Edge）打印成 PDF。"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

from ._style import CSS, f1, f2, hbars, pct
from .analysis import RATE_ITEMS, Result


# 发布版可以选择导出的内容（顺序即页面顺序）。右边是页面标题里用来认出这一页的字样。
COMPARE = "和上次比"        # 只有选了和哪一次考试比，才有这一项
SECTIONS = ["成绩通报", "各科得分", "各科明细", "班级结构分", "分数段", "前N名", "教师排名", "临界生名单", COMPARE]
_TITLE_KEY = [("和上次", COMPARE), ("成绩通报", "成绩通报"), ("各科得分", "各科得分"), ("各科明细", "各科明细"), ("班级结构分", "班级结构分"),
              ("前N名", "前N名"), ("分数段", "分数段"), ("教师", "教师排名"), ("临界生", "临界生名单")]


def available_sections(R: Result, has_compare: bool = False) -> list[str]:
    """本次考试实际有的内容（没有任课教师就没有教师排名，没有结构分方案就没有结构分，没选对比就没有“和上次比”）。"""
    return [k for k in SECTIONS if not (k == "班级结构分" and not R.structure) and not (k == "教师排名" and not R.teachers)
            and not (k == COMPARE and not has_compare)]


def _mark(vals, v):
    if len(set(vals)) == 1:
        return ""
    return ' class="best"' if v == max(vals) else (' class="worst"' if v == min(vals) else "")


def build_html(R: Result, conclusions: list[str], sections: list[str] | None = None, compare: dict | None = None) -> str:
    """sections：要导出的内容名称列表（见 SECTIONS）；None = 全部。compare：和上次比的数据（compare.build），没有就不出这两页。"""
    cfg, CL = R.cfg, R.classes
    S, NS = cfg.subjects, len(cfg.subjects)
    st, on, g = R.class_stats, R.online, R.class_stats["全年级"]
    pages = []
    norm_txt = "，各科平均分折算百分制后计分" if cfg.normalize else "，平均分按原始分计算（满分高的科目比重更大）"

    def page(title, body):
        key = next(k for mark, k in _TITLE_KEY if mark in title)
        if sections is not None and key not in sections:
            return
        date = f"　考试日期：{cfg.date}" if cfg.date else ""
        pages.append(f'<div class="page"><div class="hd"><div class="t">{title}</div><div class="s">{cfg.school}<br>{cfg.exam}{date}</div></div>'
                     f'{body}<div class="ft"><span>{cfg.org}</span><span>第 {{PNO}} 页</span></div></div>')

    n = len(R.df)
    tn = {k: (line, cnt, real) for k, line, cnt, real in R.top_n}
    t50 = tn.get(50)
    excl = f"另有 {len(R.data.excluded)} 人成绩不全，未计入" if R.data.excluded else "全部计入"
    tiles = (f'<div class="tiles"><div class="tile"><div class="k">实考人数</div><div class="v">{n}<small>人</small></div><div class="d">{excl}</div></div>'
             f'<div class="tile"><div class="k">升学参考线（全级前{pct(cfg.promote_ratio, 0)}）</div><div class="v">{R.cut:g}<small>分</small></div><div class="d">{NS}科总分，满分 {cfg.total_full:g}</div></div>'
             f'<div class="tile"><div class="k">线上人数</div><div class="v">{R.above}<small>人 · {pct(R.above / n)}</small></div><div class="d">线下 {R.below} 人</div></div>')
    if t50:
        extra = " · ".join(f"前{k}名 {tn[k][0]:g} 分" for k in (20, 100) if k in tn)
        tiles += f'<div class="tile"><div class="k">前50名分数线</div><div class="v">{t50[0]:g}<small>分</small></div><div class="d">{extra}</div></div>'
    tiles += "</div>"
    ranked = sorted(CL, key=lambda c: R.class_rank[c])
    rows = "".join(
        f'<tr{" class=hl" if R.class_rank[c] == 1 else ""}><td><span class="rk{" rk1" if R.class_rank[c] == 1 else ""}">{R.class_rank[c]}</span></td>'
        f'<td class="b">{c}</td><td class="b">{f2(R.composite[c])}</td><td>{on[c]["线上"]}/{on[c]["实考人数"]}（{pct(on[c]["上线率"])}）</td><td>{on[c]["高分"]}</td></tr>'
        for c in ranked)
    page(f"{cfg.grade}成绩通报",
         tiles + f'<div class="row"><div style="flex:1"><h3>各班综合排名</h3><table><tr><th>名次</th><th>班级</th><th>{NS}科综合得分</th><th>上线人数（上线率）</th><th>{cfg.high_line:g}分以上</th></tr>{rows}</table>'
         f'<p class="note" style="margin-top:6px">{cfg.scheme}方案：单科得分 = {cfg.formula_text()}；{NS}科综合 = {NS}科得分之和。</p></div>'
         f'<div style="flex:1.1"><h3>本次考试主要结论</h3><ul class="find" style="font-size:13px;line-height:1.6">' +
         "".join(f"<li>{c}</li>" for c in conclusions) + "</ul></div></div>")

    # 各班得分
    hdr = '<tr><th>班级</th><th>人数</th>' + "".join(f'<th>{s}<br><span style="font-weight:400;font-size:11px">满分{cfg.full[s]:g}</span></th>' for s in S) + f'<th>{NS}科综合</th><th>名次</th></tr>'
    body = ""
    for c in CL:
        body += f'<tr><td class="b">{c}</td><td>{on[c]["实考人数"]}</td>'
        for s in S:
            vals = [st[x][s]["得分"] for x in CL]
            body += f'<td{_mark(vals, st[c][s]["得分"])}>{f2(st[c][s]["得分"])}</td>'
        body += f'<td class="b">{f2(R.composite[c])}</td><td class="b">{R.class_rank[c]}</td></tr>'
    body += f'<tr class="tot"><td>全年级</td><td>{n}</td>' + "".join(f'<td>{f2(g[s]["得分"])}</td>' for s in S) + f'<td>{f2(R.composite["全年级"])}</td><td>—</td></tr>'
    feats = []
    for c in ranked:
        best = [s for s in S if st[c][s]["得分"] == max(st[x][s]["得分"] for x in CL)]
        worst = [s for s in S if st[c][s]["得分"] == min(st[x][s]["得分"] for x in CL)]
        t = f"{c}（第{R.class_rank[c]}）："
        t += ("、".join(best) + "全级最高") if best else ""
        t += ("；" if best and worst else "") + (("、".join(worst) + "全级最低") if worst else "")
        if best or worst:
            feats.append(t + "。")
    chart = hbars([(c, R.composite[c]) for c in CL], w=520, color="#1c5cab", vmax=max(R.composite[c] for c in CL) * 1.15, fmt=f2, rowh=34)
    page("各班成绩统计（一）：各科得分",
         f'<h3>各班各科得分（蓝色为该科最高，红色为该科最低）</h3><table style="font-size:13px">{hdr}{body}</table>'
         f'<p class="note" style="margin:6px 0 14px">及格线为满分的{pct(cfg.pass_ratio, 0)}，优秀线为满分的{pct(cfg.excellent_ratio, 0)}；单科得分 = {cfg.formula_text()}{norm_txt}。</p>'
         f'<div class="row"><div style="flex:1"><h3>{NS}科综合得分</h3>{chart}</div><div style="flex:1"><h3>各班特点</h3><ul class="find" style="font-size:13px">'
         + "".join(f"<li>{x}</li>" for x in feats) + "</ul></div></div>")

    # 各科明细
    hdr = '<tr><th>学科</th><th>指标</th>' + "".join(f"<th>{c}</th>" for c in CL) + '<th>全年级</th></tr>'
    body = ""
    for s in S:
        for m, (lab, fmt) in enumerate([("平均分", f2), ("及格率", pct), ("优秀率", pct)]):
            vals = [st[c][s][lab] for c in CL]
            body += "<tr>" + (f'<td rowspan=3 class="b">{s}<br><span style="font-weight:400;font-size:11px;color:#5b6168">满分{cfg.full[s]:g}</span></td>' if m == 0 else "")
            body += f"<td>{lab}</td>" + "".join(f'<td{_mark(vals, st[c][s][lab])}>{fmt(st[c][s][lab])}</td>' for c in CL) + f'<td class="b">{fmt(g[s][lab])}</td></tr>'
    lines = "、".join(f"{s}{cfg.pass_line(s):g}" for s in S)
    exc = "、".join(f"{s}{cfg.excellent_line(s):g}" for s in S)
    pad = "2.5px 4px" if NS > 5 else "5px 4px"
    page("各班成绩统计（二）：各科明细",
         f'<table style="font-size:12px">{hdr}{body}</table>'.replace("<td>", f'<td style="padding:{pad}">') +
         f'<p class="note" style="margin-top:5px">每行蓝色为最高、红色为最低。及格线：{lines}；优秀线：{exc}。</p>')

    # 班级结构分
    if R.structure:
        page("班级结构分评比", _structure_body(R))

    # 分数段（一）
    bh = '<tr><th>分数段</th>' + "".join(f"<th>{c}</th>" for c in CL) + '<th>全级</th><th>占线上</th><th>累计</th></tr>'
    bb, cum = "", 0
    for lab, lo, hi, cnt, total in R.bands:
        cum += total
        bb += f'<tr><td class="b">{lab}</td>' + "".join(f"<td>{cnt[c] or '–'}</td>" for c in CL) + f'<td class="b">{total}</td><td>{pct(total / R.above)}</td><td>{cum}</td></tr>'
    bb += f'<tr style="background:#fdecec"><td class="b dn">线下（{R.cut:g}分以下）</td>' + "".join(f"<td>{R.below_counts[c]}</td>" for c in CL) + f'<td class="b">{R.below}</td><td colspan=2>占全级 {pct(R.below / n)}</td></tr>'
    bb += '<tr class="tot"><td>实考人数</td>' + "".join(f'<td>{on[c]["实考人数"]}</td>' for c in CL) + f"<td>{n}</td><td colspan=2></td></tr>"
    rate = hbars([(c, on[c]["上线率"]) for c in CL], w=390, color="#1c5cab", vmax=1.0, fmt=lambda v: pct(v),
                 ref=on["全年级"]["上线率"], reflab=f'全级 {pct(on["全年级"]["上线率"])}', rowh=36)
    ot = f'<tr><th>班级</th><th>线上</th><th>线下</th><th>线上临界</th><th>线下临界</th><th>{cfg.high_line:g}分以上</th></tr>' + "".join(
        f'<tr><td class="b">{c}</td><td>{on[c]["线上"]}</td><td>{on[c]["线下"]}</td><td>{on[c]["线上临界"]}</td><td class="dn">{on[c]["线下临界"]}</td><td>{on[c]["高分"]}</td></tr>' for c in CL)
    big = max(R.bands, key=lambda b: b[4])
    page("分数段成绩统计（一）",
         f'<div class="row"><div style="flex:1.25"><h3>线上学生分数段分布（{NS}科总分，升学线 {R.cut:g} 分）</h3><table>{bh}{bb}</table>'
         f'<p class="note" style="margin-top:6px">{big[0]} 分人数最多（{big[4]} 人）。</p></div>'
         f'<div style="flex:1"><h3>各班上线率</h3>{rate}<h3 style="margin-top:10px">各班上线与临界情况</h3><table style="font-size:12px">{ot}</table>'
         f'<p class="note" style="margin-top:4px">临界生：升学线上下 {cfg.near_range:g} 分以内。线上临界需稳住，线下临界是冲线重点。</p></div></div>')

    # 分数段（二）前N名
    th = '<tr><th>名次段</th><th>最低总分</th>' + "".join(f"<th>{c}</th>" for c in CL) + '<th>实际人数</th></tr>'
    tb = ""
    for k, line, cnt, real in R.top_n:
        is_cut = k == R.cut_rank
        lab = f"前{pct(cfg.promote_ratio, 0)}（升学线）" if is_cut else f"前{k}名"
        mx = max(cnt.values())
        tb += f'<tr{" style=background:#eef5ee" if is_cut else ""}><td class="b">{lab}</td><td>{line:g}</td>' + "".join(
            f'<td{" style=font-weight:700;color:#104281" if cnt[c] == mx and mx > 0 else ""}>{cnt[c]}<span style="color:#8a9096;font-size:11px;font-weight:400">（{pct(cnt[c] / on[c]["实考人数"], 0)}）</span></td>' for c in CL) + f"<td>{real}</td></tr>"
    sc = R.subj_compare
    stt = '<tr><th>学科</th><th>满分</th><th>线上均分</th><th>线下均分</th><th>线上及格率</th><th>线下及格率</th></tr>' + "".join(
        f'<tr><td class="b">{s}</td><td>{cfg.full[s]:g}</td><td>{f1(sc[s]["线上平均分"])}</td><td>{f1(sc[s]["线下平均分"])}</td>'
        f'<td class="{"dn" if sc[s]["线上及格率"] < 0.5 else ""}">{pct(sc[s]["线上及格率"])}</td><td>{pct(sc[s]["线下及格率"])}</td></tr>' for s in S)
    points = []
    if 10 in tn:
        c10 = tn[10][1]
        mx = max(CL, key=lambda c: c10[c])
        zero = [c for c in CL if c10[c] == 0]
        points.append(f"前10名{mx}占{c10[mx]}人" + (f"，{'、'.join(zero)}无人进入" if zero else "") + "。")
    if 50 in tn:
        c50 = tn[50][1]
        mx = max(CL, key=lambda c: c50[c])
        points.append(f"前50名{mx}有{c50[mx]}人最多。")
    weak = [s for s in S if sc[s]["线上及格率"] < 0.5]
    if weak:
        points.append(f"线上学生{'、'.join(weak)}及格率不到50%，是提分主要空间。")
    page("分数段成绩统计（二）：前N名分布",
         f'<h3>前 N 名各班人数（累计；括号内为占本班人数比例；每行最多的班级加粗）</h3><table style="font-size:12.5px">{th}{tb}</table>'
         f'<p class="note" style="margin:4px 0 10px">同分学生全部计入，因此“实际人数”可能多于名次。</p>'
         f'<div class="row"><div style="flex:1.2"><h3>线上 / 线下学生各科对比</h3><table style="font-size:12px">{stt}</table></div>'.replace("<td>", '<td style="padding:3px 4px">') +
         f'<div style="flex:1"><h3>看点</h3><ul class="find" style="font-size:13px">' + "".join(f"<li>{p}</li>" for p in points) + "</ul></div></div>")

    # 教师排名
    if R.teachers:
        blocks = ""
        close = []
        for s in S:
            T = sorted([t for t in R.teachers if t["学科"] == s], key=lambda t: t["名次"])
            if not T:
                continue
            for a, b in zip(T, T[1:]):
                if a["得分"] - b["得分"] < 1:
                    close.append(f"{s}第{a['名次']}、{b['名次']}名仅差 {a['得分'] - b['得分']:.2f} 分")
            rr = "".join(
                f'<tr{" class=hl" if t["名次"] == 1 else ""}><td><span class="rk{" rk1" if t["名次"] == 1 else ""}">{t["名次"]}</span></td><td class="b">{t["教师"]}</td>'
                f'<td>{t["班级"]}</td><td>{t["人数"]}</td><td>{f2(t["平均分"])}</td><td>{pct(t["及格率"])}</td><td>{pct(t["优秀率"])}</td><td class="b">{f2(t["得分"])}</td></tr>' for t in T)
            blocks += (f'<div class="tb"><h4>{s}（满分{cfg.full[s]:g}）</h4><table><tr><th>名次</th><th>教师</th><th>任教班级</th><th>人数</th>'
                       f'<th>均分</th><th>及格率</th><th>优秀率</th><th>综合得分</th></tr>{rr}</table></div>')
        blocks += ('<div class="tb"><h4>说明</h4><p class="note">教师成绩按所教班级的学生合并计算，算法与班级相同，只在同一学科内比较名次。'
                   + ("<br>" + "；".join(close) + "，成绩相当。" if close else "") + "</p></div>")
        page("教师成绩排名", f'<div class="grid2">{blocks}</div>')

    # 临界生
    near = (R.near.assign(_k=R.near["班级"].str[1:].astype(int))
            .sort_values(["_k", "总分"], ascending=[True, False]).drop(columns="_k"))
    dn = near[near["位置"] == "线下"]
    up = near[near["位置"] == "线上"]
    abbr = {s: s[0] for s in S}

    def ntbl(df):
        h = '<tr><th>班级</th><th>姓名</th>' + "".join(f"<th>{abbr[s]}</th>" for s in S) + '<th>总分</th><th>距线</th><th>薄弱</th></tr>'
        for _, r in df.iterrows():
            h += (f'<tr><td>{r["班级"]}</td><td class="b">{r["姓名"]}</td>' + "".join(f"<td>{r[s]:g}</td>" for s in S) +
                  f'<td class="b">{r["总分"]:g}</td><td class="{"up" if r["距线"] >= 0 else "dn"}">{"+" if r["距线"] > 0 else ""}{r["距线"]}</td><td class="weak">{r["薄弱学科"]}</td></tr>')
        return f'<table style="font-size:11.5px">{h}</table>'.replace("<td>", '<td style="padding:3px">')
    wc = Counter(R.near["薄弱学科"]).most_common()
    legend = "、".join(f"{abbr[s]}={s}" for s in S)
    PER = 18
    chunks = max((len(dn) + PER - 1) // PER, (len(up) + PER - 1) // PER, 1)
    for i in range(chunks):
        d_, u_ = dn.iloc[i * PER:(i + 1) * PER], up.iloc[i * PER:(i + 1) * PER]
        tip = ""
        if i == chunks - 1:
            tip = ('<h3 style="margin-top:12px">帮扶建议</h3><ul class="find" style="font-size:12.5px;line-height:1.6">'
                   f'<li>按“薄弱”学科安排任课教师个别辅导' + (f"，{wc[0][0]}是重点（{wc[0][1]}人）" if wc else "") + '。</li>'
                   '<li>线下临界生冲线，线上临界生防下滑。</li><li>下次考试对照本名单检查变化。</li></ul>')
        page("临界生名单" + (f"（{i + 1}）" if chunks > 1 else ""),
             f'<div class="row" style="gap:18px"><div style="flex:1"><h3>线下临界生（{len(dn)}人，冲线重点）</h3>{ntbl(d_) if len(d_) else "<p class=note>无</p>"}</div>'
             f'<div style="flex:1"><h3>线上临界生（{len(up)}人，需稳住）</h3>{ntbl(u_) if len(u_) else "<p class=note>无</p>"}{tip}</div></div>'
             f'<p class="note" style="position:absolute;bottom:40px;left:44px;right:44px">临界生：升学参考线（全级前{pct(cfg.promote_ratio, 0)}，{NS}科总分 {R.cut:g} 分）上下 {cfg.near_range:g} 分以内的学生，按班级排列。'
             f'表头：{legend}。“薄弱”按得分率比较。本名单仅供教师使用。</p>')

    if compare:
        for title, body in _compare_pages(compare):
            page(title, body)

    html = "".join(p.replace("{PNO}", str(i + 1)) for i, p in enumerate(pages))
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>{html}</body></html>'


def _structure_body(R: Result) -> str:
    cfg, CL, W, X = R.cfg, R.classes, R.cfg.structure, R.structure
    items = [k for k in RATE_ITEMS if k in W]
    rate_lab = {"平均成绩": "得分率", "全科合格率": "比率", "全科优秀率": "比率", "参考率": "比率", "进线率": "比率"}
    small = '<span style="color:#8a9096;font-size:11px;font-weight:400">'
    hdr = "<tr><th>名次</th><th>班级</th>" + "".join(
        f'<th>{k}<br>{small}{W[k]:g}分</span></th>' for k in items)
    if "增值评价" in W:
        hdr += f'<th>小计</th><th>增值评价<br>{small}{W["增值评价"]:g}分</span></th>'
    if "前10名每人加分" in W:
        hdr += f'<th>前10名加分<br>{small}每人{W["前10名每人加分"]:g}分</span></th>'
    hdr += "<th>结构分</th></tr>"
    body = ""
    for c in sorted(CL, key=lambda c: X[c]["名次"]):
        x = X[c]
        first = x["名次"] == 1
        body += (f'<tr{" class=hl" if first else ""}><td><span class="rk{" rk1" if first else ""}">{x["名次"]}</span></td>'
                 f'<td class="b">{c}</td>')
        for k in items:
            vals = [X[y][k + "分"] for y in CL]
            body += f'<td{_mark(vals, x[k + "分"])}>{f2(x[k + "分"])}<br>{small}{pct(x[k])}</span></td>'
        if "增值评价" in W:
            prev = f"，上次第{x['上次名次']}" if x["上次名次"] else ""
            body += (f'<td>{f2(x["小计"])}</td>'
                     f'<td>{f2(x["增值评价分"])}<br>{small}两率一分第{x["两率一分名次"]}{prev}</span></td>')
        if "前10名每人加分" in W:
            body += f'<td>{f2(x["加分"])}<br>{small}{x["前10名人数"]}人</span></td>'
        body += f'<td class="b" style="font-size:15px">{f2(x["结构分"])}</td></tr>'
    chart = hbars([(c, X[c]["结构分"]) for c in CL], w=470, color="#1c5cab",
                  vmax=max(X[c]["结构分"] for c in CL) * 1.15, fmt=f2, rowh=30)
    rule = [f"各项直接按比例折分（比率 × 分值）；表中小字为比率。平均成绩 = 本班总分平均分 ÷ 总分满分（{cfg.total_full:g}）。",
            "全科合格 / 全科优秀：每一科都达到合格线 / 优秀线的学生 ÷ 本班实考人数。"]
    if "参考率" in W:
        rule.append("参考率 = 实考人数 ÷ 应考人数（实考人数 = 各科成绩齐全、计入统计的学生）。"
                    + ("本次没有填应考人数的班，按实考人数计算（参考率 100%）。" if any(X[c].get("应考按实考") for c in CL) else ""))
    if "进线率" in W:
        rule.append(f"进线率 = 本班升学线（全级前{pct(cfg.promote_ratio, 0)}，{R.cut:g}分）以上人数 ÷ 实考人数。")
    if "增值评价" in W:
        step, mv = W.get("增值名次差", 0.1), W.get("增值进退步", 0.05)
        rule.append(f"增值评价：按两率一分（平均成绩、全科合格率、全科优秀率三项得分之和）排名，第1名 {W['增值评价']:g} 分，每低一名减 {step:g} 分；"
                    f"与上次的两率一分名次比，每进步一名加 {mv:g} 分、退步一名减 {mv:g} 分，最高 {W['增值评价']:g} 分。"
                    + ("" if cfg.prev_rank else "本次未输入上次名次，只按本次名次给分。"))
    if "前10名每人加分" in W:
        rule.append(f"前10名加分：全级总分前10名（同分计入）每有一人，本班加 {W['前10名每人加分']:g} 分。")
    return (f'<h3>班级结构分（{cfg.scheme}）</h3><table style="font-size:13px">{hdr}{body}</table>'
            f'<p class="note" style="margin:4px 0 12px">每列蓝色为最高、红色为最低。</p>'
            f'<div class="row"><div style="flex:1"><h3>结构分对比</h3>{chart}</div>'
            f'<div style="flex:1.1"><h3>计算方法</h3><ul class="find" style="font-size:12.5px;line-height:1.6">'
            + "".join(f"<li>{t}</li>" for t in rule) + "</ul></div></div>")


def find_browser() -> str | None:
    """找电脑上已经装好的 Edge / Chrome（Windows 自带 Edge）。"""
    import os
    import shutil
    cands = []
    if os.name == "nt":
        for base in (os.environ.get("PROGRAMFILES(X86)"), os.environ.get("PROGRAMFILES"), os.environ.get("LOCALAPPDATA")):
            if base:
                cands += [Path(base) / "Microsoft/Edge/Application/msedge.exe",
                          Path(base) / "Google/Chrome/Application/chrome.exe"]
    elif sys.platform == "darwin":
        cands += [Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
                  Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
                  Path("/Applications/Chromium.app/Contents/MacOS/Chromium")]
    for c in cands:
        if c.is_file():
            return str(c)
    for name in ("msedge", "microsoft-edge", "google-chrome", "chrome", "chromium", "chromium-browser"):
        w = shutil.which(name)
        if w:
            return w
    return None


def _pdf_by_browser(browser: str, html_file: Path, pdf_path: Path, timeout: float = 90) -> bool:
    """让浏览器在后台把网页打印成 PDF。浏览器有时打印完不自己退出，所以等文件写好就把它关掉。"""
    import subprocess
    import tempfile
    import time
    pdf_path.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as prof:      # 用临时资料夹，不碰用户自己的浏览器
        proc = subprocess.Popen(
            [browser, "--headless", "--disable-gpu", "--no-first-run", "--no-default-browser-check", "--disable-extensions",
             "--no-pdf-header-footer", "--print-to-pdf-no-header", f"--user-data-dir={prof}",
             f"--print-to-pdf={pdf_path}", html_file.resolve().as_uri()],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        t0, last, stable = time.time(), -1, 0
        while time.time() - t0 < timeout:
            if proc.poll() is not None:
                break
            size = pdf_path.stat().st_size if pdf_path.exists() else 0
            stable = stable + 1 if (size > 0 and size == last) else 0
            if stable >= 4:                                   # 文件大小连续 1 秒不变 = 写完了
                break
            last = size
            time.sleep(0.25)
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(5)
            except subprocess.TimeoutExpired:
                proc.kill()
    return pdf_path.exists() and pdf_path.stat().st_size > 1000


def html_to_pdf(html: str, pdf_path: Path):
    """把发布版网页打印成 PDF：先用电脑上已有的 Edge / Chrome；没有的话再试 Playwright（如果装了）。"""
    pdf_path = Path(pdf_path)
    tmp = pdf_path.with_suffix(".html")
    tmp.write_text(html, encoding="utf-8")
    try:
        browser = find_browser()
        if browser and _pdf_by_browser(browser, tmp, pdf_path):
            return pdf_path
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise RuntimeError("没有找到 Edge 或 Chrome 浏览器，这次没法生成 PDF（Excel 不受影响）。"
                               "装一个 Edge 或 Chrome 后再算一次即可。") from None
        with sync_playwright() as p:
            b = None
            for kw in ({}, {"channel": "msedge"}, {"channel": "chrome"}):
                try:
                    b = p.chromium.launch(**kw)
                    break
                except Exception:
                    continue
            if b is None:
                raise RuntimeError("没有找到可用的浏览器，这次没法生成 PDF（Excel 不受影响）。")
            page = b.new_page()
            page.goto(tmp.resolve().as_uri())
            page.wait_for_timeout(300)
            page.pdf(path=str(pdf_path), width="297mm", height="210mm", print_background=True)
            b.close()
        return pdf_path
    finally:
        tmp.unlink(missing_ok=True)


def _move(a, b):
    """名次从 a 变到 b：↑ 进步、↓ 退步。"""
    if a == b:
        return '<span class="note">持平</span>'
    return f'<span class="up">↑{a - b}</span>' if b < a else f'<span class="dn">↓{b - a}</span>'


def _delta(v, d=2, unit=""):
    if abs(v) < 0.5 * 10 ** -d:
        return '<span class="note">持平</span>'
    return f'<span class="{"up" if v > 0 else "dn"}">{"+" if v > 0 else "−"}{abs(v):.{d}f}{unit}</span>'


def _compare_pages(C: dict) -> list[tuple[str, str]]:
    """和上次比：两页。一页各班、各科；一页学生名次进退和上次的临界生。"""
    pv, nw = C["prev"], C["now"]
    small = '<span style="color:#8a9096;font-size:11px">'
    who = (f'<p class="note" style="margin-bottom:10px">上次：{pv["exam"]}{"（" + pv["date"] + "）" if pv["date"] else ""}，{pv["scheme"]}方案，'
           f'实考 {pv["n"]} 人，升学线 {pv["cut"]:g} 分；本次：{nw["scheme"]}方案，实考 {nw["n"]} 人，升学线 {nw["cut"]:g} 分。箭头 ↑ 为进步、↓ 为退步。</p>')
    hs = C["hasStructure"]
    hdr = "<tr><th>班级</th>" + ("<th>结构分</th><th>结构分名次</th>" if hs else "") + "<th>综合得分名次</th><th>上线率</th></tr>"
    body = ""
    for x in C["classes"]:
        if not x["found"]:
            body += f'<tr><td class="b">{x["name"]}</td><td colspan="{4 if hs else 2}" class="note">上次没有这个班</td></tr>'
            continue
        body += f'<tr><td class="b">{x["name"]}</td>'
        if hs:
            body += (f'<td>{small}{f2(x["prevStruct"])} →</span> <b>{f2(x["struct"])}</b> {_delta(x["struct"] - x["prevStruct"])}</td>'
                     f'<td>{small}{x["prevStructRank"]} →</span> <b>{x["structRank"]}</b> {_move(x["prevStructRank"], x["structRank"])}</td>')
        body += (f'<td>{small}{x["prevRank"]} →</span> <b>{x["rank"]}</b> {_move(x["prevRank"], x["rank"])}</td>'
                 f'<td>{small}{pct(x["prevOnlineRate"])} →</span> <b>{pct(x["onlineRate"])}</b></td></tr>')
    cls_tbl = f'<h3>各班</h3><table>{hdr}{body}</table>'
    sh = "<tr><th>科目</th><th>平均得分率</th><th>及格率</th><th>优秀率</th></tr>"
    cell = lambda a, b: f'<td>{small}{pct(a)} →</span> <b>{pct(b)}</b> {_delta((b - a) * 100, 1, "点")}</td>'
    sb = "".join(f'<tr><td class="b">{x["subject"]}</td>{cell(x["prevAvgRate"], x["avgRate"])}{cell(x["prevPass"], x["pass"])}{cell(x["prevExc"], x["exc"])}</tr>'
                 for x in C["subjects"]) or '<tr><td colspan="4" class="note">两次考试没有相同的科目</td></tr>'
    sub_tbl = (f'<h3>各科（全年级）</h3><table>{sh}{sb}</table>'
               '<p class="note" style="margin-top:6px">两次满分可能不同，所以比得分率（平均分 ÷ 满分）。“点”是百分点。</p>')
    page1 = who + f'<div class="row"><div style="flex:1.15">{cls_tbl}</div><div style="flex:1">{sub_tbl}</div></div>'

    S = C["students"]
    ups = [x for x in S if x["change"] > 0][:12]                          # 每边 12 人，给下面的临界生留出位置
    dns = sorted([x for x in S if x["change"] < 0], key=lambda x: (x["change"], x["rank"]))[:12]

    def stbl(L):
        if not L:
            return '<p class="note">无</p>'
        h = "<tr><th>班级</th><th>姓名</th><th>上次级名次</th><th>本次级名次</th><th>进退</th></tr>"
        h += "".join(f'<tr><td>{x["cls"]}</td><td class="b">{x["name"]}</td><td>{x["prevRank"]}</td><td class="b">{x["rank"]}</td>'
                     f'<td>{_move(x["prevRank"], x["rank"])}</td></tr>' for x in L)
        return f'<table style="font-size:12px">{h}</table>'.replace("<td>", '<td style="padding:3px">')
    a, b = C["nearUp"], C["nearDown"]
    moved = lambda goal: "、".join(f'{x["cls"]}{x["name"]}（{x["prevTotal"]:g}→{x["total"]:g}）' for x in C["near"] if x["where"] == goal) or "无"
    near = (f'<h3 style="margin-top:14px">上次的临界生</h3><ul class="find" style="font-size:12.5px;line-height:1.6">'
            f'<li>上次<b>线下</b>临界生 {a["all"]} 人，这次能对上 {a["found"]} 人，其中 <b class="up">{a["moved"]}</b> 人上线了：{moved("上线了")}。</li>'
            f'<li>上次<b>线上</b>临界生 {b["all"]} 人，这次能对上 {b["found"]} 人，其中 <b class="dn">{b["moved"]}</b> 人掉到了线下：{moved("掉到线下")}。</li></ul>')
    page2 = (f'<div class="row" style="gap:18px"><div style="flex:1"><h3>进步最多的学生</h3>{stbl(ups)}</div>'
             f'<div style="flex:1"><h3>退步最多的学生</h3>{stbl(dns)}</div></div>{near}'
             f'<p class="note" style="position:absolute;bottom:40px;left:44px;right:44px">学生按“班级 + 姓名”对应：两次都参加并且能对上的共 {len(S)} 人；'
             f'换了班、同班重名、缺考的不在比较之内。名单含学生姓名，仅限教师使用。</p>')
    return [("和上次考试比较（一）：各班与各科", page1), ("和上次考试比较（二）：学生名次与临界生", page2)]
