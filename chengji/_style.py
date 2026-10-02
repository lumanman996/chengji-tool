"""发布版的样式与小图表（SVG）。"""


def pct(x, d=1):
    return f"{x * 100:.{d}f}%"


def f1(x):
    return f"{x:.1f}"


def f2(x):
    return f"{x:.2f}"


CSS = '''
@page { size: 297mm 210mm; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: "Microsoft YaHei", "微软雅黑", "PingFang SC", "Noto Sans CJK SC", sans-serif; color: #1f2328; background: #fff; }
.page { width: 1123px; height: 794px; padding: 34px 44px 30px; position: relative; overflow: hidden; page-break-after: always; background: #fff; }
.hd { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 3px solid #1c5cab; padding-bottom: 10px; margin-bottom: 18px; }
.hd .t { font-size: 26px; font-weight: 700; color: #104281; }
.hd .s { font-size: 13px; color: #5b6168; text-align: right; line-height: 1.5; }
.tag { display: inline-block; background: #c62828; color: #fff; font-size: 12px; padding: 2px 8px; border-radius: 3px; margin-left: 10px; vertical-align: middle; font-weight: 500; }
.ft { position: absolute; bottom: 14px; left: 44px; right: 44px; font-size: 11px; color: #8a9096; display: flex; justify-content: space-between; border-top: 1px solid #e3e6ea; padding-top: 6px; }
h3 { font-size: 16px; color: #104281; margin: 4px 0 8px; padding-left: 9px; border-left: 4px solid #1c5cab; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th { background: #e8f0fb; color: #104281; font-weight: 600; padding: 6px 4px; border: 1px solid #c9d6e8; }
td { padding: 5px 4px; border: 1px solid #dde3ea; text-align: center; }
tr.hl td { background: #fff6e0; }
tr.tot td { background: #f1f3f5; font-weight: 600; }
.b { font-weight: 700; }
.note { font-size: 12px; color: #5b6168; line-height: 1.6; }
.tiles { display: flex; gap: 14px; margin-bottom: 18px; }
.tile { flex: 1; border: 1px solid #dde3ea; border-radius: 8px; padding: 12px 16px; }
.tile .k { font-size: 13px; color: #5b6168; }
.tile .v { font-size: 34px; font-weight: 700; color: #104281; line-height: 1.2; }
.tile .v small { font-size: 15px; font-weight: 500; color: #5b6168; margin-left: 3px; }
.tile .d { font-size: 12px; color: #8a9096; }
.find { font-size: 14px; line-height: 1.75; }
.find li { margin-left: 20px; margin-bottom: 4px; }
.row { display: flex; gap: 22px; }
.legend { font-size: 12px; color: #5b6168; display: flex; gap: 14px; margin-bottom: 4px; }
.legend i { display: inline-block; width: 11px; height: 11px; border-radius: 2px; margin-right: 4px; vertical-align: -1px; }
.rk { display: inline-block; width: 22px; height: 22px; line-height: 22px; border-radius: 50%; background: #e3e6ea; font-weight: 700; font-size: 12px; }
.rk1 { background: #1c5cab; color: #fff; }
.card { flex: 1; border: 1px solid #dde3ea; border-radius: 10px; padding: 16px 18px; position: relative; }
.card.first { border: 2px solid #1c5cab; background: #f5f9ff; }
.card .no { position: absolute; top: 12px; right: 16px; font-size: 40px; font-weight: 800; color: #c9d6e8; }
.card.first .no { color: #1c5cab; }
.card .nm { font-size: 22px; font-weight: 700; }
.card .cl { font-size: 13px; color: #5b6168; margin-bottom: 10px; }
.card .sc { font-size: 30px; font-weight: 700; color: #104281; }
.card .sc small { font-size: 13px; color: #5b6168; font-weight: 400; }
.card .m { display: flex; justify-content: space-between; font-size: 13px; margin-top: 8px; border-top: 1px dashed #dde3ea; padding-top: 8px; }
.card .m div { text-align: center; } .card .m b { display: block; font-size: 16px; }
.up { color: #1c5cab; font-weight: 600; } .dn { color: #c62828; font-weight: 600; }
.weak { color: #c62828; font-weight: 600; }
'''


CSS += '''
.best { color:#104281; font-weight:700; background:#e8f0fb; } .worst { color:#c62828; font-weight:700; background:#fdecec; }
.grid2 { display:grid; grid-template-columns:1fr 1fr; gap:10px 26px; }
.tb h4 { font-size:14px; margin-bottom:4px; padding-left:8px; border-left:4px solid #1c5cab; color:#104281; }
.tb table { font-size:12px; } .tb td,.tb th { padding:4px 3px; }
'''


def grouped_bars(data, series, colors, w=1030, h=250, vmax=None, fmt=f1):
    # data: list of (label, [v1,v2,v3])
    vmax = vmax or max(max(v) for _, v in data) * 1.15
    L, R, T, B = 40, 10, 16, 28
    pw, ph = w - L - R, h - T - B
    gw = pw / len(data); bw = min(34, (gw - 30) / len(series))
    s = [f'<svg width="{w}" height="{h}" xmlns="http://www.w3.org/2000/svg" font-family="Microsoft YaHei, Noto Sans CJK SC, sans-serif" font-size="11">']
    for t in range(0, 6):
        v = vmax * t / 5; y = T + ph - ph * t / 5
        s.append(f'<line x1="{L}" x2="{w-R}" y1="{y:.1f}" y2="{y:.1f}" stroke="#eceef1"/><text x="{L-6}" y="{y+4:.1f}" text-anchor="end" fill="#8a9096">{v:.0f}</text>')
    for i, (lab, vals) in enumerate(data):
        x0 = L + gw * i + (gw - bw * len(series) - 2 * (len(series) - 1)) / 2
        for j, v in enumerate(vals):
            bh = ph * v / vmax; x = x0 + j * (bw + 2); y = T + ph - bh
            s.append(f'<path d="M{x:.1f},{T+ph} V{y+4:.1f} Q{x:.1f},{y:.1f} {x+4:.1f},{y:.1f} H{x+bw-4:.1f} Q{x+bw:.1f},{y:.1f} {x+bw:.1f},{y+4:.1f} V{T+ph} Z" fill="{colors[j]}"/>')
            s.append(f'<text x="{x+bw/2:.1f}" y="{y-4:.1f}" text-anchor="middle" fill="#3d4349" font-size="10.5">{fmt(v)}</text>')
        s.append(f'<text x="{L+gw*i+gw/2:.1f}" y="{h-8}" text-anchor="middle" fill="#1f2328" font-size="13" font-weight="600">{lab}</text>')
    s.append(f'<line x1="{L}" x2="{w-R}" y1="{T+ph}" y2="{T+ph}" stroke="#8a9096"/></svg>')
    return ''.join(s)

def hbars(data, w=480, color='#2a78d6', vmax=None, fmt=f1, ref=None, reflab='', hl=None, rowh=34):
    vmax = vmax or max(v for _, v in data) * 1.12
    L, R = 60, 60; pw = w - L - R; h = rowh * len(data) + 26
    s = [f'<svg width="{w}" height="{h}" xmlns="http://www.w3.org/2000/svg" font-family="Microsoft YaHei, Noto Sans CJK SC, sans-serif" font-size="12">']
    for i, (lab, v) in enumerate(data):
        y = 8 + i * rowh; bw = pw * v / vmax; bh = rowh - 12
        c = color if (hl is None or lab == hl) else '#86b6ef'
        s.append(f'<text x="{L-8}" y="{y+bh/2+4:.1f}" text-anchor="end" fill="#1f2328" font-size="13" font-weight="600">{lab}</text>')
        s.append(f'<path d="M{L},{y} H{L+bw-4:.1f} Q{L+bw:.1f},{y} {L+bw:.1f},{y+4} V{y+bh-4} Q{L+bw:.1f},{y+bh} {L+bw-4:.1f},{y+bh} H{L} Z" fill="{c}"/>')
        if ref is not None: s.append(f'<text x="{L+bw-8:.1f}" y="{y+bh/2+4:.1f}" text-anchor="end" fill="#ffffff" font-weight="600">{fmt(v)}</text>')
        else: s.append(f'<text x="{L+bw+6:.1f}" y="{y+bh/2+4:.1f}" fill="#3d4349" font-weight="600">{fmt(v)}</text>')
    if ref is not None:
        x = L + pw * ref / vmax
        s.insert(1, f'<line x1="{x:.1f}" x2="{x:.1f}" y1="2" y2="{h-18}" stroke="#c62828" stroke-dasharray="4 3" stroke-width="1.5"/>'); s.append(f'<text x="{x:.1f}" y="{h-4}" text-anchor="middle" fill="#c62828" font-size="11">{reflab}</text>')
    s.append('</svg>')
    return ''.join(s)

