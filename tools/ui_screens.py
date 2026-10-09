"""界面截图：把每一页、每个标签在几种常见的窗口大小、明暗主题、字号下各截一张图，和上一次比哪些变了。

用来防止“修好这里，那里歪了”：改界面之前截一次，改完再截一次，打开生成的「对比.html」逐张看变了的图。
同时检查有没有东西被挤出窗口（出现横向滚动条、按钮跑到窗口外面），有就列出来。

只用 samples/ 里的虚构数据，在临时文件夹里运行，不碰真实的 data/ 和 output/ 里的考试。
截图存在 output/界面截图/<时间>/（output 不上传）。需要 playwright 和它的浏览器内核。

用法：
  python tools/ui_screens.py            全套：各种窗口大小 × 明暗 × 字号
  python tools/ui_screens.py --快速      只截最常见的一种（1366×768 屏幕、明亮、100%）
"""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
OUT = ROOT / "output" / "界面截图"

# 窗口里能放网页的部分（宽、高、系统缩放）：去掉了任务栏和窗口标题栏
SIZES = {
    "1366x768": (1366, 690, 1),          # 学校里最常见的笔记本、旧显示器
    "1366x768-125%": (1093, 545, 1.25),   # 同一块屏幕，Windows 缩放 125%（很多笔记本默认）
    "1920x1080-150%": (1280, 650, 1.5),  # 1080p 屏幕，Windows 缩放 150%
    "1920x1080": (1440, 900, 1),
}
FULL = [("1366x768", "light", 100), ("1366x768", "dark", 100), ("1366x768-125%", "light", 100),
        ("1366x768-125%", "light", 150), ("1366x768-125%", "dark", 130), ("1920x1080-150%", "light", 100),
        ("1920x1080", "light", 100), ("1920x1080", "dark", 120)]
QUICK = [("1366x768", "light", 100)]

# 检查“有没有东西跑出窗口”：本来就该横向滚动的地方（表格外框、标签栏）不算
OVERFLOW_JS = """() => {
  const W = document.documentElement.clientWidth, bad = [];
  const inScroll = e => e.closest('.scroll, .tabbar, .tt, #ted .body, .updNotes, .fchips, .drawer table');
  const main = document.querySelector('.main');
  if (main && main.scrollWidth > main.clientWidth + 1) bad.push('主区出现横向滚动条（多出 ' + (main.scrollWidth - main.clientWidth) + 'px）');
  for (const e of document.querySelectorAll('body *')) {
    const s = getComputedStyle(e); if (s.display === 'none' || s.visibility === 'hidden' || !e.offsetParent && s.position !== 'fixed') continue;
    if (inScroll(e) || e.closest('.drawer:not(.on)')) continue;
    const r = e.getBoundingClientRect(); if (!r.width || !r.height) continue;
    if (r.right > W + 1 && r.left < W) bad.push((e.id ? '#' + e.id : e.tagName.toLowerCase() + (e.className && typeof e.className === 'string' ? '.' + e.className.split(' ')[0] : '')) + ' 超出右边 ' + Math.round(r.right - W) + 'px');
    if (bad.length > 8) break;
  }
  return [...new Set(bad)];
}"""


def _second_sample(tmp: Path) -> Path:
    """造“第二次考试”：九年级示例里九1 每科加几分、九6 减几分，好让“和上次比”有内容。"""
    import openpyxl
    wb = openpyxl.load_workbook(ROOT / "samples" / "示例登分表_九年级.xlsx"); ws = wb.active
    for r in range(3, ws.max_row + 1):
        cls = str(ws.cell(r, 2).value)
        for c in range(5, 12):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):           # 加分不超过 50（最小的满分），免得超过满分
                d = 6 if cls == "九1" and v + 6 <= 50 else -5 if cls == "九6" and v >= 5 else 0
                ws.cell(r, c).value = v + d
    p = tmp / "第二次月考.xlsx"; wb.save(p)
    return p


def _setup(tmp: Path):
    os.environ["CHENGJI_NO_OPEN"] = "1"
    (tmp / "下载").mkdir()
    os.environ["CHENGJI_DOWNLOADS"] = str(tmp / "下载")
    for d in ("config", "templates", "samples"):
        shutil.copytree(ROOT / d, tmp / d)
    (tmp / "data").mkdir()
    shutil.copy2(ROOT / "samples" / "示例任课总表.xlsx", tmp / "data" / "任课总表.xlsx")
    (tmp / "data" / "本校设置.yaml").write_text("已引导: true\n学校: 示例县第一初级中学\n", encoding="utf-8")
    from chengji.server import App, start
    return start(App(tmp))


def _idle(pg):
    pg.wait_for_function("!document.querySelector('#busy').classList.contains('on')", timeout=120000)


def _compute(pg, url, f: Path, exam: str):
    pg.goto(url); pg.wait_for_selector("#recent .li")
    pg.set_input_files("#file", str(f)); pg.wait_for_selector("#wizBody:not(.hide)")
    pg.fill("#exam", exam); pg.click("#run"); _idle(pg); pg.wait_for_selector("#resBody:not(.hide)")


def _look(pg, theme: str, zoom: int):
    """用程序自己的设置接口改主题、字号（新界面）；老界面没有这些接口就退回浏览器模拟。"""
    pg.emulate_media(color_scheme=theme)
    pg.evaluate("""async z => { try { await post('ui_save', {theme: '跟随系统', zoom: z}); } catch (e) {}
      try { await post('ui_save', {fontSize: z >= 125 ? '特大' : z > 100 ? '大' : '标准'}); } catch (e) {} }""", zoom)


def shoot(pg, url, size: str, theme: str, zoom: int, out: Path, problems: dict):
    w, h, dpr = SIZES[size]
    tag = f"{size}_{'明亮' if theme == 'light' else '暗色'}_{zoom}%"
    ctx_page = pg
    ctx_page.set_viewport_size({"width": w, "height": h})
    ctx_page.goto(url); ctx_page.wait_for_selector("#recent .li")
    _look(ctx_page, theme, zoom)
    ctx_page.goto(url); ctx_page.wait_for_selector("#recent .li"); ctx_page.wait_for_timeout(300)
    n = [0]

    def snap(name):
        ctx_page.wait_for_timeout(260)
        n[0] += 1
        fn = f"{tag}__{n[0]:02d}_{name}.png"
        ctx_page.screenshot(path=str(out / fn))
        bad = ctx_page.evaluate(OVERFLOW_JS)
        if bad:
            problems[fn] = bad

    def page(p):
        ctx_page.evaluate(f"go('{p}')"); ctx_page.evaluate("document.querySelector('.main').scrollTop = 0")

    snap("开始")
    ctx_page.evaluate("document.querySelector('.main').scrollTop = 99999"); snap("开始-下半")
    # 核算向导：导入示例登分表
    ctx_page.set_input_files("#file", str(ROOT / "samples" / "示例登分表_九年级.xlsx")); ctx_page.wait_for_selector("#wizBody:not(.hide)")
    _idle(ctx_page); snap("核算")
    ctx_page.evaluate("document.querySelector('.main').scrollTop = 99999"); snap("核算-下半")
    # 结果：打开最近的一次（有上一次可比）
    page("home"); ctx_page.wait_for_selector("#recent .li[data-n]")
    ctx_page.click("#recent .li[data-n]"); _idle(ctx_page); ctx_page.wait_for_selector("#resBody:not(.hide)")
    tabs = ctx_page.locator("#tabs [data-t]")
    for i in range(tabs.count()):
        t = tabs.nth(i); label = t.get_attribute("data-t"); name = t.inner_text().split("\n")[0].strip()
        t.click(); ctx_page.evaluate("document.querySelector('.main').scrollTop = 0")
        snap(f"结果-{name or label}")
    ctx_page.locator("#tabs [data-t]").first.click()
    ctx_page.locator("#ovRank .clk").first.click(); ctx_page.wait_for_selector("#drawer.on"); snap("班级详情")
    ctx_page.keyboard.press("Escape")
    page("export"); snap("导出")
    page("settings"); ctx_page.wait_for_selector("#setChips .chip")
    stabs = ctx_page.locator("#stabs [data-s]:not(.hide)")
    if stabs.count():
        for i in range(stabs.count()):
            t = stabs.nth(i); t.click(); ctx_page.evaluate("document.querySelector('.main').scrollTop = 0")
            snap("设置-" + t.inner_text().split("\n")[0].strip())
    else:
        snap("设置"); ctx_page.evaluate("document.querySelector('.main').scrollTop = 1400"); snap("设置-中段")
        ctx_page.evaluate("document.querySelector('.main').scrollTop = 99999"); snap("设置-下半")
    if stabs.count():
        ctx_page.click('#stabs [data-s="teacher"]')
    ctx_page.click("#tEdit"); ctx_page.wait_for_selector("#ted.on"); snap("任课总表编辑")
    ctx_page.click("#tedNo"); ctx_page.wait_for_timeout(200)
    page("help"); snap("帮助")
    ctx_page.evaluate("openAct()"); snap("激活窗口"); ctx_page.evaluate("closeAct()")


def compare(out: Path) -> list[tuple[str, float]]:
    """和上一次截图比：返回 [(文件名, 变化的像素占比)]，新出现的图占比记为 1。"""
    from PIL import Image, ImageChops
    runs = sorted(p for p in OUT.iterdir() if p.is_dir() and p != out and p.name < out.name)
    if not runs:
        return []
    prev, res = runs[-1], []
    for f in sorted(out.glob("*.png")):
        g = prev / f.name
        if not g.is_file():
            res.append((f.name, 1.0)); continue
        a, b = Image.open(f).convert("RGB"), Image.open(g).convert("RGB")
        if a.size != b.size:
            res.append((f.name, 1.0)); continue
        bbox = ImageChops.difference(a, b).point(lambda v: 255 if v > 24 else 0).convert("L")
        hist = bbox.histogram()
        res.append((f.name, hist[255] / (a.size[0] * a.size[1])))
    return res


def write_sheet(out: Path, problems: dict, diffs: list):
    """生成 对比.html：每张图一格，变了的放前面，左边上一次、右边这一次。"""
    prev = sorted(p for p in OUT.iterdir() if p.is_dir() and p.name < out.name)
    prev = prev[-1] if prev else None
    changed = {n: r for n, r in diffs if r > 0.0005}
    names = sorted((p.name for p in out.glob("*.png")), key=lambda n: (n not in changed, n))
    cells = []
    for n in names:
        tag = f"<b class='chg'>变了 {changed[n] * 100:.1f}%</b>" if n in changed else "<span>没变</span>" if prev else ""
        warn = "".join(f"<li>{p}</li>" for p in problems.get(n, []))
        old = f"<img src='../{prev.name}/{n}' loading=lazy>" if prev and n in changed and (prev / n).is_file() else ""
        cells.append(f"<section><h3>{n[:-4]} {tag}</h3>{f'<ul>{warn}</ul>' if warn else ''}<div class='p'>{old}<img src='{n}' loading=lazy></div></section>")
    html = f"""<!doctype html><meta charset=utf-8><title>界面截图 {out.name}</title>
<style>body{{font:14px -apple-system,"Microsoft YaHei",sans-serif;margin:24px;background:#F5F4F0;color:#15171C}}
section{{margin:0 0 28px}} h3{{font-size:14px;margin:0 0 6px}} .chg{{color:#B0452F}} span{{color:#7A7F87;font-weight:400}}
ul{{color:#B0452F;margin:4px 0 8px}} .p{{display:flex;gap:10px;align-items:flex-start}} img{{max-width:min(100%,900px);border:1px solid #ddd}}
.p img+img{{outline:2px solid #0E5A4B}}</style>
<h1>界面截图 · {out.name}</h1><p>共 {len(names)} 张；和上一次（{prev.name if prev else '无'}）比，变了 {len(changed)} 张；
发现 {len(problems)} 张有东西跑出窗口。变了的图左边是上一次、右边（绿框）是这一次。</p>{''.join(cells)}"""
    (out / "对比.html").write_text(html, encoding="utf-8")


def main():
    from playwright.sync_api import sync_playwright
    combos = QUICK if "--快速" in sys.argv else FULL
    out = OUT / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    out.mkdir(parents=True)
    tmp = Path(tempfile.mkdtemp(prefix="chengji_shots_"))
    httpd, url, _ = _setup(tmp)
    problems: dict = {}
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1366, "height": 690})
            _compute(pg, url, ROOT / "samples" / "示例登分表_九年级.xlsx", "九年级第一次月考")
            _compute(pg, url, _second_sample(tmp), "九年级第二次月考")
            pg.close()
            for size, theme, zoom in combos:
                w, h, dpr = SIZES[size]
                ctx = b.new_context(viewport={"width": w, "height": h}, device_scale_factor=dpr)
                pg = ctx.new_page()
                errs = []
                pg.on("pageerror", lambda e: errs.append(str(e)))
                shoot(pg, url, size, theme, zoom, out, problems)
                if errs:
                    problems[f"{size}_{theme}_{zoom}（脚本出错）"] = errs
                ctx.close()
                print(f"  截好：{size} {theme} {zoom}%")
            b.close()
    finally:
        httpd.shutdown()
        shutil.rmtree(tmp, ignore_errors=True)
    diffs = compare(out)
    write_sheet(out, problems, diffs)
    (out / "问题.json").write_text(json.dumps(problems, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len(list(out.glob("*.png")))
    print(f"\n截了 {n} 张，在：{out}")
    if diffs:
        print(f"和上一次比，变了 {sum(r > 0.0005 for _, r in diffs)} 张。")
    print(f"有东西跑出窗口的：{len(problems)} 张" + ("" if not problems else "（详见 问题.json）"))
    for k, v in list(problems.items())[:12]:
        print("  ", k, "：", "；".join(v[:3]))
    print("打开 对比.html 逐张看。")


if __name__ == "__main__":
    main()
