"""界面的端到端测试：用自动化浏览器像真人一样操作（导入 → 设置 → 计算 → 查看 → 导出 → 试用到期 → 激活）。
需要 playwright（没装就跳过）。只用 samples/ 里的虚构数据，在临时文件夹里进行，不碰真实的 data/ 和 output/。"""
import datetime as dt
import shutil
import sys
from pathlib import Path

import pytest

pw = pytest.importorskip("playwright.sync_api")
ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture()
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("CHENGJI_NO_OPEN", "1")
    for d in ("config", "templates", "samples"):
        shutil.copytree(ROOT / d, tmp_path / d)
    (tmp_path / "data").mkdir()
    shutil.copy2(ROOT / "samples" / "示例任课总表.xlsx", tmp_path / "data" / "任课总表.xlsx")
    from chengji.server import App, start
    httpd, url, a = start(App(tmp_path))
    yield tmp_path, url
    httpd.shutdown()


@pytest.fixture()
def page():
    with pw.sync_playwright() as p:
        try:
            b = p.chromium.launch()
        except Exception as e:                      # 没装浏览器内核
            pytest.skip(f"没有可用的浏览器内核：{e}")
        pg = b.new_page(viewport={"width": 1320, "height": 880})
        errs = []
        pg.on("pageerror", lambda e: errs.append(str(e)))
        yield pg
        b.close()
        assert not errs, errs


def idle(pg):
    pg.wait_for_function("!document.querySelector('#busy').classList.contains('on')", timeout=120000)


def import_and_compute(pg, url, sample, exam, fill=None):
    pg.goto(url)
    pg.wait_for_selector("#recent .li")
    pg.set_input_files("#file", str(ROOT / "samples" / sample))
    pg.wait_for_selector("#wizBody:not(.hide)")
    pg.fill("#exam", exam)
    if fill:
        fill(pg)
    pg.click("#run")


def test_full_flow(app, page):
    root, url = app
    def fill(pg):
        for c, v in zip(["九1", "九2", "九3", "九4", "九5", "九6"], [3, 6, 1, 5, 4, 2]):
            pg.fill(f'#prev input[data-c="{c}"]', str(v))
    import_and_compute(page, url, "示例登分表_九年级.xlsx", "界面测试", fill)
    page.wait_for_selector("#resBody:not(.hide)")
    assert page.inner_text("#rTitle") == "界面测试"
    first = page.inner_text("#ovRank .rk").split()
    assert first[:3] == ["1", "九6", "46.45"]                      # 与命令行、Excel 公式核对过的结果一致
    assert "进步 1 名" in page.inner_text("#ovRank .rk")
    page.click('#tabs button[data-t="de"]')
    page.click('#fc button:has-text("九2")')
    assert page.inner_text("#cnt") == "共 45 人"
    assert page.inner_text("#deT tbody tr td") == "1"              # 筛选后序号从 1 排起
    page.click('#tabs button[data-t="te"]')
    assert page.locator("#teG .tcard").count() == 7
    page.evaluate("go('export')")
    page.click("#xlsBtn"); idle(page)
    out = root / "output" / "界面测试"
    assert (out / "界面测试_各班综合统计.xlsx").stat().st_size > 20000
    assert (out / "结果.json").is_file() and (out / "登分表副本.xlsx").is_file()
    # 重新打开程序后，能从“最近的考试”里回看并导出
    page.goto(url); page.wait_for_selector('#recent .li[data-n="界面测试"]')
    page.click('#recent .li[data-n="界面测试"]'); page.wait_for_selector("#resBody:not(.hide)")
    assert page.inner_text("#ovRank .rk").split()[:3] == ["1", "九6", "46.45"]


def test_zhongkao_needs_full_marks(app, page):
    """中考 10 科：生物、地理、体育没定满分 → 先拦住并说明；填上后算出进线率、前 10 名加分。"""
    root, url = app
    def fill(pg):
        pg.click('#scheme button:has-text("中考核算")')
    import_and_compute(page, url, "示例登分表_中考.xlsx", "中考测试", fill)
    page.wait_for_selector("#wizErr:not(.hide)")
    assert "生物、地理、体育" in page.inner_text("#wizErr") and "满分" in page.inner_text("#wizErr")
    for s in ("生物", "地理", "体育"):
        page.fill(f'#fullChips input[data-s="{s}"]', "50")
    page.click("#run"); page.wait_for_selector("#resBody:not(.hide)")
    page.click('#tabs button[data-t="st"]')
    head = page.inner_text("#stT thead")
    assert "进线率" in head and "前 10 名加分" in head and "增值评价" not in head


def test_settings_saved_locally(app, page):
    root, url = app
    page.goto(url); page.wait_for_selector("#recent .li")
    page.evaluate("go('settings')"); page.wait_for_selector("#setChips .chip")
    page.fill("#sSchool", "测试中学"); page.fill('#setChips input[data-s="生物"]', "50")
    page.click("#saveBtn"); page.wait_for_selector("#saved.on")
    text = (root / "data" / "本校设置.yaml").read_text(encoding="utf-8")
    assert "测试中学" in text and "生物: 50" in text
    assert "××中学" in (root / "config" / "学校设置.yaml").read_text(encoding="utf-8")   # 公共设置文件不被改动


def test_trial_expired_then_activate(app, page, tmp_path, monkeypatch):
    """试用到期：能算能看、不能导出；输入本机的激活码后可以导出。（需要本机有许可模块和生成器）"""
    impl = pytest.importorskip("chengji._license_impl")
    if not (ROOT / "admin" / "keygen.py").is_file():
        pytest.skip("没有激活码生成器")
    sys.path.insert(0, str(ROOT / "admin"))
    import keygen
    monkeypatch.setenv("CHENGJI_LICENSE_ENFORCE", "1")
    monkeypatch.setenv("CHENGJI_LICENSE_DIR", str(tmp_path / "lic"))
    monkeypatch.setenv("CHENGJI_MACHINE_RAW", "ui-test-machine")
    monkeypatch.setattr(impl, "_MACHINE", None)
    monkeypatch.setattr(keygen, "LOG_FILE", tmp_path / "记录.csv")
    impl.status()                                                   # 开始试用
    for d in impl._dirs():                                          # 把“第一次使用”挪到 4 天前
        st = impl._unseal((d / "state.json").read_text()); st["first"] -= 4 * 86400
        (d / "state.json").write_text(impl._seal(st))
    root, url = app
    import_and_compute(page, url, "示例登分表_九年级.xlsx", "到期测试")
    page.wait_for_selector("#resBody:not(.hide)")                   # 到期后仍然能算、能看
    assert page.inner_text("#licText") == "试用已结束"
    page.evaluate("go('export')")
    assert page.is_visible("#export .lock")                         # 导出被锁住
    assert not (root / "output" / "到期测试" / "到期测试_各班综合统计.xlsx").exists()
    page.click("#export .lock button"); page.wait_for_selector("#mask.on")
    assert page.inner_text("#machine") == impl.machine_code()
    page.fill("#code", "AAAAA-BBBBB"); page.click("#actBtn")
    page.wait_for_function("document.querySelector('#actErr').textContent.length > 0")
    code = keygen.make_code(impl.machine_code(), dt.date.today() + dt.timedelta(days=365))
    page.fill("#code", code); page.click("#actBtn")
    page.wait_for_function("!document.querySelector('#mask').classList.contains('on')")
    assert page.inner_text("#licText") == "已激活" and not page.is_visible("#export .lock")
    page.click("#xlsBtn"); idle(page)
    assert (root / "output" / "到期测试" / "到期测试_各班综合统计.xlsx").is_file()
