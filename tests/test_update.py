"""检查新版本与一键更新（不联网：用假的发布页数据和自己造的安装包）。
真正的“下载 → 替换 → 重新打开”在打包自检（packaging/smoke_test.py）里用打好的程序走一遍。"""
import io
import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

from chengji import updater

ROOT = Path(__file__).resolve().parent.parent
PREFIX = "https://github.com/lumanman996/chengji-tool/releases/download/"


@pytest.fixture(autouse=True)
def _log_to_tmp(tmp_path, monkeypatch):
    monkeypatch.setattr(updater, "LOG", tmp_path / "更新日志.txt")          # 不往项目文件夹里写日志


def _release(tag="v9.9.9", assets=None, body="## 更新内容\n\n- 新功能一\n- 新功能二\n\n## 下载\n\n下载说明"):
    if assets is None:
        assets = [{"name": f"chengji-tool-{tag}-windows.zip", "browser_download_url": PREFIX + f"{tag}/chengji-tool-{tag}-windows.zip", "size": 10},
                  {"name": f"chengji-tool-{tag}-mac.zip", "browser_download_url": PREFIX + f"{tag}/chengji-tool-{tag}-mac.zip", "size": 10}]
    return {"tag_name": tag, "body": body, "assets": assets}


@pytest.fixture()
def fake_github(monkeypatch):
    """把“访问 GitHub”换成返回准备好的数据。"""
    box = {"rel": _release(), "urls": []}

    class Resp(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            self.close()

    def fake_open(url, timeout, accept=""):
        box["urls"].append(url)
        if isinstance(box["rel"], Exception):
            raise box["rel"]
        return Resp(json.dumps(box["rel"]).encode("utf-8"))
    monkeypatch.setattr(updater, "_open", fake_open)
    monkeypatch.delenv("CHENGJI_NO_UPDATE_CHECK", raising=False)
    return box


def test_version_order():
    assert updater.parse_version("v2.0.10") > updater.parse_version("2.0.9") > updater.parse_version("2.0")
    assert updater.parse_version("2.0.2") == (2, 0, 2)


def test_pick_asset_by_suffix(monkeypatch):
    """安装包名字带版本号，按 -windows.zip / -mac.zip 结尾认；别的仓库的下载地址不认。"""
    rel = _release("v2.0.2")
    monkeypatch.setattr(sys, "platform", "win32")
    assert updater.pick_asset(rel["assets"])["name"] == "chengji-tool-v2.0.2-windows.zip"
    monkeypatch.setattr(sys, "platform", "darwin")
    assert updater.pick_asset(rel["assets"])["name"] == "chengji-tool-v2.0.2-mac.zip"
    bad = [{"name": "chengji-tool-v2.0.2-mac.zip", "browser_download_url": "https://example.com/chengji-tool-v2.0.2-mac.zip"}]
    assert updater.pick_asset(bad) is None
    monkeypatch.setattr(sys, "platform", "linux")
    assert updater.pick_asset(rel["assets"]) is None


def test_check_newer(fake_github, monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    r = updater.check()
    assert r["ok"] and r["newer"] and r["latest"] == "9.9.9"
    assert r["notes"] == "· 新功能一\n· 新功能二"                      # 只给“更新内容”，不带下载说明
    assert fake_github["urls"] == ["https://api.github.com/repos/lumanman996/chengji-tool/releases/latest"]
    assert not r["canAuto"] and "源码" in r["reason"]                 # 用源码运行时不自己替换


def test_check_not_newer_and_offline(fake_github):
    from chengji import __version__
    fake_github["rel"] = _release("v" + __version__)
    r = updater.check()
    assert r["ok"] and not r["newer"]
    fake_github["rel"] = OSError("没有网络")
    r = updater.check()
    assert not r["ok"] and not r["newer"] and "没有连上" in r["reason"]


def test_check_switched_off_for_tests(monkeypatch):
    monkeypatch.setenv("CHENGJI_NO_UPDATE_CHECK", "1")
    monkeypatch.setattr(updater, "_open", lambda *a, **k: pytest.fail("不该联网"))
    assert updater.check()["newer"] is False


def test_release_notes_from_changelog(tmp_path):
    """发布说明 = CHANGELOG 里这个版本的内容 + 固定的下载说明；程序里显示的就是前一段。"""
    out = tmp_path / "notes.md"
    subprocess.run([sys.executable, str(ROOT / "packaging" / "release_notes.py"), "v2.0.1", str(out)], check=True, capture_output=True)
    text = out.read_text(encoding="utf-8")
    assert text.startswith("## 更新内容") and "联系作者" in text and "## 下载" in text
    assert "## v2.0.0" not in text and "第一个正式版本" not in text           # 只取这一个版本
    assert "联系作者" in updater._changes(text) and "下载" not in updater._changes(text)


def _fake_package(tmp_path) -> Path:
    """按这台电脑的系统，造一个和正式安装包结构一样的小压缩包。"""
    z = tmp_path / "pkg.zip"
    if sys.platform == "darwin":
        stage = tmp_path / "stage" / "分寸"
        exe = stage / "分寸.app" / "Contents" / "MacOS" / "分寸"
        exe.parent.mkdir(parents=True)
        exe.write_text("new")
        (stage / "使用说明.txt").write_text("说明")
        subprocess.run(["ditto", "-c", "-k", "--keepParent", str(stage), str(z)], check=True)
    else:
        with zipfile.ZipFile(z, "w") as f:
            for name, text in [("分寸/分寸.exe", "new"), ("分寸/_internal/x.dll", "new"), ("分寸/使用说明.txt", "说明"),
                               ("分寸/config/学校设置.yaml", "新版默认设置"), ("分寸/config/新加的设置.yaml", "新文件"),
                               ("分寸/data/把任课总表放这里.txt", "提示"), ("分寸/templates/登分表_九年级.xlsx", "模板")]:
                f.writestr(name, text)
    return z


def test_prepare_package(tmp_path):
    up = updater.Updater()
    up.prepare(_fake_package(tmp_path))
    assert up.status()["phase"] == "ready"
    assert up.new_root.name == ("分寸.app" if sys.platform == "darwin" else "分寸")
    bad = tmp_path / "bad.zip"
    with zipfile.ZipFile(bad, "w") as f:
        f.writestr("别的东西/readme.txt", "x")
    with pytest.raises(ValueError, match="没有找到程序"):
        updater.Updater().prepare(bad)


def test_windows_update_keeps_user_data(tmp_path, monkeypatch):
    """Windows 的数据就在程序文件夹里：新版本里的 data、output 整个去掉，config 里已有的文件不覆盖、新增的补上。"""
    monkeypatch.setattr(sys, "platform", "win32")
    up = updater.Updater()
    new = tmp_path / "new" / "分寸"
    for name in ("分寸.exe", "config/学校设置.yaml", "config/新加的设置.yaml", "data/把任课总表放这里.txt", "output/x.txt", "templates/t.xlsx"):
        (new / name).parent.mkdir(parents=True, exist_ok=True)
        (new / name).write_text("new")
    target = tmp_path / "已经装好的"
    (target / "config").mkdir(parents=True)
    (target / "config" / "学校设置.yaml").write_text("本机改过")
    up.new_root = new
    up._keep_user_files(target)
    left = sorted(str(p.relative_to(new)).replace("\\", "/") for p in new.rglob("*") if p.is_file())
    assert left == ["config/新加的设置.yaml", "templates/t.xlsx", "分寸.exe"]


def test_apply_refused_when_not_installed(tmp_path):
    """用源码运行时不会去替换任何文件。"""
    up = updater.Updater()
    up.prepare(_fake_package(tmp_path))
    with pytest.raises(ValueError, match="源码"):
        up.apply()


def test_download_only_from_this_project(tmp_path):
    up = updater.Updater()
    up._download({"browser_download_url": "https://example.com/evil.zip", "size": 1})
    assert up.status()["phase"] == "error" and "下载地址不对" in up.status()["error"]
