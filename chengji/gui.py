"""图形界面入口：开一个独立的程序窗口，里面显示界面（界面由本机的小服务提供）。

电脑上没有可用的窗口组件时，退而用系统默认浏览器打开同一个界面。
"""
from __future__ import annotations

import sys
import threading
import time
import webbrowser

from . import APP_TITLE
from .server import start

TITLE = APP_TITLE


def selftest() -> int:
    """打包后的自检：界面文件在不在、本地服务通不通、许可状态是什么。不开窗口。"""
    import json
    import re
    import urllib.request
    httpd, url, app = start()
    html = urllib.request.urlopen(url, timeout=10).read().decode("utf-8")
    token = re.search(r'const TOKEN = "([^"]+)"', html).group(1)
    req = urllib.request.Request(url + "api/state", data=b"{}", headers={"X-Token": token, "Content-Type": "application/json"})
    st = json.loads(urllib.request.urlopen(req, timeout=10).read().decode("utf-8"))
    httpd.shutdown()
    ok = "分寸" in html and st.get("ok")
    from . import license as lic
    info = {"ok": bool(ok), "version": st["data"]["version"], "license": st["data"]["license"]["state"], "root": str(app.root),
            "licenseModule": lic.module_kind(), "canExport": st["data"]["license"]["canExport"]}
    (app.root / "自检结果.json").write_text(json.dumps(info, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(info, ensure_ascii=False))
    return 0 if ok else 1


def _screen_fit(webview) -> tuple[int, int, bool]:
    """窗口开多大：屏幕够大就 1320×880；屏幕小（如 1366×768，或者系统缩放 125%、150% 后放不下）就直接铺满，免得窗口比屏幕还大。

    pywebview 的窗口尺寸按“缩放后的”算；屏幕尺寸在 Windows 上可能是没缩放的，所以再除一下系统缩放比例
    （万一已经是缩放后的，除了只会显得屏幕更小，结果是铺满窗口，也不碍事）。算不出来就当屏幕够大。"""
    try:
        sc = webview.screens[0]
        w, h = sc.width, sc.height
        if sys.platform == "win32":
            import ctypes
            k = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100 or 1
            w, h = w / k, h / k
    except Exception:
        return 1320, 880, False
    return min(1320, int(w) - 40), min(880, int(h) - 80), w < 1300 or h < 800


def main(browser: bool = False):
    httpd, url, app = start()
    if not browser:
        try:
            import webview
            width, height, small = _screen_fit(webview)
            app.window = webview.create_window(TITLE, url, width=max(width, 860), height=max(height, 540),
                                               min_size=(860, 540), maximized=small,
                                               background_color="#111315" if app.dark_start() else "#F5F4F0")
            webview.start()
            httpd.shutdown()
            return
        except Exception as e:                               # 没有窗口组件（或启动失败）：改用浏览器
            app.window = None
            print(f"窗口组件不可用（{type(e).__name__}），改用浏览器打开。")
    webbrowser.open(url)
    print(f"界面已在浏览器里打开：{url}\n关掉浏览器里的那个页面，本程序稍后会自动退出。")
    app.last_ping = time.time() + 20                         # 给浏览器一点启动时间
    try:
        while time.time() - app.last_ping < 30:              # 页面每 5 秒报一次到；30 秒没动静就退出
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    httpd.shutdown()


if __name__ == "__main__":
    main(browser="--browser" in sys.argv)
