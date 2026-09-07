# -*- coding: utf-8 -*-
"""
Edge 调试实例管理（launcher 与 /api/edge/open 共用）
============================================================
- 探测 9222 调试端口是否可达
- 找不到运行中的调试实例时，用独立 profile 拉起带 --remote-debugging-port=9222 的 Edge，
  这样不影响用户日常用的 Edge（登录态、收藏夹互不干扰）。
用法：r = edgeutil.ensure_edge_open(); r["ok"] / r["already"] / r["msg"]
"""
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

CDP_PORT = 9222
BAN_ORIGIN = "*"
SEARCH_HOME_URL = "https://www.zhipin.com/web/geek/job"   # BOSS 求职端首页，方便用户登录后直接搜索


def _app_home() -> Path:
    """打包（frozen）后取 exe 同级目录；源码运行取本模块同级目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROFILE_DIR = _app_home() / "edge_cdp_profile"   # 独立 Edge 配置目录（放登录态）


def cdp_ok(timeout: float = 1.0) -> bool:
    """9222 调试端口是否已经有人监听。"""
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=timeout)
        s.close()
        return True
    except OSError:
        return False


def _find_edge() -> str | None:
    """按平台查找 Edge 可执行文件；找不到返回 None。"""
    if sys.platform.startswith("win"):
        for c in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                  r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
            if os.path.exists(c):
                return c
        return None
    if sys.platform == "darwin":
        p = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        return p if os.path.exists(p) else None
    return None


def ensure_edge_open() -> dict:
    """确保 9222 已有调试版 Edge；没有则拉起。返回 {ok, already, msg}。"""
    if cdp_ok():
        return {"ok": True, "already": True, "msg": "调试端口已就绪"}

    edge = _find_edge()
    if not edge:
        return {"ok": False, "already": False,
                "msg": "找不到 Microsoft Edge，请先安装 Edge 浏览器再运行本程序"}

    try:
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return {"ok": False, "already": False, "msg": f"无法创建 Edge 配置目录：{e}"}

    cmd = [edge,
           f"--remote-debugging-port={CDP_PORT}",
           f"--remote-allow-origins={BAN_ORIGIN}",
           f"--user-data-dir={str(PROFILE_DIR)}",
           SEARCH_HOME_URL]
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError as e:
        return {"ok": False, "already": False, "msg": f"启动 Edge 失败：{e}"}

    # 等 it 开始监听 9222（最多 15 秒）
    for _ in range(15):
        if cdp_ok(timeout=0.5):
            return {"ok": True, "already": False,
                    "msg": "已拉起调试版 Edge，请在 Edge 窗口里登录 BOSS 直聘并打开搜索结果页"}
        time.sleep(1)
    return {"ok": False, "already": False,
            "msg": "Edge 已启动但未检测到调试端口，请确认本机只运行了一个 Edge 调试实例"}