# -*- coding: utf-8 -*-
"""
BOSSone 启动器（打包后为 BOSSone.exe）
============================================================
流程：
  1) 确保 9222 有调试版 Edge（没有则由 edgeutil 自动拉起，独立 profile）
  2) 打开前端页面 http://127.0.0.1:5000
  3) 启动 uvicorn 后端

打包：PyInstaller 入口脚本用本文件（launcher.py），app.py / controller / edgeutil
     会作为依赖模块一并收集；index.html 通过 --add-data 随包分发。
"""
import json
import socket
import sys
import webbrowser

import uvicorn

import app         # noqa: F401  确保 app 模块（FastAPI 应用）被加载
import edgeutil

# 交付模板：首次运行时若程序文件夹没有 user_need.json，则生成一份干净的（model 留空，等用户填）
_TEMPLATE_USER_NEED = {
    "provider": "ark",
    "query": "AI产品经理",
    "city_code": "101010100",
    "target_num": 3,
    "xinzixiaxian": 15000,
    "qiwangdidian": "北京",
    "daiyuyaoqiu": "双休",
    "gangweiyaoqiu": "AI产品经理",
    "model": "",
    "workflow_id": "",
}


def _ensure_default_files():
    """补齐首次运行需要的文件：user_need.json（用户自填 model）。"""
    try:
        if not app.USER_NEED_FILE.exists():
            app.USER_NEED_FILE.write_text(
                json.dumps(_TEMPLATE_USER_NEED, ensure_ascii=False, indent=2),
                encoding="utf-8")
            print(" · 已生成默认 user_need.json（请按提示填上你的 model 接入点）")
    except Exception as e:
        print(" ⚠ 初始化 user_need.json 失败：" + str(e))


def _pick_port(start: int = 5000) -> int:
    """从 start 起找一个本机空闲端口（避免与其他程序/旧后端抢占）。"""
    for p in range(start, start + 20):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", p))
            s.close()
            return p
        except OSError:
            continue
    return start


def main():
    print("=" * 60)
    print(" BOSSone 智能筛岗助手 启动中…（请勿关闭本窗口）")
    print("=" * 60)

    _ensure_default_files()

    # 1) 确保调试版 Edge 在跑（登录 BOSS 是第一步）
    try:
        r = edgeutil.ensure_edge_open()
        if r["ok"]:
            if r["already"]:
                print(" · Edge 调试端口已就绪（9222）")
            else:
                print(" · 已拉起调试版 Edge（独立配置，不影响你日常的 Edge）")
                print("   请在弹出的 Edge 窗口里登录 BOSS 直聘，并打开岗位搜索结果列表页")
        else:
            print(" ⚠ 启动 Edge 失败：" + r["msg"])
            print("   仍将继续启动网页服务；可在页面顶端点击「打开 Edge（登录 BOSS）」重试")
    except Exception as e:
        print(" ⚠ 启动 Edge 时出现异常：" + str(e))

    # 2) 选空闲端口（默认 5000，被占用自动后移）
    port = _pick_port(5000)
    url = f"http://127.0.0.1:{port}"
    try:
        webbrowser.open(url)
    except Exception:
        pass

    # 3) 后端服务（FastAPI + uvicorn）
    print(f" · 网页服务启动中：{url}")
    print("-" * 60)
    uvicorn.run(app.app, host="127.0.0.1", port=port, log_level="info")


if __name__ == "__main__":
    main()