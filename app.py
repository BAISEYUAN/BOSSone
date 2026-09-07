#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSSone API 壳（FastAPI）
============================================================
把本地的"抓取 + 评估"管线暴露成 HTTP 接口，供前端网页调用。

启动方式（在脚本同级目录执行）：
    uvicorn app:app --host 127.0.0.1 --port 5000

接口：
    GET  /           健康检查，确认服务在跑
    GET  /docs       自动生成的调试页，可在浏览器里直接点按钮测试接口
    POST /run        接收用户需求 → 抓取 BOSS 岗位 → 豆包评估 → 返回 JSON

说明：
    本壳只做"翻译"：把前端 JSON 转成 user_need.json，复用 bossone_controller
    里的 run_scraper / evaluate_one / clean_jd_text，业务逻辑全在 controller 里。
    抓取依赖本机已登录 BOSS 的 Edge（localhost:9222），因此本服务只能在本机跑。
"""

# ---- 引入标准库 ----
import json
import os
import re
import sys
import socket
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# ---- 引入第三方库 ----
import openpyxl
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

# ---- 复用主控脚本：常量、run_scraper、evaluate_one、clean_jd_text ----
import bossone_controller as ctrl
import edgeutil

# 数据目录：打包（frozen）后取 exe 同级（用户可写、好找）；源码运行取本文件同级
def _app_home() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

BASE_DIR = _app_home()
USER_NEED_FILE = BASE_DIR / "user_need.json"
# 前端页面：打包后从 PyInstaller 资源目录（_MEIPASS）读，源码运行直接用项目目录里的 index.html
_INDEX_HTML = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "index.html"

# 6 维顺序，与前端一致
DIM_ORDER = ["薪资", "位置", "待遇", "岗位相关性", "无外包风险", "技能匹配度"]

# ---- FastAPI 应用与跨域 ----
app = FastAPI(title="BOSSone API", version="0.1.0")

# 允许所有来源：本地联调时前端端口可能和 API 不同，属于跨域，必须放开
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---- 请求体定义：对应前端表单的 4 个必需字段 + 3 个可选抓取参数 ----
class RunRequest(BaseModel):
    gangweiyaoqiu: str = ""            # 目标岗位（必需）
    xinzixiaxian: int = 15000          # 薪资下限（元/月，必需）
    qiwangdidian: str = ""             # 期望地点（必需）
    daiyuyaoqiu: str = ""              # 待遇要求（必需）
    query: str = ""                    # 抓取关键词（可选，默认取目标岗位）
    city_code: str = "101010100"       # 城市码（可选，默认北京）
    target_num: int = 3                # 抓取条数（可选，默认 3）


def update_user_need(req: RunRequest) -> dict:
    """把前端填的需求写进 user_need.json（保留 provider / model 等原有配置）。"""
    cfg = {}
    if USER_NEED_FILE.exists():
        cfg = json.loads(USER_NEED_FILE.read_text(encoding="utf-8"))
    cfg.update({
        "query": req.query or req.gangweiyaoqiu,   # 抓取关键词：没单独填就用目标岗位
        "city_code": req.city_code,
        "target_num": req.target_num,
        "gangweiyaoqiu": req.gangweiyaoqiu,
        "xinzixiaxian": req.xinzixiaxian,
        "qiwangdidian": req.qiwangdidian,
        "daiyuyaoqiu": req.daiyuyaoqiu,
    })
    USER_NEED_FILE.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return cfg


def _exit_to_http(e: SystemExit):
    """把 controller 里 sys.exit(1) 的错误转成 HTTP 500，避免进程崩溃。"""
    msg = str(e) or "本地脚本退出（原因见后端日志）"
    raise HTTPException(status_code=500, detail=f"本地管线失败：{msg}")


def clean_location(loc):
    """清洗 location：
    ① 混入薪资/经验/学历等列表页噪音（如"ai产品经理 20-30K 经验不限 本科 联想弘扬 北京·海淀区·上地"）→ 置空；
    ② 混入公司名前缀（如"美团外卖 北京·朝阳区·大望路"）→ 只保留含「·」的地点片段；
    ③ 纯地点短串（如"北京·海淀区·上地"）→ 原样保留。"""
    if not loc:
        return ""
    if "经验" in loc or "学历" in loc or "K" in loc or re.search(r"\d+-\d+", loc):
        return ""
    if "·" in loc:
        for tok in loc.split():
            if "·" in tok:
                return tok.strip()
        return ""
    if len(loc) <= 20:
        return loc.strip()
    return ""


def save_report(cfg, items):
    """把 /run 的评估结果（含 6 维 + 总评）落盘为 Excel + MD 到「筛选报告」目录。"""
    ctrl.REPORT_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = f"BOSSone筛选报告_{cfg.get('query', '')}_{ts}"

    # ---- Excel ----
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "筛选结果汇总"
    dims = DIM_ORDER
    ws.append(["岗位名", "公司", "薪资", "经验", "学历", "地点", "URL", "JD原文",
               "总评", "总评理由"] + dims + [d + "理由" for d in dims])
    for it in items:
        ws.append([
            it["name"], it["company"], it["salary"], it["experience"],
            it["education"], it["location"], it["url"], it["jd"],
            it["verdict"], it["verdict_reason"],
        ] + [it["dims"].get(d, "") for d in dims]
          + [it.get("dim_reasons", {}).get(d, "") for d in dims])
    excel_file = ctrl.REPORT_DIR / f"{stem}.xlsx"
    wb.save(str(excel_file))

    # ---- MD ----
    cnt = {"建议投": 0, "谨慎": 0, "不建议": 0}
    for it in items:
        cnt[it["verdict"]] = cnt.get(it["verdict"], 0) + 1
    L = []
    L.append(f"# BOSSone 筛选报告 · {cfg.get('query', '')}")
    L.append("")
    L.append(f"- 时间：{ts} ｜ 共 {len(items)} 条")
    L.append(f"- 目标岗位：{cfg.get('gangweiyaoqiu', '')} ｜ 薪资下限：{cfg.get('xinzixiaxian', '')} 元/月 ｜ 期望地点：{cfg.get('qiwangdidian', '')}")
    L.append(f"- 待遇要求：{cfg.get('daiyuyaoqiu', '')}")
    L.append("")
    L.append(f"**建议投：{cnt['建议投']} ｜ 谨慎：{cnt['谨慎']} ｜ 不建议：{cnt['不建议']}**")
    L.append("")
    for i, it in enumerate(items, 1):
        L.append(f"## {i}. {it['name']} — {it['verdict']}")
        L.append("")
        L.append(f"- 公司：{it['company'] or '—'} ｜ 薪资：{it['salary'] or '—'} ｜ 地点：{it['location'] or '—'}")
        L.append(f"- 经验：{it['experience'] or '—'} ｜ 学历：{it['education'] or '—'}")
        L.append(f"- 链接：{it['url']}")
        L.append("")
        L.append("| 维度 | 判断 | 判断理由 |")
        L.append("| --- | --- | --- |")
        for d in dims:
            L.append(f"| {d} | {it['dims'].get(d, '')} | {it.get('dim_reasons', {}).get(d, '')} |")
        L.append("")
        L.append(f"**总评理由：** {it['verdict_reason']}")
        L.append("")
        L.append("<details><summary>JD 原文</summary>")
        L.append("")
        L.append(it["jd"])
        L.append("")
        L.append("</details>")
        L.append("")
    md_file = ctrl.REPORT_DIR / f"{stem}.md"
    md_file.write_text("\n".join(L), encoding="utf-8")

    return excel_file, md_file


# ---- 接口 1：首页 —— 直接返回前端页面（打包后随 exe 一起分发，用户只开 http://127.0.0.1:5000） ----
@app.get("/")
def home():
    if _INDEX_HTML.exists():
        return HTMLResponse(_INDEX_HTML.read_text(encoding="utf-8"))
    # 兜底：页面文件缺失时仍返回健康检查 JSON，便于排查
    return {"status": "ok", "service": "BOSSone API", "version": "0.1.0", "warning": "index.html 未找到"}


# ---- 接口 2：启动自检 —— 前端页面打开时调用，把缺 Key / Edge 未连等前提问题提示给用户 ----
@app.get("/api/status")
def api_status():
    warnings = []

    # 1) 火山方舟 API Key 是否就绪
    key_ok = bool(os.environ.get("ARK_API_KEY")) or (BASE_DIR / "ark_api_key.txt").exists()
    if not key_ok:
        warnings.append("还没有填火山方舟 API Key：在程序文件夹里新建 <b>ark_api_key.txt</b>，粘贴你的 Key 后重试")

    # 2) 火山方舟模型接入点（model）是否配置
    model_set = False
    try:
        if USER_NEED_FILE.exists():
            _cfg = json.loads(USER_NEED_FILE.read_text(encoding="utf-8"))
            model_set = bool(str(_cfg.get("model", "")).strip())
    except Exception:
        pass
    if not model_set:
        warnings.append("还没配置模型接入点：在程序文件夹的 <b>user_need.json</b> 里，把 <b>model</b> 填成你的火山方舟接入点ID（ep-xxx）或模型ID")

    # 3) Edge 调试端口（9222）是否连上
    edge_ok = False
    try:
        s = socket.create_connection(("127.0.0.1", 9222), timeout=1)
        s.close()
        edge_ok = True
    except OSError:
        pass
    if not edge_ok:
        warnings.append("还没连上调试版 Edge：请关闭程序后重新打开（会自动拉起 Edge），或用程序自带的「打开 Edge」按钮")

    # 3) 兜底：确保报告目录存在
    try:
        ctrl.RESULT_DIR.mkdir(exist_ok=True)
        ctrl.REPORT_DIR.mkdir(exist_ok=True)
    except OSError:
        pass

    return {
        "key_ok": key_ok,
        "edge_ok": edge_ok,
        "model_set": model_set,
        "key_file": str(BASE_DIR / "ark_api_key.txt"),
        "warnings": warnings,
    }


# ---- 接口 3：手动打开调试版 Edge（页面顶端「打开 Edge」按钮调用） ----
@app.get("/api/edge/open")
def edge_open():
    return edgeutil.ensure_edge_open()


# ---- 接口 4：跑一次完整筛选 ----
@app.post("/run")
def run_screening(req: RunRequest):
    # 1) 校验必需字段
    if not req.gangweiyaoqiu or not req.qiwangdidian:
        raise HTTPException(status_code=400, detail="目标岗位和期望地点不能为空")

    # 2) 把需求写进 user_need.json
    cfg = update_user_need(req)

    # 3) 读 API 密钥（provider=ark 读 ark_api_key.txt）
    try:
        token = ctrl.load_api_key(cfg.get("provider", "ark"))
    except SystemExit as e:
        _exit_to_http(e)

    # 4) 调 V15B 抓取（阻塞，抓 3 条约 1-2 分钟）
    try:
        excel_path, cfg = ctrl.run_scraper(cfg)
    except SystemExit as e:
        _exit_to_http(e)

    # 5) 读取抓取结果 Excel
    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    # 6) 并发评估（复用 controller 的 evaluate_one，含超时重试）
    max_workers = max(1, min(ctrl.COZE_CONCURRENCY, len(rows)))
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(ctrl.evaluate_one, idx, row, cfg, token): idx
                   for idx, row in enumerate(rows, start=1)}
        for fut in as_completed(futures):
            results.append(fut.result())
    results.sort(key=lambda r: r["行号"])

    # 7) 组装前端需要的 JSON（列映射见 build_jd_text 注释）
    items = []
    for r in results:
        row = rows[r["行号"] - 1]
        dim_map = {}
        dim_reason_map = {}
        for _d in r["维度"]:
            dim_map[_d["评估项"]] = _d["匹配"]
            dim_reason_map[_d["评估项"]] = _d.get("理由", "")
        items.append({
            "name": row[3] or "",
            "company": row[7] or "",
            "salary": row[4] or "",
            "experience": row[5] or "",
            "education": row[6] or "",
            "location": clean_location(row[8] or ""),
            "url": str(row[9] or "").strip(chr(96)),   # 剥掉 BOSS 列表页塞进来的反引号
            "jd": ctrl.clean_jd_text(row[10] or ""),
            "verdict": r["总评"],
            "verdict_reason": r["总评理由"],
            "dims": {d: dim_map.get(d, "") for d in DIM_ORDER},
            "dim_reasons": {d: dim_reason_map.get(d, "") for d in DIM_ORDER},
        })

    # 8) 落盘：把评估结果（含 6 维 + 总评）存成 Excel + MD，方便离线查看
    report = {}
    try:
        excel_file, md_file = save_report(cfg, items)
        report = {"excel": str(excel_file), "md": str(md_file)}
    except Exception as e:
        print(f"[落盘] 报告保存失败（不影响返回）：{e}")

    return {
        "total": len(items),
        "query": cfg["query"],
        "city_code": cfg["city_code"],
        "items": items,
        "report": report,
    }

# ---- 接口 3：历史记录（读「筛选报告」目录里已落盘的 Excel/MD）----
REPORT_STEM_RE = re.compile(r"^BOSSone筛选报告_(.+)_(\d{8}_\d{6})$")


def _read_report_items(excel_path) -> list:
    """把一份历史报告的 Excel 读回成与 /run 返回 items 一致的 JSON（便于前端复用渲染）。"""
    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c) if c is not None else "" for c in rows[0]]
    items = []
    for row in rows[1:]:
        if row is None or all(c in (None, "") for c in row):
            continue
        rec = dict(zip(header, row))
        items.append({
            "name": rec.get("岗位名", "") or "",
            "company": rec.get("公司", "") or "",
            "salary": rec.get("薪资", "") or "",
            "experience": rec.get("经验", "") or "",
            "education": rec.get("学历", "") or "",
            "location": clean_location(rec.get("地点", "") or ""),
            "url": str(rec.get("URL", "") or "").strip(chr(96)),
            "jd": rec.get("JD原文", "") or "",
            "verdict": rec.get("总评", "") or "",
            "verdict_reason": rec.get("总评理由", "") or "",
            "dims": {d: rec.get(d, "") or "" for d in DIM_ORDER},
            "dim_reasons": {d: rec.get(d + "理由", "") or "" for d in DIM_ORDER},
        })
    return items


@app.get("/history")
def list_history():
    """列出「筛选报告」目录下所有历史报告（按时间倒序）。"""
    ctrl.REPORT_DIR.mkdir(exist_ok=True)
    files = sorted(ctrl.REPORT_DIR.glob("BOSSone筛选报告_*.xlsx"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    records = []
    for p in files:
        m = REPORT_STEM_RE.match(p.stem)
        query = m.group(1) if m else p.stem
        md = p.with_suffix(".md")
        records.append({
            "file": p.name,
            "query": query,
            "ts": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "has_md": md.exists(),
        })
    return {"total": len(records), "records": records}


@app.get("/history/{file_name}")
def get_history(file_name: str):
    """读取某一份历史报告的完整内容（岗位列表），返回与 /run 相同的 items 结构。"""
    excel_file = (ctrl.REPORT_DIR / file_name).resolve()
    # 防路径穿越：必须落在 REPORT_DIR 内
    if not str(excel_file).startswith(str(ctrl.REPORT_DIR.resolve())) or not excel_file.exists():
        raise HTTPException(status_code=404, detail="历史记录不存在")
    items = _read_report_items(excel_file)
    m = REPORT_STEM_RE.match(excel_file.stem)
    query = m.group(1) if m else excel_file.stem
    return {
        "total": len(items),
        "query": query,
        "items": items,
        "report": {"excel": str(excel_file),
                   "md": str(excel_file.with_suffix(".md"))},
    }
