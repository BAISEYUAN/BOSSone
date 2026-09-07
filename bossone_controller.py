#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSSone 本地主控脚本（路线 A）
功能：读用户需求 → 调用 V15B 抓取 BOSS 岗位 → 逐条调大模型 API 评估 → 汇总成 Excel 报告。
支持两种评估后端（user_need.json 里 provider 字段切换）：
  - provider = "coze"：调 Coze 工作流 API（原方案，需 COZE_TOKEN + workflow_id）
  - provider = "ark" ：调火山方舟豆包（OpenAI 兼容接口，需 ARK_API_KEY + model）
使用方式：
    1. 确保 user_need.json 已按需求填写（provider / model 等）。
    2. 准备 API 密钥：Coze 用 COZE_TOKEN / coze_token.txt；火山方舟用 ARK_API_KEY / ark_api_key.txt。
    3. 运行：python bossone_controller.py
"""

# ---- 第 1 行：引入标准库 ----
import json                     # 解析 JSON：Coze 返回、user_need.json、结果报告都用它
import os                       # 操作文件路径、读环境变量
import re                       # 正则表达式：从非标准文本里抠评估结果
import subprocess               # 调用外部程序：运行 V15B 抓取脚本
import sys                      # 控制脚本退出、打印错误到 stderr
import time                     # 控制请求间隔，避免触发 Coze 频率限制
import urllib.request           # 不用额外安装 requests，用标准库发 HTTP POST
import urllib.error             # 捕获 HTTP 错误（4xx、5xx）
from datetime import datetime   # 给输出报告文件名加时间戳
from concurrent.futures import ThreadPoolExecutor, as_completed  # 并发调 Coze 评估
from pathlib import Path        # 现代化路径处理，比 os.path 更直观

# openpyxl 是 V15B 已经在用的库，主控脚本复用它读写 Excel
try:
    import openpyxl
except ImportError:
    print("错误：缺少 openpyxl。请用你运行 V15B 的那个 Python 执行：pip install openpyxl")
    sys.exit(1)


# ---- 第 2 块：配置常量 ----
def _app_home() -> Path:
    """数据目录定位：打包（PyInstaller frozen）后取 exe 同级目录（用户可写、好找）；
    源码直接运行时取本文件同级目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

BASE_DIR = _app_home()                                  # 数据根目录（Key/需求/报告都放这里）
USER_NEED_FILE = BASE_DIR / "user_need.json"            # 用户需求配置文件
SCRAPER_SCRIPT = BASE_DIR / "V15B_终极纯CDP_DOM抓_0内部API_不超时.py"  # 抓取脚本（源码运行用）
SCRAPER_EXE = BASE_DIR / "scraper.exe"                  # 抓取子进程（打包后 V15B 的独立 exe）
RESULT_DIR = BASE_DIR / "抓取记录"                    # V15B 抓好的 Excel 落在这里
REPORT_DIR = BASE_DIR / "筛选报告"                       # 最终汇总报告 Excel 落在这里
COZE_API_BASE = "https://api.coze.cn"                   # 扣子国内版 API 域名
ARK_API_BASE = "https://ark.cn-beijing.volces.com/api/v3"   # 火山方舟 OpenAI 兼容接口
COZE_CONCURRENCY = 3                                          # 并发评估条数（受 Coze 并发配额限制，遇 429 调小）

# 豆包评估提示词：把原 Coze 工作流（A/B/C 三 Agent）的评估逻辑合并成一段，直接喂给豆包
EVAL_PROMPT = """你是一名资深招聘筛选助手，根据求职者需求和岗位JD原文，做6个维度评估并给出总评。严格按要求输出。

## 求职者需求
- 目标岗位：{gangweiyaoqiu}
- 薪资下限：{xinzixiaxian} 元/月
- 期望地点：{qiwangdidian}
- 待遇要求：{daiyuyaoqiu}

## 岗位JD原文
{jd_wenben}

## 评估规则（务必严格遵守）

### 三态定义（量的是"信息确定性"，不是岗位好坏）
- 达标：JD信息足够且确定满足
- 不达标：JD信息足够且确定不满足
- 存疑：JD信息不足无法判断；此档理由必须以「信息不足：」开头

### 6个评估项
1. 薪资：比较JD薪资区间下限与求职者薪资下限。
   - JD区间下限 >= 薪资下限 → 达标
   - JD区间下限 < 薪资下限 但 上限 >= 薪资下限 → 不达标（理由注明"上限有空间，面试可谈"）
   - JD上限 < 薪资下限 → 不达标
   - JD未写明薪资 → 存疑
2. 位置：判断JD工作地点（城市/区域）是否与期望地点匹配。
3. 待遇：只检查JD是否明确写出求职者要求的待遇（如双休、五险一金）；未明确写出的判存疑，不要自行推断。
4. 岗位相关性：判断JD实际职责是否与目标岗位相符（防岗位名与JD名实不符）。
5. 无外包风险：检查JD是否出现外包/派遣/转包/借调/灵活用工/派驻/人力资源服务等词；出现则 不达标。
6. 技能匹配度：判断JD要求的能力/经验是否与目标岗位常规要求匹配。

### 总评规则
- 建议投：6维全部达标
- 谨慎：存在不达标但理由偏软（如薪资上限有空间），或个别维度存疑
- 不建议：存在关键维度硬性不达标（如岗位相关性、无外包风险、薪资硬性不达标）
总评理由必须能回查JD原文对应句，不得编造JD没有的信息。

## 输出格式（只输出一个JSON对象，不要输出其他任何文字）
{{"岗位名": "JD中的岗位名称", "总评": "建议投/谨慎/不建议", "总评理由": "…", "维度": [
  {{"评估项": "薪资", "匹配": "达标/不达标/存疑", "理由": "…"}},
  {{"评估项": "位置", "匹配": "…", "理由": "…"}},
  {{"评估项": "待遇", "匹配": "…", "理由": "…"}},
  {{"评估项": "岗位相关性", "匹配": "…", "理由": "…"}},
  {{"评估项": "无外包风险", "匹配": "…", "理由": "…"}},
  {{"评估项": "技能匹配度", "匹配": "…", "理由": "…"}}
]}}"""


# ---- 第 3 块：加载用户配置 ----
def load_user_need():
    """读取 user_need.json，把用户画像和抓取参数一起返回成字典。"""
    # 检查配置文件是否存在，防止用户还没填就运行
    if not USER_NEED_FILE.exists():
        print(f"错误：找不到用户需求文件 {USER_NEED_FILE}")
        sys.exit(1)

    # 打开并解析 JSON，encoding='utf-8' 保证中文不乱码
    with open(USER_NEED_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 必要字段校验：缺一个就没法正确评估
    required = ["query", "city_code", "target_num", "gangweiyaoqiu",
                "xinzixiaxian", "qiwangdidian", "daiyuyaoqiu"]
    # workflow_id 仅 Coze 后端需要，model 仅火山方舟需要，都放到对应后端里校验
    missing = [k for k in required if k not in cfg]
    if missing:
        print(f"错误：user_need.json 缺少字段 {missing}")
        sys.exit(1)

    return cfg


# ---- 第 4 块：读取 API 密钥 ----
def load_api_key(provider):
    """按后端读取 API 密钥：优先环境变量，其次本地文件。
    provider="coze" 用 COZE_TOKEN / coze_token.txt；provider="ark" 用 ARK_API_KEY / ark_api_key.txt。"""
    if provider == "coze":
        token = os.environ.get("COZE_TOKEN")                  # 先看环境变量，安全且不暴露到代码
        if token:
            return token

        token_file = BASE_DIR / "coze_token.txt"            # 备选：本地放一个只有 token 的 txt
        if token_file.exists():
            return token_file.read_text(encoding="utf-8-sig").strip()

        print("错误：找不到 Coze API Token。请执行以下任一操作：")
        print("  1. 设置环境变量 COZE_TOKEN=你的令牌")
        print("  2. 在脚本同级目录创建 coze_token.txt，写入令牌")
        sys.exit("还没有填 Coze API Token：请在程序文件夹里创建 coze_token.txt，把令牌粘贴进去后重试")

    elif provider == "ark":
        key = os.environ.get("ARK_API_KEY")
        if key:
            return key

        key_file = BASE_DIR / "ark_api_key.txt"
        if key_file.exists():
            return key_file.read_text(encoding="utf-8-sig").strip()

        print("错误：找不到火山方舟 API Key。请执行以下任一操作：")
        print("  1. 设置环境变量 ARK_API_KEY=你的密钥")
        print("  2. 在脚本同级目录创建 ark_api_key.txt，写入密钥")
        sys.exit("还没有填火山方舟 API Key：请在程序文件夹里创建 ark_api_key.txt 文件，把你的密钥粘贴进去后重试")

    else:
        print(f"错误：user_need.json 的 provider 必须是 coze 或 ark，当前是：{provider}")
        sys.exit(f"配置错误：user_need.json 的 provider 只能是 coze 或 ark，当前是：{provider}")


# ---- 第 5 块：调用抓取脚本 ----
def run_scraper(cfg):
    """调用 V15B 完成抓取，返回 (excel路径, 修正后cfg)。
    V15B 方案B按页面URL实际关键词抓取，这里解析 stdout 的 ACTUAL_QUERY/ACTUAL_CITY 同步 cfg，保证报告命名一致。
    打包（frozen）后 V15B 被打成独立 scraper.exe 子进程，需求改为命令行参数传入；
    源码运行保持原逻辑：python V15B.py（它自己读 user_need.json）。"""
    if getattr(sys, "frozen", False):
        cmd = [str(SCRAPER_EXE),
               "--query", str(cfg.get("query", "")),
               "--city", str(cfg.get("city_code", "")),
               "--num", str(int(cfg.get("target_num", 3) or 3)),
               "--min-salary", str(int(cfg.get("xinzixiaxian", 0) or 0)),
               "--city-name", str(cfg.get("qiwangdidian", "")),
               "--benefit", str(cfg.get("daiyuyaoqiu", "")),
               "--target", str(cfg.get("gangweiyaoqiu", ""))]
    else:
        python = sys.executable
        cmd = [python, str(SCRAPER_SCRIPT)]

    # 记录运行前最新的抓取文件（按名字），用于判断 V15B 是否真的抓到了新数据
    before = sorted(RESULT_DIR.rglob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    before_names = {p.name for p in before}

    print(f"\n[抓取] 启动 {SCRAPER_SCRIPT.name}")
    print(f"       （配置）关键词：{cfg['query']}  城市：{cfg['city_code']}  目标条数：{cfg['target_num']}")
    print(f"       （说明）方案B模式下，实际关键词/城市以 V15B 当前页面URL解析结果为准")

    result = subprocess.run(cmd, cwd=str(BASE_DIR), capture_output=True, text=True)
    if result.returncode != 0:
        print("抓取脚本出错：")
        print(result.stdout)
        print(result.stderr)
        sys.exit(1)

    print(result.stdout)

    # ⚠️ 关键检查：V15B 是否真的生成了新的抓取文件（成功抓到至少 1 条才会保存）。
    # 若没有新文件（典型场景：Edge 当前停在岗位详情页，列表页 0 张卡片），立即报错，
    # 避免静默复用上一次的旧 Excel，让前端把旧数据当成"这次的新筛选结果"展示给用户。
    after = sorted(RESULT_DIR.rglob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    new_xlsx = next((p for p in after if p.name not in before_names), None)
    if new_xlsx is None:
        print("抓取失败：V15B 没有生成新的抓取文件。")
        print("最可能原因：Edge 当前打开的页面不是岗位搜索结果列表页（例如停留在某个岗位详情页），DOM 抓不到卡片。")
        print("解决办法：切回 BOSS 直聘的岗位搜索结果列表页后重新运行。")
        sys.exit("抓取失败：V15B 未生成新抓取文件，Edge 当前页不是搜索结果列表页，请先切回 BOSS 搜索列表页再试")

    # 解析 stdout 里 V15B 结尾打印的实际关键词/城市
    q_m = re.search(r'ACTUAL_QUERY=([^\s]+)', result.stdout)
    c_m = re.search(r'ACTUAL_CITY=([^\s]+)', result.stdout)
    actual_q = q_m.group(1).strip('"\'') if q_m else cfg["query"]
    actual_c = c_m.group(1).strip('"\'') if c_m else cfg["city_code"]
    if q_m or c_m:
        print(f"[抓取] （实际）关键词：{actual_q}  城市：{actual_c}（与配置不一致为正常）")
        cfg2 = dict(cfg)
        cfg2["query"] = actual_q
        cfg2["city_code"] = actual_c
    else:
        cfg2 = cfg

    return find_latest_excel(actual_q, actual_c), cfg2


# ---- 第 6 块：找到最新抓取结果 ----
def find_latest_excel(expected_query=None, expected_city=None):
    """找最新（城市码+关键词）匹配的.xlsx；匹配不到才回退全局最新。
    避免：配置是A岗位、页面实际搜B时，结果文件被误读成另一关键词的旧文件。"""
    if not RESULT_DIR.exists():
        print(f"错误：抓取结果目录不存在 {RESULT_DIR}")
        sys.exit(1)

    files_all = sorted(RESULT_DIR.rglob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files_all:
        print(f"错误：{RESULT_DIR} 里没有 .xlsx 文件")
        sys.exit(1)

    if expected_query or expected_city:
        filt = list(files_all)
        if expected_city:
            filt = [p for p in filt if expected_city in p.name]
        if expected_query:
            q_filt = [p for p in filt if expected_query in p.name]
            if q_filt: filt = q_filt
        if filt:
            latest = filt[0]
            print(f"[抓取] 使用匹配结果：{latest.name}（城市码{expected_city or '-'}+关键词{expected_query or '-'}）")
            return latest
        else:
            print(f"[抓取] _ 没找到匹配 城市码{expected_city or '-'}+关键词{expected_query or '-'} 的xlsx，回退全局最新：{files_all[0].name}")

    latest = files_all[0]
    print(f"[抓取] 使用最新结果：{latest.name}")
    return latest


# ---- 第 7 块：调用 Coze 工作流 API ----
def call_coze_workflow(token, cfg, jd_text, workflow_id):
    """调一次 Coze 工作流，返回 Agent C 的原始输出字符串。"""
    url = f"{COZE_API_BASE}/v1/workflow/run"              # 非流式接口，直接返回最终结果

    # 构造请求体：字段名必须和 Coze 工作流开始节点的变量名完全一致
    payload = {
        "workflow_id": workflow_id,
        "parameters": {
            "gangweiyaoqiu": cfg["gangweiyaoqiu"],
            "xinzixiaxian": cfg["xinzixiaxian"],
            "qiwangdidian": cfg["qiwangdidian"],
            "daiyuyaoqiu": cfg["daiyuyaoqiu"],
            "jd_wenben": jd_text
        }
    }

    # 把字典转成字节流，utf-8 编码
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    # 构造 HTTP 请求对象，设置 Header：认证 + 内容类型 + 内容长度
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    try:
        # 工作流（豆包大模型）单次推理可能 30~300 秒，读超时放宽到 300 秒
        with urllib.request.urlopen(req, timeout=300) as resp:
            raw = resp.read().decode("utf-8")
            # 识别 Coze 业务错误：成功时 code=0，业务失败（如 4028 积分不足/参数错）时 code!=0，但仍是 HTTP 200
            try:
                _j = json.loads(raw)
                if isinstance(_j, dict) and "code" in _j and _j.get("code") not in (0, None, ""):
                    _msg = _j.get("msg") or _j.get("message") or "未知错误"
                    raise RuntimeError(f"Coze 业务错误 code={_j['code']}: {_msg}")
            except json.JSONDecodeError:
                pass  # 非 JSON 就直接当原始文本返回
            return raw
    except urllib.error.HTTPError as e:
        # HTTP 错误时把响应体也读出来，方便定位 400/401/429 原因
        err_body = e.read().decode("utf-8") if e.fp else ""
        raise RuntimeError(f"Coze HTTP {e.code}: {err_body}") from e
    except urllib.error.URLError as e:
        # 读超时/连接中断：统一转成 TimeoutError，供外层自动重试判断
        reason = getattr(e, "reason", e)
        msg = str(reason)
        if "timed out" in msg.lower() or "timeout" in msg.lower():
            raise TimeoutError(f"Coze 工作流响应超时（>300秒）：{msg}") from e
        raise RuntimeError(f"Coze 网络错误：{msg}") from e


# ---- 第 7.5 块：调用火山方舟豆包（OpenAI 兼容接口）----
def call_doubao_chat(api_key, model, jd_text, cfg):
    """直接调豆包 chat/completions，返回模型输出的文本（含 JSON）。model 为接入点ID(ep-xxx)或模型ID。"""
    url = f"{ARK_API_BASE}/chat/completions"

    # 把评估提示词和 JD 拼进 user 消息
    system_prompt = EVAL_PROMPT.format(
        gangweiyaoqiu=cfg["gangweiyaoqiu"],
        xinzixiaxian=cfg["xinzixiaxian"],
        qiwangdidian=cfg["qiwangdidian"],
        daiyuyaoqiu=cfg["daiyuyaoqiu"],
        jd_wenben=jd_text,
    )
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "你是一名严格的招聘筛选评估助手，严格按用户指令输出 JSON。"},
            {"role": "user", "content": system_prompt},
        ],
        "temperature": 0.2,
        "max_tokens": 2000,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            raw = resp.read().decode("utf-8")
            try:
                obj = json.loads(raw)
                return obj["choices"][0]["message"]["content"]
            except (KeyError, IndexError, json.JSONDecodeError) as e:
                raise RuntimeError(f"火山方舟返回格式异常：{raw[:500]}") from e
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8") if e.fp else ""
        raise RuntimeError(f"火山方舟 HTTP {e.code}: {err_body}") from e
    except urllib.error.URLError as e:
        reason = getattr(e, "reason", e)
        msg = str(reason)
        if "timed out" in msg.lower() or "timeout" in msg.lower():
            raise TimeoutError(f"火山方舟响应超时（>300秒）：{msg}") from e
        raise RuntimeError(f"火山方舟网络错误：{msg}") from e


# ---- 第 8 块：解析 Agent C 输出 ----
def _unwrap(obj):
    """把 Coze 多层包裹（{code,data} / {output1:"<json str>"} / 列表）递归剥到最内层。"""
    if isinstance(obj, dict):
        # Coze 标准包裹：{code, msg, data:"<json>"}
        if "data" in obj and isinstance(obj["data"], (str, dict)):
            return _unwrap(obj["data"])
        # 形如 {output1: "<json 字符串>"}：尝试把每个字符串值也解析一层
        new = {}
        changed = False
        for k, v in obj.items():
            if isinstance(v, str):
                try:
                    new[k] = _unwrap(json.loads(v))
                    changed = True
                except json.JSONDecodeError:
                    new[k] = v
            else:
                new[k] = _unwrap(v)
        return new if changed else obj
    if isinstance(obj, str):
        try:
            return _unwrap(json.loads(obj))
        except json.JSONDecodeError:
            return obj
    if isinstance(obj, list):
        return [_unwrap(x) for x in obj]
    return obj


def extract_structured(d):
    """递归遍历任意深度的结构，提取总评 / 总评理由 / 岗位名 / 各维度。"""
    result = {"岗位名": "", "总评": "", "总评理由": "", "维度": []}

    def walk(obj):
        if isinstance(obj, dict):
            if "总评" in obj and not result["总评"]:
                result["总评"] = obj["总评"]
            if "总评理由" in obj and not result["总评理由"]:
                result["总评理由"] = obj["总评理由"]
            if "岗位名" in obj and not result["岗位名"]:
                result["岗位名"] = obj["岗位名"]
            # 标准对照表元素：{评估项, 匹配, 理由}
            if "评估项" in obj and "匹配" in obj and "理由" in obj:
                result["维度"].append({
                    "评估项": obj["评估项"],
                    "匹配": obj["匹配"],
                    "理由": obj["理由"]
                })
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for it in obj:
                walk(it)

    walk(d)
    return result


def parse_agent_c_output(raw_text):
    """
    把 Agent C 的输出解析成结构化字典。
    优先按 Coze 真实返回结构（多层包裹的 JSON）提取；
    兼容纯 JSON、markdown 代码块、以及树状文本三种情况。
    """
    text = raw_text.strip()

    # 优先：剥开 Coze 包裹，尝试结构化提取
    try:
        wrapped = json.loads(text)
        obj = _unwrap(wrapped)
        if isinstance(obj, dict):
            structured = extract_structured(obj)
            if structured["维度"] or structured["总评"]:
                return structured
    except json.JSONDecodeError:
        pass

    # 尝试 2：从 markdown 代码块里抠 JSON
    code_block = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if code_block:
        try:
            obj = json.loads(code_block.group(1))
            structured = extract_structured(obj)
            if structured["维度"] or structured["总评"]:
                return structured
        except json.JSONDecodeError:
            pass

    # 尝试 3：兜底正则，解析树状文本
    return parse_tree_like_output(text)


def parse_tree_like_output(text):
    """兜底：从类似 Coze 试运行树状文本里抠出关键字段。"""
    result = {
        "岗位名": "",
        "总评": "",
        "总评理由": "",
        "维度": []
    }

    # 抠岗位名
    m_name = re.search(r'岗位名\s*[:：]\s*"?([^"\n]+)"?', text)
    if m_name:
        result["岗位名"] = m_name.group(1).strip()

    # 抠总评与总评理由
    m_total = re.search(r'总评\s*[:：]\s*"?([^"\n]+)"?', text)
    if m_total:
        result["总评"] = m_total.group(1).strip()
    m_reason = re.search(r'总评理由\s*[:：]\s*"?([^"\n]+)"?', text)
    if m_reason:
        result["总评理由"] = m_reason.group(1).strip()

    # 抠每个维度：评估项 / 匹配 / 理由
    # 用正则找到所有 "评估项 : " 开头的块
    items = re.findall(
        r'评估项\s*[:：]\s*"?([^"\n]+)"?\s*\n\s*匹配\s*[:：]\s*"?([^"\n]+)"?\s*\n\s*理由\s*[:：]\s*"?([^"\n]+)"?',
        text
    )
    for item, status, reason in items:
        result["维度"].append({
            "评估项": item.strip(),
            "匹配": status.strip(),
            "理由": reason.strip()
        })

    return result


# ---- 第 9 块：从 Excel 行构造完整 JD 文本 ----
def build_jd_text(row):
    """把 Excel 里的多列合并成一段完整 JD，供工作流评估。"""
    # 列对应：0=# 1=来源层级 2=岗位ID 3=岗位名称 4=薪资范围 5=经验 6=学历 7=公司 8=地址 9=URL 10=JD原文
    title = row[3] or ""
    salary = row[4] or ""
    company = row[7] or ""
    address = row[8] or ""
    jd_body = row[10] or ""

    # 清洗 JD 原文：剔除 BOSS 页面噪音、招聘者信息、重复段落（见 clean_jd_text）
    jd_body = clean_jd_text(jd_body)

    full_jd = f"""【职位】{title}
【公司】{company}
【地点】{address}
【薪资】{salary}

【JD原文】
{jd_body}
"""
    return full_jd.strip()


# ---- 第 9.4 块：JD 原文清洗（剔除 BOSS 页面噪音）----
def clean_jd_text(text):
    """把 V15B 抓回的 JD 原文去噪：页面导航、招聘者信息、工商/安全提示区块、重复段落。

    基于 2026-08-31 实际抓取样本观察到的噪音模式编写：
    1. 开头「微信扫 码 分享 举报 职位描述」（页面导航）
    2. 招聘者信息行（如「曾女士 今日活跃 智乐活 · HR」）
    3. 「竞争力分析…极好」区块
    4. 「BOSS 安全提示…请立即举报」区块
    5. 「工商信息…点击查看地图」区块
    6. 「工作地址…点击查看地图」区块
    7. 残留短语（查看全部 / 点击查看地图 / 查看完整个人竞争力）
    8. 整段重复（【职位描述】【任职要求】【公司介绍】等重复两遍）
    """
    if not text:
        return text

    # 1) 去掉页面头部导航（兼容空格与无空格两种形态）
    text = re.sub(r"^微\s*信\s*扫\s*码\s*分\s*享\s*举\s*报\s*职位描述\s*", "", text.strip())
    text = re.sub(r"^微信扫码分享举报\s*职位描述\s*", "", text)

    # 2) 去掉招聘者信息行（含「今日活跃」「刚刚活跃」「N日内活跃」「· HR」等特征）
    text = re.sub(r"[^\n]*(今日|刚刚|\d+\s*日内)活跃[^\n]*\n?", "", text)
    text = re.sub(r"\s*·\s*(HR|猎头顾问|招聘者|招聘经理|招聘专员|HRBP)[^\n]*\n?", "\n", text)

    # 3) 去掉「竞争力分析」区块
    text = re.sub(r"竞争力分析.*?极好\s*", "", text, flags=re.S)
    # 4) 去掉「BOSS 安全提示」区块
    text = re.sub(r"BOSS\s*安全提示.*?请立即举报\s*", "", text, flags=re.S)
    # 5) 去掉「工商信息」区块（页脚，删到文本结束，不依赖结尾按钮）
    text = re.sub(r"工商信息[\s\S]*\Z", "", text)
    # 6) 去掉「工作地址」区块
    text = re.sub(r"工作地址.*?点击查看地图\s*", "", text, flags=re.S)

    # 7) 去掉固定残留短语
    for noise in ("查看全部", "点击查看地图", "查看完整个人竞争力"):
        text = text.replace(noise, "")

    # 8) 按 BOSS 常见段标题切块；同一标题的重复段落只保留第一次出现的完整版。
    #    BOSS 页面会把标题重复渲染（如"岗位职责岗位职责："或"职位描述【职位描述】"），因此：
    #    ① 裸标题加 (?<!【) 防止在【】内部误切；② 只有标题没内容的碎片块跳过，不占用标题，
    #    保证真实内容块被保留。注意保留第一份——外包风险信号（人力资源服务许可证等）常在首份。
    #    【注意】不要加通用 [一-龿]{2,12}： 锚点——它会把段内子标题切碎，导致去重失效。
    TITLES = (r"(【[^】]+】|(?<!【)关于岗位|(?<!【)岗位职责|(?<!【)任职要求|(?<!【)岗位描述|"
              r"(?<!【)职位描述|(?<!【)加分项|(?<!【)优先考虑|(?<!【)公司介绍|(?<!【)公司简介|"
              r"(?<!【)团队介绍|(?<!【)工作地址|(?<!【)工商信息)")
    blocks = re.split(r"(?=" + TITLES + r")", text)
    seen_titles = set()
    seen_body = set()
    deduped = []
    for blk in blocks:
        key = re.sub(r"\s+", "", blk)   # 归一化空白后作为去重 key
        if not key:
            continue
        tm = re.match(TITLES, key)
        if tm:
            title = tm.group(1)
            if len(key) <= len(title) + 2:
                continue                 # 只有标题没内容的碎片块：跳过，不占用标题
            if title in seen_titles:
                continue                 # 同标题已保留第一份完整版
            seen_titles.add(title)
            deduped.append(blk)
        else:
            if key in seen_body:         # 无标题正文片段（开头导航等）：逐字去重
                continue
            seen_body.add(key)
            deduped.append(blk)
    text = "".join(deduped)

    # 9) 去掉「更多职位/精选职位」推荐区块（在 JD 末尾，删到文件结尾）
    text = re.sub(r"(更多职位|看过该职位的人还看了|精选职位).*\Z", "", text, flags=re.S)
    # 10) 去掉页面底部「城市招聘/热门职位/热门企业 XX招聘」广告区块（到文件结尾）
    text = re.sub(r"城市招聘.*\Z", "", text, flags=re.S)
    # 10) 去掉搜索框残留（如「北京 北京 搜索」）
    text = re.sub(r"\s*北京\s*北京\s*搜索\s*", "\n", text)

    # 11) 压缩多余空行、去首尾空白
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ---- 第 9.5 块：单条岗位评估（供并发调用）----
def evaluate_one(idx, row, cfg, token):
    """评估单条岗位：含超时自动重试（最多 3 次），返回结果字典。供 ThreadPool 并发调用。"""
    print(f"\n[评估] 第 {idx} 条：{row[3]}（并发任务已提交）")
    jd_text = build_jd_text(row)

    last_err = None
    for attempt in range(1, 4):
        try:
            provider = cfg.get("provider", "coze")
            if provider == "coze":
                wf_id = cfg.get("workflow_id")
                if not wf_id:
                    raise RuntimeError("user_need.json 缺少 workflow_id，请填入你 Coze 工作流的 ID")
                raw = call_coze_workflow(token, cfg, jd_text, wf_id)
            elif provider == "ark":
                model = cfg.get("model", "").strip()
                if not model:
                    raise RuntimeError("user_need.json 缺少 model，请填入火山方舟的接入点ID(ep-xxx)或模型ID")
                raw = call_doubao_chat(token, model, jd_text, cfg)
            else:
                raise RuntimeError(f"user_need.json 的 provider 必须是 coze 或 ark，当前是：{provider}")
            parsed = parse_agent_c_output(raw)
            print(f"   ✅ 第 {idx} 条评估完成：总评={parsed.get('总评','')}")
            return {
                "行号": idx,
                "岗位名称": row[3],
                "公司名称": row[7],
                "薪资范围": row[4],
                "岗位URL": row[9],
                "原始返回": raw,
                "总评": parsed.get("总评", ""),
                "总评理由": parsed.get("总评理由", ""),
                "维度": parsed.get("维度", [])
            }
        except TimeoutError as e:
            last_err = e
            print(f"       ⚠️ 第 {idx} 条 Coze 超时（第 {attempt}/3 次尝试）→ 等待 20 秒后自动重试...")
            time.sleep(20)
        except Exception as e:
            # 非超时错误（参数错/限频/解析失败等）：直接记为失败，不重试
            last_err = e
            break
    print(f"       ⚠️ 第 {idx} 条评估失败：{last_err}")
    return {
        "行号": idx,
        "岗位名称": row[3],
        "公司名称": row[7],
        "薪资范围": row[4],
        "岗位URL": row[9],
        "原始返回": f"ERROR: {last_err}",
        "总评": "ERROR",
        "总评理由": str(last_err),
        "维度": []
    }


# ---- 第 10 块：主流程 ----
def main():
    """主控入口：配置 → 抓取 → 评估 → 写报告。"""
    print("=" * 50)
    print("BOSSone 自动筛选主控脚本")
    print("=" * 50)

    # 1) 加载配置
    cfg = load_user_need()
    token = load_api_key(cfg.get("provider", "coze"))

    # 2) 创建报告目录
    REPORT_DIR.mkdir(exist_ok=True)  # 不存在就建，存在也不报错

    # 3) 运行抓取脚本，拿到最新 Excel
    excel_path, cfg = run_scraper(cfg)

    # 4) 读取 Excel
    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    rows = list(ws.iter_rows(min_row=2, values_only=True))  # 从第 2 行开始读数据行
    print(f"[读取] 共 {len(rows)} 条岗位")

    # 5) 并发调 Coze 评估（线程池：COZE_CONCURRENCY 条同时跑，跑完按行号排回顺序）
    max_workers = max(1, min(COZE_CONCURRENCY, len(rows)))
    print(f"\n[评估] 开始并发评估：共 {len(rows)} 条，同时并发 {max_workers} 条（若遇 429 限频，把 COZE_CONCURRENCY 调小）")
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(evaluate_one, idx, row, cfg, token): idx
                   for idx, row in enumerate(rows, start=1)}
        for fut in as_completed(futures):
            try:
                results.append(fut.result())
            except Exception as e:
                # 理论上 evaluate_one 内部已兜底，这里再兜一层
                results.append({
                    "行号": futures[fut],
                    "岗位名称": "",
                    "公司名称": "",
                    "薪资范围": "",
                    "岗位URL": "",
                    "原始返回": f"ERROR: {e}",
                    "总评": "ERROR",
                    "总评理由": str(e),
                    "维度": []
                })
    results.sort(key=lambda r: r["行号"])
    print(f"\n[评估] 并发评估结束：成功 {sum(1 for r in results if r['总评'] != 'ERROR')} 条，失败 {sum(1 for r in results if r['总评'] == 'ERROR')} 条")

    # 6) 写汇总报告 Excel
    report_wb = openpyxl.Workbook()
    report_ws = report_wb.active
    report_ws.title = "筛选结果汇总"

    # 表头：原始信息 + 评估结果 + 6 维明细
    extra_headers = ["原始返回", "总评", "总评理由"]
    dim_names = ["薪资", "位置", "待遇", "岗位相关性", "无外包风险", "技能匹配度"]
    report_ws.append(headers + extra_headers + dim_names)

    # 按行写入
    for r in results:
        # 先把原始 Excel 的列放前面
        base_row = list(rows[r["行号"] - 1])
        # 追加总评、总评理由
        base_row.extend([r.get("原始返回", ""), r["总评"], r["总评理由"]])
        # 追加 6 维匹配结果
        dim_map = {d["评估项"]: d["匹配"] for d in r["维度"]}
        base_row.extend([dim_map.get(d, "") for d in dim_names])
        report_ws.append(base_row)

    # 文件名带时间戳
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = REPORT_DIR / f"BOSSone筛选报告_{cfg['query']}_{ts}.xlsx"
    report_wb.save(str(report_file))

    print("\n" + "=" * 50)
    print(f"[完成] 汇总报告已保存：{report_file}")
    print(f"       总条数：{len(results)}  成功：{sum(1 for r in results if r['总评'] != 'ERROR')}  失败：{sum(1 for r in results if r['总评'] == 'ERROR')}")
    print("=" * 50)


# ---- 第 11 行：程序入口 ----
if __name__ == "__main__":
    main()
