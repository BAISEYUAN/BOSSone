# BOSSone 智能筛岗助手

一个「填需求 → 自动抓取 BOSS 直聘岗位 → AI 逐条评估 6 维匹配度 → 输出投递建议 + 报告」的本地桌面工具。

- 纯本地运行：数据只在你自己的电脑上，不上传任何服务器
- AI 评估：对接火山方舟（豆包）大模型，需自备 API Key
- 输出：每条岗位给出「建议投 / 谨慎 / 不建议」+ 六维打分理由，并自动生成 Excel 和 Markdown 报告

> ⚠️ **免责声明**：本工具通过浏览器自动化抓取 BOSS 直聘公开页面，仅供个人学习与研究使用，禁止用于商业用途或大规模爬取。使用本工具产生的任何后果（账号风控、封禁、法律法规问题等）由使用者自行承担。请遵守 BOSS 直聘用户协议及当地法律法规。

---

## 快速开始（Windows，即开即用 ✅）

**1. 下载** —— 到右侧 **Releases** 页，下载最新的 `BOSSone_win_v*.*.zip` 压缩包

**2. 解压** —— 解压到任意文件夹，双击 `BOSSone.exe` 即可运行（会弹出黑色命令窗口，**别关它**）

**3. 准备两个东西（都是你自己的账号，约 5 分钟）**

| 需要的 | 去哪里拿 | 放哪里 |
|---|---|---|
| 火山方舟 API Key | [火山方舟控制台](https://console.volcengine.com/ark) → API Key 管理 → 创建 | 程序文件夹新建 `ark_api_key.txt` 粘贴进去 |
| 模型接入点 ID | 火山方舟 → 在线推理 → 创建推理接入点（如 `ep-xxxxxxxx-xxxxx`） | 用记事本改 `user_need.json` 里的 `"model"` 字段 |

**4. 使用流程**
1. 双击 `BOSSone.exe`，自动弹出调试版 Edge + 网页 `http://127.0.0.1:5000`
2. 在弹出的 Edge 里扫码登录 BOSS 直聘，搜索目标岗位，停在**搜索结果列表页**
3. 回到网页填需求（岗位、薪资下限、地点、待遇要求、抓取条数）→ 点「开始智能筛选」
4. 1-2 分钟出结果，报告保存在程序文件夹的「筛选报告」目录

> Windows 10/11 自带 Edge；首次运行若被杀毒软件误报，请选择「保留/信任」（PyInstaller 打包程序的特征，具体见包内《使用说明.md》）。

---

## 从源码运行（开发者 / Mac 用户）

```bash
# 需要 Python 3.10+
pip install -r requirements.txt   # 见下方依赖清单
python launcher.py
```

主要依赖：`fastapi`、`uvicorn`、`requests`、`pandas`、`openpyxl`、`markdown`、`pywin32`（Windows 专用，Mac 自动跳过）、`websocket-client`

Mac 打包方法见包内《使用说明.md》第五节，或参考 `build_win.bat`（Windows 打包脚本）。

## 项目结构

```
app.py                      网页后端（FastAPI）
bossone_controller.py       筛选引擎（调用 AI 评估）
V15B_...py                  抓取模块（浏览器 CDP 自动化）
index.html                  网页界面
edgeutil.py                 Edge 调试版启动工具
launcher.py                 一键启动入口
user_need.json              用户需求配置模板
build_win.bat               Windows 打包脚本
```

## License

本项目仅供学习交流，无开源 License 授权，请勿用于任何形式的商业用途。