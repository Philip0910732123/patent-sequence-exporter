# 专利生物序列导出器 (Patent Sequence Exporter)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-Flask-blue.svg)](https://www.python.org/)
[![PatSnap Bio API](https://img.shields.io/badge/PatSnap-Bio%20API-2EAD33.svg)](https://open.zhihuiya.com/)
[![Release v6.0](https://img.shields.io/badge/Release-v6.0-blue.svg)](https://github.com/Philip0910732123/patent-sequence-exporter/releases/tag/v6.0)
[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2B-blue.svg)](https://www.microsoft.com/)

> 通过智慧芽 Bio OpenAPI 自动提取专利中的生物序列，支持 Excel (.xlsx) 和 FASTA 格式导出。两种使用方式，按需选择。

## 两种使用方式

| 方式 | 适合谁 | 说明 |
|------|--------|------|
| **📦 桌面 exe** | 所有用户 | 双击即用，无需 Python，无需 LLM，浏览器内操作（单专利） |
| **🤖 AI Skill** | AI 平台开发者 | 导入 AI Coding 平台，支持批量多专利提取，可嵌入工作流，支持二创 |

### 桌面 exe 版本（免 LLM）

- 双击 `patent-sequence-exporter-v6.exe` → 浏览器自动打开 → 输入专利号 + API Key → 提取 → 下载 Excel/FASTA
- 无需安装任何环境，仅需要智慧芽 API Key
- 下载地址：[最新 Release](https://github.com/Philip0910732123/patent-sequence-exporter/releases)
- 注意：exe 版本为**单专利提取**，如需批量多专利提取请使用 AI Skill 版本

### AI Skill 版本（需 LLM）

- 将 `skill/` 目录导入 AI Coding 平台即可作为 Skill 调用
- LLM 负责解析用户意图（专利号、筛选条件）、调用脚本、展示结果
- Python 脚本负责专利号标准化、分页全量序列提取、Excel/FASTA/JSON 导出
- **支持批量多专利提取**（exe 版本不支持），带断点恢复和进度回调
- 消耗大模型 token，由用户按需选择
- 其他开发者可 fork 做二创，嵌入自己的工作流

---

## 功能概览

- **自动提取**：输入专利号 + Seq ID 范围，自动调用智慧芽 Bio OpenAPI 提取序列
- **多格式导出**：Excel (.xlsx) 表格预览 / FASTA 序列文件 / JSON
- **高级过滤**：按序列类型（蛋白质/核酸）和长度范围筛选
- **历史记录**：每次提取结果自动保存到 exe 同目录 `history/` 文件夹，支持查看、重新下载、删除
- **自动打开浏览器**：双击 exe 后自动打开浏览器，无需手动输入地址
- **本地运行**：Flask Web 应用，所有数据在本地处理，API Key 不上传到任何第三方服务器

## 桌面 exe 使用方式

### 第一步：获取 API Key

访问 [智慧芽开放平台](https://open.zhihuiya.com) → 注册/登录 → 获取 API Key

### 第二步：启动程序

双击 `patent-sequence-exporter-v6.exe`，会弹出一个命令行窗口（**不要关闭**），程序会自动打开浏览器。

```
专利生物序列导出器 - Web 版 (v6)
输出目录：...\outputs
浏览器已自动打开，如未打开请手动访问 http://127.0.0.1:5000
按 Ctrl+C 退出
 * Running on http://127.0.0.1:5000
```

### 第三步：填写表单

| 字段 | 必填 | 说明 |
|------|:----:|------|
| **API Key** | ✅ | 粘贴智慧芽 API Key（支持直接输入或从文件读取） |
| **专利号** | ✅ | 必须是完整公开号，带字母后缀（如 `WO2004050828A2`、`US20210000123A1`、`CN11205304B1`）；可点击"核对专利号"按钮验证 |
| **Seq ID 范围** | ✅ | 如 `1-50`、`1,5,10,23,24`，也支持空格分隔（`1 2 3` 自动转为 `1,2,3`） |
| **输出格式** | — | 下拉选择 `Excel (.xlsx)` 或 `FASTA` |
| **序列类型过滤** | — | 可选，下拉选择"全部"/"蛋白质"/"核酸" |
| **长度范围** | — | 可选，设置起始和结束长度过滤 |

点击 **"开始提取序列"** 按钮即可。

### 第四步：查看历史记录

提取完成后，文件保存在 exe 同目录下的 `outputs/` 文件夹中。同时，每次提取的结果会自动保存到 `history/` 文件夹（带时间戳命名），点击页面上的 **"📋 历史记录"** 按钮可查看、重新下载或删除之前的提取记录。

## AI Skill 使用方式

### 安装

将 `skill/` 目录导入 AI Coding 平台（如 Eureka）：

```
skill/
├── SKILL.md                    # Skill 文档（skill-auditor v3.4 认证）
├── skill.manifest.json          # 依赖声明
└── scripts/
    ├── batch_extract.py         # 批量提取引擎（677行）
    ├── extract_sequences.py      # 单专利提取引擎
    ├── selftest.py              # 17 项自检
    └── main.py                  # CLI 入口
```

### 调用示例

在 AI 平台中直接说出需求：
- "提取 WO2020123456 的全部序列"
- "批量提取 WO2020123456, WO2021123456, US2022001234A1 的蛋白质序列"
- "提取专利 WO2020123456 中 Seq ID 1-50 的 DNA 序列，长度 100-1000"

### Skill 依赖

```
requests    # 智慧芽 API 调用
openpyxl     # Excel 导出
```

### Skill 与 exe 的能力对比

| 能力 | 桌面 exe | AI Skill |
|------|---------|---------|
| 单专利提取 | ✅ | ✅ |
| **批量多专利提取** | ❌ | ✅ 逗号/分号/换行分隔 |
| Excel 导出 | ✅ | ✅ |
| FASTA 导出 | ✅ | ✅ |
| JSON 导出 | ❌ | ✅ |
| 历史记录 | ✅ history/ | ❌ 由 AI 平台管理 |
| 断点恢复 | ❌ | ✅ batch_state.json |
| 进度回调 | ❌ | ✅ progress_callback |
| NDJSON 即时落盘 | ❌ | ✅ |
| 二创定制 | 改源码 | ✅ fork skill |
| token 消耗 | 免费 | 消耗大模型 token |
| 交互式专利号选择 | ✅ 多结果手动选择 | ❌ 取第一个（非交互式） |

## 环境要求

- **操作系统**：Windows 10/11（64 位）
- **网络**：需要互联网连接（调用智慧芽 API）
- **浏览器**：Chrome / Edge / Firefox 等现代浏览器
- **无需安装**：exe 内置 Python + Flask 运行时，双击即用

## 仓库结构

```
patent-sequence-exporter/
├── app.py                      # exe 后端（Flask）
├── templates/
│   └── index.html              # exe 前端
├── build_exe.py                # PyInstaller 构建脚本
├── requirements.txt            # exe 依赖
├── README.md
├── LICENSE
└── skill/                      # AI Skill 版本
    ├── SKILL.md                # Skill 文档（skill-auditor v3.4 认证）
    ├── skill.manifest.json
    └── scripts/
        ├── batch_extract.py     # 批量提取引擎
        ├── extract_sequences.py # 单专利提取引擎
        ├── selftest.py          # 自检脚本
        └── main.py             # CLI 入口
```

## 关于 API Key

- API Key 在本地使用，不上传到任何第三方服务器
- 获取地址：[https://open.zhihuiya.com](https://open.zhihuiya.com)
- 如果 Key 无效，程序会提示鉴权失败

## 常见问题

| 问题 | 解决方案 |
|------|----------|
| 命令行窗口不小心关了 | 重新双击 exe 即可 |
| 浏览器没有自动打开 | 检查命令行窗口是否还在运行，或手动访问 http://127.0.0.1:5000 |
| 专利号提示格式错误 | 确保带完整后缀（如 `A1`、`B1`、`A2`），不要只填数字 |
| 提取报错"鉴权失败" | 检查 API Key 是否正确，是否已过期 |
| 提取结果为空 | 检查 Seq ID 范围是否在该专利中存在 |
| 历史记录不见了 | 历史记录保存在 exe 同目录的 `history/` 文件夹中，移动 exe 时请一并移动 |

## 技术栈

- **后端**：Python + Flask（本地 Web 服务）
- **API**：智慧芽 Bio OpenAPI
- **前端**：HTML + JavaScript（表单交互）
- **导出**：openpyxl（Excel）+ 原生 FASTA 写入
- **历史记录**：JSON 持久化 + 时间戳命名 + 文件名安全校验

## 变更记录

### v6.0 (2026-09-13)

#### 新增功能

- **历史记录**：每次提取成功后自动保存结果到 `history/` 文件夹，文件名格式 `history_YYYYMMDD_HHMMSS.json`，包含完整序列数据 + 元信息（时间戳、专利号、Seq ID 范围、类型分布等）
- **历史记录管理**：新增 5 个 API 路由（列表 / 详情 / 下载 Excel·FASTA·JSON / 删除），前端新增历史记录弹窗，支持查看详情、重新下载、删除
- **自动打开浏览器**：双击 exe 后延迟 1.5 秒自动打开浏览器（`threading.Timer` + `webbrowser`），无需手动输入地址

#### 优化

- exe 体积从 32.3 MB 增至 43.3 MB（新增 `openpyxl` 数据收集 + `threading`/`webbrowser` 模块）
- 命令行窗口新增"按 Ctrl+C 退出"提示
- 文件名安全校验：历史记录文件名正则校验 `^history_\d{8}_\d{6}\.json$`，防止路径穿越攻击

### v5.0 (2026-09-12)

- 初始版本
- Flask Web 应用，双击 exe 启动本地服务
- 支持专利号标准化（非标准专利号智能查找）
- Excel (.xlsx) / FASTA 双格式导出
- 按序列类型和长度范围过滤

## License

MIT © Philip0910732123

## 相关项目

- [compound-structure-lookup](https://github.com/Philip0910732123/compound-structure-lookup) — 化合物结构信息批量查询工具（exe + AI Skill）

---

## 💬 交流与合作

如需技术交流、问题反馈或商业合作，欢迎扫描下方微信二维码联系作者。

<p align="center">
  <img src="wechat_qr.jpg" width="200" alt="作者微信二维码" />
</p>

> 添加时请注明来自 GitHub 仓库，我会优先通过。
