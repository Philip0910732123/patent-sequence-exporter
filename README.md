# 专利生物序列导出器 (Patent Sequence Exporter)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-Flask-blue.svg)](https://www.python.org/)
[![PatSnap Bio API](https://img.shields.io/badge/PatSnap-Bio%20API-2EAD33.svg)](https://open.zhihuiya.com/)
[![Release v6.0](https://img.shields.io/badge/Release-v6.0-blue.svg)](https://github.com/Philip0910732123/patent-sequence-exporter/releases/tag/v6.0)
[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2B-blue.svg)](https://www.microsoft.com/)

> 通过智慧芽 Bio OpenAPI 自动提取专利中的生物序列，支持 Excel (.xlsx) 和 FASTA 格式导出。本地运行，API Key 不上传。

## 功能概览

- **自动提取**：输入专利号 + Seq ID 范围，自动调用智慧芽 Bio OpenAPI 提取序列
- **多格式导出**：Excel (.xlsx) 表格预览 / FASTA 序列文件
- **高级过滤**：按序列类型（蛋白质/核酸）和长度范围筛选
- **历史记录**：每次提取结果自动保存到 exe 同目录 `history/` 文件夹，支持查看、重新下载、删除
- **自动打开浏览器**：双击 exe 后自动打开浏览器，无需手动输入地址
- **本地运行**：Flask Web 应用，所有数据在本地处理，API Key 不上传到任何第三方服务器

## 快速开始

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

## 使用示例

```
专利号：WO2004050828A2
Seq ID 范围：1-50
输出格式：Excel (.xlsx)
```

提取完成后，文件保存在 exe 同目录下的 `outputs/` 文件夹中。

## 环境要求

- **操作系统**：Windows 10/11（64 位）
- **网络**：需要互联网连接（调用智慧芽 API）
- **浏览器**：Chrome / Edge / Firefox 等现代浏览器
- **无需安装**：exe 内置 Python + Flask 运行时，双击即用

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

- [patsnap-patent-skill-suite](https://github.com/Philip0910732123/patsnap-patent-skill-suite) — 专利撰写全流程 Skill 套件
- [patent-oa-analysis-skills](https://github.com/Philip0910732123/patent-oa-analysis-skills) — 专利审查意见分析技能


---

## 💬 交流与合作

如需技术交流、问题反馈或商业合作，欢迎扫描下方微信二维码联系作者。

<p align="center">
  <img src="wechat_qr.jpg" width="200" alt="作者微信二维码" />
</p>

> 添加时请注明来自 GitHub 仓库，我会优先通过。
