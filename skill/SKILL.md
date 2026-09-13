---
name: patent-sequence-exporter
description: 专利生物序列批量提取工具。输入多个专利号和智慧芽 API Key，批量提取专利中的生物序列并导出 Excel/FASTA/JSON。支持序列类型筛选、长度筛选、Seq ID 范围筛选、断点恢复和进度回调。
measurable_outcome:
  - 批量提取成功率 ≥ 95%（API 可达前提下）
  - 单专利序列提取完整率 100%（分页全量获取）
  - 断点恢复正确率 100%（中断后跳过已完成专利）
compatibility:
  - 需要智慧芽 Bio OpenAPI Key（connect.zhihuiya.com）
  - 需要网络访问智慧芽 API
  - 依赖 requests + openpyxl（无 MCP 工具依赖）
  - exe 版本（v6）为单专利 Web UI，本 skill 版本为批量 CLI
---

# 专利生物序列批量提取工具 v1.0

## 功能说明

从智慧芽 Bio OpenAPI 批量提取专利中的生物序列（DNA/RNA/蛋白质），支持：
- 多专利批量提取（逗号/分号/换行分隔）
- 序列类型筛选（DNA / RNA / Protein / All）
- 序列长度筛选（min ~ max）
- Seq ID 范围筛选（all / 具体编号 / 范围）
- 专利号智能标准化（非标准号 → 标准公开号 + UUID）
- 分页全量获取（PAGE_SIZE=100，自动翻页）
- 断点恢复（batch_state.json 记录已完成专利，中断后跳过）
- NDJSON 即时落盘（每条序列提取后立即 append）
- 进度回调（供 AI 平台实时展示提取进度）
- 多格式导出（Excel / FASTA / JSON / both）

## When to use

以下场景触发本工具：
- 用户给出一个或多个专利号，需要提取其中的生物序列
- 用户需要批量提取 10+ 专利的序列（exe 版本只支持单专利）
- 用户需要按序列类型/长度/Seq ID 范围筛选序列
- 用户需要将专利序列导出为 Excel 或 FASTA 格式
- 与其他 skill 配合时，需要 FASTA 文件作为下游分析输入

## 使用方式

在 AI 平台中直接说出需求，例如：
- "提取 WO2020123456 的全部序列"
- "批量提取 WO2020123456, WO2021123456, US2022001234A1 的蛋白质序列"
- "提取专利 WO2020123456 中 Seq ID 1-50 的 DNA 序列，长度 100-1000"

## 输入变量

### 必填

| 变量 | 类型 | 说明 |
|---|---|---|
| `patents` | 字符串 | 专利号列表，逗号/分号/换行分隔 |
| `api_key` | 字符串 | 智慧芽 Bio OpenAPI Key（sk-xxx 格式） |

### 可选

| 变量 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `seq_id_range` | 字符串 | `all` | `all` / 具体编号(逗号分隔) / 范围(如 `1-50`) |
| `sequence_type` | 字符串 | `null` | `DNA` / `RNA` / `Protein` / `null`(全部) |
| `seq_length_start` | 整数 | `null` | 序列长度下限 |
| `seq_length_end` | 整数 | `null` | 序列长度上限 |
| `output_format` | 枚举 | `excel` | `excel` / `fasta` / `json` / `both` |
| `output_dir` | 路径 | 脚本同目录 `outputs/` | 输出目录 |

## 输出

- 每专利独立 Excel 文件：`{patent_id}_sequences.xlsx`（Summary + Sequences 两个 Sheet）
- 每专利独立 FASTA 文件：`{patent_id}_sequences.fasta`
- 批量汇总 JSON：`batch_summary.json`（每专利提取状态 + 序列统计）
- NDJSON 中间文件：`sequences_all.ndjson`（即时落盘）
- 断点状态文件：`batch_state.json`（记录已完成专利）

## 架构分离

```
LLM 负责
├── 解析用户意图（专利号、筛选条件）
├── 调用 batch_extract.py（通过 run_script）
├── 读取 batch_summary.json 向用户展示结果
└── 生成 Markdown 表格摘要

Python 负责
├── 专利号标准化（patent-search-pn API）
├── 分页全量序列提取（extract/single API）
├── NDJSON 即时落盘 + 断点恢复
├── Excel/FASTA/JSON 导出
└── 进度回调推送
```

## 数据流决策树

```
输入专利号列表
    │
    ├─ Phase 1: 专利号标准化
    │   ├── 单个专利号 → search_patent_pn → 标准公开号 + UUID
    │   ├── 多个专利号 → 逐个标准化 → patent_ids.json
    │   └── 标准化失败 → 记录错误，跳过该专利
    │
    ├─ Phase 2: 批量序列提取
    │   ├── 读取 batch_state.json（如有）→ 跳过已完成专利
    │   ├── 逐专利调用 extract/single API
    │   │   ├── 分页获取（PAGE_SIZE=100，自动翻页）
    │   │   ├── 每页结果即时 append 到 sequences_all.ndjson
    │   │   └── 进度回调（current/total/patent/phase）
    │   └── 更新 batch_state.json（标记该专利完成）
    │
    └─ Phase 3: 合并导出
        ├── 读取 sequences_all.ndjson
        ├── 按专利分组 → 生成 Excel/FASTA
        └── 生成 batch_summary.json
```

## 强制执行层

以下条件为硬性流程，不满足时**必须暂停**：

1. **API Key 必须非空**：空 Key 时停止，提示用户提供智慧芽 API Key
2. **专利号列表必须非空**：空列表时停止，提示用户输入专利号
3. **API 连接必须成功**：首次 API 调用失败时暂停，提示检查 API Key 或网络
4. **输出目录必须可写**：无法创建 output_dir 时暂停

### 暂停规则

遇到以下情况必须暂停并向用户报告：

- **API Key 为空或格式错误** → 暂停，提示用户提供 sk-xxx 格式的 Key
- **专利号标准化全部失败** → 暂停，提示检查专利号格式
- **API 连续 3 次重试失败** → 暂停，提示检查网络或 API 服务状态
- **输出目录不可写** → 暂停，提示更换输出路径

## 边路径定义

| 步骤 | 条件 | 处理 | 后果传播 |
|------|------|------|---------|
| 专利号标准化 | 非标准号匹配到多个结果 | skill 版本取第一个（非交互式） | 可能提取到错误专利的序列 |
| 专利号标准化 | 未找到匹配专利 | 记录 error，跳过该专利 | 该专利无输出，batch_summary 标记 failed |
| API 调用 | HTTP 429 限流 | 指数退避重试 3 次 | 若仍失败 → 该专利标记 failed |
| API 调用 | HTTP 401 认证失败 | 立即暂停，提示 API Key 错误 | 全部专利无法提取 |
| API 调用 | 返回 0 条序列 | 正常完成，输出空文件 | Excel/FASTA 文件为空，summary 标记 0 序列 |
| 分页获取 | total 字段缺失 | 假定仅 1 页 | 可能遗漏后续页序列 |
| NDJSON 落盘 | 磁盘空间不足 | 捕获 IOError，暂停 | 已落盘数据保留，可断点恢复 |
| 断点恢复 | batch_state.json 损坏 | 捕获 JSON 解析错误，从头开始 | 重复提取已完成专利（幂等，无数据损坏） |
| Excel 导出 | PermissionError（文件被占用） | 自动加时间戳重命名 | 文件名变更，用户需注意 |
| FASTA 导出 | 序列为空字符串 | 写入 header 但不写序列行 | FASTA 文件含空记录 |

## 多层穿透分析（DID-002）

**穿透路径 1：API Key 错误 → 全部专利提取失败**
- API Key 无效 → 专利号标准化失败 → 无 UUID → 序列提取无法调用
- **后果**：全部专利标记 failed，无任何输出
- **缓解**：首次 API 调用前做一次 ping 测试（search_patent_pn 用一个已知专利号）

**穿透路径 2：网络不可达 → 标准化 + 提取同时失效**
- 专利号标准化依赖网络 → 失败
- 序列提取依赖网络 → 失败
- **后果**：无任何输出
- **缓解**：启动时探测 API 连通性

### 防御层覆盖矩阵

| 功能 \ 防御层 | API Key 校验 | 专利号标准化 | 分页获取 | 断点恢复 | NDJSON 落盘 |
|--------------|-------------|-------------|---------|---------|------------|
| 认证 | ✅ 唯一 | — | — | — | — |
| 专利定位 | — | ✅ 唯一 | — | — | — |
| 序列获取 | — | — | ✅ 唯一 | — | — |
| 中断恢复 | — | — | — | ✅ 唯一 | — |
| 数据持久化 | — | — | — | — | ✅ 唯一 |

## Phase 契约（前置/后置条件）

| Phase | 前置条件 | 后置条件 |
|-------|---------|----------|
| Phase 1: 专利号标准化 | patents 非空 AND api_key 非空 AND 网络可达 | patent_ids.json 存在 AND len(patent_ids) ≥ 1 |
| Phase 2: 序列提取 | Phase 1 后置满足 AND api_key 有效 | sequences_all.ndjson 行数 ≥ 0 AND batch_state.json 记录所有专利状态 |
| Phase 3: 合并导出 | Phase 2 后置满足 AND output_dir 可写 | 每个成功专利有 Excel/FASTA 文件 AND batch_summary.json 存在 |

**契约违反处理**：任一前置条件不满足时暂停并报错；后置条件不满足时记录 WARN 但继续执行。

**全程不变量**：
- `len(patent_ids) == len(set(patent_ids))`（无重复专利）
- `batch_state.json 中的 completed 列表 ⊆ patent_ids 列表`
- `sequences_all.ndjson 中的 patent_id 字段 ⊆ patent_ids 列表`

## 规模化临界点

| 规模 | 临界点 | 风险 | 缓解措施 |
|------|--------|------|----------|
| 专利数量 | > 20 个 | 单次运行时间过长、API 限流 | 分批处理，每批 10 个，批间间隔 5 秒 |
| 单专利序列数 | > 5000 条 | 分页次数多、内存堆积 | NDJSON 即时落盘，不堆积在内存 |
| API 请求频率 | > 10 请求/分钟 | 可能触发限流 | 专利间间隔 2 秒 |
| batch_state.json | > 100 个专利 | JSON 文件过大 | 考虑改用 SQLite 存储 |

## 分批处理策略

批量提取（>10 个专利）时采用分批处理：

1. **批次大小**：每批 10 个专利
2. **断点恢复**：`batch_state.json` 记录已完成专利列表，中断后从断点继续
3. **即时落盘**：每条序列提取完成后 append 到 `sequences_all.ndjson`
4. **批次间隔**：每批完成后间隔 5 秒，避免 API 限流
5. **最终合并**：所有批次完成后按专利分组导出 Excel/FASTA

## 进度回调接口

```python
def progress_callback(current_patent, total_patents, patent_id, sequences_count, phase):
    # phase: 'standardizing' | 'extracting' | 'exporting' | 'done' | 'error'
    pass
```

AI 平台调用时传入回调函数即可实时获取每个专利的提取进度和结果。

## 参数纠正表

| 常见错误输入 | 正确方式 | 说明 |
|---|---|---|
| 专利号含空格 | 去除空格后传入 | 如 `WO 2020/123456` → `WO2020123456` |
| 专利号含斜杠 | 去除斜杠 | 如 `WO2020/123456` → `WO2020123456` |
| 多专利用换行分隔 | 逗号/分号/换行均可 | 脚本用 re.split(r'[,;\n]+') 分割 |
| API Key 含引号 | 去除引号 | 脚本自动 strip |
| seq_id_range 含空格 | 空格自动转逗号 | 如 `1 2 3` → `1,2,3` |
| sequence_type 小写 | 自动转大写 | 如 `dna` → `DNA` |

## 失效模式枚举（FMEA-001）

| 失效模式 | 触发条件 | 后果 | 检测能力 | RPN |
|---------|---------|------|---------|-----|
| API Key 过期 | Key 有效期结束 | 全部专利失败 | 首次调用 401 | S=9, O=3, D=3, RPN=81 |
| API 服务降级 | 智慧芽服务端问题 | 间歇性失败 | 重试 3 次后报错 | S=5, O=4, D=5, RPN=100 |
| 专利号错误 | 用户输入错误专利号 | 该专利无结果 | 标准化时未找到 | S=3, O=5, D=7, RPN=105 |
| 断点文件损坏 | 异常中断 | 从头开始 | JSON 解析失败 | S=2, O=2, D=8, RPN=32 |
| 磁盘空间不足 | 大量序列写入 | 落盘失败 | IOError 捕获 | S=7, O=1, D=9, RPN=63 |

## 融合钩子

| 衔接 skill | 输入 → 输出 | 典型工作流 |
|------------|------------|-----------|
| bio-sequence-toolkit | FASTA 文件 → 序列比对/引物设计 | 提取序列 → 下游分析 |
| compound-structure-lookup | 专利号 → 化合物信息交叉 | 序列提取 + 化合物查询 |
| biomarker-investigation | 序列 → 文献检索 | 序列 → 标志物关联 |
| patent-sequence-exporter (exe) | exe 单专利 → skill 批量 | exe 提取 → skill 批量二创 |

**调用方式**：AI 平台中先调用本 skill 获取 FASTA 文件，将文件路径作为下游 skill 的输入参数。

## 与 exe 版本的对比

| 维度 | exe v6 | skill v1.0 |
|------|--------|---------|
| 界面 | Web UI（Flask + 浏览器） | CLI / AI 平台调用 |
| 批量专利 | ❌ 单专利 | ✅ 多专利批量 |
| 交互式选择 | ✅ 多结果手动选择 | ❌ 取第一个 |
| 历史记录 | ✅ history/ 文件夹 | ❌ 由 AI 平台管理 |
| 自动打开浏览器 | ✅ 双击 exe | ❌ 不适用 |
| 输出 | Excel / FASTA | Excel / FASTA / JSON / both |
| 断点恢复 | ❌ | ✅ batch_state.json |
| 进度回调 | ❌ | ✅ progress_callback |
| 上下文管理 | N/A | ✅ NDJSON 即时落盘 |
