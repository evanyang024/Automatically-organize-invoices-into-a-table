---
name: invoice-auto-processor
description: "Use when the user needs to auto-process Chinese VAT e-invoice PDFs — extract fields and write to Excel ledger with folder watching, smart renaming, dedup, and archiving."
version: 5.0.0
author: Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [invoice, ocr, pdf, excel, automation, chinese-vat]
    related_skills: [skill-creator]
---

# 发票自动处理工具 (Invoice Auto Processor)

自动识别中国增值税电子发票 PDF，提取字段并写入 Excel 台账。支持文件夹持续监控、智能重命名、发票去重、自动归档。

## Overview

处理中国增值税电子发票（PDF 格式）的完整自动化流水线。从 PDF 文本提取开始，经过智能重命名预处理（格式：`YYYY-MM-DD-项目-金额.pdf`），正则匹配解析关键字段，最后写入 Excel 台账并归档。支持持续监控模式（watchdog）和单次批处理模式。

## 核心原则

**金额必须正确。** 发票格式千差万别（横排、竖排、散列文本），正则不可能一次写对。处理流程：先提取 PDF 原文 → 当前正则尝试解析 → 金额不对则主动 debug（打印原文、¥位置、金额上下文）→ 根据实际文本调正则（多策略回退）→ 重跑直到金额正确 → 更新 `parser.py`。发现新格式时主动适配，不等用户纠错。

## 铁律（不可违反）

| # | 原则 | 说明 |
|---|------|------|
| 1 | **不删原文件** | `input/` 文件只读不删，`archive/` 用 `copy2` 不用 `move` |
| 2 | **不改旧数据** | Excel 只追加行，不覆盖已有；`archive/` 不重复归档 |
| 3 | **改动前告知** | 修改解析器前用 `clarify` 说明 (1)改什么 (2)为什么 (3)影响哪些 (4)后果，用户同意再执行 |
| 4 | **加新不丢旧** | 每次改 parser 后回测全部已成功发票，验证金额准确性 |
| 5 | **金额必须正确** | 价税合计 = max(¥金额)，遇新格式主动 debug 适配 |

## When to Use

- 用户提到"发票"、"增值税发票"、"电子发票"、"invoice"
- 用户要求"处理发票"、"提取发票字段"、"发票自动录入"、"发票台账"
- 用户引用项目路径 `E:/桌面E/待整理发票` 或同名项目
- 用户说"把发票放进表格"、"发票整理"

Don't use for:
- 国际商业发票（英文 Invoice，格式不同）
- 图片扫描件（拍照发票，非 PDF 电子发票）
- 纸质发票手写 OCR 识别

## Workflow

### 部署（一次性）

```bash
# 1. 克隆仓库
git clone https://github.com/evanyang024/Automatically-organize-invoices-into-a-table.git
cd Automatically-organize-invoices-into-a-table

# 2. 安装依赖
pip install pymupdf openpyxl watchdog pyyaml

# 3. 配置（可选，config.yaml 已有默认值）
# 完成。程序首次运行会自动创建 input/ output/ archive/ 三个文件夹。
```

### 使用

**第 1 步：放入发票**

把增值税电子发票 PDF 拖进 `input/` 文件夹。

**第 2 步：运行程序**

```bash
python run.py --once    # 处理完即退出
python run.py           # 持续监控，有新发票自动处理
```

**第 3 步：程序自动完成**

```
input/ 中 PDF 被逐个处理：
  ├─ 提取文本 (pymupdf)
  ├─ 解析字段 (正则匹配)
  │   ├─ 价税合计 = max(所有 ¥ 金额)
  │   ├─ 开票日期 = YYYY年MM月DD日
  │   ├─ 货物名称 = 星号或表格中提取
  │   └─ 发票号码（去重用）
  ├─ 智能重命名 (YYYY-MM-DD-项目-金额.pdf)
  ├─ 去重检查（按发票号码，已处理则跳过台账）
  ├─ 写入 Excel（只追加，不覆盖旧数据）
  │   ├─ 价税合计 = 数字类型，可 SUM
  │   ├─ 开票日期 = MM月DD日（去年份）
  │   └─ 无边框，多明细合并为一行
  └─ 归档副本到 archive/（input/ 原文件保留不动）
```

**第 4 步：查看结果**

打开 `output/发票台账.xlsx`，表头三列：

| 货物或服务名称 | 价税合计 | 开票日期 |
|------|------|------|
| 住宿费 | 392.00 | 05月14日 |
| 铁路客运 | 124.00 | 06月03日 |

**第 5 步：对照检查**

程序自动核对 `input/` 和 `archive/` 按发票号码是否一一对应，如有差异会报告。

## File Behavior

- **`input/`**: 原始文件保留不动，不会被删除或移动
- **`archive/`**: 以智能重命名后的文件名存入副本。即使去重跳过了台账写入（发票号已存在），仍会归档到 archive/
- **`output/`**: Excel 台账（只追加不覆盖）+ `.processed_invoices.json`（去重跟踪）

```
放入:    input/澈界域名71元.pdf
              │  提取 → 解析 → 写入 Excel
              │  计算新名: 2026-06-03-云服务费-71.00.pdf
              │
              ▼
保留:    input/澈界域名71元.pdf                        ← 原封不动
归档:    archive/2026-06-03-云服务费-71.00.pdf         ← 重命名副本
输出:    output/发票台账.xlsx
```

### Excel 列（默认 3 列）

| 货物或服务名称 | 价税合计 | 开票日期 |
|---|---|---|
| 云服务费 | 71.00 | 06月03日 |

- 价税合计为**数字类型**（float），可直接 SUM 求和
- **无边框**，简洁干净
- 开票日期格式 `MM月DD日`（去年份）
- `config.yaml` 中 `columns` 列表控制输出列

### 扩展/修改

- 调整解析规则 → 修改 `invoice_processor/parser.py`（含 v4 竖排文本适配）
- 增减 Excel 列 → 修改 `config.yaml` columns（`writer.py` 用映射表，列名匹配即生效）
- 关闭重命名 → `config.yaml` 设 `rename_enabled: false`

## Common Pitfalls

1. **pymupdf 装不上** — Windows 需 Python 3.10+；Linux 先 `apt install libmupdf-dev`。
2. **解析字段为空** — 查看 `invoice_processor.log`，根据实际发票文本在 `parser.py` 中调正则。
3. **中文乱码** — `config.yaml` 和所有 .py 文件以 UTF-8 保存。
4. **重命名失败 / 同名冲突** — 非法字符自动过滤；同名自动加时间戳。归档用 `copy2`，原文件保留。
5. **重复发票被重复写入** — `skip_duplicates` 依赖发票号码去重，确认 `config.yaml` 中开启。
6. **价税合计取错** — 很多发票 ¥金额列是 `¥价税合计, ¥不含税, ¥税额`。规则：**价税合计 = max(¥金额)**，不能取最后/第一个。v5 已内置此逻辑。
7. **改解析器时破坏旧格式** — 修改正则后必须用全部已成功发票回测。已验证 3 种格式：(A)完整跨行明细 (B)简单星号 (C)竖排散列。加新规则不丢旧规则，每次改动后跑全量 `--once`。
8. **修改前必须告知用户** — 解析器改动方案先用 `clarify` 告知：(1)改什么 (2)为什么 (3)影响哪些已成功发票 (4)预估后果。用户同意后才执行。
9. **模块缓存导致修改不生效** — `watcher.py` 在模块顶层 `from .parser import parse_invoice`，修改 `parser.py` 后必须重启进程（子进程或清缓存）才能生效。在同一个 Python 进程中反复 import 不会更新。
10. **Excel 文件被锁导致追加而非重建** — 运行时如果 Excel 正被打开，`output/发票台账.xlsx` 无法被写入。此时用 `pywin32` COM 自动化：连接 Excel → 找到台账 → 保存 → 关闭 Excel → 再跑程序。**禁止 `taskkill /F`**（会丢失用户的修改）。pywin32 已安装可用。

11. **watcher 静态导入导致 parser 更新不生效** — `watcher.py` 中 `from .parser import parse_invoice` 在模块加载时绑定函数引用，即使重载 parser 模块，watcher 中的旧引用仍指向旧函数。修复：改为 `from . import parser as _parser`，调用 `_parser.parse_invoice()` 实现动态查找。

13. **¥ 独立成行导致金额为空** — 部分发票 ¥ 符号独占一行，数字在别处（如通信费发票）。v6 已加入三级回退：同行¥ → 跨行¥ → (小写)附近 → 最大数字。遇此格式自动适配，不必重写解析器。

14. **renamer 缺年份导致文件名缺日期** — 日期去年份必须在 writer（写 Excel 时），不能在 parser（提取时）。parser 返回完整 `YYYY年MM月DD日`，renamer 才能拼出 `YYYY-MM-DD-项目-金额.pdf`。若在 parser 层砍年份，归档文件名会丢失日期前缀。

15. **多明细发票拆成多行** — 一张发票多条明细时，台账只写一行：用第一项名称 + "等N项"，金额为价税合计。不拆行，不重复增加数据行。此规则在 `writer.py` 的 `append_invoice_rows` 中实现。

16. **金额存为文本导致 Excel 无法计算** — `writer.py` 写入时转数字：`float(val.replace(",", ""))`，并设 `number_format = '#,##0.00'`。价税合计列可正常求和。

17. **表格有边框** — 用户偏好无边框表格。`writer.py` 不设 `cell.border`，不导入 `Border`/`Side`。

18. **火车票身份证号被误识别为货物名** — 12306 电子客票含身份证号如 `****1819`，标准正则 `*分类*名称` 会匹配到 `****1819` 并提取 `1819` 作为货物名。v7 已加入「铁路电子客票」检测，文本含此特征时跳过标准解析，走 `_parse_train_ticket()` 专用函数。

19. **火车票乘车日期≠开票日期** — 12306 票面有乘车日期（如 `2025年12月07日`）和开票日期（`开票日期:2026年06月03日`）两个日期。`_parse_train_ticket` 明确提取「开票日期」字段，避免取到乘车日期。

20. **MIME 邮件头编码不统一** — 如果未来从邮件下载发票附件，邮件主题可能用 GB2312/GBK 编码（如 `=?gb2312?B?...?=`）。`email.header.decode_header()` 返回 `(bytes, charset)` 元组，必须按声明的 charset 解码，不能用 UTF-8 强解。`safe_decode_header()` 模式可参考 `references/encoding-pitfalls.md`。

21. **⚠️ patch 工具会双重转义反斜杠** — 用 `patch` 工具编辑包含大量 `\d` `\s` `\n` 等正则的 Python 文件时，反斜杠可能被二次转义（`\d` → `\\d`），导致正则失效。症状：函数逻辑正确但正则全不匹配。修复：改用 `execute_code` 直接读写文件，或用 `write_file` 重写整个文件。

## Parser Evolution Summary

| 版本 | 解决的问题 | 关键规则 |
|------|-----------|----------|
| v1-v2 | 横排标准发票 | 字段名值同行 |
| v3-v4 | 竖排散列文本、星号跨行 | `\s*` 跨行匹配 |
| v5-v6 | ¥独立成行（金额在别处）、三级回退 | ¥同行→¥跨行→(小写)→最大数字 |
| v7 | 铁路电子客票（12306 火车票） | 检测「铁路电子客票」→专用解析函数，项目名固定「铁路客运」，开票日期≠乘车日期 |
| 铁律 | 不删原文件、不改旧数据、改动前告知、回测全量 | 记入 .processed_invoices.json 去重 |

### 金额规则（v5 核心）

发票的 ¥金额列顺序不固定，**价税合计永远取最大的 ¥金额**。日期输出格式为 `MM月DD日`（去年份），通过 `re.sub` 在 parser 中处理。

### 去重机制

独立跟踪文件 `output/.processed_invoices.json`（JSON 数组），按发票号码去重，不依赖台账列结构。每处理一张写入一次。

## Verification Checklist

- [ ] 重跑前关闭 Excel（否则文件被锁，追加旧数据）
- [ ] 子进程运行：修改代码后务必用 `subprocess.run` 测试，避免模块缓存
- [ ] 金额验证：抽检 2-3 张，确认价税合计 = max(¥金额)
- [ ] input/ 原文件不动：处理后原文件仍在
- [ ] archive/ 重命名格式：`YYYY-MM-DD-项目-金额.pdf`
- [ ] 去重验证：同号发票第二次跳过
- [ ] 回测全量已成功发票
- [ ] 🔍 input/ ↔ archive/ 按发票号码一一对照，报告差异

## Supporting Files

| 文件 | 内容 |
|------|------|
| `references/parser.py` | 发票字段解析 (v6: ¥三级回退 + 竖排 + 星号跨行) |
| `references/extractor.py` | PDF 文本提取 (pymupdf) |
| `references/renamer.py` | 智能重命名逻辑 |
| `references/writer.py` | Excel 写入 (数字类型、无边框、多明细合并) |
| `references/watcher.py` | 文件夹监控 (动态导入 parser、去重、归档) |
| `references/config.yaml` | 配置文件模板 (3 列) |
| `scripts/save_close_excel.py` | Excel 锁住时保存+关闭 (pywin32 COM) |

其他人 clone 仓库后，复制 `skills/invoice-auto-processor/` 到 `~/.hermes/skills/productivity/` 即可使用。
| `references/encoding-pitfalls.md` | 🆕 MIME 邮件头编码坑点（GB2312/GBK） |
