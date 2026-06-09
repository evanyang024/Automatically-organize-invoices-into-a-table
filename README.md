# 发票自动处理工具

自动识别中国增值税电子发票 PDF，提取关键字段并写入 Excel 台账。

---

## 功能

把发票 PDF 丢进 `input/`，自动完成：**文本提取 → 字段解析 → 智能重命名 → 去重 → 写入 Excel → 归档副本**

| 功能 | 说明 |
|------|------|
| PDF 识别 | 支持横排、竖排、散列、铁路客票 4 种格式 |
| 金额解析 | 价税合计 = max(¥金额)，三级回退，数字类型可求和 |
| 重命名 | `YYYY-MM-DD-项目-金额.pdf`，多明细合并为"等N项" |
| Excel | 3 列（货物名称 / 价税合计 / 开票日期），无边框，只追加 |
| 去重 | 按发票号码，独立跟踪文件 |
| 归档 | input/ 原文件不动，archive/ 存副本 |
| 对照 | 整理后自动核对 input/ ↔ archive/ |

---

## 快速开始

**环境：** Python 3.10+，Windows/macOS/Linux

```bash
git clone https://github.com/evanyang024/Automatically-organize-invoices-into-a-table.git
cd Automatically-organize-invoices-into-a-table
mkdir input output archive
pip install pymupdf openpyxl watchdog pyyaml

# 放入发票 PDF 到 input/，然后：
python run.py --once

# 或持续监控：
python run.py
```

---

## 项目结构

```
├── run.py                    # 主入口
├── config.yaml               # 配置
├── input/                    # 放入发票
├── output/                   # Excel 台账
├── archive/                  # 归档副本
└── invoice_processor/
    ├── extractor.py          # PDF 提取
    ├── parser.py             # 字段解析（v7，含铁路客票）
    ├── renamer.py            # 重命名
    ├── writer.py             # Excel 写入
    └── watcher.py            # 文件夹监控
```

---

## 完整流程

```
放入 input/ → 提取文本 → 解析字段(正则) → 重命名(日期-项目-金额)
→ 去重(发票号) → 写 Excel(追加/数字/无边框) → 归档(copy 副本) → 对照检查
```

### 金额规则

发票上 ¥ 数字位置不固定，**价税合计永远取最大的 ¥ 金额**。

### Excel 格式

| 货物或服务名称 | 价税合计 | 开票日期 |
|------|------|------|
| 云服务费 | 71.00 | 06月03日 |
| 铁路客运 | 124.00 | 06月03日 |

- 价税合计为数字类型，可直接 SUM
- 无边框
- 日期格式 MM月DD日
- 多明细合并一行

---

## 不同用户

| 用户 | 使用方式 |
|------|----------|
| Hermes Agent | 复制 `skills/` 到 `~/.hermes/skills/productivity/` |
| Claude Code / Codex | 直接用项目代码，纯 Python 独立运行 |
| 通用 | 6 个核心文件，跨平台可用 |

---

## 平台

| 功能 | Windows | macOS | Linux |
|------|:--:|:--:|:--:|
| PDF 提取 | ✅ | ✅ | ✅ |
| 字段解析 | ✅ | ✅ | ✅ |
| 重命名 | ✅ | ✅ | ✅ |
| Excel 写入 | ✅ | ✅ | ✅ |
| 文件夹监控 | ✅ | ✅ | ✅ |
| Excel 锁自动处理 | ✅ | ❌ | ❌ |

---

## 自定义

修改 `config.yaml`：增减列、关闭重命名/去重/归档。修改 `parser.py` 适配新发票格式。

## 铁律

1. 不删原文件  2. 不改旧数据  3. 金额必须正确  4. 加新不丢旧  5. 整理完对照
