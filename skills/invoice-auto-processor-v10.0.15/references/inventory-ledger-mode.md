# 进销存台账模式（inventory ledger mode）

Provenance: absorbed from the former standalone `invoice-inventory` skill (archived 2026-11-12).
Use this when the invoice target is an **进销存 (purchase/sales/inventory) workbook** rather
than a flat expense ledger. The parsing rules, denoising, and dedup keys are identical to the
parent skill — only the *write target and matching logic* differ.

## When this mode applies

- Target workbook is `/Users/yangtao/Desktop/产品进销存管理.xlsx` (or a similar 3-sheet inventory book)
- User says 进销存 / 库存管理 / 录采购 / 录销售 / 发票记账
- Deliverable is stock movement, not just an expense row

## Workbook structure (3 sheets)

### Sheet 1 基础档案 (product master)
| Col | Field | Notes |
|-----|-------|-------|
| A | 物品编码 | unique numeric id |
| B | 名称 | product name |
| C | 规格型号 | e.g. `0.10-10mg/l，50支/盒` |
| D | 货号 | vendor part no |
| E | 品牌 | brand |
| F | RMB 单价 | purchase unit price |
| G | 采购厂家 | supplier |

### Sheet 2 出入库流水明细 (unified transaction log)
| Col | Field |
|-----|-------|
| A | 日期 (YYYY-MM-DD) |
| B | 单据类型 (`入库` / `出库`) |
| C | 物品编码 |
| D | 物品名称 |
| E | 规格型号 |
| F | 变动类型 (`采购` / `销售`) |
| G | 数量 |
| H | 关联合同号/备注 |
| I | 金额 |
| J | 经办人 (optional) |

**Key design:** inbound and outbound share ONE sheet. A purchase of N products on one
date = one date, N rows. Never create per-product sheets.

### Sheet 3 库存看板 (auto-computed)
`=SUMIFS(出入库流水明细!G:G, 出入库流水明细!C:C, A{row}, 出入库流水明细!B:B, "入库")`
minus the same for `"出库"`. Never write hand-computed stock numbers — write the formula
so the board stays live.

## Write workflow (10 steps)

1. Ensure the workbook exists; `scripts/init_workbook.py` creates a fresh one.
2. Read the invoice (PDF or image).
3. Extract fields (same parser rules as parent SKILL.md).
4. Decide if it's in scope (采购 vs 销售 vs out-of-scope 劳务/服务).
5. Classify invoice type.
6. Match to an existing 物品编码 using the 4-dimensional rule below.
7. **Per-row pre-write check (mandatory, do not skip).**
8. Append to 出入库流水明细.
9. Report the result (what was written, what was skipped and why).
10. Archive the source file into `录入成功/采购发票/`, `录入成功/销售发票/`, or `被过滤发票/`.

`scripts/add_record.py` implements steps 8 + product registration:

```bash
# add a movement
python3 add_record.py add --code 3 --date "2025-06-01" --direction 入库 \
    --change-type 采购 --product 氨氮试剂 --spec "0.10-10mg/l" \
    --unit 盒 --quantity 10 --amount 3990 --invoice FP20250601-001

# register a new product (auto-assigns code)
python3 add_record.py register --product 氨氮试剂 --spec "0.10-10mg/l" \
    --unit 盒 --brand 量子 --price 399 --supplier 重庆联庆仪器

# list existing codes
python3 add_record.py query-codes
```

## 4-dimensional product matching (anti-mismatch)

Spec + name similarity + unit price together. Name similarity alone will silently merge
different products that share a size.

| 规格 | 名称相似度 | 单价 | 判定 |
|------|-----------|------|------|
| 相同 | ≥80% | any | 同一产品（价格不同时警告） |
| 相同 | 50–80% | 相同 | 同一产品，名称有变化 |
| 相同 | 50–80% | 不同 | **不同产品**（接头盒80A vs 法兰盖80A） |
| 相同 | <50% | any | **不同产品** |
| 无规格 | ≥80% | 相同 | 匹配 |

## Import (进口) business rules

Imported goods have no domestic VAT invoice. Three documents must be read **together** —
reading only the commercial invoice produces names that disagree with the customs record.

| Priority | Document | Extract |
|---|---|---|
| 1 | INVOICE & PACKING LIST | English name, spec, qty, unit price, total, supplier, invoice no, date |
| 2 | 合同 | Chinese name, spec description, Chinese supplier name |
| 3 | 报关单 | Customs Chinese name, HS code, spec |

Rules:
- Product name authority order: **报关单 > 合同 > 自主翻译**
- Quantity/price authority: **INVOICE** (it is the settlement document)
- Register with bilingual naming: `报关单中文名（INVOICE英文名）`, e.g. `透明PVC接头盒（CLEAR PVC UNION BOX）`
- Tag both the master record and the movement with `--note "进口"`
- 进口关税/增值税 are part of purchase cost — note them, do not create separate movement rows
- No 劳务/服务 filtering for imports (all imports are physical goods)

## Data safety rules (highest priority)

1. **Read before write.** Load the whole 流水 + 基础档案 before adding anything.
2. **Never overwrite.** Only append to new trailing rows. Master data may update 单价, but
   never 编码 / 名称 / 规格.
3. **Block duplicates:**
   - same 发票号 + 编码 + 变动类型 → hard duplicate, refuse
   - same 日期 + 编码 + 数量 + 金额 → suspected duplicate, ask the user
4. **Report duplicates aloud** — which row, which sheet line, why skipped. Silent skipping is
   a bug.
5. New products must be registered in 基础档案 before any movement is written.
6. 数量 must be numeric; 日期 must be `YYYY-MM-DD`; 金额 to 2 decimals.

## Relationship to this skill's core dedup rules

These rules are an application of the parent skill's 六条核心规则 — specifically rule 1
(explicit append vs overwrite), rule 2 (two-layer dedup) and rule 3 (business key, not
similar fields). Keep `output/.processed_invoices.json` in sync: if the user clears the
workbook, the tracking file must be cleared in the same breath, or every row will be
skipped as "already processed".
