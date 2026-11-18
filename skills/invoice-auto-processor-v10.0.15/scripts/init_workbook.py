#!/usr/bin/env python3
"""初始化进销存管理 Excel 表格（三表结构）。

创建：
1. 基础档案 - 产品主数据（物品编码、名称、规格、单位、品牌、参考单价、采购厂家）
2. 出入库流水 - 合并的采购+销售记录（日期、单据类型、编码、名称、变动类型、数量、发票号/合同号、金额、备注）
3. 库存看板 - 自动汇总每种产品的库存（SUMIFS 按编码匹配）
"""

import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_PATH = os.path.expanduser("~/Desktop/Claude code/CC自动库存管理.xlsx")

# 样式
HEADER_FONT = Font(name="苹方-简", bold=True, size=12, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
DATA_FONT = Font(name="苹方-简", size=12)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")


def setup_sheet(ws, headers, widths, fill=HEADER_FILL):
    """设置工作表表头和列宽"""
    for col, (name, _) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
    for col, (_, w) in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"


def create_workbook(filepath):
    wb = Workbook()

    # ===== Sheet 1: 基础档案 =====
    ws1 = wb.active
    ws1.title = "基础档案"
    # 物品编码 | 名称 | 规格型号 | 单位 | 品牌 | 参考单价 | 采购厂家
    headers1 = [
        ("物品编码", 10), ("名称", 20), ("规格型号", 22), ("单位", 8),
        ("品牌", 12), ("参考单价", 12), ("采购厂家", 20), ("备注", 10),
    ]
    setup_sheet(ws1, headers1, [w for _, w in headers1])
    ws1.auto_filter.ref = ws1.dimensions

    # ===== Sheet 2: 出入库流水 =====
    ws2 = wb.create_sheet("出入库流水")
    # 日期 | 单据类型 | 物品编码 | 物品名称 | 变动类型 | 数量 | 关联发票号/合同号 | 金额 | 备注
    headers2 = [
        ("日期", 14), ("单据类型", 10), ("物品编码", 10), ("物品名称", 20),
        ("变动类型", 10), ("数量", 10), ("关联发票号/合同号", 22), ("金额", 14), ("备注", 20),
    ]
    setup_sheet(ws2, headers2, [w for _, w in headers2])
    ws2.auto_filter.ref = ws2.dimensions

    # ===== Sheet 3: 库存看板 =====
    ws3 = wb.create_sheet("库存看板")
    # 物品编码 | 物品名称 | 累计入库 | 累计出库 | 当前库存
    headers3 = [
        ("物品编码", 10), ("物品名称", 20),
        ("累计入库", 12), ("累计出库", 12), ("当前库存", 12),
    ]
    setup_sheet(ws3, headers3, [w for _, w in headers3])

    wb.save(filepath)
    return filepath


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH

    if os.path.exists(path):
        print(f"文件已存在: {path}")
        print("如果要重新创建，请先删除或重命名现有文件。")
        sys.exit(1)

    create_workbook(path)
    print(f"进销存表格已创建: {path}")
    print(f"包含工作表: 基础档案 | 出入库流水 | 库存看板")


if __name__ == "__main__":
    main()
