#!/usr/bin/env python3
"""向进销存表格添加出入库记录并刷新库存看板。

支持三种操作模式：
  1. add    - 添加出入库流水记录
  2. register - 注册新产品到基础档案
  3. query-codes - 查看已有产品编码列表

用法:
    # 添加记录
    python3 add_record.py add --code 3 --date "2025-06-01" --direction "入库" \
        --change-type "采购" --product "氨氮试剂" --spec "0.10-10mg/l" \
        --unit "盒" --quantity 10 --amount 3990 --invoice "FP20250601-001"

    # 注册新产品（自动分配编码）
    python3 add_record.py register --product "氨氮试剂" --spec "0.10-10mg/l" \
        --unit "盒" --brand "量子" --price 399 --supplier "重庆联庆仪器"

    # 查看编码
    python3 add_record.py query-codes
"""

import os
import sys
import argparse
import difflib
import subprocess
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def close_wps_file(filepath):
    """如果 WPS 打开了指定文件，先保存再关闭，释放文件锁。"""
    basename = os.path.basename(filepath)
    try:
        # 告诉 WPS 保存并关闭指定文档
        subprocess.run([
            "osascript", "-e",
            f'tell application "wpsoffice" to close (every document whose name is "{basename}") saving yes'
        ], timeout=10)
    except Exception:
        pass  # WPS 没运行或文件没打开，忽略

DEFAULT_PATH = os.path.expanduser("~/Desktop/Claude code/CC自动库存管理.xlsx")

# 样式
DATA_FONT = Font(name="苹方-简", size=12)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")


def style_cell(cell, align=CENTER_ALIGN):
    cell.font = DATA_FONT
    cell.border = THIN_BORDER
    cell.alignment = align


# ==================== 产品档案操作 ====================

def get_all_codes(wb):
    """获取所有已注册的物品编码（整列匹配，含空行空字符串）"""
    ws = wb["基础档案"]
    codes = {}
    for row in range(2, ws.max_row + 1):
        code_val = ws.cell(row=row, column=1).value
        if code_val is not None and str(code_val).strip():
            try:
                code = int(code_val)
            except ValueError:
                code = str(code_val).strip()
            name = str(ws.cell(row=row, column=2).value or "").strip()
            spec = str(ws.cell(row=row, column=3).value or "").strip()
            unit = str(ws.cell(row=row, column=4).value or "").strip()
            brand = str(ws.cell(row=row, column=5).value or "").strip()
            price = ws.cell(row=row, column=6).value
            supplier = str(ws.cell(row=row, column=7).value or "").strip()
            codes[code] = {
                "row": row, "name": name, "spec": spec, "unit": unit,
                "brand": brand, "price": price, "supplier": supplier,
            }
    return codes


def name_similarity(name1, name2):
    """计算两个名称的相似度（0-1）"""
    return difflib.SequenceMatcher(None, name1.strip(), name2.strip()).ratio()


def find_code(codes, product_name, spec="", price=None):
    """三维匹配编码（规格 + 名称相似度 + 单价）。

    返回: (code, info, warning) 或 None

    核心原则：规格相同时不代表一定是同一产品。对于尺寸型规格（如管径80A）
    不同品类（接头盒vs法兰盖）应视为不同产品。

    匹配规则：
    1. 规格相同 + 名称≥80%相似 → 同一产品（价格不同时警告）
    2. 规格相同 + 名称50-80%相似 + 单价相同 → 同一产品，名称有变化
    3. 规格相同 + 名称50-80%相似 + 单价不同 → 不同产品！不匹配
    4. 规格相同 + 名称<50%相似 → 不同产品！不匹配
    5. 无规格 → 名称≥80%相似 + 单价相同 → 匹配
    """
    s = spec.strip() if spec else ""
    p = product_name.strip()

    def _price_same(existing_price):
        if price is None or existing_price is None:
            return None  # 无价格信息，无法判断
        return abs(float(price) - float(existing_price)) < 0.01

    # 有规格型号
    if s:
        best_candidate = None  # (code, info, sim, price_same)

        for code, info in codes.items():
            if info["spec"].strip() != s:
                continue

            sim = name_similarity(info["name"], p)
            existing_price = info.get("price")
            ps = _price_same(existing_price)

            # 名称≥80% → 确认同一产品
            if sim >= 0.8:
                warning = "price_differs" if ps is False else None
                return code, info, warning

            # 名称<50% → 完全不同产品，跳过
            if sim < 0.5:
                continue

            # 名称50-80%：看价格辅助判断
            if best_candidate is None:
                best_candidate = (code, info, sim, ps)

        # 处理名称50-80%的候选项
        if best_candidate:
            code, info, sim, ps = best_candidate
            # 单价相同 → 同一产品；单价不同且都有价格 → 不同产品
            if ps is True:
                return code, info, "name_differs"
            elif ps is False:
                return None  # 名称50-80% + 价格不同 → 不是同一产品
            else:
                # 无价格信息，名称50-80% → 算同一产品但提醒
                return code, info, "name_differs"

        return None  # 规格匹配但名称<50% → 不同产品

    # 无规格：名称≥80% + 价格辅助
    for code, info in codes.items():
        sim = name_similarity(info["name"], p)
        if sim >= 0.8:
            existing_price = info.get("price")
            ps = _price_same(existing_price)
            if ps is True or ps is None:
                return code, info, None

    return None


def get_next_code(codes):
    """获取下一个可用的物品编码"""
    if not codes:
        return 1
    numeric = [c for c in codes if isinstance(c, int)]
    return max(numeric) + 1 if numeric else 1


def register_product(filepath, product, spec, unit="", brand="", price=None, supplier="", note=""):
    """注册新产品到基础档案"""
    wb = load_workbook(filepath)
    codes = get_all_codes(wb)

    # 三维匹配检查（规格+名称相似度+单价）
    existing = find_code(codes, product, spec, price)
    if existing:
        existing_code, info, warning = existing
        existing_price = info.get("price")
        wb.close()
        return {
            "status": "exists",
            "code": existing_code,
            "product": product,
            "spec": spec,
            "warning": warning,
            "existing_price": existing_price,
            "new_price": price,
            "existing_name": info.get("name"),
        }

    # 分配新编码
    new_code = get_next_code(codes)
    ws = wb["基础档案"]
    next_row = ws.max_row + 1
    if ws.cell(row=2, column=1).value is None:
        next_row = 2

    values = [new_code, product, spec, unit, brand, price, supplier, note]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=next_row, column=col, value=val)
        if col == 1:  # 编码
            style_cell(cell, CENTER_ALIGN)
        elif col == 4:  # 单位
            style_cell(cell, CENTER_ALIGN)
        elif col == 6:  # 单价
            style_cell(cell, RIGHT_ALIGN)
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
        elif col == 8:  # 备注
            style_cell(cell, CENTER_ALIGN)
        else:
            style_cell(cell, LEFT_ALIGN)

    wb.save(filepath)
    wb.close()
    return {"status": "ok", "code": new_code, "product": product, "spec": spec}


# ==================== 流水记录操作 ====================

def check_duplicate(wb, code, invoice, change_type):
    """检查是否已存在相同记录（同编码+同发票号+同变动类型）"""
    ws = wb["出入库流水"]
    for row in range(2, ws.max_row + 1):
        c = ws.cell(row=row, column=3).value
        inv = str(ws.cell(row=row, column=7).value or "")
        ct = str(ws.cell(row=row, column=5).value or "")
        try:
            c = int(c)
        except (ValueError, TypeError):
            pass
        if c == code and inv == invoice and ct == change_type:
            return row
    return None


def add_transaction(filepath, code, date_str, direction, change_type,
                    product, quantity, amount, invoice="", note=""):
    """向出入库流水添加一条记录"""
    wb = load_workbook(filepath)
    ws = wb["出入库流水"]

    # 检查重复
    dup_row = check_duplicate(wb, code, invoice, change_type)
    if dup_row:
        wb.close()
        return {"status": "duplicate", "row": dup_row, "code": code}

    # 找下一个空行
    next_row = ws.max_row + 1
    if ws.cell(row=2, column=1).value is None:
        next_row = 2

    values = [date_str, direction, code, product, change_type, quantity, invoice, amount, note]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=next_row, column=col, value=val)
        if col in (1, 2, 3, 5):  # 日期、单据类型、编码、变动类型
            style_cell(cell, CENTER_ALIGN)
        elif col in (6, 8):  # 数量、金额
            style_cell(cell, RIGHT_ALIGN)
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00' if col == 8 else '#,##0'
        else:
            style_cell(cell, LEFT_ALIGN)

    wb.save(filepath)
    wb.close()
    return {"status": "ok", "row": next_row, "code": code, "product": product}


# ==================== 库存看板操作 ====================

def refresh_inventory(filepath):
    """刷新库存看板——按编码从流水表汇总数量"""
    wb = load_workbook(filepath)
    ws_inv = wb["库存看板"]

    # 收集基础档案中的产品编码
    codes = get_all_codes(wb)

    # 清空旧数据（保留表头）
    for row in range(2, ws_inv.max_row + 1):
        for col in range(1, 6):
            ws_inv.cell(row=row, column=col).value = None

    # 流水表最后一行
    ws_flow = wb["出入库流水"]
    flow_last = max(ws_flow.max_row, 2)

    row = 2
    for code in sorted(codes.keys(), key=lambda x: int(x) if isinstance(x, int) else 9999):
        info = codes[code]
        # 编码
        cell_c = ws_inv.cell(row=row, column=1, value=code)
        style_cell(cell_c, CENTER_ALIGN)

        # 名称（带规格）
        display_name = info["name"]
        if info["spec"]:
            display_name += f" ({info['spec']})"
        cell_n = ws_inv.cell(row=row, column=2, value=info["name"])
        style_cell(cell_n, LEFT_ALIGN)

        # 累计入库 = SUMIFS(流水!F:F, 流水!C:C, code, 流水!B:B, "入库")
        formula_in = (
            f'=SUMIFS(出入库流水!F:F,出入库流水!C:C,A{row},出入库流水!B:B,"入库")'
        )
        cell_in = ws_inv.cell(row=row, column=3, value=formula_in)
        style_cell(cell_in, RIGHT_ALIGN)

        # 累计出库 = SUMIFS(流水!F:F, 流水!C:C, code, 流水!B:B, "出库")
        formula_out = (
            f'=SUMIFS(出入库流水!F:F,出入库流水!C:C,A{row},出入库流水!B:B,"出库")'
        )
        cell_out = ws_inv.cell(row=row, column=4, value=formula_out)
        style_cell(cell_out, RIGHT_ALIGN)

        # 当前库存
        cell_stock = ws_inv.cell(row=row, column=5, value=f"=C{row}-D{row}")
        style_cell(cell_stock, RIGHT_ALIGN)

        row += 1

    # 列宽
    widths = [10, 22, 14, 14, 12]
    for col, w in enumerate(widths, 1):
        ws_inv.column_dimensions[get_column_letter(col)].width = w

    wb.save(filepath)
    wb.close()
    return len(codes)


def list_codes(filepath):
    """列出所有产品编码"""
    wb = load_workbook(filepath)
    codes = get_all_codes(wb)
    wb.close()
    return codes


# ==================== 命令行入口 ====================

def cmd_add(args):
    """添加出入库记录"""
    close_wps_file(args.file)
    if not os.path.exists(args.file):
        print(f"文件不存在: {args.file}")
        sys.exit(1)

    result = add_transaction(
        args.file, args.code, args.date, args.direction, args.change_type,
        args.product, args.quantity, args.amount, args.invoice, args.note or "",
    )

    if result["status"] == "duplicate":
        print(f"重复记录: 编码 {args.code} 发票 {args.invoice} 已存在（第 {result['row']} 行），跳过。")
        return

    product_count = refresh_inventory(args.file)

    print(f"已录入 [{args.direction}/{args.change_type}] 编码{args.code} {args.product}")
    print(f"   日期: {args.date} | 数量: {args.quantity} | 金额: {args.amount}")
    print(f"   库存看板已刷新，共 {product_count} 种产品。")


def cmd_register(args):
    """注册新产品"""
    close_wps_file(args.file)
    if not os.path.exists(args.file):
        print(f"文件不存在: {args.file}")
        sys.exit(1)

    result = register_product(
        args.file, args.product, args.spec or "", args.unit or "",
        args.brand or "", args.price, args.supplier or "", args.note or "",
    )

    if result["status"] == "exists":
        w = result.get("warning")
        if w == "price_differs":
            old_p = result.get("existing_price")
            new_p = result.get("new_price")
            old_n = result.get("existing_name", "")
            print(f" 产品已匹配: 编码 {result['code']} {old_n} → {result['product']} ({result['spec']})")
            print(f" 价格变动: 档案 ¥{old_p} → 发票 ¥{new_p}")
        elif w == "name_differs":
            old_n = result.get("existing_name", "")
            print(f"产品已匹配(规格相同): 编码 {result['code']} {old_n} → {result['product']} ({result['spec']})")
        else:
            print(f"产品已存在: 编码 {result['code']} {result['product']} ({result['spec']})")
        return

    product_count = refresh_inventory(args.file)
    print(f"新注册产品: 编码 {result['code']} {result['product']} ({result['spec']})")
    print(f"   库存看板已刷新，共 {product_count} 种产品。")


def cmd_query_codes(args):
    """查询编码"""
    close_wps_file(args.file)
    if not os.path.exists(args.file):
        print(f"文件不存在: {args.file}")
        sys.exit(1)

    codes = list_codes(args.file)
    if not codes:
        print("暂无产品。")
        return

    print(f"{'编码':<6} {'名称':<20} {'规格型号':<24} {'单位':<6} {'厂家'}")
    print("-" * 80)
    for code in sorted(codes.keys(), key=lambda x: int(x) if isinstance(x, int) else 9999):
        info = codes[code]
        print(f"{str(code):<6} {info['name']:<20} {info['spec']:<24} {info['unit']:<6} {info['supplier']}")


def main():
    parser = argparse.ArgumentParser(description="进销存管理工具")
    subparsers = parser.add_subparsers(dest="command")

    # ---- add ----
    p_add = subparsers.add_parser("add", help="添加出入库记录")
    p_add.add_argument("--file", default=DEFAULT_PATH, help="Excel 文件路径")
    p_add.add_argument("--code", type=int, required=True, help="物品编码")
    p_add.add_argument("--date", required=True, help="日期 YYYY-MM-DD")
    p_add.add_argument("--direction", required=True, choices=["入库", "出库"], help="单据类型")
    p_add.add_argument("--change-type", required=True, choices=["采购", "销售"], help="变动类型")
    p_add.add_argument("--product", required=True, help="产品名称")
    p_add.add_argument("--quantity", type=float, required=True, help="数量")
    p_add.add_argument("--amount", type=float, required=True, help="金额")
    p_add.add_argument("--invoice", default="", help="发票号/合同号")
    p_add.add_argument("--note", default="", help="备注")

    # ---- register ----
    p_reg = subparsers.add_parser("register", help="注册新产品")
    p_reg.add_argument("--file", default=DEFAULT_PATH, help="Excel 文件路径")
    p_reg.add_argument("--product", required=True, help="产品名称")
    p_reg.add_argument("--spec", default="", help="规格型号")
    p_reg.add_argument("--unit", default="", help="单位")
    p_reg.add_argument("--brand", default="", help="品牌")
    p_reg.add_argument("--price", type=float, default=None, help="参考单价")
    p_reg.add_argument("--supplier", default="", help="采购厂家")
    p_reg.add_argument("--note", default="", help="备注（如：进口）")

    # ---- query-codes ----
    p_q = subparsers.add_parser("query-codes", help="查看产品编码列表")
    p_q.add_argument("--file", default=DEFAULT_PATH, help="Excel 文件路径")

    args = parser.parse_args()

    if args.command == "add":
        cmd_add(args)
    elif args.command == "register":
        cmd_register(args)
    elif args.command == "query-codes":
        cmd_query_codes(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
