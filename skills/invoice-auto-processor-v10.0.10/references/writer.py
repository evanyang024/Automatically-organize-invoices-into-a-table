"""Excel 写入模块 — 金额为数字、无边框、多明细合并 + WPS锁检测(v10.0.10)"""
import os, re, json as _json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill


def get_output_path(output_dir, filename):
    return os.path.join(output_dir, filename)


def init_or_load_workbook(filepath, columns):
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "发票台账"
        for col_idx, col_name in enumerate(columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        for col_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
    return wb, ws


def append_invoice_rows(filepath, columns, invoice_data):
    # v10.0.10: 检测 Excel/WPS 锁。若被占用，中止写入并提示用户关闭，避免覆盖丢失数据
    if _close_excel_if_open(os.path.basename(filepath), os.path.dirname(filepath)):
        raise PermissionError("台账文件被 Excel/WPS 占用，已中止写入。请手动关闭后重试。")
    
    wb, ws = init_or_load_workbook(filepath, columns)
    items = invoice_data.get("items", [])
    next_row = _find_empty_row(ws)

    if items:
        # 一张发票只写一行
        item = items[0]
        if len(items) > 1:
            item["货物或服务名称"] = f"{item['货物或服务名称']}等{len(items)}项"
        row_data = _build_row(columns, item, invoice_data)
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.alignment = Alignment(vertical='center')
            if col_idx == 2 and isinstance(value, (int, float)):
                cell.number_format = '#,##0.00'
    else:
        row_data = _build_row(columns, {}, invoice_data)
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.alignment = Alignment(vertical='center')

    for col_idx in range(1, len(columns) + 1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 16

    # 按项目名称分类排序，同类型发票挨在一起
    _sort_rows_by_name(ws)

    wb.save(filepath)
    wb.close()


def _sort_rows_by_name(ws):
    """按第一列（货物或服务名称）排序，同类型发票挨在一起"""
    rows_data = []
    for row in range(2, ws.max_row + 1):
        row_values = []
        has_data = False
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            row_values.append(val)
            if val is not None:
                has_data = True
        if not has_data:
            continue
        rows_data.append(row_values)
    
    if not rows_data:
        return
    
    rows_data.sort(key=lambda r: str(r[0]) if r[0] is not None else "zzz")
    
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).value = None
    
    for i, row_data in enumerate(rows_data, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=i, column=col_idx, value=value)
            cell.alignment = Alignment(vertical='center')
            if col_idx == 2 and isinstance(value, (int, float)):
                cell.number_format = '#,##0.00'


def _find_empty_row(ws):
    """找到第一个空行（所有列都没有数据）"""
    # 从第2行开始找（第1行是表头）
    for row in range(2, ws.max_row + 2):
        is_empty = True
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=row, column=col).value is not None:
                is_empty = False
                break
        if is_empty:
            return row
    return ws.max_row + 1


def _build_row(columns, item, invoice_data):
    import re
    mapping = {
        "货物或服务名称": item.get("货物或服务名称", ""),
        "规格型号": item.get("规格型号", ""),
        "数量": item.get("数量", ""),
        "单价": item.get("单价", ""),
        "金额": item.get("金额", ""),
        "税率": item.get("税率", ""),
        "不含税金额合计": invoice_data.get("amount_excluding_tax", ""),
        "税额合计": invoice_data.get("total_tax", ""),
        "价税合计": _to_number(invoice_data.get("total_amount", "")),
        "开票日期": re.sub(r'[0-9]{4}[ ]*年[ ]*', '', invoice_data.get("invoice_date", "")),
        "备注": invoice_data.get("remarks", ""),
    }
    return [mapping.get(col, "") for col in columns]


def _to_number(val):
    if val:
        try:
            return float(str(val).replace(",", ""))
        except ValueError:
            return val
    return val


def _close_excel_if_open(target_filename, output_dir=""):
    """检测 Excel/WPS 是否打开了目标文件（占用锁）。

    v10.0.10 修复：不再模拟按键强杀 WPS（会导致未落盘数据被覆盖丢失）。
    改为检测到锁文件就返回 True（表示文件被占用），由调用方停止并提示用户手动关闭。

    Returns:
        bool: True = 文件被 Excel/WPS 占用（写入应中止）；False = 可安全写入。
    """
    import platform
    if platform.system() != "Darwin":
        return False

    full_path = target_filename if os.path.isabs(target_filename) else os.path.join(output_dir, target_filename)

    # 检测 WPS/Excel 的锁文件（macOS 上 WPS/Excel 会用 .~文件名 作为锁文件）
    base_dir = os.path.dirname(full_path) if os.path.dirname(full_path) else "."
    lock_names = [
        ".~" + os.path.basename(full_path),       # WPS 式锁文件
        "~$" + os.path.basename(full_path).replace(" ", ""),  # Office 式临时锁
        os.path.basename(full_path) + ".lock",
    ]
    for ln in lock_names:
        lock_path = os.path.join(base_dir, ln)
        if os.path.exists(lock_path):
            import logging
            logging.getLogger(__name__).warning(
                f"⚠️ 台账文件被 Excel/WPS 占用（检测到锁文件 {ln}）。"
                f"为免覆盖丢失数据，已中止写入。请手动关闭 Excel/WPS 中打开的台账后重新运行。"
            )
            return True

    # 兜底：尝试以读+写方式打开检测（openpyxl 写需要独占，被锁会抛 PermissionError）
    try:
        from openpyxl import load_workbook
        if os.path.exists(full_path):
            wb = load_workbook(full_path)
            wb.close()
            return False
    except PermissionError:
        import logging
        logging.getLogger(__name__).warning(
            "⚠️ 台账文件无写权限或正被占用（PermissionError），请关闭 Excel/WPS 后重试。"
        )
        return True
    except Exception:
        # openpyxl 读失败不一定代表锁，不误判
        return False
    return False


# === 去重跟踪 ===
def _tracking_file(output_dir):
    return os.path.join(output_dir, ".processed_invoices.json")


def get_processed_invoice_numbers(output_dir):
    tf = _tracking_file(output_dir)
    if not os.path.exists(tf):
        return set()
    try:
        with open(tf, "r", encoding="utf-8") as f:
            return set(_json.load(f))
    except Exception:
        return set()


def mark_invoice_processed(output_dir, invoice_number):
    if not invoice_number:
        return
    tf = _tracking_file(output_dir)
    existing = get_processed_invoice_numbers(output_dir)
    existing.add(invoice_number)
    with open(tf, "w", encoding="utf-8") as f:
        _json.dump(sorted(existing), f, ensure_ascii=False)
