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
        "开票日期": _format_date_display(invoice_data.get("invoice_date", "")),
        "备注": invoice_data.get("remarks", ""),
    }
    return [mapping.get(col, "") for col in columns]


def _format_date_display(date_str):
    """统一开票日期为「M月D日」格式（去年份，月/日不补零）。

    支持输入格式：2026年09月06日 / 2026-09-06 / 2026/09/06 / 2026-7-9 /
                  09月06日（已去年份）/ 2026年9月6日
    统一输出：9月6日（不带年份，月/日不补零）
    """
    if not date_str:
        return ""
    s = str(date_str).strip()
    import re
    # 提取 年月日（允许 年月日 或 y-m-d 或 y/m/d 分隔）
    m = re.search(r'(\d{4})\s*年?\s*(\d{1,2})\s*月?\s*(\d{1,2})\s*日?', s)
    if m:
        return f"{int(m.group(2))}月{int(m.group(3))}日"
    # 已经是 M月D日 形式（可能带补零，如 09月06日）
    m = re.search(r'(\d{1,2})\s*月\s*(\d{1,2})\s*日', s)
    if m:
        return f"{int(m.group(1))}月{int(m.group(2))}日"
    # 仅 M/D 或 M-D 形式（无说明符）
    m = re.search(r'(\d{1,2})[-/](\d{1,2})$', s)
    if m:
        return f"{int(m.group(1))}月{int(m.group(2))}日"
    return s


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
            # v10.0.13：锁文件存在 ≠ 真的被占用。用 lsof 判断是否有进程真正持有目标文件。
            # 僵尸锁文件（进程已结束但文件残留）应清理并放行，否则会永远卡死。
            if _is_really_locked(full_path):
                import logging
                logging.getLogger(__name__).warning(
                    f"⚠️ 台账文件被 Excel/WPS 占用（检测到锁文件 {ln} 且进程持有）。"
                    f"为免覆盖丢失数据，已中止写入。请手动关闭 Excel/WPS 中打开的台账后重新运行。"
                )
                return True
            else:
                # 僵尸锁文件：无进程持有，尝试清理后继续
                import logging
                logging.getLogger(__name__).warning(
                    f"清理僵尸锁文件（无进程持有）: {ln}"
                )
                try:
                    os.remove(lock_path)
                except OSError:
                    pass
                break

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


def _is_really_locked(full_path):
    """判断目标文件是否真的被某个进程持有（用 lsof 检测）。

    锁文件存在但 lsof 查不到持有进程 → 僵尸锁文件，返回 False（应清理放行）。
    """
    import subprocess
    try:
        r = subprocess.run(
            ["lsof", full_path],
            capture_output=True, text=True, timeout=5
        )
        # lsof 输出含除自己外的打开记录 → 真持有
        lines = [l for l in r.stdout.splitlines() if l and 'COMMAND' not in l]
        return len(lines) > 0
    except Exception:
        # lsof 不可用时，保守按"锁文件存在=锁定"处理（宁可不写，不丢数据）
        return True


# === 去重跟踪（v10.0.12: 号码 + 金额 + 日期 三元组）===
def _tracking_file(output_dir):
    return os.path.join(output_dir, ".processed_invoices.json")


def get_processed_invoices(output_dir):
    """读取去重记录，返回 list of dict。

    v10.0.12 升级：记录结构为 {"invoice_number","total_amount","invoice_date"}。
    兼容旧格式（旧清单是纯号码数组 ['xxx','yyy']），旧条目金额/日期为空，只能按号码匹配。
    """
    tf = _tracking_file(output_dir)
    if not os.path.exists(tf):
        return []
    try:
        with open(tf, "r", encoding="utf-8") as f:
            raw = _json.load(f)
        # 新格式：list of dict
        if isinstance(raw, list) and raw and isinstance(raw[0], dict):
            return raw
        # 旧格式：list of str（纯号码）
        if isinstance(raw, list):
            return [{"invoice_number": str(x), "total_amount": "", "invoice_date": ""} for x in raw if x]
    except Exception:
        return []
    return []


def get_processed_invoice_numbers(output_dir):
    """兼容旧接口：只返回号码 set（部分旧代码仍用这个读号码）"""
    return set(r["invoice_number"] for r in get_processed_invoices(output_dir) if r.get("invoice_number"))


def _normalize_amount(amount):
    """金额归一化：去掉货币符号/千分位，转成可比的字符串"""
    if amount is None:
        return ""
    s = str(amount).strip()
    s = s.replace("¥", "").replace("￥", "").replace(",", "").replace(" ", "")
    try:
        return f"{float(s):.2f}"
    except (ValueError, TypeError):
        return s


def _normalize_date(date_str):
    """日期归一化：转成 YYYYMMDD 便于比较（处理 2026年09月06日 / 2026-09-06 / 2026/09/06）"""
    if not date_str:
        return ""
    import re as _re
    s = str(date_str).strip()
    m = _re.search(r'(\d{4})\s*年?\s*(\d{1,2})\s*月?\s*(\d{1,2})\s*日?', s)
    if m:
        try:
            return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"
        except ValueError:
            return ""
    m = _re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m:
        try:
            return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"
        except ValueError:
            return ""
    return ""


def is_duplicate(record, processed):
    """判断 record 是否重复（相对已处理清单 processed）。

    v10.0.12 修正判定规则：
    - **唯一重复依据 = 发票号码相同** → 重复
    - 若号码不同，即使金额+日期都相同，也是**两张独立发票**，不判重复
    - 金额+日期相同仅作为"疑似重复"信号，用于比对号码；号码不符则认定独立

    record / processed 元素：{"invoice_number","total_amount","invoice_date"}
    Returns: (bool, matched_record_or_None)
    """
    rec_num = (record.get("invoice_number") or "").strip()
    if not rec_num:
        # 无号码，无法凭号码判定 → 不判重复（当作独立发票，宁多勿漏）
        return False, None
    rec_amt = _normalize_amount(record.get("total_amount"))
    rec_date = _normalize_date(record.get("invoice_date"))

    for p in processed:
        p_num = (p.get("invoice_number") or "").strip() if isinstance(p, dict) else ""
        # 唯一重复依据：发票号码相同
        if p_num and rec_num == p_num:
            return True, p
    return False, None


def mark_invoice_processed(output_dir, invoice_number, total_amount="", invoice_date=""):
    """标记一张发票已处理，写入 (号码,金额,日期)。"""
    if not invoice_number and not total_amount:
        return
    tf = _tracking_file(output_dir)
    existing = get_processed_invoices(output_dir)
    # 去重追加：同号码或同金额日期的不重复添加
    rec = {"invoice_number": invoice_number or "",
           "total_amount": total_amount if total_amount else "",
           "invoice_date": invoice_date if invoice_date else ""}
    for e in existing:
        e_num = (e.get("invoice_number") or "").strip()
        if invoice_number and e_num and e_num == invoice_number:
            # 更新补全金额/日期
            if not e.get("total_amount") and total_amount:
                e["total_amount"] = total_amount
            if not e.get("invoice_date") and invoice_date:
                e["invoice_date"] = invoice_date
            with open(tf, "w", encoding="utf-8") as f:
                _json.dump(existing, f, ensure_ascii=False)
            return
    existing.append(rec)
    with open(tf, "w", encoding="utf-8") as f:
        _json.dump(existing, f, ensure_ascii=False)
