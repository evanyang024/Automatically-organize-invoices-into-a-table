"""发票文件智能重命名模块 — 文件名格式: YYYY-MM-DD-项目-金额.pdf"""
import os, re, logging

logger = logging.getLogger(__name__)
_INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def build_new_filename(invoice_data, original_ext=".pdf"):
    """根据发票数据构建新文件名（parser 必须返回完整日期）"""
    date_str = _format_date(invoice_data.get("invoice_date", ""))
    item_str = _format_items(invoice_data.get("items", []))
    amount_str = invoice_data.get("total_amount", "")
    if not amount_str:
        amount_str = invoice_data.get("amount_excluding_tax", "")
    if not date_str and not amount_str:
        return None

    amount_str = re.sub(r'[￥¥$,]', '', amount_str).strip()
    parts = [p for p in [date_str, item_str, amount_str] if p]
    if not parts:
        return None

    name = "-".join(parts)
    name = _sanitize_filename(name)
    return f"{name}{original_ext}"


def _format_date(date_str):
    """'2026年04月21日' → '2026-04-21'（需要完整日期含年份）"""
    if not date_str:
        return ""
    m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', date_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', date_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def _format_items(items):
    """单项目→名称，多项目→'第一项等N项'"""
    if not items:
        return ""
    first_name = items[0].get("货物或服务名称", "").strip()
    if not first_name:
        return ""
    if len(items) == 1:
        return first_name
    return f"{first_name}等{len(items)}项"


def _sanitize_filename(name):
    name = _INVALID_CHARS.sub('', name)
    name = name.strip('. ')
    if len(name) > 180:
        name = name[:180]
    return name


def rename_invoice_file(filepath, invoice_data):
    """重命名单个发票文件（原地改名），返回新路径"""
    new_name = build_new_filename(invoice_data, os.path.splitext(filepath)[1])
    if not new_name:
        logger.warning(f"无法生成新文件名，保留原名: {os.path.basename(filepath)}")
        return filepath

    dirname = os.path.dirname(filepath)
    new_path = os.path.join(dirname, new_name)
    if os.path.normcase(new_path) != os.path.normcase(filepath) and os.path.exists(new_path):
        name, ext = os.path.splitext(new_name)
        counter = 1
        while os.path.exists(os.path.join(dirname, f"{name}({counter}){ext}")):
            counter += 1
        new_path = os.path.join(dirname, f"{name}({counter}){ext}")
    if os.path.normcase(new_path) == os.path.normcase(filepath):
        return filepath
    try:
        os.rename(filepath, new_path)
        logger.info(f"重命名: {os.path.basename(filepath)} → {os.path.basename(new_path)}")
        return new_path
    except OSError as e:
        logger.error(f"重命名失败: {e}")
        return filepath
