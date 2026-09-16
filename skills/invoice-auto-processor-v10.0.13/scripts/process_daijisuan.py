#!/usr/bin/env python3
"""发票台账整理 — 手动触发模式（处理 待统计/ 文件夹）

与 run.py 的区别：
- run.py 处理 input/ 并归档；本脚本处理 待统计/，只写台账，不归档
- 每次生成一个新的台账文件：发票台账-<YYYY-MM-DD>.xlsx
"""
import os, sys, logging, glob
import yaml
from datetime import date

# 确保能 import 项目模块
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from invoice_processor.extractor import extract_text_from_pdf
from invoice_processor.parser import parse_invoice
from invoice_processor.writer import get_output_path, init_or_load_workbook, append_invoice_rows, get_processed_invoices, is_duplicate, mark_invoice_processed
from invoice_processor.watcher import collect_invoice_files, is_archive_file

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", stream=sys.stdout)
logger = logging.getLogger(__name__)


def load_config(config_path="config.yaml"):
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def main():
    config = load_config()
    supported = [e.lower() for e in config["processing"]["supported_extensions"]]

    # 待统计 文件夹（相对项目根）
    watch_dir = os.path.join(BASE_DIR, "待统计")
    if not os.path.isdir(watch_dir):
        logger.error(f"待统计 文件夹不存在: {watch_dir}")
        return

    # 收集待统计里的所有 PDF（含子文件夹、压缩包）
    pdf_files, tmp_dirs, archives, _ = collect_invoice_files(watch_dir, supported)

    if not pdf_files:
        logger.info("待统计 文件夹里没有 PDF，跳过。")
        return

    logger.info(f"待统计 发现 {len(pdf_files)} 个 PDF、{len(archives)} 个压缩包，开始整理...")

    # 本次台账文件名：发票台账-<今天日期>.xlsx（独立新文件，不覆盖旧台账）
    today = date.today().strftime("%Y-%m-%d")
    output_dir = config["paths"]["output_dir"]
    filename = f"发票台账-{today}.xlsx"
    output_path = get_output_path(output_dir, filename)

    # 若该文件已存在（今天已生成过一次），追加到它而不是覆盖
    wb, ws = init_or_load_workbook(output_path, config["excel"]["columns"])

    processed = 0
    failed = []
    total_amount = 0.0
    skipped_dup = 0
    dup_details = []
    existing_records = []   # 本次批次内已处理的发票记录（号码+金额+日期），用于批次内查重
    for pdf in sorted(pdf_files):
        try:
            extracted = extract_text_from_pdf(pdf)
            data = parse_invoice(extracted["full_text"])
            record = {
                "invoice_number": data.get("invoice_number", ""),
                "total_amount": data.get("total_amount", ""),
                "invoice_date": data.get("invoice_date", ""),
            }
            # 查重（v10.0.12）：唯一依据 = 发票号码相同。号码不同即使金额日期相同也独立
            dup, matched = is_duplicate(record, existing_records)
            if dup:
                skipped_dup += 1
                dup_details.append((os.path.basename(pdf), record["invoice_number"], "同号码"))
                logger.warning(f"⏭️ 跳过重复(同号码): {record['invoice_number']} | {os.path.basename(pdf)}")
                continue
            # 写入台账
            append_invoice_rows(output_path, config["excel"]["columns"], data)
            processed += 1
            # 记录到本次去重清单（防本次内重复统计）
            existing_records.append(record)
            try:
                total_amount += float(data.get("total_amount", 0) or 0)
            except (ValueError, TypeError):
                pass
            logger.info(f"✅ {os.path.basename(pdf)} → {data.get('total_amount','')}")
        except Exception as e:
            failed.append((os.path.basename(pdf), str(e)))
            logger.error(f"❌ {os.path.basename(pdf)} 失败: {e}")

    # 清理解压临时目录
    for tmp in tmp_dirs:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)

    logger.info(f"整理完成: 成功 {processed} 张, 跳过重复 {skipped_dup} 张, 失败 {len(failed)} 张")
    logger.info(f"台账已写入: {output_path}")
    logger.info(f"本次台账金额合计: {total_amount:.2f}")

    if dup_details:
        logger.info("=== 已跳过重复发票详情 ===")
        for name, no, reason in dup_details:
            logger.info(f"  - [{reason}] {no} | {name}")

    # === 台账输出后查重校验（v10.0.13）：检查本次台账是否有"同日期+同金额"项，比对号码 ==="
    check_ledger_duplicates(existing_records)


def check_ledger_duplicates(records):
    """校验本次输出的台账：找出"同日期+同金额"的项，比对发票号码。

    规则（按用户定义）：
    - 同日期+同金额 → 疑似重复信号 → 比对发票号码
    - 号码相同 → 真重复（需向用户报告）
    - 号码不同 → 两张独立发票，不重复
    """
    from collections import defaultdict
    groups = defaultdict(list)
    for r in records:
        d = _norm_date(r.get("invoice_date", ""))
        amt = _norm_amount(r.get("total_amount", ""))
        groups[(d, amt)].append({
            "number": (r.get("invoice_number") or "").strip(),
            "file": r.get("invoice_number") or "",
            "date": r.get("invoice_date", ""),
            "amount": r.get("total_amount", ""),
        })

    suspect_groups = [(k, v) for k, v in groups.items() if len(v) > 1]
    if not suspect_groups:
        logger.info("✅ 台账查重校验：无同日期同金额项")
        return

    logger.info("=== ⚠️ 台账查重校验：发现同日期同金额项，需比对发票号码 ===")
    for (d, amt), items in sorted(suspect_groups, key=lambda x: str(x[0])):
        nums = [it["number"] for it in items]
        all_same = len(set(nums)) == 1 and nums[0] != ""
        if all_same:
            verdict = "🔴 真重复（号码相同）"
        else:
            verdict = "🟢 独立发票（号码不同）"
        logger.info(f"  日期:{d} 金额:{amt} 共{len(items)}项 → {verdict}")
        for it in items:
            logger.info(f"     - 号码:{it['number']} 日期:{it['date']} 金额:{it['amount']}")


def _norm_date(s):
    """日期归一化到 YYYYMMDD 用于分组"""
    import re
    s = str(s).strip()
    # 完整日期: 2026年09月04日 或 2026-09-04 或 2026/09/04
    m = re.search(r'(\d{4})[-/年]\s*(\d{1,2})[-/月]\s*(\d{1,2})', s)
    if m:
        return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"
    # 仅 月/日 形式（如 9月4日）
    m = re.search(r'(\d{1,2})\s*月\s*(\d{1,2})\s*日', s)
    if m:
        return f"0000{int(m.group(1)):02d}{int(m.group(2)):02d}"
    return s


def _norm_amount(amount):
    """金额归一化到两位小数用于分组"""
    s = str(amount).strip().replace("¥", "").replace("￥", "").replace(",", "")
    try:
        return f"{float(s):.2f}"
    except (ValueError, TypeError):
        return s

    # 失败清单
    if failed:
        logger.info("失败清单:")
        for name, err in failed:
            logger.info(f"  - {name}: {err}")


if __name__ == "__main__":
    main()
