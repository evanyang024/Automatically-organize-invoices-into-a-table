#!/usr/bin/env python3
"""第一部分：input/ → 智能重命名 + 归档（不写台账，不删源文件）

流程：
- 收集 input/ 里所有 PDF（含子文件夹、压缩包）
- 解析每张发票，生成智能新名（YYYY-MM-DD-项目-金额.pdf）
- 拷贝一份到 archive/（源文件保留在 input/ 不动）
- 不做台账写入（台账归第二部分 process_daijisuan.py）
"""
import os, sys, logging, shutil, glob
import yaml

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from invoice_processor.extractor import extract_text_from_pdf
from invoice_processor.parser import parse_invoice
from invoice_processor.renamer import build_new_filename
from invoice_processor.watcher import collect_invoice_files
from invoice_processor.writer import get_processed_invoices, is_duplicate, mark_invoice_processed

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", stream=sys.stdout)
logger = logging.getLogger(__name__)


def load_config(config_path="config.yaml"):
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def main():
    config = load_config()
    supported = [e.lower() for e in config["processing"]["supported_extensions"]]
    archive_dir = os.path.join(BASE_DIR, config["paths"]["archive_dir"])
    watch_dir = os.path.join(BASE_DIR, config["paths"]["watch_dir"])  # input/
    output_dir = os.path.join(BASE_DIR, config["paths"]["output_dir"])
    os.makedirs(archive_dir, exist_ok=True)

    if not os.path.isdir(watch_dir):
        logger.error(f"input 文件夹不存在: {watch_dir}")
        return

    pdf_files, tmp_dirs, archives, _ = collect_invoice_files(watch_dir, supported)
    if not pdf_files:
        logger.info("input/ 里没有待归档 PDF。")
        return

    logger.info(f"input/ 发现 {len(pdf_files)} 个 PDF、{len(archives)} 个压缩包，开始重命名归档...")

    # 载入已处理清单（号码+金额+日期 三元组）
    processed = get_processed_invoices(output_dir)

    archived = 0
    skipped_dup = 0
    dup_details = []   # 记录跳过详情供汇报（精确到命中原因）
    batch_records = []  # 本次批次内已归档的发票记录，用于同日期同金额校验
    failed = []
    for pdf in sorted(pdf_files):
        try:
            filename = os.path.basename(pdf)
            extracted = extract_text_from_pdf(pdf)
            data = parse_invoice(extracted["full_text"])
            record = {
                "invoice_number": data.get("invoice_number", ""),
                "total_amount": data.get("total_amount", ""),
                "invoice_date": data.get("invoice_date", ""),
            }

            # 查重（v10.0.12）：唯一依据 = 发票号码相同。号码不同即使金额日期相同也独立
            dup, matched = is_duplicate(record, processed)
            if dup:
                skipped_dup += 1
                dup_details.append((filename, record["invoice_number"], "同号码"))
                logger.warning(f"⏭️ 跳过重复(同号码): {record['invoice_number']} | {filename}")
                continue

            new_filename = build_new_filename(data, os.path.splitext(filename)[1]) or filename
            # 拷贝到 archive/（源文件保留在 input/)
            dest = os.path.join(archive_dir, new_filename)
            if os.path.exists(dest):
                name, ext = os.path.splitext(new_filename)
                import time
                dest = os.path.join(archive_dir, f"{name}_{int(time.time())}{ext}")
            shutil.copy2(pdf, dest)
            archived += 1
            # 归档成功后标记到去重清单（号码+金额+日期）
            mark_invoice_processed(output_dir, record["invoice_number"], record["total_amount"], record["invoice_date"])
            batch_records.append(record)  # 加入批次内记录，供同日期同金额校验
            logger.info(f"✅ {filename} → archive/{os.path.basename(dest)}  ({record['invoice_number']})")
        except Exception as e:
            failed.append((os.path.basename(pdf), str(e)))
            logger.error(f"❌ {os.path.basename(pdf)} 失败: {e}")

    # 清理解压临时目录
    for tmp in tmp_dirs:
        shutil.rmtree(tmp, ignore_errors=True)

    logger.info(f"归档完成: 成功 {archived} 张, 跳过重复 {skipped_dup} 张, 失败 {len(failed)} 张")
    logger.info("源文件保留在 input/（未删除）。台账未写入（归第二部分 process_daijisuan.py）。")

    # 汇报：重复跳过的详情
    if dup_details:
        logger.info("=== 已跳过重复发票（同号码已在清单/archive）详情 ===")
        for name, no in dup_details:
            logger.info(f"  - 发票号 {no} | {name}")

    # === 批次内同日期同金额校验（v10.0.13）：检查本次 input 已归档的发票是否有同日期同金额项，比对号码 ===
    if batch_records:
        try:
            from process_daijisuan import check_ledger_duplicates
            logger.info("--- 第一部分批次内查重校验 ---")
            check_ledger_duplicates(batch_records)
        except ImportError:
            logger.info("(跳过同日期同金额校验：无法导入校验模块)")
        except Exception as e:
            logger.info(f"(同日期同金额校验异常: {e})")

    if failed:
        logger.info("失败清单:")
        for name, err in failed:
            logger.info(f"  - {name}: {err}")


if __name__ == "__main__":
    main()
