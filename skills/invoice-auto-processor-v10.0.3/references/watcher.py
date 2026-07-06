"""文件夹监控模块 — 动态导入 + 去重 + 归档"""
import os, time, logging, shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from .extractor import extract_text_from_pdf
from . import parser as _parser
from .renamer import build_new_filename
from .writer import append_invoice_rows, get_output_path, get_processed_invoice_numbers, mark_invoice_processed

logger = logging.getLogger(__name__)


class InvoiceHandler(FileSystemEventHandler):
    def __init__(self, config):
        self.config = config
        self.watch_dir = config["paths"]["watch_dir"]
        self.output_dir = config["paths"]["output_dir"]
        self.archive_dir = config["paths"]["archive_dir"]
        self.excel_filename = config["excel"]["filename"]
        self.columns = config["excel"]["columns"]
        self.archive = config["processing"]["archive_after_process"]
        self.skip_duplicates = config["processing"]["skip_duplicates"]
        self.rename_enabled = config["processing"].get("rename_enabled", True)
        self.supported_exts = [e.lower() for e in config["processing"]["supported_extensions"]]
        for d in [self.watch_dir, self.output_dir, self.archive_dir]:
            os.makedirs(d, exist_ok=True)

    def on_created(self, event):
        if event.is_directory:
            return
        filepath = event.src_path
        if os.path.splitext(filepath)[1].lower() not in self.supported_exts:
            return
        time.sleep(0.5)
        if not os.path.exists(filepath):
            return
        logger.info(f"发现新文件: {filepath}")
        self._process_file(filepath)

    def _process_file(self, filepath):
        filename = os.path.basename(filepath)
        output_path = get_output_path(self.output_dir, self.excel_filename)
        try:
            # 1. 提取
            extracted = extract_text_from_pdf(filepath)
            text = extracted["full_text"]
            # 2. 解析
            invoice_data = _parser.parse_invoice(text)
            invoice_no = invoice_data.get("invoice_number", "")
            logger.info(f"[{filename}] 发票号码: {invoice_no}, 金额: {invoice_data.get('total_amount','')}")
            # 3. 计算新名
            new_filename = None
            if self.rename_enabled:
                new_filename = build_new_filename(invoice_data, os.path.splitext(filepath)[1])
            # 4. 去重 — v9: 重复发票同时跳过台账写入和归档
            if self.skip_duplicates and invoice_no:
                existing = get_processed_invoice_numbers(self.output_dir)
                if invoice_no in existing:
                    logger.warning(f"[{filename}] 发票号码 {invoice_no} 已处理过，跳过台账写入和归档")
                    return
            # 5. 写入 Excel
            append_invoice_rows(output_path, self.columns, invoice_data)
            if self.skip_duplicates and invoice_no:
                mark_invoice_processed(self.output_dir, invoice_no)
            logger.info(f"[{filename}] ✅ 处理完成")
            # 6. 归档
            if self.archive:
                self._archive_file(filepath, new_filename)
        except Exception as e:
            logger.error(f"[{filename}] ❌ 处理失败: {e}", exc_info=True)

    def _archive_file(self, filepath, new_filename=None):
        try:
            dest = os.path.join(self.archive_dir, new_filename or os.path.basename(filepath))
            if os.path.exists(dest):
                name, ext = os.path.splitext(new_filename or os.path.basename(filepath))
                dest = os.path.join(self.archive_dir, f"{name}_{int(time.time())}{ext}")
            shutil.copy2(filepath, dest)
            logger.info(f"[{os.path.basename(filepath)}] 已归档 → {os.path.basename(dest)}")
        except Exception as e:
            logger.warning(f"归档失败: {e}")


def start_watching(config):
    watch_dir = config["paths"]["watch_dir"]
    os.makedirs(watch_dir, exist_ok=True)
    event_handler = InvoiceHandler(config)
    observer = Observer()
    observer.schedule(event_handler, watch_dir, recursive=False)
    observer.start()
    logger.info(f"监控: {watch_dir} | Ctrl+C 停止")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logger.info("监控已停止")
    observer.join()


def process_existing_files(config):
    watch_dir = config["paths"]["watch_dir"]
    supported = [e.lower() for e in config["processing"]["supported_extensions"]]
    if not os.path.isdir(watch_dir):
        return
    files = [f for f in os.listdir(watch_dir) if os.path.splitext(f)[1].lower() in supported]
    if files:
        logger.info(f"发现 {len(files)} 个已有文件，开始批量处理...")
        handler = InvoiceHandler(config)
        for f in files:
            handler._process_file(os.path.join(watch_dir, f))
        logger.info("批量处理完成")
