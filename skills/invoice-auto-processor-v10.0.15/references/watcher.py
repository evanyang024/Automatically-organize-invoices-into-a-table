"""文件夹监控模块 — 动态导入 + 去重 + 归档 + 递归扫描(子文件夹/压缩包)"""
import os, time, logging, shutil, zipfile, tempfile

try:
    import py7zr
except ImportError:
    py7zr = None
try:
    import rarfile
except ImportError:
    rarfile = None

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from .extractor import extract_text_from_pdf
from . import parser as _parser
from .renamer import build_new_filename
from .writer import append_invoice_rows, get_output_path, get_processed_invoice_numbers, mark_invoice_processed

logger = logging.getLogger(__name__)

# 支持的压缩包后缀
ARCHIVE_EXTS = {".zip", ".rar", ".7z"}
# rarfile 解压需要系统 unar/unrar
if rarfile is not None:
    rarfile.UNRAR_TOOL = shutil.which("unrar") or shutil.which("unar") or shutil.which("rar") or "unar"


def is_archive_file(path):
    """判断是否为支持的压缩包文件"""
    return os.path.splitext(path)[1].lower() in ARCHIVE_EXTS


def extract_archive(archive_path, extract_dir):
    """按压缩包类型解压到 extract_dir，返回解压出的文件列表"""
    os.makedirs(extract_dir, exist_ok=True)
    ext = os.path.splitext(archive_path)[1].lower()
    extracted_files = []

    if ext == ".zip":
        with zipfile.ZipFile(archive_path, "r") as zf:
            # 安全解压：防止路径穿越
            for member in zf.infolist():
                # 规范化成员路径，拒绝绝对路径 / .. 穿越
                normal = os.path.normpath(member.filename)
                if normal.startswith("..") or os.path.isabs(normal):
                    logger.warning(f"跳过不安全zip成员: {member.filename}")
                    continue
                member.filename = normal
            zf.extractall(extract_dir)
        # 收集所有文件
        for root, _, files in os.walk(extract_dir):
            for f in files:
                extracted_files.append(os.path.join(root, f))

    elif ext == ".7z" and py7zr is not None:
        with py7zr.SevenZipFile(archive_path, "r") as z:
            z.extractall(extract_dir)
        for root, _, files in os.walk(extract_dir):
            for f in files:
                extracted_files.append(os.path.join(root, f))

    elif ext == ".rar" and rarfile is not None:
        with rarfile.RarFile(archive_path, "r") as rf:
            rf.extractall(extract_dir)
        for root, _, files in os.walk(extract_dir):
            for f in files:
                extracted_files.append(os.path.join(root, f))

    else:
        logger.warning(f"不支持的压缩包或缺少解压库: {archive_path}")
        return []

    # 只返回 PDF（压缩包里可能还有其他文件）
    pdf_files = [f for f in extracted_files if os.path.splitext(f)[1].lower() == ".pdf"]
    logger.info(f"解压 {os.path.basename(archive_path)} → 找到 {len(pdf_files)} 个PDF (共{len(extracted_files)}个文件)")
    return pdf_files


def collect_invoice_files(watch_dir, supported_exts):
    """递归收集 input/ 下的所有待处理文件。

    策略：
    - os.walk 递归遍历子文件夹，收集所有 supported_exts 后缀的文件
    - 遇到压缩包（zip/rar/7z）→ 解压到临时目录 → 收集其中的 PDF
    - 返回 (文件列表, 待清理的临时目录列表, 待清理的压缩包列表)
    """
    if not os.path.isdir(watch_dir):
        return [], [], []

    pdf_files = []
    tmp_dirs = []          # 解压产生的临时目录（处理完要清理）
    archives = []          # 遇到的压缩包（归档后要清理）
    subdirs_to_clean = []  # 含PDF的子文件夹（处理完清理）
    archive_exts = ARCHIVE_EXTS
    skip = set()
    queue = list(os.walk(watch_dir))

    for root, dirs, files in queue:
        rel_root = os.path.relpath(root, watch_dir)
        # 子文件夹（非 watch_dir 本身，即 rel_root != "."）
        if rel_root != ".":
            subdirs_to_clean.append(root)

        for f in files:
            path = os.path.join(root, f)
            ext = os.path.splitext(f)[1].lower()

            # 支持的文档（pdf）
            if ext in supported_exts:
                pdf_files.append(path)
            # 压缩包
            elif ext in archive_exts:
                archives.append(path)
                # 解压到临时目录
                tmp = tempfile.mkdtemp(prefix="invoice_extract_")
                tmp_dirs.append(tmp)
                extracted_pdfs = extract_archive(path, tmp)
                pdf_files.extend(extracted_pdfs)

    # 排序，去重（同路径只处理一次）
    seen, unique_pdfs = set(), []
    for p in sorted(pdf_files):
        if p not in seen:
            seen.add(p)
            unique_pdfs.append(p)

    return unique_pdfs, tmp_dirs, archives, subdirs_to_clean


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
            # 4. 去重
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

    def _archive_then_remove(self, src, archive_dir):
        """归档一份副本到 archive_dir，然后删除源文件"""
        try:
            dest = os.path.join(archive_dir, os.path.basename(src))
            if os.path.exists(dest):
                name, ext = os.path.splitext(os.path.basename(src))
                dest = os.path.join(archive_dir, f"{name}_{int(time.time())}{ext}")
            shutil.copy2(src, dest)
            logger.info(f"[{os.path.basename(src)}] 已备份归档 → {os.path.basename(dest)}")
            os.remove(src)
            logger.info(f"[{os.path.basename(src)}] 源压缩包已删除")
        except Exception as e:
            logger.warning(f"压缩包归档/删除失败: {e}")


def start_watching(config):
    watch_dir = config["paths"]["watch_dir"]
    os.makedirs(watch_dir, exist_ok=True)
    event_handler = InvoiceHandler(config)
    observer = Observer()
    observer.schedule(event_handler, watch_dir, recursive=True)
    observer.start()
    logger.info(f"监控: {watch_dir} (递归子文件夹+压缩包) | Ctrl+C 停止")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logger.info("监控已停止")
    observer.join()


def process_existing_files(config):
    """批量处理 input/ 下所有待处理文件（含子文件夹和压缩包内PDF）"""
    watch_dir = config["paths"]["watch_dir"]
    archive_dir = config["paths"]["archive_dir"]
    supported = [e.lower() for e in config["processing"]["supported_extensions"]]
    if not os.path.isdir(watch_dir):
        return

    pdf_files, tmp_dirs, archives, subdirs = collect_invoice_files(watch_dir, supported)

    if not pdf_files and not archives:
        logger.info("没有找到待处理的发票文件")
        return

    logger.info(f"发现 {len(pdf_files)} 个PDF、{len(archives)} 个压缩包、{len(subdirs)} 个子文件夹，开始批量处理...")
    handler = InvoiceHandler(config)

    processed_pdfs = 0
    failed_pdfs = 0
    for pdf_path in pdf_files:
        try:
            handler._process_file(pdf_path)
            processed_pdfs += 1
        except Exception as e:
            failed_pdfs += 1
            logger.error(f"处理失败: {pdf_path}: {e}")

    # 清理：压缩包归档到 archive/ 后删除 + 清理临时解压目录
    # 注意：按铁律「不删原文件」，子文件夹里的源PDF保留不删，子文件夹不做自动清理
    handler.archive = True  # 确保归档生效
    for archive_path in archives:
        handler._archive_then_remove(archive_path, archive_dir)

    # 清理临时解压目录（解压出的PDF已处理归档，临时目录可安全删除）
    for tmp in tmp_dirs:
        try:
            shutil.rmtree(tmp, ignore_errors=True)
            logger.info(f"已清理临时解压目录: {tmp}")
        except Exception as e:
            logger.warning(f"清理临时目录失败: {tmp}: {e}")

    logger.info(f"批量处理完成: 成功 {processed_pdfs} 张, 失败 {failed_pdfs} 张")
    if subdirs:
        logger.info(f"子文件夹中的源PDF已保留（不删原文件铁律），共 {len(subdirs)} 个子文件夹. 已处理发票清单见 output/.processed_invoices.json")
