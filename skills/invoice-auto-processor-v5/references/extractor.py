"""PDF文本提取模块 - 使用 pymupdf"""
import fitz  # pymupdf


def extract_text_from_pdf(pdf_path):
    """从PDF中提取全部文本"""
    doc = fitz.open(pdf_path)
    pages = []
    full_text_parts = []

    for i, page in enumerate(doc):
        text = page.get_text()
        full_text_parts.append(text)
        pages.append({"page_num": i + 1, "text": text})

    doc.close()

    import os
    return {
        "full_text": "\n".join(full_text_parts),
        "pages": pages,
        "page_count": len(pages),
        "filename": os.path.basename(pdf_path)
    }
