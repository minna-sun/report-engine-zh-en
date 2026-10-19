"""PDF 提取服务（新增 · 2026-09-15）。

职责：从 PDF 中提取文本和嵌入图片，供 IonContaminationStrategy 使用。
依赖：PyMuPDF（pip install PyMuPDF）
"""

import os
import tempfile

from common.logger import info, warning

# 临时目录（支持环境变量覆盖，用于多进程隔离）
_TASK_ID = os.environ.get("REPORT_PLATFORM_TASK_ID", "")
_TEMP_BASE = os.environ.get("REPORT_PLATFORM_TEMP", tempfile.gettempdir())
_TEMP_DIR = os.path.join(
    _TEMP_BASE,
    "report_platform_pdf" + (f"_{_TASK_ID}" if _TASK_ID else "")
)
os.makedirs(_TEMP_DIR, exist_ok=True)


def extract_text(pdf_path: str) -> str:
    """提取 PDF 所有页的文本，返回拼接后的字符串。失败返回空字符串。"""
    try:
        import fitz
    except ImportError:
        warning("[pdf_extractor] 未安装 PyMuPDF，无法读取 PDF")
        return ""
    try:
        doc = fitz.open(pdf_path)
        texts = []
        for page in doc:
            texts.append(page.get_text())
        doc.close()
        return "\n".join(texts)
    except Exception as e:
        warning(f"[pdf_extractor] 读取 PDF 失败 {pdf_path}：{e}")
        return ""


def extract_main_image(pdf_path: str, output_path: str = None):
    """从 PDF 中截取坐标图区域（按固定页面比例），保存为 PNG。

    源报告为纯扫描件，无文字层和完整嵌入图片，因此采用整页截图 + 比例裁切。
    截图范围：页面高度 54.6% ~ 90.9%，页面全宽。

    Args:
        pdf_path: PDF 文件路径
        output_path: 输出 PNG 路径；None 时自动生成到临时目录

    Returns:
        成功：输出 PNG 路径
        失败：None
    """
    try:
        import fitz
    except ImportError:
        warning("[pdf_extractor] 未安装 PyMuPDF，无法提取图片")
        return None
    try:
        doc = fitz.open(pdf_path)
        if output_path is None:
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            output_path = os.path.join(_TEMP_DIR, f"{base}_main.png")
        # 取第一页（源报告通常只有 1 页）
        page = doc[0]
        page_rect = page.rect
        # 按比例计算裁切区域
        # 按 18 等分取第 10.5~15.5 份（整体上移 0.5 份）
        y_start = page_rect.height * (9.5 / 18)
        y_end = page_rect.height * (15.5 / 18)
        clip = fitz.Rect(0, y_start, page_rect.width, y_end)
        # 300 DPI 高分辨率截图
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, clip=clip)
        pix.save(output_path)
        doc.close()
        info(f"[pdf_extractor] 已截取坐标图 {output_path} "
             f"({pix.width}×{pix.height}, 裁切范围 y={y_start:.1f}~{y_end:.1f})")
        return output_path
    except Exception as e:
        warning(f"[pdf_extractor] 截取坐标图失败 {pdf_path}：{e}")
        return None


def extract_full_page(pdf_path: str, output_path: str = None):
    """截取 PDF 第一页整页，保存为 PNG（300 DPI）。

    Args:
        pdf_path: PDF 文件路径
        output_path: 输出 PNG 路径；None 时自动生成到临时目录

    Returns:
        成功：输出 PNG 路径
        失败：None
    """
    try:
        import fitz
    except ImportError:
        warning("[pdf_extractor] 未安装 PyMuPDF，无法截取整页")
        return None
    try:
        doc = fitz.open(pdf_path)
        if output_path is None:
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            output_path = os.path.join(_TEMP_DIR, f"{base}_full.png")
        page = doc[0]
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        pix.save(output_path)
        doc.close()
        info(f"[pdf_extractor] 已截取整页 {output_path} ({pix.width}×{pix.height})")
        return output_path
    except Exception as e:
        warning(f"[pdf_extractor] 截取整页失败 {pdf_path}：{e}")
        return None
