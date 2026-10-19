"""全自动图片型策略（新增 · 2026-09-11）。
职责：在图片型章节中，自动扫描指定图片文件夹下的所有图片，
无需依赖物料配置表的占位符和图片名称映射关系，自动在
Result description 段落后创建表格，将图片依次插入。
表格结构：
  - 图片宽度 ≤ 9cm：2 列，总宽 18cm，每列 9cm
  - 图片宽度 > 9cm：1 列，总宽 18cm
  - 每张图片占用 2 行：上方单元格填图片名称（灰色打底、居中），
    下方单元格插入图片（居中）
  - 图片数量为奇数时，2 列模式的最后一格填 "/"
适用章节：所有类型为 ImageAuto 的测试章节。
边界：
- 不导入 python-docx / lxml / DocumentEngine。
- 不负责章节分页（由 GenericProcessor 统一处理）。
- 不负责公共字段填充（由 GenericProcessor 负责）。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
"""
import os
import re
from typing import Any
from strategies.base import BaseStrategy
from common.logger import info, warning
from template_engine import _doc, insert_image_auto_table
from models.result import ChapterResult
import re as _re


def _extract_width_from_name(name: str, default_cm: float) -> float:
    """从文件名提取宽度：匹配最后一个括号内的数字/小数（半角/全角）。

    例：'1-1(10)'    → 10.0
        '1-1（10.5）' → 10.5
        '1-1(1)(10)' → 10.0   (取最后一个)
        '1-1'        → default_cm
        '1-1(abc)'   → default_cm
    """
    matches = _re.findall(r"[（(](\d+(?:\.\d+)?)[）)]", name)
    if matches:
        try:
            return float(matches[-1])
        except (ValueError, TypeError):
            pass
    return default_cm


def _clean_name_for_display(name: str) -> str:
    """去掉文件名最后一个括号及其内容（半角/全角）。

    例：'1-1(10)'    → '1-1'
        '1-1（10.5）' → '1-1'
        '1-1(1)(10)' → '1-1(1)'   (只去最后一个)
        '1-1'        → '1-1'
    """
    m = _re.search(r"\s*[（(]\d+(?:\.\d+)?[）)]\s*$", name)
    if m:
        return name[:m.start()].rstrip()
    return name


def _natural_sort_key(name: str):
    """自然排序 key：将文件名中的数字段按整数比较。"""
    parts = re.split(r"(\d+)", name)
    return [int(p) if p.isdigit() else p.lower() for p in parts]
# 兼容全角【xxx】和半角 [xxx]，取文件名开头第一个
_MARKER_RE = _re.compile(r"^(?:【([^】]*)】|\[([^\]]*)\])(.*)$")


def _split_sort_prefix_and_display(name: str):
    """解析文件名开头的第一个标记（【xxx】或 [xxx]）。

    Returns:
        (sort_key, display_name):
          - 有标记 → (标记内容, 剥掉标记后的剩余)
          - 无     → (name, name)
    """
    m = _MARKER_RE.match(name)
    if m:
        marker = (m.group(1) if m.group(1) is not None else m.group(2)) or ""
        rest = m.group(3) or ""
        return marker.strip(), rest.strip()
    return name, name

def _sort_key_with_marker(name: str):
    """排序键：方括号项全部排前面，其余排后面；同组内按自然排序。"""
    marker, _ = _split_sort_prefix_and_display(name)
    if marker != name:                       # 有【...】
        return (0, _natural_sort_key(marker))
    return (1, _natural_sort_key(name))

def _process_image_auto_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理全自动图片章节：扫描图片 → 创建表格 → 插入图片。"""
    doc = _doc(None)
    config = chapter_ctx.config
    test_item = config.test_item
    source_folder = chapter_ctx.source_folder_name or config.folder_name
    width_cm = config.image_width or 7.0
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")
    if not source_folder:
        base.status = "skipped"
        base.error = "source_folder 为空"
        warning(f"[ImageAuto] {test_item}：source_folder 为空，跳过")
        return base
    img_dir = os.path.join(task_ctx.folder_path, "原始记录", source_folder)
    if not os.path.isdir(img_dir):
        base.status = "skipped"
        base.error = f"图片目录不存在 {img_dir}"
        warning(f"[ImageAuto] {test_item}：{base.error}")
        return base
    # 扫描顶层图片文件（不递归子目录）
    imgs = []
    try:
        for f in os.listdir(img_dir):
            low = f.lower()
            if low.startswith("~$"):
                continue
            if any(s in low for s in ("_processed", "thumbs.db", "desktop.ini")):
                continue
            if low.endswith((".jpg", ".jpeg", ".png", ".bmp", ".gif")):
                imgs.append(os.path.join(img_dir, f))
    except Exception as e:
        base.status = "failed"
        base.error = f"读取图片目录失败：{e}"
        warning(f"[ImageAuto] {test_item}：{base.error}")
        return base
    if not imgs:
        base.status = "skipped"
        base.error = f"图片目录无有效图片 {img_dir}"
        warning(f"[ImageAuto] {test_item}：{base.error}")
        return base
    # 自然排序
    imgs.sort(key=lambda p: _sort_key_with_marker(
        os.path.splitext(os.path.basename(p))[0]))

    # 为每张图提取：显示名称（去最后一个括号）+ 宽度（括号 > 配置 > 默认）
    image_list = []
    for p in imgs:
        raw_label = os.path.splitext(os.path.basename(p))[0]
        _, label_after_marker = _split_sort_prefix_and_display(raw_label)
        display_name = _clean_name_for_display(label_after_marker)
        img_width = _extract_width_from_name(raw_label, width_cm)
        image_list.append((display_name, p, img_width))
    info(f"[ImageAuto] {test_item}：共 {len(image_list)} 张图片，配置宽度={width_cm}cm")
    try:
        ok = insert_image_auto_table(doc, test_item, image_list, width_cm)
        if not ok:
            base.status = "failed"
            base.error = "插入表格失败"
    except Exception as e:
        base.status = "failed"
        base.error = str(e)
        warning(f"[ImageAuto] {test_item} 插入失败：{e}")
    return base


class ImageAutoStrategy(BaseStrategy):
    """全自动图片型策略——扫描文件夹下的图片并自动构建表格插入。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        try:
            result = _process_image_auto_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[ImageAutoStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
