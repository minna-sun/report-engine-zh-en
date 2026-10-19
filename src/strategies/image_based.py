"""图片型策略（模块 11 · v5.0）。

职责：在图片型测试章节中，将图片插入到对应的占位符位置。

实现方式（v5.0）：原委托 template_engine.process_image_chapter 的业务编排逻辑
已内联到本模块（_process_image_chapter），图片匹配辅助函数（_list_images /
_parse_image_mapping / _match_one）一并迁为本模块私有函数，execute 直接调用
本模块逻辑并补全 ChapterResult 公共字段，降低对 template_engine 间接层依赖。

边界（见 重构.md §6.4）：
- 不导入 python-docx / lxml / DocumentEngine。
- 不负责设备/测试标准/Test Result 查询（由 reference_service / GenericProcessor 处理）。
- 不负责图片裁剪/缩放（由 image_service 处理）。
- 不负责图片的 Word 插入操作细节（由 template_engine 执行）。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
- 不负责解析物料类别配置（由 section_manager 预先组装 ChapterContext）。
"""

import os
from typing import Any

from strategies.base import BaseStrategy
from common.logger import info, warning
from template_engine import (
    _doc,
    _eng,
    _strip_brackets,
    insert_image_at_placeholder,
)
from models.result import ChapterResult
from strategies.image_auto import _extract_width_from_name   # ← 新增


# ───────────────────────── 图片匹配辅助（私有） ─────────────────────────

def _parse_image_mapping(mapping) -> dict:
    """解析 image_mapping（"1-1→front, 1-2→back"）为 {关键词: 占位符}。"""
    out = {}
    if not mapping:
        return out
    for part in str(mapping).split(","):
        part = part.strip()
        if "→" in part:
            k, v = part.split("→", 1)
        elif "->" in part:
            k, v = part.split("->", 1)
        else:
            continue
        out[_strip_brackets(k.strip())] = _strip_brackets(v.strip())
    return out


def _list_images(folder: str) -> list:
    """列出当前文件夹下的图片文件（不递归子目录，排除系统文件）。"""
    exts = (".jpg", ".jpeg", ".png", ".bmp", ".gif")
    skip = ("_processed", "thumbs.db", "desktop.ini")
    out = []
    try:
        entries = sorted(os.listdir(folder))
    except OSError:
        return out
    for f in entries:
        full = os.path.join(folder, f)
        if not os.path.isfile(full):
            continue
        low = f.lower()
        if low.startswith("~$"):
            continue
        if any(s in low for s in skip):
            continue
        if low.endswith(exts):
            out.append(full)
    return out


def _match_one(imgs: list, keyword: str):
    """三轮匹配：精确 → 子串 → 去后缀/同义词模糊。返回图片路径或 None。"""
    names = [(p, os.path.splitext(os.path.basename(p))[0]) for p in imgs]
    # 第一轮：精确匹配（去扩展名）
    for p, n in names:
        if n == keyword:
            return p
    # 第二轮：子串匹配
    for p, n in names:
        if keyword in n or n in keyword:
            return p
    # 第三轮：去后缀 + 同义词模糊
    suffixes = ["照片", "图片", "photo", "image", "pic", "图"]
    kw = keyword
    for s in suffixes:
        kw = kw.replace(s, "")
    synonyms = [("正面", "反面"), ("焊前", "可焊前"), ("前", "焊前")]
    cands = [kw]
    for a, b in synonyms:
        if a in kw:
            cands.append(kw.replace(a, b))
    for p, n in names:
        nn = n
        for s in suffixes:
            nn = nn.replace(s, "")
        for c in cands:
            if c and (c in nn or nn in c):
                return p
    return None


def _process_image_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理图片型章节：配置驱动定位，按占位符插入图片。"""
    doc = _doc(None)
    eng = _eng()
    config = chapter_ctx.config
    test_item = config.test_item
    source_folder = chapter_ctx.source_folder_name or config.folder_name
    placeholders = config.placeholders or []
    width_cm = config.image_width or 7.0
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")

    if not source_folder:
        base.status = "skipped"
        base.error = "source_folder 为空"
        warning(f"[process_image_chapter] {test_item}：source_folder 为空，跳过")
        return base

    img_dir = os.path.join(task_ctx.folder_path, "原始记录", source_folder)
    if not os.path.isdir(img_dir):
        base.status = "skipped"
        base.error = f"图片目录不存在 {img_dir}"
        warning(f"[process_image_chapter] {test_item}：{base.error}")
        return base

    imgs = _list_images(img_dir)
    kw_to_ph = _parse_image_mapping(config.image_mapping)
    ph_to_kw = {v: k for k, v in kw_to_ph.items()}

    matched = {}
    for ph in placeholders:
        kw = ph_to_kw.get(ph, ph)
        img = _match_one(imgs, kw)
        if img:
            matched[ph] = img
        else:
            warning(f"[process_image_chapter] {test_item}：占位符 {ph} 未匹配到图片")

    for ph, img in matched.items():
        try:
            _stem = os.path.splitext(os.path.basename(img))[0]
            _w = _extract_width_from_name(_stem, width_cm)   # 括号 → 配置表 → 7.0
            insert_image_at_placeholder(ph, img, _w, section_name=test_item)
        except Exception as e:
            base.status = "failed"
            base.error = str(e)
            warning(f"[process_image_chapter] {test_item} 插入失败：{e}")
    return base          # ← 新增这一行（缩进 4 空格）


class ImageBasedStrategy(BaseStrategy):
    """图片型策略——按占位符插入图片（逻辑内联）。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行图片型策略：本模块内联逻辑完成图片匹配与插入。

        Args:
            template_engine: template_engine 模块/实例（兼容形参，文档操作经内部 helper 完成）。
            chapter_ctx: 章节执行上下文（含 config.placeholders, config.image_mapping 等）。
            task_ctx: 任务执行上下文（含 folder_path 等）。

        Returns:
            ChapterResult：章节执行结果（status / error 等由本模块填充，公共字段经 _finalize 补全）。
        """
        try:
            result = _process_image_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[ImageBasedStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
