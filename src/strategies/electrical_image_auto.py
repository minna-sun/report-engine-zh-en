"""电性 + 全自动图片复合策略（新增 · 2026-09-11）。
职责：在同一章节中依次执行电性图片插入（复用 ElectricalTableStrategy 逻辑）
和全自动图片表格插入（复用 ImageAutoStrategy 逻辑）。两部分相互独立，任一失败
仅记录警告，不影响另一部分。
实现方式（v5.5）：两部分均直接调用对应策略的 execute 方法；靠 addnext 的自然
插入行为实现最终顺序（先电性后图片）。
  - 先调 ImageAutoStrategy → 图片表格紧跟在 Result description 后
  - 再调 ElectricalTableStrategy → 电性图片插在 Result description 和图片表格之间
  - 最终顺序：Result description → 电性图片 → 图片表格
适用章节：同时含图片和电性 Excel 数据的章节（物料类别配置表中 strategy 列填
ElectricalImageAutoStrategy，类型列填 ElectricalImageAuto）。
边界（同 ImageAndElectricalStrategy）：
- 不导入 python-docx / lxml / DocumentEngine。
- 不负责章节分页（由 GenericProcessor 统一处理）。
- 不负责公共字段填充（由 GenericProcessor 负责）。
- 不负责设备数据 / 测试标准查询（由 reference_service 处理）。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
"""
from typing import Any

from strategies.base import BaseStrategy
from strategies.image_auto import ImageAutoStrategy
from strategies.electrical_table import ElectricalTableStrategy
from common.logger import info, warning
from template_engine import _doc, _eng, insert_empty_para_after
from services import language_pack
from services.ooxml_service import _q


class ElectricalImageAutoStrategy(BaseStrategy):
    """电性 + 全自动图片复合策略——先插电性图片，再插图片表格。"""

    def __init__(self):
        self._image_auto_strategy = ImageAutoStrategy()
        self._electrical_strategy = ElectricalTableStrategy()

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行复合策略，合并两部分结果。

        最终顺序：Result description → 电性表格 → 空行 → 图片表格
        """
        test_item = chapter_ctx.config.test_item

        # ── 0. 预先定位 Result description 段落（插空行用）──
        doc = _doc(None)
        eng = _eng()
        rd_para = None
        try:
            section = eng.find_section(doc, test_item) if doc is not None else None
            if section is not None:
                lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
                label = language_pack.labels(lang)["chapter_result_description"]
                rd_para = eng.find_paragraph(doc, section, label)
                if rd_para is None:
                    for _lbl in ("Result description", "结果说明"):
                        rd_para = eng.find_paragraph(doc, section, _lbl)
                        if rd_para:
                            break
        except Exception as e:
            warning(f"[ElectricalImageAuto] 定位 RD 段落失败：{e}")

        # ── 第一部分：图片表格（先插入，会被后插的挤到后面）──
        img_result = None
        try:
            img_result = self._image_auto_strategy.execute(
                template_engine, chapter_ctx, task_ctx)
            info(f"[ElectricalImageAuto] 图片表格插入完成: {test_item}")
        except Exception as e:
            warning(f"[ElectricalImageAuto] 图片表格插入异常，跳过 ({test_item}): {e}")

        # 图片成功 → 在 RD 后插空行（会位于电性表格和图片表格之间）
        img_ok = (img_result is not None and img_result.status == "success")
        blank_inserted = False
        if img_ok and rd_para is not None:
            try:
                insert_empty_para_after(doc, rd_para)
                blank_inserted = True
            except Exception as e:
                warning(f"[ElectricalImageAuto] 插空行失败：{e}")

        # ── 第二部分：电性表格（后插入，挤到 RD 和空行之间）──
        elec_result = None
        try:
            elec_result = self._electrical_strategy.execute(
                template_engine, chapter_ctx, task_ctx)
            info(f"[ElectricalImageAuto] 电性表格插入完成: {test_item}")
        except Exception as e:
            warning(f"[ElectricalImageAuto] 电性表格插入异常，跳过 ({test_item}): {e}")

        # 电性失败 → 删掉刚插的空行（避免留下孤立空行）
        elec_ok = (elec_result is not None and elec_result.status == "success")
        if blank_inserted and not elec_ok:
            try:
                first_el = rd_para._p.getnext()
                if first_el is not None and first_el.tag == _q("p"):
                    text = ''.join((t.text or "") for t in first_el.iter(_q("t"))).strip()
                    if not text:
                        parent = first_el.getparent()
                        if parent is not None:
                            parent.remove(first_el)
            except Exception as e:
                warning(f"[ElectricalImageAuto] 删空行失败：{e}")

        return self._merge_results(test_item, chapter_ctx, task_ctx,
                                   img_result, elec_result)

    # ───────────────────────── 结果合并 ─────────────────────────

    def _merge_results(self, test_item, chapter_ctx, task_ctx,
                        img_result, elec_result):
        from models.result import ChapterResult
        statuses = [r.status for r in (img_result, elec_result) if r is not None]
        if not statuses:
            merged = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                                   test_result="", status="skipped",
                                   error="两部分均未执行（无数据）")
        else:
            if any(s == "success" for s in statuses):
                status, error = "success", None
            elif any(s == "failed" for s in statuses):
                status, error = "failed", "图片或电性插入失败"
            else:
                status, error = "skipped", None
            merged = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                                   test_result="", status=status, error=error)
        return self._finalize(merged, chapter_ctx, task_ctx)
