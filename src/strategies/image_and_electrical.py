"""图片+电性复合策略（模块 15-ext · v5.0）。

职责：在同一章节中依次执行图片插入（复用 ImageBasedStrategy 逻辑）和电性表格插入
（复用 ElectricalTableStrategy 逻辑）。两部分相互独立，任一失败仅记录警告，不影响另一部分。

实现方式（v5.3）：两部分分别委托给 ImageBasedStrategy 和 ElectricalTableStrategy 的 execute 方法。

适用章节：同时含图片和电性 Excel 数据的章节（在物料类别配置表中 strategy 列填写
``ImageAndElectricalStrategy``）。

边界（见 重构.md §6.4）：
- 不导入 python-docx / lxml / DocumentEngine。
- 不负责测试条件 / 测试标准 / 设备表的查询（由 reference_service / GenericProcessor 处理）。
- 不负责 Word 插入细节（由 template_engine 执行）。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
"""

from typing import Any

from strategies.base import BaseStrategy
from strategies.image_based import ImageBasedStrategy
from strategies.electrical_table import ElectricalTableStrategy
from common.logger import info, warning
from template_engine import _doc, _eng, insert_empty_para_after
from services import language_pack


class ImageAndElectricalStrategy(BaseStrategy):
    """图片+电性复合策略——先插入图片，再插入电性 Excel 表格。"""

    def __init__(self):
        self._image_strategy = ImageBasedStrategy()
        self._electrical_strategy = ElectricalTableStrategy()

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行图片+电性复合策略，合并两部分结果。

        最终顺序：Result description → 电性表格 → 空行 → 图片
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
            warning(f"[ImageAndElectricalStrategy] 定位 RD 段落失败：{e}")

        # ── 第一部分：图片插入（先插入，会被后插的电性挤到后面）──
        img_result = None
        try:
            img_result = self._image_strategy.execute(template_engine, chapter_ctx, task_ctx)
            info(f"[ImageAndElectricalStrategy] 图片插入完成: {test_item}")
        except Exception as e:
            warning(f"[ImageAndElectricalStrategy] 图片插入异常，跳过 ({test_item}): {e}")

        # 图片成功 → 在 RD 后插空行（会位于电性表格和图片之间）
        img_ok = (img_result is not None and img_result.status == "success")
        blank_inserted = False
        if img_ok and rd_para is not None:
            try:
                insert_empty_para_after(doc, rd_para)
                blank_inserted = True
            except Exception as e:
                warning(f"[ImageAndElectricalStrategy] 插空行失败：{e}")

        # ── 第二部分：电性表格插入（后插入，挤到 RD 和空行之间）──
        elec_result = None
        try:
            elec_result = self._electrical_strategy.execute(template_engine, chapter_ctx, task_ctx)
            info(f"[ImageAndElectricalStrategy] 电性表格插入完成: {test_item}")
        except Exception as e:
            warning(f"[ImageAndElectricalStrategy] 电性表格插入异常，跳过 ({test_item}): {e}")

        # 电性失败 → 删掉刚插的空行（避免留下孤立空行）
        elec_ok = (elec_result is not None and elec_result.status == "success")
        if blank_inserted and not elec_ok:
            try:
                from services.ooxml_service import _q as _q_local
                first_el = rd_para._p.getnext()
                if first_el is not None and first_el.tag == _q_local("p"):
                    text = ''.join((t.text or "") for t in first_el.iter(_q_local("t"))).strip()
                    if not text:
                        parent = first_el.getparent()
                        if parent is not None:
                            parent.remove(first_el)
            except Exception as e:
                warning(f"[ImageAndElectricalStrategy] 删空行失败：{e}")

        return self._merge_results(test_item, chapter_ctx, task_ctx, img_result, elec_result)
    # ───────────────────────── 结果合并 ─────────────────────────

    def _merge_results(self, test_item, chapter_ctx, task_ctx, img_result, elec_result):
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
                status, error = "failed", "图片或电性表格插入失败"
            else:
                # 全部为 skipped（无数据）→ 整体视为 skipped，不污染 PASS/FAIL
                status, error = "skipped", None
            merged = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                                   test_result="", status=status, error=error)
        return self._finalize(merged, chapter_ctx, task_ctx)
