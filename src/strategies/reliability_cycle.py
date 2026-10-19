"""可靠性策略（模块 12 · v5.3）。

职责：在可靠性循环（ReliabilityCycle）测试章节中，将电性 Excel 数据以
PNG 图片形式插入到该章节 ``Result description`` 行的下一行。

实现方式（v5.3）：
  1. 调用 ``data_extractor.extract_electrical`` 读取 Excel 并生成 PNG（300 DPI）。
  2. 根据列数决定显示宽度（≤5 列 → 8 cm，>5 列 → 18 cm）。
  3. 调用 ``image_service.process_image`` 裁剪白边。
  4. 调用 ``template_engine.insert_image_after_text`` 将图片插入 Word。
  5. 在 finally 块中清理两个临时文件。

适用章节：所有类型为 ``ReliabilityCycle`` 的测试章节（如 HTRB / HTGB /
H3TRB / TC / HAST / UHAST / 双85 / 高温存储 等，见 1014.md）。

边界（见 设计说明书 §6.4）：
  - 不导入 python-docx / lxml / DocumentEngine。
  - 不直接操作 Excel（由 data_extractor 负责）。
  - 不直接操作 Word（由 template_engine.insert_image_after_text 负责）。
  - 不直接处理图片白边（由 image_service 负责）。
  - 不负责测试条件 / 测试标准 / 设备表的查询（由 reference_service / GenericProcessor 处理）。
  - 不负责判断章节是否应该存在（由 section_manager 决定）。
  - ❌ 不得调用 excel_table_builder。
"""

import os
from typing import Any

from strategies.base import BaseStrategy
from common.logger import info, warning
import data_extractor
from services import image_service
from services import language_pack
from template_engine import insert_image_after_text
from models.result import ChapterResult


def _process_reliability_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理可靠性章节：Excel → PNG 图片（v5.3：改回图片模式）。

    数据流：data_extractor → image_service（裁剪白边）→ insert_image_after_text。
    列数 ≤5 → 8 cm；>5 → 18 cm。临时文件在 finally 块中清理。
    """
    test_item = chapter_ctx.config.test_item
    source_folder = chapter_ctx.source_folder_name or chapter_ctx.config.folder_name
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    anchor = language_pack.labels(lang)["chapter_result_description"]
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")

    if not source_folder:
        base.status = "skipped"
        base.error = "source_folder 为空"
        return base

    original_path = None
    processed_path = None
    try:
        result = data_extractor.extract_electrical(
            task_ctx.folder_path, source_folder,
            force_header_row=16)
        if result is None:
            base.status = "skipped"
            base.error = "电性数据导出失败"
            warning(f"[process_reliability_chapter] {test_item}：{base.error}")
            return base

        original_path = result["image_path"]
        col_count = result["col_count"]
        width_cm = 8.0 if col_count <= 5 else 18.0

        try:
            processed_path = image_service.process_image(original_path)
        except Exception as e:
            base.status = "failed"
            base.error = f"图片处理失败: {e}"
            warning(f"[process_reliability_chapter] {test_item}：{base.error}")
            return base

        ok = insert_image_after_text(None, test_item, anchor,
                                     processed_path, width_cm)
        if not ok:
            base.status = "failed"
            base.error = "图片插入失败"
            warning(f"[process_reliability_chapter] {test_item}：{base.error}")
        return base

    except Exception as e:
        base.status = "failed"
        base.error = str(e)
        warning(f"[process_reliability_chapter] {test_item} 异常: {e}")
        return base

    finally:
        if processed_path:
            image_service.cleanup_processed(processed_path)
        if original_path and os.path.exists(original_path):
            try:
                os.remove(original_path)
            except Exception as e:
                warning(f"[process_reliability_chapter] 清理原始图片失败: {e}")


class ReliabilityCycleStrategy(BaseStrategy):
    """可靠性循环策略——电性 Excel 转 PNG 图片后插入 Word。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行可靠性策略：本模块内联逻辑完成 Excel→PNG 插入。

        Args:
            template_engine: template_engine 模块（兼容形参）。
            chapter_ctx: 章节执行上下文（含 config.test_item, config.folder_name 等）。
            task_ctx: 任务执行上下文（含 folder_path, report_language 等）。

        Returns:
            ChapterResult：章节执行结果（status / error 等由本模块填充，公共字段经 _finalize 补全）。
        """
        try:
            result = _process_reliability_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[ReliabilityCycleStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
