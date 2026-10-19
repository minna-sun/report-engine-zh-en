"""电性策略（模块 13 · v5.0 表格模式）。

职责：在电性测试（Electrical）章节中，从电性 Excel 读取全量数据，经
``excel_table_builder`` 生成表格数据，并插入到该章节 ``Result description``
锚点之后（Word 原生表格，宽度 18cm）。

实现方式（v5.0）：原委托 template_engine.process_electrical_chapter 的业务编排
逻辑已内联到本模块（_process_electrical_chapter / _process_excel_chapter），
execute 直接调用本模块逻辑并补全 ChapterResult 公共字段。

适用章节：所有类型为 ``Electrical`` 的测试章节。

边界：
- 不导入 python-docx / lxml / DocumentEngine。
- 不调用 data_extractor / image_service（表格模式，非图片模式）。
- 不负责测试条件 / 测试标准 / 设备表的查询（由 reference_service / GenericProcessor 处理）。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
"""

from typing import Any

from services import language_pack
from strategies.base import BaseStrategy
from models.result import ChapterResult
from common.logger import warning
from template_engine import _doc, _eng, _insert_table_after
import excel_table_builder


def _process_excel_chapter(chapter_ctx, task_ctx, skip_defaults: bool,
                           require_header: bool = True) -> ChapterResult:
    doc = _doc(None)
    eng = _eng()
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
    skip_columns = (2, 3) if skip_defaults else ()
    skip_rows = tuple(range(22, 28)) if skip_defaults else ()
    table_data = excel_table_builder.build_table_data(
        task_ctx.folder_path, source_folder,
        table_width_cm=18.0, skip_columns=skip_columns, skip_rows=skip_rows,
        require_header=require_header)
    if table_data is None:
        base.status = "skipped"
        base.error = "未读取到 Excel 数据"
        warning(f"[process_*_chapter] {test_item}：{base.error}")
        return base
    ok = _insert_table_after(test_item, anchor, table_data, 18.0)
    if not ok:
        base.status = "failed"
        base.error = "插入表格失败"
    return base


def _process_electrical_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理电性章节：Excel → Word 表格（全量，不跳过行列，不查找表头）。"""
    return _process_excel_chapter(chapter_ctx, task_ctx, skip_defaults=False,
                                  require_header=False)


class ElectricalTableStrategy(BaseStrategy):
    """电性策略——Excel → Word 原生表格（逻辑内联）。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行电性策略：本模块内联逻辑完成 Excel→Word 表格插入。

        Args:
            template_engine: template_engine 模块/实例（兼容形参）。
            chapter_ctx: 章节执行上下文（含 config.test_item, config.folder_name 等）。
            task_ctx: 任务执行上下文（含 folder_path, report_language 等）。

        Returns:
            ChapterResult：章节执行结果（status / error 等由本模块填充，公共字段经 _finalize 补全）。
        """
        try:
            result = _process_electrical_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[ElectricalTableStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
