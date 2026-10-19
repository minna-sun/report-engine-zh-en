"""电性策略（模块 13 · Word 表格版）。

职责：在电性测试（Electrical）章节中，从源 Excel 读取数据，
      构建 Word 表格并插入到该章节 Result description 段落的下一行，
      表格宽度固定 18cm，字体统一 Times New Roman 五号。

适用章节：所有类型为 Electrical 的测试章节。

边界：
- 不导入 python-docx / lxml / DocumentEngine。
- 不渲染 PNG。
- 不负责测试条件 / 测试标准 / 设备表的查询。
- 不负责判断章节是否应该存在（由 section_manager 决定）。
"""

import os
from typing import Any

import excel_table_builder
from services import language_pack
from strategies.base import BaseStrategy
from models.result import ChapterResult
from common.logger import warning
from template_engine import _doc, _eng, insert_word_table_after_para


class ElectricalTableStrategy(BaseStrategy):
    """电性策略——Excel → Word 表格（固定宽度 18cm）。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        test_item = chapter_ctx.config.test_item
        source_folder = chapter_ctx.source_folder_name or chapter_ctx.config.folder_name
        base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                             test_result="", status="success")

        if not source_folder:
            base.status = "skipped"
            base.error = "source_folder 为空"
            return self._finalize(base, chapter_ctx, task_ctx)

        try:
            # 1. 定位源文件夹里第一个 xlsx
            src_dir = os.path.join(task_ctx.folder_path, "原始记录", source_folder)
            if not os.path.isdir(src_dir):
                base.status = "skipped"
                base.error = f"目录不存在 {src_dir}"
                warning(f"[ElectricalTableStrategy] {test_item}：{base.error}")
                return self._finalize(base, chapter_ctx, task_ctx)

            xlsx_files = [f for f in os.listdir(src_dir)
                          if f.lower().endswith(".xlsx") and not f.startswith("~$")]
            if not xlsx_files:
                base.status = "skipped"
                base.error = "未找到 .xlsx"
                warning(f"[ElectricalTableStrategy] {test_item}：{base.error}")
                return self._finalize(base, chapter_ctx, task_ctx)

            xlsx_path = os.path.join(src_dir, sorted(xlsx_files)[0])

            # 2. 构建表格数据（字体统一 Times New Roman 五号）
            table_data = excel_table_builder.build_table_data(
                xlsx_path=xlsx_path,
                table_width_cm=18.0,
                skip_columns=(), skip_rows=(),
                require_header=False,
                force_font_name="Times New Roman",
                force_font_size=10.5,
            )
            if table_data is None:
                base.status = "skipped"
                base.error = "表格数据构建失败"
                warning(f"[ElectricalTableStrategy] {test_item}：{base.error}")
                return self._finalize(base, chapter_ctx, task_ctx)

            # 3. 定位 Result description 段落
            doc = _doc(None)
            eng = _eng()
            section = eng.find_section(doc, test_item)
            if section is None:
                base.status = "failed"
                base.error = f"未找到章节 {test_item}"
                return self._finalize(base, chapter_ctx, task_ctx)

            lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
            anchor_label = language_pack.labels(lang)["chapter_result_description"]
            anchor_para = eng.find_paragraph(doc, section, anchor_label)
            if anchor_para is None:
                for _lbl in ("Result description", "结果说明"):
                    anchor_para = eng.find_paragraph(doc, section, _lbl)
                    if anchor_para:
                        break
            if anchor_para is None:
                base.status = "failed"
                base.error = "未找到 Result description 段落"
                warning(f"[ElectricalTableStrategy] {test_item}：{base.error}")
                return self._finalize(base, chapter_ctx, task_ctx)

            # 4. 插入 Word 表格
            ok = insert_word_table_after_para(anchor_para, table_data, 18.0)
            if not ok:
                base.status = "failed"
                base.error = "表格插入失败"
                warning(f"[ElectricalTableStrategy] {test_item}：{base.error}")
            else:
                # 5. 表格下方插图片（如果 Excel 里有嵌入图片）
                excel_images = table_data.get("images") or []
                if excel_images:
                    img_width = chapter_ctx.config.image_width or 7.0
                    try:
                        from template_engine import insert_excel_images_after_table
                        insert_excel_images_after_table(
                            anchor_para, excel_images, img_width)
                    except Exception as e:
                        warning(f"[ElectricalTableStrategy] {test_item}："
                                f"Excel 图片插入失败（不影响表格）：{e}")

            return self._finalize(base, chapter_ctx, task_ctx)

        except Exception as e:
            base.status = "failed"
            base.error = str(e)
            warning(f"[ElectricalTableStrategy] {test_item} 异常：{e}")
            return self._finalize(base, chapter_ctx, task_ctx)