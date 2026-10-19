"""报告生成引擎（模块 21，v5.0 重构 · 纯调度器）。

按固定顺序串联 Processor / Strategy / template_engine 完成报告生成。v5.0 起
**不再使用 win32com / Word COM**，文档统一经 ``engine_factory.get_engine()``
（python-docx 实现）打开与保存；COM 不进入数据模型层（TaskContext 仅纯数据）。

一票否决项（spec §七，跨技术栈仍适用）：
1. 单章节失败继续，不中断其它章节；
2. 结论只依赖 test_result（不受章节填充成败影响）；
3. 任一路径最终都返回统一的 ReportResult（成功必有 report_path + final_verdict）。
"""

import os
import shutil

from common.logger import info, error, warning
from models.result import ReportResult
from models.context import TaskContext
from processors.sample_processor import execute as sample_execute
from section_manager import build_chapters, get_chapter_template_name
from services.engine_factory import get_engine
from template_engine import (
    sync_chapters,
    highlight_expired_calibrations,
    set_active_doc,
)
import template_engine
from processor_registry import get_processor
from processors.summary_processor import execute as summary_execute
from services import language_pack


def generate_report(template_path, record, folder_path, output_path, config,
                    report_language: str = "en", standalone_mode: bool = False):
    """生成单份报告（纯调度器）。步骤 1→12（v5.0，去 COM）：复制模板→python-docx
    打开→建TaskContext→sample→build_chapters→sync→遍历GenericProcessor→summary→
    标红→保存→返回结果。

    Args:
        report_language: 报告语言（``"zh"`` / ``"en"``）。由 ``main`` 按「物料表指定 >
            模板自动识别 > 默认 en」决策后传入；config.yaml 不提供任何语言配置。
            各 Processor 从 TaskContext 读取，不再逐层透传。
    """
    if not config or not config.get("default_chapter_template"):
        return ReportResult(status="failed", error="配置缺失 default_chapter_template")

    # 语言归一（非法取值回退英文）
    report_language = language_pack.normalize(report_language)

    # 路径归一
    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(output_path)

    # 步骤 1：复制模板到输出路径
    if not os.path.isfile(template_path):
        return ReportResult(status="failed", error="报告模板不存在")
    try:
        shutil.copy2(template_path, output_path)
    except Exception as exc:
        return ReportResult(status="failed", error=f"模板复制失败：{exc}")

    doc = None
    eng = None
    hard_error = None
    final_verdict = None

    try:
        # 步骤 2：经 DocumentEngine 打开文档（python-docx，跨平台，无 COM 进程）
        eng = get_engine()
        doc = eng.open(output_path)
        if doc is None:
            return ReportResult(status="failed", error="报告模板打开失败")
        # 设置模块级活跃文档，供 Strategy 经 template_engine.process_*_chapter 操作
        set_active_doc(doc)

        # 步骤 3：创建 TaskContext（纯数据，COM 对象不放入）
        # report_language 供各 Processor / template_engine 切换锚点与翻译行为
        task_ctx = TaskContext(record=record, folder_path=folder_path,
                               config=config, output_path=output_path,
                               report_language=report_language,
                               standalone_mode=standalone_mode)
        info(f"[report_engine] 报告语言: {report_language}")

        # 步骤 4：sample_processor（步骤1~3：页眉/封面/Sample Description）
        try:
            sample_result = sample_execute(doc, task_ctx)
            if sample_result is not None and sample_result.status == "failed":
                hard_error = f"步骤1~3填充失败：{sample_result.error}"
                error(f"[report_engine] {hard_error}")
                return ReportResult(status="failed", error=hard_error)
        except Exception as exc:
            hard_error = f"步骤1~3填充失败：{exc}"
            error(f"[report_engine] {hard_error}")
            return ReportResult(status="failed", error=hard_error)

        # 步骤 5：section_manager 建立章节框架
        try:
            chapter_contexts = build_chapters(folder_path, record, lang=report_language)
        except Exception as exc:
            hard_error = f"章节框架建立失败：{exc}"
            error(f"[report_engine] {hard_error}")
            return ReportResult(status="failed", error=hard_error)

        # 步骤 6：template_engine 同步章节（保留/删除/新增）
        # 模板章节名取值优先级：
        #   1. 物料类别配置表中的「新增章节范本」列（推荐，按类别配置）
        #   2. config.yaml 的 default_chapter_template_{lang}
        #   3. config.yaml 的 default_chapter_template（向后兼容单语配置）
        try:
            # 6.1 选模板章节名
            _chapter_tpl = ""
            try:
                _material_category = getattr(record, "material_category", "") or ""
                _chapter_tpl = get_chapter_template_name(
                    _material_category, report_language)
            except Exception as _e:
                warning(f"[report_engine] 读物料配置表的章节范本失败：{_e}")

            if not _chapter_tpl:
                _lang_key = f"default_chapter_template_{report_language}"
                _chapter_tpl = (config.get(_lang_key) or "").strip() \
                               or config["default_chapter_template"]

            # 6.2 记录"模板原本的第一个章节名"（供 add_page_break 判断是否跳过）
            try:
                _sections_before = eng.list_sections(doc)
                _first_name = (_sections_before[0][0] or "").strip() \
                              if _sections_before else ""
                template_engine.set_first_section_name(_first_name)
            except Exception as _e:
                warning(f"[report_engine] 记录模板首章节失败（忽略）：{_e}")

            # 6.3 执行同步
            sync_chapters(doc, chapter_contexts, _chapter_tpl, lang=report_language)
        except Exception as exc:
            hard_error = f"章节框架操作失败：{exc}"
            error(f"[report_engine] {hard_error}")
            return ReportResult(status="failed", error=hard_error)

        # 步骤 7：遍历章节，调 GenericProcessor 填充（单章节失败继续，不中断）
        chapter_results = []
        for chapter_ctx in chapter_contexts:
            test_item = chapter_ctx.config.test_item
            try:
                processor = get_processor(chapter_ctx.config.category)
                if processor is None:
                    warning(f"[report_engine] 章节 '{test_item}' 无对应 Processor，跳过")
                    continue
                result = processor.execute(doc, chapter_ctx, task_ctx)
                if result is not None:
                    chapter_results.append(result)
                    if result.status == "failed":
                        error(f"[report_engine] 章节 {test_item} 填充失败：{result.error}")
            except Exception as exc:
                error(f"[report_engine] 章节 {test_item} 填充失败：{exc}")

        # 步骤 8：summary_processor 生成汇总页+最终结论
        # （结论只依赖 test_result，不受章节填充成败影响）
        try:
            final_verdict = summary_execute(doc, chapter_results, task_ctx)
        except Exception as exc:
            hard_error = f"汇总页生成失败：{exc}"
            error(f"[report_engine] {hard_error}")
            return ReportResult(status="failed", error=hard_error)

    except Exception as exc:
        if hard_error is None:
            hard_error = f"报告生成异常：{exc}"

    finally:
        # 步骤 9：校准有效期标红（成功路径执行；失败仅警告，不中断保存）
        if doc is not None and hard_error is None:
            try:
                highlight_expired_calibrations(doc)
            except Exception as exc:
                warning(f"[report_engine] 校准有效期标红失败：{exc}")

        # 步骤 10：保存文档（仅无硬失败且文档存在时；保存失败记录为硬失败）
        saved = False
        if doc is not None and hard_error is None:
            try:
                saved = eng.save(doc, output_path)
            except Exception as exc:
                error(f"[report_engine] 报告保存失败：{exc}")
                hard_error = hard_error or f"报告保存失败：{exc}"

        if doc is not None and eng is not None:
            try:
                eng.close(doc)
            except Exception as e:
                warning(f"[report_engine] 文档关闭失败：{e}")

    # 步骤 11~12：返回统一结果
    if hard_error or not saved:
        return ReportResult(status="failed", error=hard_error or "报告保存失败")
    return ReportResult(status="success", report_path=output_path,
                        pdf_path=None, final_verdict=final_verdict, error=None)


__all__ = ["generate_report"]
