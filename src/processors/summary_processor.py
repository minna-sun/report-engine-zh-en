"""汇总处理器（模块 19，v5.0 重构）。

职责：步骤5 汇总页生成，详见《业务执行规范》第 4.2.5 节。直接由
``core/report_engine`` 调用（不参与 ``processor_registry`` 调度，属于非类型化
章节处理器，与 ``sample_processor`` 同级）。

v5.0 迁移要点：
- 彻底移除 win32com/COM（旧版 ``doc.Tables`` / ``cell.Range.Text`` /
  ``para.Range.Find`` 等）；所有 Word 操作经 ``template_engine``（底层
  DocumentEngine / python-docx）。
- 业务过滤（排除 Sample Description）保留在本模块；表格填充、统计段落更新、
  Notes 填写、结论着色等统一委托 ``template_engine.fill_summary``。

边界：
- 不持有 COM 对象（doc 由 report_engine 打开并传入）
- 不负责判断 Test Result 正确性（由 GenericProcessor 按文件夹后缀判定后写入
  ``chapter_results``）
- 不负责生成 Test Basis 内容（由 reference_service 查询，GenericProcessor 写入）
- 不负责报告重命名（由 file_renamer.py 执行，本模块仅返回 final_verdict）

输入：
- ``doc``：已打开的 Word Document（由 report_engine 持有）
- ``chapter_results``：List[ChapterResult]，每个含 test_item / test_basis /
  sample_number / test_result / result_description / status / error
- ``task_ctx``：TaskContext（本模块不使用，保留以便未来扩展）

输出：
- ``final_verdict``：``"PASS"`` 或 ``"FAIL"``，供 report_engine → file_renamer 重命名
"""

from typing import List

from common.logger import info
from models.result import ChapterResult
from services import language_pack
import template_engine as te

# 汇总页章节标题 KEY（真实值按语言从 language_pack 取：英文 "Summary of Results"、
# 中文 "结果汇总"）。严禁在此硬编码任何语言的锚点文本。
_SECTION_KEY = "summary_title"


def _filtered_results(chapter_results: List[ChapterResult]) -> List[ChapterResult]:
    """过滤掉 Sample Description / 样品描述（不参与汇总统计，见 spec §4.2.5）。"""
    return [r for r in chapter_results if r.test_item != "Sample Description"]


def execute(doc, chapter_results: List[ChapterResult], task_ctx) -> str:
    """生成汇总页并返回最终结论（内部判定值固定为 PASS / FAIL）。

    v5.0：委托 ``template_engine.fill_summary`` 完成 Test Contents 表填充
    （兼容 4/5 列表头）、统计段落更新与结论着色、Notes 填写，并返回最终结论。
    本模块仅负责业务过滤（排除 Sample Description）与语言分派。

    双语：汇总页章节标题、统计段落关键词、汇总表表头、判定结论引导语与判定词
    全部按 ``task_ctx.report_language`` 切换（中文报告判定词为「合格」/「不合格」），
    但**返回给上层的 final_verdict 仍是 ``"PASS"`` / ``"FAIL"``**（供 file_renamer
    按既有规范重命名，不随语言变化）。

    Args:
        doc: 已打开的 Word Document。
        chapter_results: 各章节执行结果列表（含 Sample Description）。
        task_ctx: 任务上下文（取 report_language）。

    Returns:
        final_verdict: ``"PASS"`` 或 ``"FAIL"``。
    """
    results = _filtered_results(chapter_results)
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))

    # 文档填充 + 结论（best-effort：定位失败不影响返回结论；fill_summary 内部已
    # 容错返回 "PASS" 作为极端兜底，正常路径下结论由 test_result 决定）。
    verdict = te.fill_summary(doc, results, lang)
    info(f"[summary_processor] 汇总完成（lang={lang}），最终结论: {verdict}")
    return verdict


__all__ = ["execute"]
