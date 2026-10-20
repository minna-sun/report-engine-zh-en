"""通用处理器（模块 16 · v5.0）。

职责（编排器，非业务黑洞）：处理所有类型化测试章节（ImageInspection / ReliabilityCycle /
Electrical / Compliance）。负责 **统一的章节分页** 与 **跨类型公共字段填充**，再根据
``chapter_ctx.config.strategy`` 调度具体 Strategy 执行类型-specific 操作。

核心边界（spec §6.5）：
- 只做跨类型公共逻辑：分页、Test Standards、设备表、Test Result、Result description、
  Test Condition、Sample Number；所有类型-specific 逻辑（图片匹配 / 电性 Excel / 跨文档
  复制表格）全部下沉到 Strategy（经 template_engine 执行）。
- 禁止在本模块硬编码任何测试项目名、设备编号、文件夹名。
- 章节分页统一由本模块处理，策略层不再各自实现分页。
- Test Result 由 ``reference_service.determine_test_result`` 判定，与执行成败无关。
- 单个章节失败不向上抛异常（返回 status="failed" 并记录 error），由 report_engine 决定
  是否继续其它章节；test_result 不受影响。

v5.0 变更：所有 Word 操作经 ``template_engine``（python-docx + DocumentEngine）完成，
不再使用任何 win32com / COM 调用；Strategy 调度时把 ``template_engine`` 模块作为首参传入，
策略不再直接持有 doc。
"""

import re

from common.logger import info, warning
from models.result import ChapterResult
from services import reference_service
from services import language_pack
from services.engine_factory import get_engine
from template_engine import (
    fill_equipment_table,
    insert_text_after,
    add_page_break,
)
import template_engine as te
from strategy_registry import get_strategy

# 公共字段标签 KEY（模板中多为中文全角冒号；insert_text_after 内部按实际冒号自适应，
# 故此处只传基础标签串，由 template_engine 负责冒号后替换）。
# 严禁在此硬编码任何语言的锚点文本 —— 一律经 language_pack 按 report_language 取用。
_KEY_TEST_STANDARDS = "chapter_test_standards"
_KEY_TEST_CONDITION = "chapter_test_condition"
_KEY_SAMPLE_NUMBER = "chapter_sample_number"
_KEY_RESULT_DESCRIPTION = "chapter_result_description"
_KEY_TEST_RESULT = "chapter_test_result"

# 过期 / FAIL 标红用 RGB 整数（python-docx 为 RGB 顺序：0xFF0000 = 红）
_COLOR_RED = 0xFF0000


def _read_field_value(doc, test_item: str, base: str) -> str:
    """从模板章节中读取某字段标签所在段落、冒号之后的文本内容（样品编号回退规则用）。

    v5.0：经 DocumentEngine 定位段落并读取文本，不再使用 win32com 的 Range.Text。
    """
    try:
        eng = get_engine()
        section = eng.find_section(doc, test_item)
        if section is None:
            return ""
        para = eng.find_paragraph(doc, section, base)
        if para is None:
            return ""
        text = para.text or ""
        for sep in ("：", ":"):
            prefix = base + sep
            if text.startswith(prefix):
                return text[len(prefix):].strip()
        return text.strip()
    except Exception:
        return ""


def _fill_equipment(doc, test_item: str, device, lang: str = "en") -> None:
    """填充设备表：支持单个或逗号分隔的多个设备编号。

    设备编号为空时，跳过填充，模板保留原样。
    device 可以是单个编号字符串（如 "yy-ll-gh1"）或逗号分隔的多个编号
    （如 "yy-ll-tf1, yy-ll-jh42"），每个编号独立查询并填入单独一行。

    Args:
        lang: 报告语言；设备表锚点与表头均按语言切换（中文：检测设备 / 设备名称 / …）。
    """
    if not device:
        warning(f'[GenericProcessor] 章节 "{test_item}" 的设备编号为空，跳过设备表填充')
        return

    # 拆分多个设备编号：支持逗号、中文逗号、斜杠、加号、空格等分隔符
    raw = str(device)
    ids = [d.strip() for d in re.split(r"[,，/+\s]+", raw) if d.strip()]
    if not ids:
        warning(f'[GenericProcessor] 章节 "{test_item}" 的设备编号解析为空，跳过设备表填充')
        return

    details = []
    for did in ids:
        try:
            detail = reference_service.get_equipment_detail(did, lang=lang)
            details.append(detail)
        except Exception as e:
            warning(f'[GenericProcessor] 查询设备详情失败({did}): {e}')
            details.append({"name": "/", "model": "/", "no": "/", "expiry": "/"})

    fill_equipment_table(doc, test_item, details, lang)


def _dispatch_strategy(chapter_ctx, task_ctx, status: str, error):
    """按 strategy 名称查找并实例化 Strategy，执行类型-specific 操作。

    首参传入 ``template_engine`` 模块，策略经它完成所有文档操作（不直接持有 doc）。

    Returns:
        (status, error) 元组；Strategy 失败或抛异常时返回 ("failed", 原因)，
        否则原样返回传入的 status/error（不影响 test_result）。
    """
    strategy = chapter_ctx.config.strategy
    test_item = chapter_ctx.config.test_item
    strategy_cls = get_strategy(strategy)
    if strategy_cls is None:
        warning(f'[GenericProcessor] 章节 "{test_item}" 的 strategy "{strategy}" 未注册，'
                f'跳过类型-specific 操作')
        return status, error
    try:
        instance = strategy_cls()
        # 关键：首参为 template_engine 模块（v5.0 不再传 doc）
        sres = instance.execute(te, chapter_ctx, task_ctx)
        if sres is not None and getattr(sres, "status", "success") == "failed":
            reason = getattr(sres, "error", "Strategy 执行失败")
            warning(f'[GenericProcessor] 章节 "{test_item}" 的 Strategy 执行失败: {reason}')
            return "failed", reason
    except Exception as e:
        warning(f'[GenericProcessor] 章节 "{test_item}" 的 Strategy 抛异常: {e}')
        return "failed", f"Strategy 抛异常: {e}"
    return status, error


class GenericProcessor:
    """通用处理器：公共字段填充 + Strategy 调度。

    由 ``processor_registry.get_processor(category)`` 实例化后调用 ``execute``。
    """

    def execute(self, doc, chapter_ctx, task_ctx) -> ChapterResult:
        """填充当前章节的公共字段，并调度对应 Strategy。

        Args:
            doc: 已打开的 Word Document（由 report_engine 持有，python-docx 对象）。
            chapter_ctx: 章节执行上下文（静态配置 + 运行时数据 + 物料类别 + 原始文件夹名）。
            task_ctx: 任务执行上下文（含 record / folder_path / config /
                report_language 等）。

        Returns:
            ChapterResult: 含 test_item / test_basis / sample_number / test_result /
            status / error。test_result 由文件夹后缀判定，不受执行成败影响；
            test_basis 取自 chapter_ctx.config.standard（Compliance 亦如此，仅文档
            Test Standards 字段填充被跳过）。
        """
        # v5.0：设置模块级活跃文档，供 Strategy 经 template_engine.process_*_chapter
        # （使用 _doc(None) 回退）操作同一文档。
        te.set_active_doc(doc)

        cfg = chapter_ctx.config
        rt = chapter_ctx.runtime
        test_item = cfg.test_item
        is_compliance = cfg.category.strip().lower() == "compliance"

        # 报告语言：中文模式下所有公共字段锚点 / 设备表表头切中文
        lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
        L = language_pack.labels(lang)

        # 提前判定 test_result（与执行成败完全独立）
        try:
            test_result = reference_service.determine_test_result(
                chapter_ctx.source_folder_name)
        except Exception:
            test_result = "PASS"

        status = "success"
        error = None
        # test_basis 直接取自章节静态配置（chapter_ctx.config.standard），
        # 不再经 reference_service.get_standard 重新查询
        test_basis = cfg.standard or ""

        try:
            # 1. 统一的章节分页（v5.0：经 template_engine.add_page_break，无 COM）
            add_page_break(doc, test_item)

            # 4. Test Standards / 检测依据 文档填充（Compliance 类型跳过此步骤；
            #    test_basis 已在上方取 config.standard，与是否跳过无关）
            if not is_compliance:
                insert_text_after(doc, test_item, L[_KEY_TEST_STANDARDS], test_basis)

            # 5. 设备表填充（Compliance 跳过）
            if not is_compliance:
                _fill_equipment(doc, test_item, rt.device, lang)

            # 6. Test Condition / 测试条件 填充（Compliance 跳过；空则保留标签）
            if not is_compliance and rt.condition:
                insert_text_after(doc, test_item, L[_KEY_TEST_CONDITION], rt.condition)

            # 7. Sample Number / 样品编号 填充（空则保留标签）
            if rt.sample_number:
                insert_text_after(doc, test_item, L[_KEY_SAMPLE_NUMBER], rt.sample_number)

            # 8. Result description / 结果说明 填充（Compliance 跳过；空则保留标签）
            if not is_compliance and rt.result_description:
                insert_text_after(doc, test_item, L[_KEY_RESULT_DESCRIPTION], rt.result_description)

            # 9. Test Result / 检测结果 填充（Compliance 跳过；FAIL 时值标红）
            # 显示文本按报告语言映射；test_result 原值保留用于标红判断和 ChapterResult。
            if not is_compliance:
                _RESULT_DISPLAY = {
                    "en": {"PASS": "PASS", "FAIL": "FAIL", "Reference": "Reference"},
                    "zh": {"PASS": "合格", "FAIL": "不合格", "Reference": "供参考"},
                }
                result_display = _RESULT_DISPLAY.get(lang, _RESULT_DISPLAY["en"]).get(
                    test_result, test_result)
                insert_text_after(
                    doc, test_item, L[_KEY_TEST_RESULT], result_display,
                    font_color=_COLOR_RED if test_result == "FAIL" else None)

            # 10. Strategy 调度（类型-specific 操作）
            status, error = _dispatch_strategy(chapter_ctx, task_ctx, status, error)
        except Exception as e:
            status = "failed"
            error = f"公共字段填充异常: {e}"
            warning(f'[GenericProcessor] 章节 "{test_item}" 公共字段填充异常: {e}')

        # 样品编号规则：优先 runtime，空则从模板回退读取，再空则留空字符串
        sample_number = rt.sample_number or _read_field_value(
            doc, test_item, L[_KEY_SAMPLE_NUMBER])

        return ChapterResult(
            test_item=test_item,
            test_basis=test_basis,
            sample_number=sample_number or "",
            test_result=test_result,
            result_description=rt.result_description or "",
            status=status,
            error=error,
        )


__all__ = ["GenericProcessor"]
