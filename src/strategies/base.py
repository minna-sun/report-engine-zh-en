"""策略抽象基类（模块 10 · v5.0）。

职责：定义统一 ``execute`` 接口契约，所有具体策略必须继承本类并实现该方法。

边界（见 重构.md §6.4）：
- 不导入 python-docx / lxml / DocumentEngine。
- 不直接操作 Word 文档：通过调用方传入的 ``template_engine`` 实例完成所有文档操作
  （``template_engine`` 提供的工具函数，如 ``insert_image_at_placeholder`` 等）。
- 不持有 COM 对象（v5.0 已无 COM）。
"""

from abc import ABC, abstractmethod
import re
from typing import Any

from models.chapter import ChapterContext
from models.context import TaskContext
from models.result import ChapterResult


def _extract_width_from_name(name: str, default_cm: float) -> float:
    """从文件名提取宽度：匹配最后一个括号内的数字/小数（半角/全角）。

    例：'1-1(10)'    → 10.0
        '1-1（10.5）' → 10.5
        '1-1(1)(10)' → 10.0   (取最后一个)
        '1-1'        → default_cm
        '1-1(abc)'   → default_cm
    """
    matches = re.findall(r"[（(](\d+(?:\.\d+)?)[）)]", name)
    if matches:
        try:
            return float(matches[-1])
        except (ValueError, TypeError):
            pass
    return default_cm


class BaseStrategy(ABC):
    """策略抽象基类。

    所有具体策略（图片型、可靠性、电性、合规）必须继承本类，
    并实现 ``execute`` 方法，将类型-specific 的章节内容填充委托给
    ``template_engine`` 执行。
    """

    @abstractmethod
    def execute(self, template_engine: Any, chapter_ctx: ChapterContext,
                task_ctx: TaskContext) -> ChapterResult:
        """执行策略，填充当前章节的类型-specific 内容。

        Args:
            template_engine: ``template_engine`` 模块/实例（所有文档操作经它完成）。
            chapter_ctx: 章节执行上下文（含静态配置 + 运行时数据）。
            task_ctx: 任务执行上下文（含物料记录、文件夹路径等）。

        Returns:
            ChapterResult: 章节执行结果，含 test_item、test_basis、
                sample_number、test_result、status、error 等字段。

        实现要点：
        1. 通过 ``chapter_ctx.config`` 读取策略所需的静态配置
           （如 placeholders、image_mapping、folder_name、standard 等）。
        2. 通过 ``chapter_ctx.runtime`` 读取运行时数据
           （如 sample_number、device、condition、result_description 等）。
        3. 通过 ``task_ctx.folder_path`` 定位原始记录文件夹。
        4. 所有 Word 文档操作委托给 ``template_engine.process_*_chapter``，
           策略本身不直接接触文档对象。
        5. 捕获异常，失败时返回 ``status="failed"`` 并记录错误原因，
           不向上抛出异常（避免中断其他章节）。
        """
        ...

    # ───────────────────────── 结果填充辅助 ─────────────────────────

    def _finalize(self, result: ChapterResult, chapter_ctx: ChapterContext,
                  task_ctx: TaskContext) -> ChapterResult:
        """补全 ChapterResult 的公共字段（test_basis / sample_number）。

        - ``test_basis`` 取自 ``config.standard``（测试标准）。
        - ``sample_number`` 优先取 ``runtime.sample_number``，回退到 ``record.number``。
        - ``test_result`` 由 GenericProcessor 经 reference_service 判定后回填，
          此处保持 process_*_chapter 返回的原值（一般为空）。
        """
        result.test_basis = chapter_ctx.config.standard or ""
        sn = chapter_ctx.runtime.sample_number
        if (not sn) and task_ctx.record is not None:
            sn = getattr(task_ctx.record, "number", None)
        result.sample_number = sn or ""
        return result

