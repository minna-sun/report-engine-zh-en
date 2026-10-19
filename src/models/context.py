"""任务执行上下文数据模型。

本模块仅定义一次报告生成过程中的任务执行上下文数据结构，不包含任何业务逻辑、
COM 对象或文件操作。``TaskContext`` 保持纯数据，Word COM 对象（``Application`` /
``doc``）由 ``core/report_engine.py`` 作为局部变量持有，调用各 ``Processor`` 时显式
传入（见 spec「决策 3：为什么新增 TaskContext」）。

所有模块通过显式导入引用，例如::

    from models.context import TaskContext
"""

from dataclasses import dataclass
from typing import Optional

from models.material import MaterialRecord


@dataclass
class TaskContext:
    """任务执行上下文：贯穿一次报告生成过程（纯数据，不含 COM 对象）。

    Attributes:
        record: 当前物料记录（MaterialRecord）。
        folder_path: 目标文件夹路径。
        config: 全局配置字典（来自 config.yaml）。
        output_path: 报告输出路径；未指定时为 ``None``。
        report_language: 最终确定的报告语言（``"zh"`` / ``"en"``）。由 ``main`` 按
            「物料表指定 > 模板自动识别 > 默认 en」三级优先级决策后传入；
            各 Processor / template_engine 据此切换锚点与翻译行为。
    """

    record: MaterialRecord
    folder_path: str
    config: dict
    output_path: Optional[str] = None
    report_language: str = "en"
    standalone_mode: bool = False

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含关键字段名-值的字典（record 以 to_dict 展开）。
        """
        return {
            "record": self.record.to_dict(),
            "folder_path": self.folder_path,
            "config": self.config,
            "output_path": self.output_path,
            "report_language": self.report_language,
            "standalone_mode": self.standalone_mode,
        }
