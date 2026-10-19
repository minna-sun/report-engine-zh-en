"""执行结果数据模型。

本模块仅定义执行结果相关的数据结构，不包含任何业务逻辑、COM 对象或文件操作。
所有模块通过显式导入引用，例如::

    from models.result import ScanResult, ChapterResult, ReportResult, GenerationSummary

各模块必须使用本包定义的数据结构作为标准输入输出格式，禁止自行定义同名但字段
不一致的结构（见 spec「特殊规则」）。
"""

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ScanResult:
    """``folder_scanner.py`` 的返回结果。

    Attributes:
        number: 物料编号。
        folder_paths: 匹配到的目标文件夹路径列表（可能为空列表）。
        status: 扫描状态，``"ok"`` 或 ``"failed"``。
        error: 失败时的错误原因；成功时为 ``None``。
    """

    number: str
    folder_paths: List[str]
    status: str
    error: Optional[str] = None

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "number": self.number,
            "folder_paths": self.folder_paths,
            "status": self.status,
            "error": self.error,
        }


@dataclass
class ChapterResult:
    """单个章节的执行结果。

    Attributes:
        test_item: 测试项目名称。
        test_basis: 测试标准（Test Standards）。
        sample_number: 样品编号。
        test_result: 章节判定，``"PASS"`` / ``"FAIL"`` / ``"Reference"``。
        result_description: 备注信息（来自目标文件夹配置），供 summary_processor
            填写汇总页 Notes on FAIL/Reference 区域；无备注时为空字符串。
        status: 执行状态，``"success"`` 或 ``"failed"``。
        error: 失败时的错误原因；成功时为 ``None``。
    """

    test_item: str
    test_basis: str
    sample_number: str
    test_result: str
    result_description: str = ""
    status: str = "success"
    error: Optional[str] = None

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "test_item": self.test_item,
            "test_basis": self.test_basis,
            "sample_number": self.sample_number,
            "test_result": self.test_result,
            "result_description": self.result_description,
            "status": self.status,
            "error": self.error,
        }


@dataclass
class ReportResult:
    """``core/report_engine.py`` 的返回结果。

    Attributes:
        status: 生成状态，``"success"`` 或 ``"failed"``。
        report_path: Word 报告输出路径；失败时为 ``None``。
        pdf_path: PDF 报告输出路径；未导出时为 ``None``。
        final_verdict: 最终结论，``"PASS"`` 或 ``"FAIL"``；成功时有值。
        error: 失败时的错误原因；成功时为 ``None``。
    """

    status: str
    report_path: Optional[str] = None
    pdf_path: Optional[str] = None
    final_verdict: Optional[str] = None
    error: Optional[str] = None

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "status": self.status,
            "report_path": self.report_path,
            "pdf_path": self.pdf_path,
            "final_verdict": self.final_verdict,
            "error": self.error,
        }


@dataclass
class GenerationSummary:
    """``main.py`` 的汇总结果。

    Attributes:
        status: 汇总状态，``"success"`` / ``"partial"`` / ``"failed"``。
        total: 总编号数。
        succeeded: 成功生成数。
        failed: 失败数。
        details: 每个编号的执行明细（list of dict）。
    """

    status: str
    total: int
    succeeded: int
    failed: int
    details: List[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "status": self.status,
            "total": self.total,
            "succeeded": self.succeeded,
            "failed": self.failed,
            "details": self.details,
        }
