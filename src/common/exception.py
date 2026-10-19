"""项目自定义异常类。

本模块统一异常分类，所有业务模块抛出的异常都应继承自根异常
``ReportPlatformError``，便于 ``main.py`` 统一捕获与分类输出（见 spec
「common/exception.py」职责说明）。

本模块**不包含业务逻辑**，也**不处理异常输出**（输出由 ``common/logger.py``
或 ``main.py`` 负责）。
"""

from typing import Optional


class ReportPlatformError(Exception):
    """项目根异常：所有自定义业务异常的基类。

    所有模块抛出的业务异常都应继承此类（直接或间接），以便 ``main.py``
    用 ``except ReportPlatformError`` 一次性捕获并分类。

    可携带原始异常（``cause``）以保留链路，但不在此处做任何日志或输出。
    """

    def __init__(self, message: str = "", *, cause: Optional[Exception] = None):
        super().__init__(message)
        self.message = message
        self.cause = cause

    def __str__(self) -> str:  # pragma: no cover - 纯展示逻辑
        if self.cause is not None:
            return f"{self.message} (cause: {self.cause})"
        return self.message


class ConfigError(ReportPlatformError):
    """配置相关异常：config.yaml 缺失、字段缺失或解析失败（config_loader）。"""


class MaterialTableError(ReportPlatformError):
    """物料信息表相关异常：文件缺失、列名无法匹配或解析失败（config_loader）。"""


class FolderNotFoundError(ReportPlatformError):
    """目标文件夹扫描相关异常：按编号未找到目标文件夹（folder_scanner）。"""


class ExcelError(ReportPlatformError):
    """电性 Excel 读取 / 转图相关异常（data_extractor）。"""


class WordOperationError(ReportPlatformError):
    """Word 文档操作相关异常：模板打开、保存失败或插入图片/表格时 python-docx 报错
    （docx_engine / report_engine）。v5.0 起替代原 WordComError（已无 COM）。"""


class TemplateError(ReportPlatformError):
    """模板引擎相关异常：书签缺失、占位符无法定位（template_engine）。"""


class ReferenceError(ReportPlatformError):
    """数据查询相关异常：reference_service 匹配 / 查询失败。"""
