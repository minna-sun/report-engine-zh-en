"""文档引擎工厂（v5.0 新增 · spec 模块 12）。

职责：根据配置/环境返回 ``DocumentEngine`` 的具体实现实例。
当前唯一实现为 python-docx 引擎（``DocxEngine``）；为未来可能增加的实现
（如 LibreOffice、纯 OOXML 流式写入）预留工厂入口，使**业务层通过
``get_engine()`` 获取实例，不感知底层技术栈**，满足 v5.0 分层目标。
"""
from services.document_engine import DocumentEngine
from services.docx_engine import DocxEngine


def get_engine() -> DocumentEngine:
    """返回默认文档引擎实例（当前为 python-docx 实现）。

    业务层统一调用此函数获取引擎，避免在多处硬编码 ``DocxEngine``。
    """
    return DocxEngine()
