"""处理器注册表（模块 17）。
职责：按 ``ChapterConfig.category`` 字段查找对应的 ``GenericProcessor`` 实例。
"""
from typing import Optional

from common.logger import warning
from processors.generic_processor import GenericProcessor

_PROCESSOR_REGISTRY: dict[str, type] = {
    "ImageInspection": GenericProcessor,
    "ImageAuto": GenericProcessor,
    "ElectricalImageAuto": GenericProcessor,
    "ReliabilityCycle": GenericProcessor,
    "Electrical": GenericProcessor,
    "Compliance": GenericProcessor,
    "ImageElectrical": GenericProcessor,
    "IonContamination": GenericProcessor,
    "EDS": GenericProcessor,
    "Halogen": GenericProcessor,
}


def get_processor(category: str) -> Optional[GenericProcessor]:
    """按测试类型查找并返回 ``GenericProcessor`` 实例。"""
    processor_cls = _PROCESSOR_REGISTRY.get(category)
    if processor_cls is None:
        warning(f"[processor_registry] 未找到处理器类型 '{category}'，跳过该章节")
        return None
    return processor_cls()


__all__ = ["get_processor", "_PROCESSOR_REGISTRY"]
