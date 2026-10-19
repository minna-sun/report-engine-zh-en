"""策略注册表（模块 15）。
职责：按 ``ChapterConfig.strategy`` 字段查找对应的具体 Strategy 类。
"""
from typing import Optional, Type

from common.logger import warning
from strategies.base import BaseStrategy
from strategies.compliance import ComplianceStrategy
from strategies.eds import EDSStrategy
from strategies.electrical_image_auto import ElectricalImageAutoStrategy
from strategies.electrical_table import ElectricalTableStrategy
from strategies.image_and_electrical import ImageAndElectricalStrategy
from strategies.image_auto import ImageAutoStrategy
from strategies.image_based import ImageBasedStrategy
from strategies.ion_contamination import IonContaminationStrategy
from strategies.reliability_cycle import ReliabilityCycleStrategy
from strategies.halogen import HalogenStrategy

_STRATEGY_REGISTRY: dict[str, Type[BaseStrategy]] = {
    "ImageBasedStrategy": ImageBasedStrategy,
    "ImageAutoStrategy": ImageAutoStrategy,
    "ElectricalImageAutoStrategy": ElectricalImageAutoStrategy,
    "ReliabilityCycleStrategy": ReliabilityCycleStrategy,
    "ElectricalTableStrategy": ElectricalTableStrategy,
    "ImageAndElectricalStrategy": ImageAndElectricalStrategy,
    "ComplianceStrategy": ComplianceStrategy,
    "EDSStrategy": EDSStrategy,
    "IonContaminationStrategy": IonContaminationStrategy,
    "HalogenStrategy": HalogenStrategy,
}


def get_strategy(strategy_name: str) -> Optional[Type[BaseStrategy]]:
    """按策略名称查找具体 Strategy 类。"""
    strategy_cls = _STRATEGY_REGISTRY.get(strategy_name)
    if strategy_cls is None:
        warning(f"[strategy_registry] 未找到策略 '{strategy_name}'，跳过该章节")
    return strategy_cls


__all__ = ["get_strategy", "_STRATEGY_REGISTRY"]
