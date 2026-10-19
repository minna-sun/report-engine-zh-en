"""策略注册表（模块 15）。
职责：按 ``ChapterConfig.strategy`` 字段查找对应的具体 Strategy 类。

注意（开源发布）：
- 本仓库仅发布**通用策略基类**与**一个示例策略** ``ImageBasedStrategy``，
  用于演示「策略复用 / 0 代码扩展」的写法与契约。
- 其余生产策略（电性表、合规、卤素、EDS、离子污染、可靠性、全自动图片等）
  属于内部私有实现，保留在本地 ``src/strategies/``（见 .gitignore），
  不随本仓库公开。新增物料类别时，参照 ``image_based.py`` 在本地实现
  对应策略并注册即可，核心引擎无需改动。
"""
from typing import Optional, Type

from common.logger import warning
from strategies.base import BaseStrategy
from strategies.image_based import ImageBasedStrategy

_STRATEGY_REGISTRY: dict[str, Type[BaseStrategy]] = {
    "ImageBasedStrategy": ImageBasedStrategy,
}


def get_strategy(strategy_name: str) -> Optional[Type[BaseStrategy]]:
    """按策略名称查找具体 Strategy 类。"""
    strategy_cls = _STRATEGY_REGISTRY.get(strategy_name)
    if strategy_cls is None:
        warning(f"[strategy_registry] 未找到策略 '{strategy_name}'，跳过该章节")
    return strategy_cls


__all__ = ["get_strategy", "_STRATEGY_REGISTRY"]
