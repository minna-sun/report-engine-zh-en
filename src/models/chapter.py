"""章节数据模型。

本模块仅定义章节相关的数据结构，不包含任何业务逻辑、COM 对象或文件操作。
所有模块通过显式导入引用，例如::

    from models.chapter import ChapterConfig, ChapterRuntime, ChapterContext

设计要点：将章节的「静态策略配置」(ChapterConfig) 与「运行时数据」(ChapterRuntime)
分离，再由 ChapterContext 组合二者，并补充物料类别与原始文件夹名。
"""

from dataclasses import dataclass
from typing import List, Optional


@dataclass
class ChapterConfig:
    """单个测试章节的静态策略配置（来自物料类别配置）。

    Attributes:
        test_item: 测试项目名称（如 ``"Appearance"``）。
        category: 测试类型（如 ``"ImageInspection"``）。
        strategy: 策略名称（如 ``"ImageBasedStrategy"``）。
        folder_name: 原始数据文件夹名（如 ``"外观"``）。
        placeholders: 图片型章节在 Word 模板中预置的占位符文本列表
            （如 ``["front", "back"]``，模板中以 ``{front}`` / ``[back]`` 等形式书写，
            脚本匹配时自动去除括号）；非图片章节为 ``None``。
        image_mapping: 图片匹配说明（图片型章节使用）；非图片章节为 ``None``。
        standard: 测试标准（可选）。``GenericProcessor`` 直接将其作为
            ``ChapterResult.test_basis`` 返回，并用于填充文档 Test Standards 字段
            （Compliance 类型跳过文档填充，但 test_basis 仍取此值）。
        image_width: 图片插入宽度（厘米），来自物料类别配置的 ``图片宽度`` 列
            （如 ``7.0``）；非图片章节或留空时为 ``None``。
    """

    test_item: str
    category: str
    strategy: str
    folder_name: str
    placeholders: Optional[List[str]] = None
    image_mapping: Optional[str] = None
    standard: Optional[str] = None
    image_width: Optional[float] = None

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "test_item": self.test_item,
            "category": self.category,
            "strategy": self.strategy,
            "folder_name": self.folder_name,
            "placeholders": self.placeholders,
            "image_mapping": self.image_mapping,
            "standard": self.standard,
            "image_width": self.image_width,
        }


@dataclass
class ChapterRuntime:
    """单个测试章节的运行时数据（来自目标文件夹配置）。

    Attributes:
        device: 测试设备编号（可能为空）。
        condition: 测试条件（可靠性章节使用，可能为空）。
        result_description: Result description 内容（可能为空）。
        sample_number: 样品编号（可能为空）。
    """

    device: Optional[str] = None
    condition: Optional[str] = None
    result_description: Optional[str] = None
    sample_number: Optional[str] = None

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有字段名-值的字典。
        """
        return {
            "device": self.device,
            "condition": self.condition,
            "result_description": self.result_description,
            "sample_number": self.sample_number,
        }


@dataclass
class ChapterContext:
    """章节执行上下文：静态配置 + 运行时数据 + 物料类别 + 原始文件夹名。

    Attributes:
        config: 静态策略配置（ChapterConfig）。
        runtime: 运行时数据（ChapterRuntime）。
        material_category: 物料类别（如 ``"1014"``），用作路由键。
        source_folder_name: 原始记录子文件夹真实名称（用于 Test Result 判定）。
    """

    config: ChapterConfig
    runtime: ChapterRuntime
    material_category: str
    source_folder_name: str

    @property
    def test_item(self) -> str:
        """便捷访问：当前章节的测试项目名称（来自 config）。"""
        return self.config.test_item

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含 config/runtime 子字典及两个路由字段的字典。
        """
        return {
            "config": self.config.to_dict(),
            "runtime": self.runtime.to_dict(),
            "material_category": self.material_category,
            "source_folder_name": self.source_folder_name,
        }
