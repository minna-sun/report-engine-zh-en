"""物料数据模型。

本模块仅定义数据结构，不包含任何业务逻辑、COM 对象或文件操作。
所有模块通过显式导入引用本结构，例如::

    from models.material import MaterialRecord

字段变更时只需修改本文件，所有引用自动生效。
"""

from dataclasses import dataclass


@dataclass
class MaterialRecord:
    """物料信息表中的一条记录。

    字段含义以 ``test-report-system-spec.md`` 系统设计 v4.10 模块 0-1 为准。
    所有字段均为字符串类型：Excel 单元格缺失时以空字符串 ``""`` 表示，
    不在此处做默认值推断（物料类别为空由 ``config_loader`` 负责跳过）。

    Attributes:
        number: 编号（直接读取）。
        material_category: 物料类别（直接读取，如 ``"DEMO-001"``）；
            作为 ``section_manager`` 路由键，决定加载哪张物料类别配置。
        material_code: 料号（从“变更描述”提取，10 位以上连续数字）。
        material_subclass: 物料子类（从“变更描述”提取中文关键词，如“二极管”）。
        material_model: 物料型号（从“变更描述”提取，如 ``"JST131V-800DS"``）。
        creator: 创建者（直接读取）。
        department: 申请部门（直接读取）。
        description: 变更描述（原始文本）。
        quantity: 样品数量（直接读取）。
        receive_date: 样品接收日期（直接读取）。
        sample_count: 送样次数（直接读取）。
        supplier: 供应商（直接读取）。
        approval_status: 审批状态（直接读取；``"1"`` 表示已审批）。
        report_language: 报告语言（直接读取；``"zh"`` / ``"en"``；空串表示未指定，
            交由模板自动识别，识别失败再回退 ``"en"``）。
    """

    number: str                      # 编号（直接读取）
    material_category: str           # 物料类别（直接读取，如 DEMO-001）
    material_code: str               # 料号（从"变更描述"中提取，10 位以上连续数字）
    material_subclass: str           # 物料子类（从"变更描述"中提取中文关键词，如"二极管"）
    material_model: str              # 物料型号（从"变更描述"中提取，如"JST131V-800DS"）
    creator: str                     # 创建者（直接读取）
    department: str                  # 申请部门（直接读取）
    description: str                 # 变更描述（原始文本）
    quantity: str                    # 样品数量（直接读取）
    receive_date: str                # 样品接收日期（直接读取）
    sample_count: str                # 送样次数（直接读取）
    supplier: str                    # 供应商（直接读取）
    approval_status: str             # 审批状态（直接读取）
    report_language: str = ""        # 报告语言（直接读取；"zh"/"en"；空=未指定，走模板自动识别）

    @property
    def is_approved(self) -> bool:
        """是否已审批（仅数据判断，不含 I/O 或外部调用）。

        Returns:
            审批状态字段为 ``"1"`` 时返回 ``True``，否则 ``False``。
        """
        return self.approval_status == "1"

    @property
    def has_material_category(self) -> bool:
        """物料类别是否非空（``config_loader`` 据此跳过空值记录）。

        Returns:
            物料类别去除首尾空白后非空时返回 ``True``，否则 ``False``。
        """
        return bool(self.material_category and self.material_category.strip())

    def to_dict(self) -> dict:
        """转换为普通字典，便于日志输出与下游序列化。

        Returns:
            包含所有 14 个字段名-值的字典。
        """
        return {
            "number": self.number,
            "material_category": self.material_category,
            "material_code": self.material_code,
            "material_subclass": self.material_subclass,
            "material_model": self.material_model,
            "creator": self.creator,
            "department": self.department,
            "description": self.description,
            "quantity": self.quantity,
            "receive_date": self.receive_date,
            "sample_count": self.sample_count,
            "supplier": self.supplier,
            "approval_status": self.approval_status,
            "report_language": self.report_language,
        }
