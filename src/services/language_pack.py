"""报告语言标签包（双语支持 · 单一数据源）。

职责：集中维护「同一处锚点在中文 / 英文模板下的不同写法」，供各 Processor /
template_engine / docx_engine 按当前报告语言动态取用。

设计约束（与本次双语改造要求一一对应）：
- **禁止在业务代码里散写语言标签**：任何锚点 / 表头 / 判定词都必须从这里取，
  否则新增一种语言时要改 N 处。
- **config.yaml 不保留任何语言配置项**：语言只来自物料表「报告语言」列或
  模板自动识别（见 ``services.language_detector``）。
- 取值统一走 ``labels(lang)``，lang 非法时回退英文。

技术栈：纯 Python（不依赖 python-docx / openpyxl / win32com）。
"""

# 支持的语言
ZH = "zh"
EN = "en"
SUPPORTED = (ZH, EN)
DEFAULT = EN


_EN = {
    # ── 页眉 ──────────────────────────────────────────────────────────
    "header_report_no": "Report No.:",
    # ── 封面 ──────────────────────────────────────────────────────────
    "cover_title": "Test Report",
    "cover_applicant": "Applicant:",
    "cover_department": "Applicant Department:",
    "cover_received_date": "Received Date:",
    "cover_inspection_date": "Inspection Date:",
    "cover_report_date": "Report Date:",
    # ── 样品描述页 ────────────────────────────────────────────────────
    "sample_title": "Sample Description",
    "sample_material_number": "Material Number:",
    "sample_material_category": "Material Category:",
    "sample_material_subcategory": "Material Subcategory:",
    "sample_model_series": "Model/Series:",
    "sample_material_description": "Material Description:",
    "sample_supplier": "Supplier:",
    "sample_quantity": "Sample Quantity:",
    "sample_label": "Sample Label",
    "sample_image": "Sample Image",
    # ── 各测试章节公共字段 ────────────────────────────────────────────
    "equip_title": "Test Equipment",
    "equip_name": "Equipment Name",
    "equip_model": "Equipment Model",
    "equip_no": "Equipment No",
    "equip_cal": "Calibration Expiration Date",
    "chapter_sample_number": "Sample Number",
    "chapter_test_standards": "Test Standards",
    "chapter_test_result": "Test Result",
    "chapter_result_description": "Result description",
    "chapter_test_condition": "Test condition",
    "chapter_test_point": "Test Point",          # ← 新增（图片表格中"取点图"标签）
    # ── 汇总页 ────────────────────────────────────────────────────────
    "summary_title": "Summary of Results",
    "summary_total_of": "total of",
    "summary_passing": "items passing",
    "summary_failing": "items failing",
    "summary_reference": "items for reference",
    "summary_judged_as": "judged as:",
    "summary_notes": "Notes on FAIL/Reference",
    "col_no": "No.",
    "col_test_item": "Test Item",
    "col_test_basis": "Test Basis",
    "col_sample_number": "Sample Number",
    "col_test_result": "Test Result",
    "verdict_pass": "Pass",
    "verdict_fail": "Fail",
    # ── 卤素表翻译（源报告英文 → 目标英文不变）──
    "halogen_col_material_no": "Material No.",
    "halogen_col_test_item": "Test item",
    "halogen_col_unit": "Unit",
    "halogen_col_rl": "RL",
    "halogen_col_customer_req": "Customer's Requirement",
    "halogen_col_test_result": "Test result",
    "halogen_val_chlorine": "Chlorine（Cl）Content",
    "halogen_val_bromine": "Bromine（Br）Content",
    "halogen_val_cl_br": "Chlorine（Cl）+ Bromine（Br）content",
}

_ZH = {
    # ── 页眉 ──────────────────────────────────────────────────────────
    "header_report_no": "报告编号：",
    # ── 封面 ──────────────────────────────────────────────────────────
    "cover_title": "检测报告",
    "cover_applicant": "委托人：",
    "cover_department": "委托单位：",
    "cover_received_date": "收样日期：",
    "cover_inspection_date": "检测日期：",
    "cover_report_date": "报告日期：",
    # ── 样品描述页 ────────────────────────────────────────────────────
    "sample_title": "样品描述",
    "sample_material_number": "物料编号：",
    "sample_material_category": "物料类别：",
    "sample_material_subcategory": "物料子类：",
    "sample_model_series": "型号/系列：",
    "sample_material_description": "物料描述：",
    "sample_supplier": "供 应 商：",
    "sample_quantity": "样品数量：",
    "sample_label": "样品标签",
    "sample_image": "样品图片",
    # ── 各测试章节公共字段 ────────────────────────────────────────────
    "equip_title": "检测设备",
    "equip_name": "设备名称",
    "equip_model": "型号",
    "equip_no": "设备编号",
    "equip_cal": "校准日期",
    "chapter_sample_number": "样品编号",
    "chapter_test_standards": "检测依据",
    "chapter_test_result": "检测结果",
    "chapter_result_description": "结果说明",
    "chapter_test_condition": "测试条件",
    "chapter_test_point": "取点图",               # ← 新增
    # ── 汇总页 ────────────────────────────────────────────────────────
    "summary_title": "结果汇总",
    "summary_total_of": "本次共测试",
    "summary_passing": "合格",
    "summary_failing": "不合格",
    "summary_reference": "供参考",
    "summary_judged_as": "根据以下的检验依据，判定本次试验样品：",
    "summary_notes": "不合格/供参考说明",
    "col_no": "序号",
    "col_test_item": "测试项目",
    "col_test_basis": "检测依据",
    "col_sample_number": "样品编号",
    "col_test_result": "检测结果",
    "verdict_pass": "合格",
    "verdict_fail": "不合格",
    # ── 卤素表翻译（中文报告：源报告英文 → 中文）──
    "halogen_col_material_no": "物料编号",
    "halogen_col_test_item": "测试项目",
    "halogen_col_unit": "单位",
    "halogen_col_rl": "RL",
    "halogen_col_customer_req": "客户要求",
    "halogen_col_test_result": "检测结果",
    "halogen_val_chlorine": "氯（Cl）含量",
    "halogen_val_bromine": "溴（Br）含量",
    "halogen_val_cl_br": "氯（Cl）+ Bromine（Br）含量",
}

_PACKS = {EN: _EN, ZH: _ZH}

# 汇总表列定位顺序（与 labels() 的 key 一一对应，顺序即优先级，勿随意调整）
_SUMMARY_COL_KEYS = (
    "col_no",
    "col_test_item",
    "col_test_basis",
    "col_sample_number",
    "col_test_result",
)


def normalize(lang) -> str:
    """把任意来源的语言值归一为 ``"zh"`` / ``"en"``；无法识别时回退英文。

    Accepts: ``"zh"`` / ``"ZH"`` / ``"中文"`` / ``"zh-CN"`` / ``"en"`` / ``None``。
    """
    s = (str(lang) or "").strip().lower()
    if not s:
        return DEFAULT
    if s.startswith("zh") or s in ("中文", "china", "chinese"):
        return ZH
    if s.startswith("en") or s in ("英文", "english"):
        return EN
    return DEFAULT


def labels(lang) -> dict:
    """返回该语言的标签字典（非法语言回退英文）。

    Returns:
        dict: 键为标签 key（见 ``_EN`` 注释分组），值为该语言的实际锚点文本。
    """
    return _PACKS.get(normalize(lang), _EN)


def pick(lang, key: str) -> str:
    """取单个标签；key 不存在时返回空串（避免把 None 写进文档）。"""
    return labels(lang).get(key, "")


def summary_col_keys() -> tuple:
    """汇总表列定位顺序（按此顺序在表头里 setdefault 匹配，勿打乱）。"""
    return _SUMMARY_COL_KEYS


__all__ = ["ZH", "EN", "SUPPORTED", "DEFAULT", "normalize", "labels",
           "pick", "summary_col_keys"]
