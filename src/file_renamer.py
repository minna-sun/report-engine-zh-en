"""报告重命名（模块20，步骤6）。

职责：从 MaterialRecord 中提取重命名所需字段，按统一格式拼接目标文件名，
调用 os.rename 在「原目录」下重命名（不移动文件）。最终结论 final_verdict 由
summary_processor 生成、经 core/report_engine -> main.py 传入，本模块不负责生成。

命名格式（spec 第4.2.5节 / 命名规范第20项）：
    {收样年月} {料号} {供应商} {物料子类} {物料型号} {流程编号} {verdict}.docx
各字段取原文（不翻译），以单空格分隔；物料子类或物料型号为空时对应位置留空
（仍保留空格分隔）。收样年月从 record.receive_date 提取为 YYYYMM。

注意（2026-08-20 用户要求）：流程编号统一带「R」前缀，即 0359044 -> R0359044；
若 Excel 读到的 number 已带 R 则不重复添加。

边界（spec 第4.2.5节「边界」）：
    * 不负责生成最终结论（由 summary_processor 生成、作为参数传入）
    * 不负责判断文件是否存在（由调用方 main.py 确保文件已生成）
    * 不负责移动文件（仅在原目录下重命名）
    * 不修改 record 其他内容，不进行数据校验
"""

import os
import re

from common.logger import info, error, warning
from models.material import MaterialRecord
from services import reference_service


def _format_yyyymm(value) -> str:
    """收样年月 -> 'YYYYMM'；无法解析返回空串。

    支持 receive_date 常见格式：``YYYY/MM/DD``、``YYYY-MM-DD``、``YYYY.MM.DD``、
    ``YYYY/MM/DD HH:MM:SS``、``YYYYMMDD``（8位纯数字）、``MM/DD/YYYY``，
    以及已是 ``YYYYMM`` 的 6 位串。
    """
    if not value:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d{6}", s):                       # 已是 YYYYMM
        return s
    if re.fullmatch(r"\d{8}", s):                       # YYYYMMDD（无分隔符）
        return s[:6]
    # MM/DD/YYYY（sample_processor 统一格式化后的收样日期）
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return f"{m.group(3)}{int(m.group(1)):02d}"
    # YYYY[/-.]MM[/-.]DD 或 YYYY[/-.]MM 开头的所有格式
    m = re.search(r"(\d{4})[/\-.](\d{1,2})", s)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}"
    return ""


def _sample_letter(value) -> str:
    """送样次数 -> 版本号字母（1→A, 2→B … 26→Z）。

    输入为空或无法解析为正整数时返回空字符串。
    """
    try:
        i = int(str(value).strip())
    except (ValueError, TypeError, AttributeError):
        return ""
    if 1 <= i <= 26:
        return chr(ord("A") + i - 1)
    return ""


def _with_r(number: str) -> str:
    """流程编号原样返回（不再添加 R 前缀）。

    说明（2026-09-18）：物料表里的流程号现在本身就是完整格式
    （如 R0359044、LAB-CE202609040015），无需脚本补 R。
    保留本函数作为兼容占位，未来若有需要再启用。
    """
    return (number or "").strip()

def _strip_chinese_paren(text: str) -> str:
    """删除中文/半角括号及其内部内容（英文报告用）。"""
    if not text:
        return ""
    return re.sub(r"[（(][^）)]*[）)]", "", text).strip()


def rename_report(file_path: str, record: MaterialRecord,
                  final_verdict: str, lang: str = "zh") -> str:
    """报告重命名（步骤6）。

    Args:
        file_path: 当前文件路径（core/report_engine 返回的临时报告路径）。
        record: 物料记录；本函数自行提取重命名所需字段。
        final_verdict: 最终结论（``"PASS"`` 或 ``"FAIL"``）。
        lang: 报告语言（``"en"`` 时物料子类经翻译映射表转换为英文、供应商去掉
            中文括号、'测试报告' 替换为 'Test Report'；中文报告保持原格式）。

    Returns:
        重命名后的完整文件路径 new_path。

    Raises:
        ValueError: file_path 为空等前置条件不满足时。
        OSError: 重命名失败（如目标文件已存在）时，打印错误后向上抛出，
                 由调用方 main.py 捕获处理。
    """
    if not file_path:
        raise ValueError("[file_renamer] file_path 为空，无法重命名")

    yyyymm = _format_yyyymm(record.receive_date)
    info(f"[file_renamer] receive_date 原始值={record.receive_date!r}  解析结果={yyyymm!r}")

    lang_norm = (lang or "zh").strip().lower()
    is_en = lang_norm.startswith("en") or lang_norm in ("英文", "english")

    if is_en:
        # 物料子类：翻译映射表精确匹配（lang 硬编码 "en"），失败保留中文
        subclass_raw = (record.material_subclass or "").strip()
        if subclass_raw:
            translated = reference_service.get_translation("material", subclass_raw, "en")
            if translated and not translated.startswith("[缺失"):
                subclass = translated
            else:
                subclass = subclass_raw
        else:
            subclass = ""
        supplier = _strip_chinese_paren(record.supplier or "")
        report_word = "Test Report"
    else:
        subclass = (record.material_subclass or "").strip()
        supplier = (record.supplier or "").strip()
        report_word = "测试报告"

    parts = [
        (record.material_code or "").strip(),
        subclass,
        supplier,
        (record.material_model or "").strip(),
        report_word,
        _with_r(record.number) + _sample_letter(record.sample_count),
        final_verdict,
    ]
    new_name = " ".join(parts) + ".docx"
    new_path = os.path.join(os.path.dirname(file_path) or ".", new_name)

    # 覆盖模式：重命名前先删除同名旧文件
    if os.path.exists(new_path) and os.path.abspath(new_path) != os.path.abspath(file_path):
        try:
            os.remove(new_path)
            info(f"[file_renamer] 目标文件已存在，已删除旧版本：{new_path}")
        except Exception as e:
            warning(f"[file_renamer] 删除旧文件失败：{e}")

    try:
        os.rename(file_path, new_path)
    except OSError as exc:
        error(f"[file_renamer] 重命名失败: {file_path} -> {new_path}: {exc}")
        raise

    info(f"[file_renamer] 重命名完成: {os.path.basename(file_path)} -> {new_name}")
    return new_path
