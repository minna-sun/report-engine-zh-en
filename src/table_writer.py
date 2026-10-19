"""物料信息表回写模块（v5.0 重构）。

职责：在报告成功生成后，将输出时间写回物料信息表的「报告输出时间」列。

v5.0 迁移：由 win32com Excel COM 改为 **openpyxl**（纯 Python、跨平台、无 Office /
COM 依赖），与 config_loader 的 Excel 读取技术栈一致。

边界：
- 只做单列回写，不修改其他列；
- 目标列不存在时自动在表头末尾追加「报告输出时间」列；
- 写回失败仅记录警告，不影响主流程；
- 仅支持 .xlsx / .xlsm（openpyxl 限制）；.xls 直接跳过并警告。
"""

import os
from datetime import datetime
from typing import Optional

from common.logger import info, warning

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover - 依赖缺失时给出明确提示而非崩溃
    load_workbook = None
    warning("[table_writer] openpyxl 未安装，回写功能不可用")

_OUTPUT_TIME_COL = "报告输出时间"
_NUMBER_COL_ALIASES = ("编号", "申请编号")


def write_output_time(table_path: str, number: str,
                      output_time: Optional[datetime] = None) -> None:
    """将报告输出时间写回物料信息表对应行。

    Args:
        table_path: 物料信息表绝对路径（.xlsx/.xlsm）。
        number:     编号（用于定位行，如 "R0359044"）。
        output_time: 输出时间，默认取当前时间。
    """
    if load_workbook is None:
        warning("[table_writer] openpyxl 不可用，跳过回写")
        return
    if not os.path.isfile(table_path):
        warning(f"[table_writer] 物料信息表不存在，跳过回写: {table_path}")
        return
    if not table_path.lower().endswith((".xlsx", ".xlsm")):
        warning(f"[table_writer] 仅支持 .xlsx/.xlsm，当前为 {table_path}，跳过回写")
        return

    output_time = output_time or datetime.now()
    time_str = output_time.strftime("%Y/%m/%d %H:%M:%S")

    try:
        wb = load_workbook(table_path)
        ws = wb.active  # 第一个工作表
        if ws is None:
            warning("[table_writer] 物料信息表无工作表，跳过回写")
            return

        total_rows = ws.max_row
        total_cols = ws.max_column

        # ── 定位表头行（第 1 行）──
        hdr_row = 1
        num_col = None
        time_col = None
        for c in range(1, total_cols + 1):
            val = ws.cell(row=hdr_row, column=c).value
            if val is None:
                continue
            v = str(val).strip()
            if v in _NUMBER_COL_ALIASES and num_col is None:
                num_col = c
            if v == _OUTPUT_TIME_COL:
                time_col = c

        if num_col is None:
            warning(f"[table_writer] 未找到编号列（{_NUMBER_COL_ALIASES}），跳过回写")
            return

        # 「报告输出时间」列不存在则在末尾追加
        if time_col is None:
            time_col = total_cols + 1
            ws.cell(row=hdr_row, column=time_col, value=_OUTPUT_TIME_COL)
            info(f"[table_writer] 自动新增列「{_OUTPUT_TIME_COL}」（第 {time_col} 列）")

        # ── 逐行查找匹配编号，写入时间 ──
        matched = False
        for r in range(hdr_row + 1, total_rows + 1):
            cell_val = ws.cell(row=r, column=num_col).value
            if cell_val is None:
                continue
            if str(cell_val).strip() == number.strip():
                ws.cell(row=r, column=time_col, value=time_str)
                matched = True
                info(f"[table_writer] 已回写报告输出时间 → 编号 {number}，时间 {time_str}")
                break

        if not matched:
            warning(f"[table_writer] 未在物料信息表中找到编号 {number}，跳过回写")
            return

        wb.save(table_path)

    except Exception as exc:
        warning(f"[table_writer] 回写物料信息表失败（{number}）: {exc}")


__all__ = ["write_output_time"]
