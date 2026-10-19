"""电性/可靠性 Excel 表格构建器（spec 模块 5 / v5.0 新增）。

职责：从电性测试 Excel 读取数据（含样式），构建结构化表格数据，
供 ``template_engine`` + ``DocumentEngine.insert_table`` 消费，最终写入 Word。

核心特点：
  - 仅依赖 ``openpyxl``，不依赖 ``win32com``，跨平台可运行。
  - 不导出任何图片、不产生任何临时文件。
  - 使用 ``try / finally`` 确保工作簿正确关闭。

典型调用（见 重构.md §6.2）：
    table_data = build_table_data(folder_path, source_folder_name)
    # -> {"rows": List[List[str]], "styles": {...}, "merged_ranges": [...]}
    # 失败返回 None
"""
import os
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

from common.exception import ExcelError
from common.logger import warning

# 默认跳过的列/行，使用「原始 1-based 坐标」表达，便于与 Excel 直观对应：
#   - 跳过 B 列、C 列（1-based 列号分别为 2、3）
#   - 跳过第 22~27 行（1-based 行号 22..27）
_DEFAULT_SKIP_COLS = (2, 3)
_DEFAULT_SKIP_ROWS = tuple(range(22, 28))


def _cell_text(cell) -> str:
    """将单元格值按显示格式规范为字符串；None / 空 → 空串。

    - 百分比格式（number_format 含 '%'）：值 × 100，保留两位小数，加 '%'。
    - 整数浮点：去掉多余 '.0'。
    - 无限循环小数：保留最多 3 位小数（去掉尾部多余的 0）。
    - 文本：清理所有可能的换行符（\n、\r、\u2028、\u2029、\v、\f）→ 空格。
    """
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        fmt = getattr(cell, "number_format", None) or ""
        if "%" in fmt:
            return f"{value * 100:.2f}%"
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return f"{value:.3f}".rstrip("0").rstrip(".")
        return str(value)

    # 文本：清理所有换行符（含 Unicode 行分隔符）→ 空格
    import re as _re
    s = str(value)
    s = _re.sub(r"[\r\n\u2028\u2029\v\f]+", " ", s)
    # 压缩连续多个空格为一个
    s = _re.sub(r" {2,}", " ", s)
    return s.strip()


def _norm_color(color) -> Optional[str]:
    """从 openpyxl 颜色对象提取 '#RRGGBB'；无法解析时返回 None。

    注意：openpyxl 对显式设置的纯色填充常存为 ARGB（前两位为 alpha），
    例如灰底 ``F0F0F0`` 可能存成 ``00F0F0F0``，故一律取**后 6 位**为真实 RGB，
    仅 ``00000000`` / ``FFFFFFFF`` 视为「无 / 自动」而丢弃。
    """
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if not rgb or not isinstance(rgb, str):
        return None
    if rgb in ("00000000", "FFFFFFFF"):
        return None
    if len(rgb) == 8:
        return "#" + rgb[2:]
    if len(rgb) == 6:
        return "#" + rgb
    return None


def _header_row(ws, max_scan: int = 30) -> Optional[int]:
    """扫描前 ``max_scan`` 行，返回首个包含 'Item Name' 的行的 1-based 行号。"""
    for r in range(1, min(max_scan, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v and "Item Name" in str(v):
                return r
    return None

def _read_excel_images(ws) -> list:
    """从 worksheet 读取所有嵌入图片。

    返回 [{"data": bytes, "name": str, "row": int, "col": int}, ...]
    按 (row, col) 从上到下、从左到右排序。
    """
    images = []
    try:
        raw = list(getattr(ws, "_images", []))
    except Exception:
        return images
    for idx, img in enumerate(raw):
        try:
            # 拿二进制（openpyxl 不同版本差异）
            data = img._data
            if callable(data):
                data = data()
            if not isinstance(data, (bytes, bytearray)):
                continue
            # 拿锚点（0-based）
            row = col = 0
            try:
                anchor = img.anchor
                if hasattr(anchor, "_from"):
                    row = int(anchor._from.row)
                    col = int(anchor._from.col)
            except Exception:
                pass
            images.append({
                "data": bytes(data),
                "name": f"image_{idx + 1}",
                "row": row,
                "col": col,
            })
        except Exception as e:
            warning(f"[excel_table_builder] 读取 Excel 图片 {idx} 失败：{e}")
    images.sort(key=lambda x: (x["row"], x["col"]))
    return images

def build_table_data(folder_path: str = None,
                     source_folder_name: str = None,
                     table_width_cm: float = 18.0,
                     skip_columns=None,
                     skip_rows=None,
                     require_header: bool = True,
                     xlsx_path: str = None,
                     force_font_name: str = None,
                     force_font_size: float = None) -> Optional[Dict]:
    """读取电性 Excel，返回结构化表格数据。

    两种调用方式：
      1) 目录模式（兼容旧调用）：build_table_data(folder_path, source_folder_name, ...)
      2) 路径模式（新增）：build_table_data(xlsx_path="C:/xxx.xlsx", require_header=False)

    Args:
        folder_path: 目标文件夹路径（目录模式必传）。
        source_folder_name: 原始记录子文件夹真实名（目录模式必传）。
        table_width_cm: 表格总宽度（默认 18.0cm）。
        skip_columns: 跳过列（1-based 元组）；None 用默认 (2,3)，传 () 表示不跳过。
        skip_rows: 跳过行（1-based 元组）；None 用默认 (22..27)，传 () 表示不跳过。
        require_header: True 时查 'Item Name' 表头行；False 从第 1 行开始读。
        xlsx_path: 若传此参数，直接读该文件，忽略 folder_path/source_folder_name。
        force_font_name: 若非 None，所有单元格字体名强制为该值（如 "Times New Roman"）。
        force_font_size: 若非 None，所有单元格字号强制为该值（如 10.5）。

    Returns:
        ``{"rows": List[List[str]],
            "styles": {"r,c": {"bg","color","bold","font_name","font_size","align"}},
            "merged_ranges": [[min_r, min_c, max_r, max_c], ...],
            "width_cm": float}``
        失败（无文件/无数据/异常）返回 ``None``。
    """
    skip_cols = set(skip_columns if skip_columns is not None else _DEFAULT_SKIP_COLS)
    skip_rows_set = set(skip_rows if skip_rows is not None else _DEFAULT_SKIP_ROWS)

    # ── 定位 xlsx 路径 ────────────────────────────────────────────────
    if xlsx_path:
        xls_path = xlsx_path
        if not os.path.isfile(xls_path):
            warning(f"[excel_table_builder] xlsx 不存在：{xls_path}")
            return None
    else:
        base = os.path.join(folder_path, "原始记录", source_folder_name)
        if not os.path.isdir(base):
            warning(f"[excel_table_builder] 原始记录子文件夹不存在：{base}")
            return None
        xlsx_files = [f for f in os.listdir(base)
                      if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
        if not xlsx_files:
            warning(f"[excel_table_builder] 未找到 .xlsx：{base}")
            return None
        xls_path = os.path.join(base, sorted(xlsx_files)[0])

    wb = None
    try:
        try:
            wb = load_workbook(xls_path, data_only=True)
        except Exception as e:
            warning(f"[excel_table_builder] openpyxl 打开失败：{xls_path} -> {e}")
            return None

        ws = wb[wb.sheetnames[0]]

        # ── 确定数据起始行 ────────────────────────────────────────────
        if require_header:
            hr = _header_row(ws)
            if hr is None:
                warning(f"[excel_table_builder] 未找到包含 'Item Name' 的表头行：{xls_path}")
                return None
            warning(f"[excel_table_builder] require_header=True，表头行={hr}，文件={xls_path}")
        else:
            hr = 1   # 直接从第 1 行开始
            warning(f"[excel_table_builder] require_header=False，从第 1 行直接读取，文件={xls_path}")

        # 有效区域：hr 起 → 最后非空行；列 从第一个非空列 → 最后非空列
        last_row = hr
        last_col = 1
        first_col = ws.max_column + 1   # 初始化为一个大值，后面取最小
        for r in range(hr, ws.max_row + 1):
            row_has = False
            for c in range(1, ws.max_column + 1):
                if ws.cell(row=r, column=c).value not in (None, ""):
                    row_has = True
                    if c > last_col:
                        last_col = c
                    if c < first_col:
                        first_col = c
            if row_has:
                last_row = r
        if first_col > last_col:
            first_col = 1   # 兜底（理论上不会）

        warning(f"[excel_table_builder] 数据区域：行 {hr}~{last_row}，列 1~{last_col}，"
                f"skip_cols={sorted(skip_cols)}，skip_rows={sorted(skip_rows_set)}")

        # 构建过滤后矩阵 + 原始→过滤坐标映射
        orig_to_new: Dict[Tuple[int, int], Tuple[int, int]] = {}
        new_rows: List[List[str]] = []
        new_styles: Dict[str, Dict] = {}
        nr = 0
        for r in range(hr, last_row + 1):
            if r in skip_rows_set:
                continue
            nrow: List[str] = []
            row_styles: List[Dict] = []
            row_coords: List[Tuple[int, int]] = []
            for c in range(first_col, last_col + 1):   # ← 从 first_col 开始
                if c in skip_cols:
                    continue
                cell = ws.cell(row=r, column=c)
                text = _cell_text(cell)
                nrow.append(text)
                font = cell.font
                fill = cell.fill

                _font_name = force_font_name if force_font_name else (
                    getattr(font, "name", None) if font is not None else None)
                _font_size = force_font_size if force_font_size else (
                    getattr(font, "size", None)
                    if (font is not None and isinstance(getattr(font, "size", None), (int, float)))
                    else None)

                style = {
                    "bg": _norm_color(getattr(fill, "fgColor", None))
                          if (fill is not None and getattr(fill, "patternType", None)) else None,
                    "color": _norm_color(getattr(font, "color", None)) if font is not None else None,
                    "bold": bool(getattr(font, "bold", False)) if font is not None else False,
                    "font_name": _font_name,
                    "font_size": _font_size,
                    "align": (cell.alignment.horizontal or "left").lower(),
                }
                row_styles.append(style)
                row_coords.append((r, c))

            # ── 跳过整行全空（合并单元格 slave 行会产生这类空行）──
            if not any((v or "").strip() for v in nrow):
                continue

            # 正式加入
            for nc, (style, (rr, cc)) in enumerate(zip(row_styles, row_coords)):
                new_styles[f"{nr},{nc}"] = style
                orig_to_new[(rr, cc)] = (nr, nc)
            new_rows.append(nrow)
            nr += 1

        if not new_rows:
            warning(f"[excel_table_builder] 过滤后无数据：{xls_path}")
            return None

        # 末尾裁剪：去除尾部全空行与全空列
        while new_rows and all(v == "" for v in new_rows[-1]):
            new_rows.pop()
        if new_rows:
            ncols = len(new_rows[0])
            while ncols > 0 and all(row[ncols - 1] == "" for row in new_rows):
                ncols -= 1
            new_rows = [row[:ncols] for row in new_rows]
            nrows_final = len(new_rows)
            new_styles = {k: v for k, v in new_styles.items()
                          if int(k.split(",")[0]) < nrows_final
                          and int(k.split(",")[1]) < ncols}

        # 合并单元格：映射到过滤后 0-based 坐标
        merged_ranges: List[List[int]] = []
        for mr in ws.merged_cells.ranges:
            coords = []
            for rr in range(mr.min_row, mr.max_row + 1):
                for cc in range(mr.min_col, mr.max_col + 1):
                    if (rr, cc) in orig_to_new:
                        coords.append(orig_to_new[(rr, cc)])
            if coords:
                min_r = min(c[0] for c in coords)
                max_r = max(c[0] for c in coords)
                min_c = min(c[1] for c in coords)
                max_c = max(c[1] for c in coords)
                merged_ranges.append([min_r, min_c, max_r, max_c])

        warning(f"[excel_table_builder] 构建完成：{len(new_rows)} 行 × "
                f"{len(new_rows[0]) if new_rows else 0} 列，"
                f"合并区间 {len(merged_ranges)} 个")
        return {
            "rows": new_rows,
            "styles": new_styles,
            "merged_ranges": merged_ranges,
            "width_cm": table_width_cm,
            "images": _read_excel_images(ws),   # 新增
        }
    except ExcelError:
        raise
    except Exception as e:
        warning(f"[excel_table_builder] 读取异常：{xls_path} -> {e}")
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass