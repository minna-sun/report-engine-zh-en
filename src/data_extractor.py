"""电性数据导出模块（v5.3 · 纯 Python 版）。

职责：从电性测试 Excel 文件读取数据，用 openpyxl + PIL 绘制表格图片（300 DPI）并
保存到系统临时目录，供 ReliabilityCycleStrategy 插入 Word。

数据流：
  Excel(.xlsx) ──openpyxl──> 单元格数据/样式
                 ──PIL──> PNG 图片（300 DPI）
                          ──> 临时目录

重要约束：
  - ❌ 不得导入 win32com；不依赖 Excel 应用程序；不需要安装 Office。
  - ✅ 只使用 openpyxl + PIL（Pillow）+ 标准库。

对外接口：
  extract_electrical(folder_path, source_folder_name)
    返回 {"image_path": str, "col_count": int}，失败返回 None。
"""

import os
import tempfile

import openpyxl
from PIL import Image, ImageDraw, ImageFont

from common.logger import info, warning

# ── 临时目录 ────────────────────────────────────────────────────────────
# 临时目录（支持环境变量覆盖，用于多进程隔离）
_TASK_ID = os.environ.get("REPORT_PLATFORM_TASK_ID", "")
_TEMP_BASE = os.environ.get("REPORT_PLATFORM_TEMP", tempfile.gettempdir())
_TEMP_DIR = os.path.join(
    _TEMP_BASE,
    "report_platform_electrical" + (f"_{_TASK_ID}" if _TASK_ID else "")
)
os.makedirs(_TEMP_DIR, exist_ok=True)

# ── 渲染参数 ────────────────────────────────────────────────────────────
_DPI = 1200
_SCALE = _DPI / 96          # 96 dpi → 1200 dpi 缩放系数（约 12.5）
_BASE_ROW_H = 20            # 每行基准高度（96 dpi 像素）
_FONT_SIZE = 9              # 字号（pt）
_PAD_X = 4                  # 单元格水平内边距（96 dpi 像素）
_PAD_Y = 2                  # 单元格垂直内边距（96 dpi 像素）

# 颜色常量
_WHITE = (255, 255, 255)
_BLACK = (0, 0, 0)
_BORDER = (180, 180, 180)
_HEADER_BG = (220, 220, 220)  # 分段标题行

# 分段标题关键词（大小写不敏感）
_SECTION_KEYWORDS = ("before", "after", "change rate")

# 跳过的列索引（Excel 1-based：B=2, C=3）
_SKIP_COLS_1BASED = {2, 3}

# 跳过的行索引（Excel 1-based：22~27）
_SKIP_ROWS_1BASED = set(range(22, 28))


def _fmt_cell_value(cell) -> str:
    """将单元格值按显示格式转为字符串。

    - 百分比格式（number_format 含 '%'）：值 × 100，保留两位小数，加 '%'。
    - 整数浮点（无小数部分）：去掉多余 '.0'。
    - 其他：直接转 str。
    """
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        fmt = getattr(cell, "number_format", None) or ""
        if "%" in fmt:
            return f"{v * 100:.2f}%"
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
    return str(v)


# ─────────────────────────────────────────────────────────────────────────
# 公开接口
# ─────────────────────────────────────────────────────────────────────────

def extract_electrical(folder_path: str, source_folder_name: str,
                       require_header: bool = True,
                       skip_columns: tuple = (2, 3),
                       skip_rows: tuple = tuple(range(22, 28)),
                       force_header_row: int = None):
    """从电性 Excel 提取数据并生成 PNG 图片。

    Args:
        require_header: True（默认）时在前 30 行查找含 'Item Name' 的表头行；
                        False 时直接以第 1 行为数据起始行（适用于无特定表头场景）。
        skip_columns: 跳过的列（1-based），默认 (2, 3)（B、C 列）。
        skip_rows: 跳过的行（1-based），默认 22~27 行。

    Returns:
        成功：``{"image_path": str, "col_count": int}``
        失败：``None``
    """
    try:
        return _extract(folder_path, source_folder_name,
                        require_header=require_header,
                        skip_cols_set=set(skip_columns),
                        skip_rows_set=set(skip_rows),
                        force_header_row=force_header_row)
    except Exception as e:
        warning(f"[data_extractor] 未预期异常 ({source_folder_name}): {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────
# 内部实现
# ─────────────────────────────────────────────────────────────────────────

def _extract(folder_path: str, source_folder_name: str,
             require_header: bool = True,
             skip_cols_set: set = None,
             skip_rows_set: set = None,
             force_header_row: int = None):
    if skip_cols_set is None:
        skip_cols_set = _SKIP_COLS_1BASED
    if skip_rows_set is None:
        skip_rows_set = _SKIP_ROWS_1BASED

    # ── 1. 定位 xlsx ──────────────────────────────────────────────────
    data_dir = os.path.join(folder_path, "原始记录", source_folder_name)
    xlsx_path = _find_first_xlsx(data_dir)
    if xlsx_path is None:
        warning(f"[data_extractor] 未找到 .xlsx 文件: {data_dir}")
        return None

    # ── 2. 打开工作簿 ─────────────────────────────────────────────────
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        warning(f"[data_extractor] 打开 Excel 失败 ({xlsx_path}): {e}")
        return None
    ws = wb.active

    # ── 3. 确定表头行 ─────────────────────────────────────────────────
    if force_header_row is not None:
        header_row = force_header_row
        info(f"[data_extractor] force_header_row={header_row}，"
             f"从该行直接读取，文件={xlsx_path}")
    elif require_header:
        header_row = _find_header_row(ws)
        info(f"[data_extractor] require_header=True，表头行={header_row}，文件={xlsx_path}")
    else:
        header_row = 1
        info(f"[data_extractor] require_header=False，从第 1 行直接读取，文件={xlsx_path}")

    # ── 4. 读取数据（含合并单元格与字体颜色）─────────────────────────────
    grid, col_count = _read_grid(ws, header_row,
                                 skip_cols_set=skip_cols_set,
                                 skip_rows_set=skip_rows_set)
    if not grid or col_count == 0:
        warning(f"[data_extractor] 数据为空 ({xlsx_path})")
        return None

    # ── 5. 绘制图片 ────────────────────────────────────────────────────
    img_name = f"electrical_{os.path.basename(data_dir)}.png"
    img_path = os.path.join(_TEMP_DIR, img_name)
    _render_table(grid, img_path)
    info(f"[data_extractor] 已导出图片: {img_path} (列数={col_count})")
    return {"image_path": img_path, "col_count": col_count}


def render_excel_to_png(xlsx_path: str,
                        output_name: str = None,
                        require_header: bool = False,
                        skip_columns: tuple = (),
                        skip_rows: tuple = (),
                        force_header_row: int = None):
    """从指定 xlsx 路径直接渲染 PNG 图片（不依赖文件夹名）。

    Args:
        xlsx_path: Excel 文件绝对路径。
        output_name: 输出文件名（不含路径）；None 时以 xlsx 文件名生成。
        require_header: True 时自动查找 'Item Name' 表头行；False 从第 1 行读。
        skip_columns: 跳过的列（1-based），默认空（不跳任何列）。
        skip_rows: 跳过的行（1-based），默认空（不跳任何行）。
        force_header_row: 强制指定表头行；非 None 时忽略 require_header。

    Returns:
        成功：PNG 文件路径（str）；失败：None。
    """
    try:
        if not xlsx_path or not os.path.isfile(xlsx_path):
            warning(f"[data_extractor] render_excel_to_png：文件不存在 {xlsx_path}")
            return None
        skip_cols_set = set(skip_columns)
        skip_rows_set = set(skip_rows)
        try:
            wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        except Exception as e:
            warning(f"[data_extractor] render_excel_to_png：打开 Excel 失败 ({xlsx_path}): {e}")
            return None
        ws = wb.active
        if force_header_row is not None:
            header_row = force_header_row
        elif require_header:
            header_row = _find_header_row(ws)
        else:
            header_row = 1
        grid, col_count = _read_grid(ws, header_row,
                                     skip_cols_set=skip_cols_set,
                                     skip_rows_set=skip_rows_set)
        if not grid or col_count == 0:
            warning(f"[data_extractor] render_excel_to_png：数据为空 ({xlsx_path})")
            return None
        if output_name is None:
            output_name = os.path.splitext(os.path.basename(xlsx_path))[0] + "_eds.png"
        img_path = os.path.join(_TEMP_DIR, output_name)
        _render_table(grid, img_path)
        info(f"[data_extractor] render_excel_to_png：已导出 {img_path}")
        return img_path
    except Exception as e:
        warning(f"[data_extractor] render_excel_to_png 未预期异常 ({xlsx_path}): {e}")
        return None


def _find_first_xlsx(dir_path: str):
    if not os.path.isdir(dir_path):
        return None
    for name in sorted(os.listdir(dir_path)):
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            return os.path.join(dir_path, name)
    return None


def _find_header_row(ws) -> int:
    for r in range(1, 31):
        for c in range(1, min(ws.max_column + 1, 20)):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            if "item name" in val.lower():
                return r
    return 16


# ─────────────────────────────────────────────────────────────────────────
# 数据读取：返回 Grid（二维列表 of CellInfo）
#
# CellInfo 字段：
#   value     str        单元格显示文本
#   bg        (r,g,b)    背景色
#   fg        (r,g,b)    字体颜色
#   bold      bool       是否加粗
#   align     str        'left' / 'center' / 'right'
#   is_section bool      是否为分段标题行
#   colspan   int        横向合并格数（>=1）
#   rowspan   int        纵向合并格数（>=1）
#   is_slave  bool       True 表示被合并主格覆盖，渲染时跳过
# ─────────────────────────────────────────────────────────────────────────

def _read_grid(ws, header_row: int,
               skip_cols_set: set = None,
               skip_rows_set: set = None):
    """返回 (grid, col_count)。grid 是二维 list[list[dict]]。"""
    if skip_cols_set is None:
        skip_cols_set = _SKIP_COLS_1BASED
    if skip_rows_set is None:
        skip_rows_set = _SKIP_ROWS_1BASED

    last_row = ws.max_row
    last_col = ws.max_column

    # ── 构建合并区间映射：(r,c) → (master_r, master_c, min_r, min_c, max_r, max_c)
    merge_master = {}   # (r,c) → (master_r, master_c)
    merge_ranges = {}   # (master_r, master_c) → (min_r, min_c, max_r, max_c)
    for rng in ws.merged_cells.ranges:
        mr, mc = rng.min_row, rng.min_col
        merge_ranges[(mr, mc)] = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merge_master[(r, c)] = (mr, mc)

    # ── 筛选行、列（剔除跳过项 + 末尾全空）
    active_cols = _filter_cols(ws, last_col, header_row, last_row, merge_master,
                               skip_cols_set)
    active_rows = _filter_rows(ws, last_row, header_row, active_cols, merge_master,
                               skip_rows_set)

    if not active_rows or not active_cols:
        return [], 0

    n_rows = len(active_rows)
    n_cols = len(active_cols)

    # ── 建立 Excel坐标 → 输出坐标 的反查映射
    excel_col_to_out = {ec: oc for oc, ec in enumerate(active_cols)}
    excel_row_to_out = {er: or_ for or_, er in enumerate(active_rows)}

    # ── 初始化 grid
    grid = []
    for _ in range(n_rows):
        row = []
        for _ in range(n_cols):
            row.append({
                "value": "", "bg": _WHITE, "fg": _BLACK,
                "bold": False, "align": "left",
                "is_section": False, "colspan": 1, "rowspan": 1,
                "is_slave": False,
            })
        grid.append(row)

    # ── 填充数据，计算合并跨度
    for or_, er in enumerate(active_rows):
        for oc, ec in enumerate(active_cols):
            cell_info = grid[or_][oc]

            # 取主格坐标
            master = merge_master.get((er, ec), (er, ec))
            mr, mc = master
            raw_cell = ws.cell(row=mr, column=mc)

            # 读取值与样式（从主格）
            cell_info["value"] = _fmt_cell_value(raw_cell)
            cell_info["bg"] = _get_bg(raw_cell)
            cell_info["fg"] = _get_fg(raw_cell)
            cell_info["bold"] = bool(raw_cell.font and raw_cell.font.bold)
            cell_info["align"] = _get_align(raw_cell)

            # 计算合并跨度（仅对主格设置，slave 格在下一步标记）
            if (er, ec) == (mr, mc) and (mr, mc) in merge_ranges:
                min_r, min_c, max_r, max_c = merge_ranges[(mr, mc)]
                # colspan：在输出列中，从 oc 开始，有多少连续列属于同一主格
                cs = 1
                for oc2 in range(oc + 1, n_cols):
                    ec2 = active_cols[oc2]
                    if merge_master.get((er, ec2), (er, ec2)) == (mr, mc):
                        cs += 1
                    else:
                        break
                # rowspan：在输出行中，从 or_ 开始，有多少连续行属于同一主格
                rs = 1
                for or2 in range(or_ + 1, n_rows):
                    er2 = active_rows[or2]
                    if merge_master.get((er2, ec), (er2, ec)) == (mr, mc):
                        rs += 1
                    else:
                        break
                cell_info["colspan"] = cs
                cell_info["rowspan"] = rs

    # ── 标记 slave 格（被主格合并覆盖的格）
    for or_, er in enumerate(active_rows):
        for oc, ec in enumerate(active_cols):
            master = merge_master.get((er, ec), (er, ec))
            mr, mc = master
            if (er, ec) != (mr, mc):
                # 是 slave
                grid[or_][oc]["is_slave"] = True

    # ── 标记分段标题行（is_section）
    for or_, row_cells in enumerate(grid):
        joined = " ".join(d["value"] for d in row_cells if not d["is_slave"]).lower()
        is_section = any(kw in joined for kw in _SECTION_KEYWORDS)
        if is_section:
            for d in row_cells:
                d["is_section"] = True
                d["bg"] = _HEADER_BG
                d["bold"] = True
                d["align"] = "center"

    return grid, n_cols


# ─────────────────────────────────────────────────────────────────────────
# 行/列筛选辅助
# ─────────────────────────────────────────────────────────────────────────

def _filter_cols(ws, last_col, header_row, last_row, merge_master,
                 skip_cols_set=None):
    if skip_cols_set is None:
        skip_cols_set = _SKIP_COLS_1BASED
    cols = [c for c in range(1, last_col + 1) if c not in skip_cols_set]
    # 去末尾全空列
    while cols:
        c = cols[-1]
        if _col_empty(ws, c, header_row, last_row, merge_master):
            cols.pop()
        else:
            break
    return cols


def _filter_rows(ws, last_row, header_row, active_cols, merge_master,
                 skip_rows_set=None):
    if skip_rows_set is None:
        skip_rows_set = _SKIP_ROWS_1BASED
    rows = [r for r in range(header_row, last_row + 1) if r not in skip_rows_set]
    # 去末尾全空行
    while rows:
        r = rows[-1]
        if _row_empty(ws, r, active_cols, merge_master):
            rows.pop()
        else:
            break
    return rows


def _col_empty(ws, col, header_row, last_row, merge_master):
    for r in range(header_row, last_row + 1):
        mr, mc = merge_master.get((r, col), (r, col))
        v = ws.cell(row=mr, column=mc).value
        if v is not None and str(v).strip():
            return False
    return True


def _row_empty(ws, row, cols, merge_master):
    for c in cols:
        mr, mc = merge_master.get((row, c), (row, c))
        v = ws.cell(row=mr, column=mc).value
        if v is not None and str(v).strip():
            return False
    return True


# ─────────────────────────────────────────────────────────────────────────
# 样式读取
# ─────────────────────────────────────────────────────────────────────────

def _get_bg(cell):
    """读取背景色（RGB），默认白色。"""
    try:
        fill = cell.fill
        if fill is None:
            return _WHITE
        # patternFill：fgColor 是填充色
        if fill.patternType and fill.patternType not in ("none", None):
            color = fill.fgColor
            if color:
                rgb = _resolve_color(color)
                if rgb:
                    return rgb
        # solidFill（部分 xlsx）
        if hasattr(fill, "start_color") and fill.start_color:
            rgb = _resolve_color(fill.start_color)
            if rgb:
                return rgb
    except Exception:
        pass
    return _WHITE


def _get_fg(cell):
    """读取字体颜色（RGB），默认黑色。支持 RGB 类型；主题色回退为黑色。"""
    try:
        font = cell.font
        if font is None:
            return _BLACK
        color = font.color
        if color is None:
            return _BLACK
        rgb = _resolve_color(color)
        if rgb:
            return rgb
    except Exception:
        pass
    return _BLACK


def _resolve_color(color_obj):
    """将 openpyxl Color 对象解析为 (r,g,b)，失败返回 None。
    支持四种类型：
      'rgb'     → color_obj.rgb 为 'FFRRGGBB' 或 '00RRGGBB'
      'theme'   → 主题色（Office 默认主题色板映射）
      'indexed' → 索引色（Excel 标准 56 色板映射）
      'auto'    → 自动色（黑色）
    """
    try:
        color_type = getattr(color_obj, "type", None)
        # ── 类型1：RGB 直接值
        if color_type == "rgb":
            rgb_str = color_obj.rgb
            if rgb_str and len(rgb_str) == 8:
                alpha = int(rgb_str[0:2], 16)
                r = int(rgb_str[2:4], 16)
                g = int(rgb_str[4:6], 16)
                b = int(rgb_str[6:8], 16)
                # 透明（alpha=0）或空（全黑且 alpha=0）→ 无效
                if alpha == 0 and r == 0 and g == 0 and b == 0:
                    return None
                return (r, g, b)
        # ── 类型2：主题色（Office 默认主题色板）
        if color_type == "theme":
            theme_idx = getattr(color_obj, "theme", None)
            _THEME_COLORS = {
                0: (255, 255, 255),   # lt1 白
                1: (0, 0, 0),         # dk1 黑
                2: (231, 230, 230),   # lt2 浅灰
                3: (68, 84, 106),     # dk2 深灰
                4: (68, 114, 196),    # accent1 蓝
                5: (237, 125, 49),    # accent2 橙
                6: (165, 165, 165),   # accent3 灰
                7: (255, 192, 0),     # accent4 金
                8: (91, 155, 213),    # accent5 浅蓝
                9: (112, 173, 71),    # accent6 绿
            }
            if theme_idx is not None and theme_idx in _THEME_COLORS:
                return _THEME_COLORS[theme_idx]
        # ── 类型3：索引色（Excel 标准色板）
        if color_type == "indexed":
            idx = getattr(color_obj, "indexed", None)
            _INDEXED_COLORS = {
                0: (0, 0, 0),         # 黑
                1: (255, 255, 255),   # 白
                2: (255, 0, 0),       # 红
                3: (0, 255, 0),       # 绿
                4: (0, 0, 255),       # 蓝
                5: (255, 255, 0),     # 黄
                6: (255, 0, 255),     # 品红
                7: (0, 255, 255),     # 青
                8: (0, 0, 0),         # 黑
                9: (255, 255, 255),   # 白
                10: (255, 0, 0),      # 红
                11: (0, 255, 0),      # 绿
                12: (0, 0, 255),      # 蓝
                13: (255, 255, 0),    # 黄
                14: (255, 0, 255),    # 品红
                15: (0, 255, 255),    # 青
                16: (128, 0, 0),      # 深红
                17: (0, 128, 0),      # 深绿
                18: (0, 0, 128),      # 深蓝
                19: (128, 128, 0),    # 深黄
                20: (128, 0, 128),    # 深品红
                21: (0, 128, 128),    # 深青
                22: (192, 192, 192),  # 浅灰
                23: (128, 128, 128),  # 中灰
                24: (153, 153, 255),  # 浅蓝
                25: (153, 51, 102),   # 深紫红
                26: (255, 255, 204),  # 浅黄
                27: (204, 255, 255),  # 浅青
                28: (102, 0, 102),    # 紫
                29: (255, 128, 128),  # 浅红
                30: (0, 102, 204),    # 深蓝
                31: (204, 204, 255),  # 浅蓝紫
                32: (0, 0, 128),      # 深蓝
                33: (255, 0, 255),    # 品红
                34: (255, 255, 0),    # 黄
                35: (0, 255, 255),    # 青
                36: (128, 0, 128),    # 紫
                37: (128, 0, 0),      # 深红
                38: (0, 128, 128),    # 深青
                39: (0, 0, 255),      # 蓝
                40: (0, 204, 255),    # 亮青
                41: (204, 255, 255),  # 浅青
                42: (204, 204, 255),  # 浅蓝
                43: (255, 255, 153),  # 浅黄
                44: (153, 204, 255),  # 浅蓝
                45: (255, 153, 204),  # 浅粉
                46: (204, 153, 255),  # 浅紫
                47: (255, 204, 153),  # 浅橙
                48: (51, 102, 255),   # 蓝
                49: (51, 204, 204),   # 青绿
                50: (153, 204, 0),    # 黄绿
                51: (255, 204, 0),    # 金黄
                52: (255, 153, 0),    # 橙
                53: (255, 102, 0),    # 深橙
                54: (102, 102, 153),  # 蓝紫
                55: (150, 150, 150),  # 中灰
                56: (0, 0, 0),        # 黑
                57: (255, 255, 255),  # 白
                58: (255, 255, 255),  # 白
                59: (255, 255, 255),  # 白
                60: (255, 255, 255),  # 白
                61: (255, 255, 255),  # 白
                62: (255, 255, 255),  # 白
                63: (255, 255, 255),  # 白
                64: (0, 0, 0),        # 系统前景（黑）
                65: (255, 255, 255),  # 系统背景（白）
            }
            if idx is not None and idx in _INDEXED_COLORS:
                return _INDEXED_COLORS[idx]
        # ── 类型4：自动色（视为黑色）
        if color_type == "auto":
            return (0, 0, 0)
        # ── 兜底：尝试读 rgb 属性（部分对象 type 未设置但 rgb 有值）
        rgb_attr = getattr(color_obj, "rgb", None)
        if rgb_attr and isinstance(rgb_attr, str) and len(rgb_attr) == 8:
            alpha = int(rgb_attr[0:2], 16)
            r = int(rgb_attr[2:4], 16)
            g = int(rgb_attr[4:6], 16)
            b = int(rgb_attr[6:8], 16)
            if not (alpha == 0 and r == 0 and g == 0 and b == 0):
                return (r, g, b)
    except Exception:
        pass
    return None


def _get_align(cell):
    try:
        align = cell.alignment
        if align and align.horizontal:
            h = align.horizontal.lower()
            if "center" in h:
                return "center"
            if "right" in h:
                return "right"
    except Exception:
        pass
    return "left"


# ─────────────────────────────────────────────────────────────────────────
# PIL 渲染
# ─────────────────────────────────────────────────────────────────────────

def _render_table(grid, out_path: str) -> None:
    """将 grid 绘制为 PNG（300 DPI）。

    合并单元格（colspan/rowspan > 1）绘制为跨列/行大矩形；
    slave 格（is_slave=True）跳过绘制。
    """
    if not grid:
        return

    n_rows = len(grid)
    n_cols = len(grid[0])
    row_height = _BASE_ROW_H

    # 计算列宽（96 dpi 基准，按内容长度动态调整）
    col_widths = _calc_col_widths(grid, n_cols)

    # 画布尺寸（96 dpi 基准）
    total_w = sum(col_widths) + 1
    total_h = n_rows * row_height + 1

    # 放大到 300 dpi
    img_w = int(total_w * _SCALE)
    img_h = int(total_h * _SCALE)

    img = Image.new("RGB", (img_w, img_h), _WHITE)
    draw = ImageDraw.Draw(img)

    font, font_bold = _load_fonts(int(_FONT_SIZE * _SCALE))

    # 计算每列、每行的起始像素坐标（300 dpi）
    col_x = []  # col_x[c] = 该列左边 x 坐标（300 dpi）
    x = 0
    for cw in col_widths:
        col_x.append(int(x * _SCALE))
        x += cw
    col_x.append(int(x * _SCALE))  # 末尾哨兵

    row_y = []  # row_y[r] = 该行上边 y 坐标（300 dpi）
    for ri in range(n_rows):
        row_y.append(int(ri * row_height * _SCALE))
    row_y.append(int(n_rows * row_height * _SCALE))  # 末尾哨兵

    # ── 绘制每个格
    for ri, row_cells in enumerate(grid):
        for ci, cell in enumerate(row_cells):
            if cell["is_slave"]:
                # slave 格：背景已被主格覆盖，无需绘制
                continue

            cs = cell["colspan"]   # 横向跨列数
            rs = cell["rowspan"]   # 纵向跨行数

            # 计算矩形坐标（考虑跨行列）
            x1 = col_x[ci]
            y1 = row_y[ri]
            x2 = col_x[min(ci + cs, n_cols)]
            y2 = row_y[min(ri + rs, n_rows)]

            # 背景
            bg = cell["bg"] if not cell["is_section"] else _HEADER_BG
            draw.rectangle([x1, y1, x2 - 1, y2 - 1], fill=bg)

            # 文字
            fnt = font_bold if cell["bold"] else font
            _draw_text(
                draw, cell["value"],
                x1, y1, x2, y2,
                cell["fg"], fnt, cell["align"],
                int(_PAD_X * _SCALE), int(_PAD_Y * _SCALE),
            )

            # 边框（整个合并格的外框 + 内部线由各格自己的外框拼合）
            draw.rectangle([x1, y1, x2 - 1, y2 - 1], outline=_BORDER)

    img.save(out_path, format="PNG", dpi=(_DPI, _DPI))


def _calc_col_widths(grid, n_cols):
    """按列最长内容估算列宽（96 dpi 像素）。"""
    widths = []
    for ci in range(n_cols):
        max_len = 0
        for row in grid:
            if ci < len(row) and not row[ci]["is_slave"]:
                # 对于合并格，内容分摊到 colspan 列
                cs = row[ci]["colspan"]
                effective = max(1, len(str(row[ci]["value"])) // cs)
                max_len = max(max_len, effective)
        # 每字符约 7px（96 dpi），最小 40，最大 200
        widths.append(max(40, min(200, max_len * 7 + _PAD_X * 2)))
    return widths


def _load_fonts(size: int):
    """加载字体（优先中文字体，回退 Pillow 内置可缩放字体）。"""
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/cour.ttf",
    ]
    font = None
    for path in candidates:
        if os.path.exists(path):
            try:
                font = ImageFont.truetype(path, size)
                break
            except Exception:
                continue
    if font is None:
        # load_default() 是固定 10px 位图字体，高 DPI 下放大后极糊。
        # Pillow >= 10.1.0 支持 load_default(size=N)，可缩放；
        # 旧版本不接受 size 参数，捕获后使用位图字体并记录警告。
        try:
            font = ImageFont.load_default(size=size)
            warning(f"[data_extractor] 未找到 TrueType 字体，使用 Pillow 内置字体（size={size}）")
        except TypeError:
            font = ImageFont.load_default()
            warning(f"[data_extractor] Pillow 版本过旧，load_default 不支持 size 参数，"
                    f"图片文字可能模糊（size={size}）；建议升级 Pillow >= 10.1.0")

    bold_candidates = [
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
    ]
    font_bold = font
    for path in bold_candidates:
        if os.path.exists(path):
            try:
                font_bold = ImageFont.truetype(path, size)
                break
            except Exception:
                continue
    return font, font_bold


def _draw_text(draw, text, x1, y1, x2, y2, color, font, align, pad_x, pad_y):
    """在矩形区域内绘制单行文字（超出截断）。"""
    if not text:
        return
    cell_w = x2 - x1 - pad_x * 2
    cell_h = y2 - y1 - pad_y * 2
    if cell_w <= 0 or cell_h <= 0:
        return

    # 估算文字尺寸
    try:
        bbox = font.getbbox(text)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
    except Exception:
        fs = getattr(font, "size", 8)
        tw = len(text) * fs // 2
        th = fs

    # 超宽截断
    if tw > cell_w and tw > 0:
        ratio = cell_w / tw
        cut = max(1, int(len(text) * ratio) - 1)
        text = text[:cut] + "…"
        try:
            bbox = font.getbbox(text)
            tw = bbox[2] - bbox[0]
        except Exception:
            tw = len(text) * 7

    # 水平对齐
    if align == "center":
        tx = x1 + pad_x + max(0, (cell_w - tw) // 2)
    elif align == "right":
        tx = x2 - pad_x - tw
    else:
        tx = x1 + pad_x

    # 垂直居中
    ty = y1 + pad_y + max(0, (cell_h - th) // 2)

    draw.text((tx, ty), text, fill=color, font=font)
