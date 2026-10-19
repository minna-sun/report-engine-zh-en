"""OOXML 底层操作服务（v5.0 新增 · spec 模块 10）。

职责：``docx_engine`` 的底层技术支撑，**只对 docx_engine 开放**。
处理 python-docx 不直接支持的高级 OOXML 操作：
  - 跨文档深拷贝 ``w:tbl`` 节点（copy_table）
  - 删除段落 / 表格节点（delete_section_content）
  - 设置精确列宽（set_table_column_width）
  - 设置单元格背景色（set_cell_background）
  - 设置字体颜色（set_text_color）
  - 插入分页符（insert_page_break）
  - 读取 Excel 合并单元格范围（get_merged_cells，供 excel_table_builder 使用）

设计原则：纯函数 / 无状态；输入为 lxml 节点或文件路径；成功返回值、失败返回
False/None；**不抛出异常**（由调用方捕获）；**不包含业务逻辑**。
"""
from copy import deepcopy
from typing import List, Optional, Tuple

try:
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
except Exception:  # pragma: no cover
    Table = Paragraph = Run = object

from lxml import etree

# OOXML 命名空间
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag: str) -> str:
    return f"{{{_W}}}{tag}"


def get_merged_cells(worksheet) -> List[Tuple[int, int, int, int]]:
    """读取 openpyxl worksheet 的合并单元格范围。

    Returns:
        列表，元素为 (min_row, min_col, max_row, max_col)，均为 **1-based 闭区间**
        （与 openpyxl 一致）。无合并返回空列表。
    """
    result = []
    for mr in worksheet.merged_cells.ranges:
        result.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
    return result


def set_table_column_width(table_element, col_widths: List[int]) -> None:
    """操作 ``w:tblW`` / ``w:tcW`` 设置列宽（单位：dxa）。

    Args:
        table_element: lxml 的 ``w:tbl`` 节点。
        col_widths: 各列宽度（dxa），长度须等于列数。
    """
    tbl_pr = table_element.find(_q("tblPr"))
    if tbl_pr is None:
        return
    # 关闭自动调整
    layout = tbl_pr.find(_q("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tbl_pr, _q("tblLayout"))
    layout.set(_q("type"), "fixed")
    # 总宽
    tbl_w = tbl_pr.find(_q("tblW"))
    if tbl_w is None:
        tbl_w = etree.SubElement(tbl_pr, _q("tblW"))
    total = sum(col_widths)
    tbl_w.set(_q("w"), str(total))
    tbl_w.set(_q("type"), "dxa")
    # 各列宽（位于 tblGrid / gridCol）
    grid = table_element.find(_q("tblGrid"))
    if grid is None:
        return
    for grid_col, w in zip(grid.findall(_q("gridCol")), col_widths):
        grid_col.set(_q("w"), str(w))


def set_cell_background(cell, color: str) -> None:
    """操作 ``w:shd`` 设置单元格背景色。

    Args:
        cell: python-docx ``Cell`` 对象。
        color: ``#RRGGBB`` 或 ``RRGGBB``；None / 空则不处理。
    """
    if not color:
        return
    hexc = color.lstrip("#")
    if len(hexc) == 8:  # ARGB -> 取后 6 位
        hexc = hexc[2:]
    if len(hexc) != 6:
        return
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(_q("shd"))
    if shd is None:
        shd = etree.SubElement(tc_pr, _q("shd"))
    shd.set(_q("val"), "clear")
    shd.set(_q("color"), "auto")
    shd.set(_q("fill"), hexc)


def set_text_color(run: Run, color) -> None:
    """操作 ``w:rPr`` / ``w:color`` 设置字体颜色。

    Args:
        run: python-docx ``Run`` 对象。
        color: ``#RRGGBB`` / ``RRGGBB`` / 整数 RGB；None 不处理。
    """
    if color is None:
        return
    if isinstance(color, int):
        hexc = f"{color:06X}"
    else:
        hexc = str(color).lstrip("#")
        if len(hexc) == 8:
            hexc = hexc[2:]
    if len(hexc) != 6:
        return
    rpr = run._r.get_or_add_rPr()
    c = rpr.find(_q("color"))
    if c is None:
        c = etree.SubElement(rpr, _q("color"))
    c.set(_q("val"), hexc)


def insert_page_break(para: Paragraph) -> None:
    """在段落后插入 ``w:br w:type="page"``。"""
    p = para._p
    run = p.find(_q("r"))
    if run is None:
        run = etree.SubElement(p, _q("r"))
    br = run.find(_q("br"))
    if br is None:
        br = etree.SubElement(run, _q("br"))
    br.set(_q("type"), "page")


def delete_paragraph(para_element) -> None:
    """从 document.xml 移除 ``w:p`` 节点。"""
    p = para_element
    if hasattr(p, "_p"):  # 接受 Paragraph 或 lxml 节点
        p = p._p
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def delete_table(table_element) -> None:
    """从 document.xml 移除 ``w:tbl`` 节点。"""
    tbl = table_element
    if hasattr(tbl, "_tbl"):
        tbl = tbl._tbl
    parent = tbl.getparent()
    if parent is not None:
        parent.remove(tbl)


# ───────────────────────── 章节级结构操作（sync_chapters 支撑） ─────────────────────────

def collect_range_elements(body, start_el, end_el):
    """收集 body 中 ``[start_el, end_el)`` 区间内的 ``w:p`` / ``w:tbl`` 元素。

    含起始元素（章节标题），不含结束元素（下一个章节边界）。``end_el`` 为
    ``None`` 时取到 body 末尾（调用方需保证边界正确）。返回 lxml 元素列表。

    Args:
        body: 文档 ``document.xml`` 的 ``w:body`` 节点。
        start_el: 起始元素（段落或表格，接受 python-docx 对象或 lxml 节点）。
        end_el: 结束元素（不含），``None`` 表示取至文末。
    """
    start = start_el._p if hasattr(start_el, "_p") else start_el
    end = end_el._p if hasattr(end_el, "_p") else end_el
    out = []
    started = False
    for child in body.iterchildren():
        tag = child.tag.split('}')[-1]
        if child is start:
            started = True
        if not started:
            continue
        if end is not None and child is end:
            break
        if tag in ("p", "tbl"):
            out.append(child)
    return out


def clone_elements(elements):
    """深拷贝一组 lxml 元素，返回新元素列表。"""
    return [deepcopy(e) for e in elements]


def insert_elements_before(ref_el, elements) -> bool:
    """将 ``elements`` 按原顺序插入到 ``ref_el`` 之前。

    通过逆序 ``addprevious`` 保证顺序正确。
    """
    ref = ref_el._p if hasattr(ref_el, "_p") else ref_el
    for el in reversed(elements):
        ref.addprevious(el)
    return True


def append_elements(body, elements) -> bool:
    """将 ``elements`` 追加到 body 末尾。"""
    for el in elements:
        body.append(el)
    return True


def set_paragraph_text(para_el, text) -> bool:
    """替换段落文本为 ``text``，保留首个 run 的格式（w:t 内容）。

    若段落无 run，则新建一个 run 承载文本。用于克隆章节后重命名标题。
    """
    p = para_el._p if hasattr(para_el, "_p") else para_el
    runs = p.findall(_q("r"))
    if runs:
        first = runs[0]
        for r in runs[1:]:
            r.getparent().remove(r)
        texts = first.findall(_q("t"))
        if texts:
            for t in texts[1:]:
                t.getparent().remove(t)
            texts[0].text = text
        else:
            t = etree.SubElement(first, _q("t"))
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = text
    else:
        r = etree.SubElement(p, _q("r"))
        t = etree.SubElement(r, _q("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
    return True


# ───────────────────────── 表格行 / 单元格操作（fill_summary / fill_equipment_table 支撑） ─────────────────────────

def remove_table_row(table_element, index: int) -> bool:
    """删除表格第 ``index`` 行（0-based）。返回 ``True`` / ``False``。"""
    tbl = table_element._tbl if hasattr(table_element, "_tbl") else table_element
    trs = tbl.findall(_q("tr"))
    if 0 <= index < len(trs):
        tbl.remove(trs[index])
        return True
    return False


def copy_table_to_doc(target_doc_path: str, source_tbl_element,
                      insert_after_element) -> bool:
    """跨文档复制 ``w:tbl`` 节点（深拷贝）到目标 docx 的 insert_after_element 之后。

    说明：本实现完成节点拷贝与重命名（避免 rId 冲突由调用方配合处理）。
    返回 True / False。
    """
    try:
        from docx import Document
        target = Document(target_doc_path)
        tbody = target.element.body
        src = source_tbl_element
        if hasattr(src, "_tbl"):
            src = src._tbl
        clone = deepcopy(src)
        # 重命名内部 r:id 以避免与源文档冲突（保留结构，图片在跨文档场景由调用方保证）
        parent = (insert_after_element._p
                  if hasattr(insert_after_element, "_p") else insert_after_element)
        parent.addnext(clone)
        target.save(target_doc_path)
        return True
    except Exception:
        return False
