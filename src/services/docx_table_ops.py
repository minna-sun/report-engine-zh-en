"""表格操作 Mixin（v5.0 · 由 docx_engine.py 拆分而来）。

职责：``DocxEngine`` 的**表格**能力——表格插入 / 跨文档复制 / 列宽与边框格式、
行增删、单元格文本读写与着色、单元格样式应用。

以 Mixin 形式被 ``DocxEngine`` 继承；``copy_table`` 通过 ``self`` 调用
``NavigationMixin`` 的 ``find_paragraph`` / ``_find_anchor_in_section`` /
``_first_table_after``。

依赖：``services.docx_helpers``、``services.ooxml_service``、``lxml``。
"""
from typing import Optional

from lxml import etree as _lxml_etree

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from common.logger import warning

from services.docx_helpers import (
    _build_table_element,
    _cm_to_dxa,
    _deepclone_tbl,
    _ensure_table_borders,
    _is_in_table,
    _para_is_bold,
)
from services.ooxml_service import (
    _q,
    delete_table,
    remove_table_row,    # 补充此项
    set_cell_background,
    set_table_column_width,
    set_text_color,
)


class TableOpsMixin:
    """表格操作 Mixin。"""

    def insert_table(self, doc: Document, anchor: object,
                     table_data, width_cm: float) -> bool:
        """锚点段落后新建段落并插入表格，总宽固定 width_cm，列宽按数据自动分配。

        ``table_data`` 接受两种形态：
          - dict：``{"rows": [[...]], "styles": {...}, "merged_ranges": [...], "width_cm": float}``
                  （由 ``excel_table_builder.build_table_data`` 产出）
          - List[List[str]]：纯文本矩阵（兼容抽象接口签名）
        失败返回 False。
        """
        try:
            if isinstance(table_data, dict):
                rows = table_data.get("rows", [])
                styles = table_data.get("styles", {}) or {}
                merged = table_data.get("merged_ranges", []) or []
                if "width_cm" in table_data and table_data["width_cm"]:
                    width_cm = float(table_data["width_cm"])
            else:
                rows = table_data
                styles = {}
                merged = []
            if not rows:
                warning("[docx_engine] insert_table：空数据")
                return False

            anchor_para = anchor
            nrows = len(rows)
            ncols = max(len(r) for r in rows)
            # 用底层 API 在锚点段落后直接插入指定尺寸的表格
            from docx.table import Table as _Table
            tbl_el = _build_table_element(nrows, ncols)
            anchor_para._p.addnext(tbl_el)
            tbl = _Table(tbl_el, doc)
            # 填充
            for i, row in enumerate(rows):
                for j in range(ncols):
                    cell = tbl.cell(i, j)
                    cell.text = row[j] if j < len(row) else ""
                    key = f"{i},{j}"
                    st = styles.get(key)
                    if st:
                        self._apply_cell_style(cell, st)
            # 合并单元格
            for (mr, mc, mr2, mc2) in merged:
                try:
                    if mr2 > mr or mc2 > mc:
                        tbl.cell(mr, mc).merge(tbl.cell(mr2, mc2))
                except Exception:
                    pass
            # 清理合并后单元格里的多余空段落
            # 从 XML 层直接遍历 w:tc，避免 python-docx 的 cells 去重不稳定问题
            try:
                for _tc in tbl._tbl.iter(_q("tc")):
                    _ps = _tc.findall(_q("p"))
                    if len(_ps) <= 1:
                        continue
                    # 保留第一段（内容段）；从第 2 段起，只删"空段落"
                    for _p in _ps[1:]:
                        # 用 XML 层提取文本，避免段落对象包装的差异
                        _text = "".join(
                            (t.text or "") for t in _p.iter(_q("t"))
                        )
                        # 强化空白判定：去全角空格、零宽字符、NBSP
                        _check = _text
                        for _ch in ("\u3000", "\u200b", "\u200c", "\u200d",
                                    "\ufeff", "\xa0"):
                            _check = _check.replace(_ch, "")
                        if _check.strip():
                            continue   # 有实际内容，保留
                        # 确认段落里没有图片/表格（防止误删）
                        _has_media = False
                        for _tag in ("drawing", "pict", "object", "tbl"):
                            if _p.find(".//" + _q(_tag)) is not None:
                                _has_media = True
                                break
                        if _has_media:
                            continue
                        _tc.remove(_p)
            except Exception as _e:
                warning(f"[docx_engine] insert_table 清理多余空段落失败（忽略）：{_e}")
            # 补边框（不依赖模板样式）
            try:
                _ensure_table_borders(tbl._tbl)
            except Exception as _e:
                warning(f"[docx_engine] insert_table 补边框失败：{_e}")
            # 让 Word 按内容自动分配列宽，总宽固定 width_cm
            self.set_table_autofit(tbl, width_cm)
            return True
        except Exception as e:
            warning(f"[docx_engine] insert_table 失败：{e}")
            return False

    def copy_table(self, doc: Document, section: str,
                   source_doc: Document, source_anchor: str,
                   target_anchor: str) -> bool:
        """跨文档复制表格：复制 source_doc 中 source_anchor 标签后第一个表格，
        替换目标文档中 target_anchor 标签后第一个表格。失败返回 False。

        锚点定位策略（v5.3 修复）：
          仅在 ``section`` 章节范围内搜索锚点；章节或锚点不存在时返回 False，
          不执行任何插入操作。
        """
        try:
            # ── 定位源文档章节：模糊匹配含 "rohs" 的章节（不区分大小写），
            # 兼容 "RoHS Testing" / "RoHS Test" / "RoHS测试" 等各种命名。
            # 若 section 参数本身不含 "rohs"，则退回精确用 section 查找。
            if "rohs" in (section or "").lower():
                src_section_para = None
                for para in source_doc.paragraphs:
                    if _is_in_table(para):
                        continue
                    if _para_is_bold(para) and "rohs" in (para.text or "").lower():
                        src_section_para = para
                        break
                if src_section_para is None:
                    warning(f"[docx_engine] copy_table：源文档中未找到含 'rohs' 的章节")
                    return False
                src_anchor = self.find_paragraph(source_doc, src_section_para, source_anchor)
                if src_anchor is None:
                    warning(f"[docx_engine] copy_table：源文档 RoHS 章节内未找到锚点 {source_anchor!r}")
                    return False
            else:
                src_anchor = self._find_anchor_in_section(source_doc, section, source_anchor)
                if src_anchor is None:
                    return False

            src_tbl = self._first_table_after(source_doc, src_anchor)
            if src_tbl is None:
                warning("[docx_engine] copy_table：源表格未找到")
                return False
            clone = _deepclone_tbl(src_tbl)
            _ensure_table_borders(clone)  # 防止跨文档复制后边框丢失

            # ── 定位目标文档锚点（严格限定在指定章节内）────────────────────
            tgt_anchor_para = self._find_anchor_in_section(doc, section, target_anchor)
            if tgt_anchor_para is None:
                return False

            old = self._first_table_after(doc, tgt_anchor_para)
            if old is not None:
                delete_table(old)
            tgt_anchor_para._p.addnext(clone)
            return True
        except Exception as e:
            warning(f"[docx_engine] copy_table 失败：{e}")
            return False

    def copy_table_by_source_para(self, doc: Document, section: str,
                                  source_doc: Document, source_para_text: str,
                                  target_anchor: str) -> bool:
        """按"源文档加粗段落"定位源表格，直接插入目标文档锚点后。

        与 ``copy_table`` 的区别：
          - 源文档定位：**全局搜索加粗段落**（如 "Sample Test Result"），
            取其后第一个表格（不依赖章节标题）。
          - 目标文档：找到锚点后**直接插入**，不删除已有表格。

        适用：卤素（Halogen）等源报告无章节标题、只有加粗字段锚点的场景。
        失败返回 False。
        """
        try:
            # ── 1. 源文档：全局搜加粗段落 ──
            needle = (source_para_text or "").strip().replace("：", ":").replace(" ", "")
            if not needle:
                warning("[docx_engine] copy_table_by_source_para：锚点文本为空")
                return False
            src_para = None
            for para in source_doc.paragraphs:
                if _is_in_table(para):
                    continue
                if not _para_is_bold(para):
                    continue
                txt = (para.text or "").strip().replace("：", ":").replace(" ", "")
                if needle in txt:
                    src_para = para
                    break
            if src_para is None:
                warning(f"[docx_engine] copy_table_by_source_para："
                        f"源文档未找到加粗段落 {source_para_text!r}")
                return False

            # ── 2. 取其后第一个表格 ──
            src_tbl = self._first_table_after(source_doc, src_para)
            if src_tbl is None:
                warning(f"[docx_engine] copy_table_by_source_para："
                        f"段落 {source_para_text!r} 后未找到表格")
                return False
            clone = _deepclone_tbl(src_tbl)
            _ensure_table_borders(clone)

            # ── 3. 目标文档：章节内找锚点 → 直接插入 ──
            tgt_anchor_para = self._find_anchor_in_section(doc, section, target_anchor)
            if tgt_anchor_para is None:
                warning(f"[docx_engine] copy_table_by_source_para："
                        f"目标章节 {section!r} 内未找到锚点 {target_anchor!r}")
                return False
            tgt_anchor_para._p.addnext(clone)
            return True
        except Exception as e:
            warning(f"[docx_engine] copy_table_by_source_para 失败：{e}")
            return False
    def set_table_format(self, table, width_cm: float) -> None:
        """设置表格总宽为 width_cm，禁用自动调整，列宽按内容视觉宽度分配，
        表格居中，中文按 2 个字符宽计。"""
        try:
            ncols = len(table.columns)
            # ── 1. 计算每列的"视觉宽度"（中文=2，英文=1）──
            col_chars = [1.0] * ncols
            for row in table.rows:
                for j, cell in enumerate(row.cells):
                    if j < ncols:
                        text = (cell.text or "").strip()
                        w = _visual_width(text)
                        if w > col_chars[j]:
                            col_chars[j] = w

            # 每列加 2 个字符的 padding（左右内边距）
            col_chars = [c + 2 for c in col_chars]
            total = sum(col_chars) or 1.0

            total_dxa = _cm_to_dxa(width_cm)
            min_dxa = _cm_to_dxa(1.2)   # 每列最小 1.2cm

            # ── 2. 按比例分配列宽 ──
            col_widths = []
            for c in col_chars:
                w = int(round(c / total * total_dxa))
                w = max(min_dxa, w)
                col_widths.append(w)

            # ── 3. 归一化：把差额补到最长列，保证总宽精确 ──
            s = sum(col_widths)
            if s != total_dxa:
                diff = total_dxa - s
                max_idx = col_chars.index(max(col_chars))
                col_widths[max_idx] = max(min_dxa, col_widths[max_idx] + diff)

            set_table_column_width(table._tbl, col_widths)

            # ── 4. 设置 tblLayout=fixed，让列宽生效 ──
            tbl = table._tbl if hasattr(table, "_tbl") else table
            tblPr = tbl.find(_q("tblPr"))
            if tblPr is None:
                tblPr = _lxml_etree.SubElement(tbl, _q("tblPr"))
                tbl.insert(0, tblPr)
            tblLayout = tblPr.find(_q("tblLayout"))
            if tblLayout is None:
                tblLayout = _lxml_etree.SubElement(tblPr, _q("tblLayout"))
            tblLayout.set(_q("type"), "fixed")

            # ── 5. 表格整体居中 ──
            jc = tblPr.find(_q("jc"))
            if jc is None:
                jc = _lxml_etree.SubElement(tblPr, _q("jc"))
            jc.set(_q("val"), "center")
        except Exception as e:
            warning(f"[docx_engine] set_table_format 失败：{e}")

    def _set_cells_nowrap(self, table) -> None:
        """给所有单元格加 w:noWrap，禁止自动换行。"""
        try:
            tbl = table._tbl if hasattr(table, "_tbl") else table
            for tc in tbl.iter(_q("tc")):
                tcp = tc.find(_q("tcPr"))
                if tcp is None:
                    tcp = _lxml_etree.SubElement(tc, _q("tcPr"))
                    tc.insert(0, tcp)
                nw = tcp.find(_q("noWrap"))
                if nw is None:
                    _lxml_etree.SubElement(tcp, _q("noWrap"))
        except Exception as e:
            warning(f"[docx_engine] _set_cells_nowrap 失败：{e}")

    def set_table_autofit(self, table, width_cm: float) -> None:
        """将表格切换为 autofit 模式，总宽固定为 width_cm，列宽由 Word 自动分配。

        OOXML 实现：
          1. ``w:tblW`` 设为固定总宽（dxa）。
          2. ``w:tblLayout type="autofit"``：启用自动列宽。
          3. 每个 ``w:tc`` 的 ``w:tcW`` 设为 ``type="auto" w="0"``：让 Word 按内容分配。
        """
        try:
            tbl = table._tbl if hasattr(table, "_tbl") else table
            total_dxa = _cm_to_dxa(width_cm)

            # ── 1. 表格属性：总宽 + autofit ─────────────────────────────────
            tblPr = tbl.find(_q("tblPr"))
            if tblPr is None:
                tblPr = _lxml_etree.SubElement(tbl, _q("tblPr"))
                tbl.insert(0, tblPr)

            # w:tblW
            tblW = tblPr.find(_q("tblW"))
            if tblW is None:
                tblW = _lxml_etree.SubElement(tblPr, _q("tblW"))
            tblW.set(_q("w"), str(total_dxa))
            tblW.set(_q("type"), "dxa")

            # w:tblLayout type="autofit"
            tblLayout = tblPr.find(_q("tblLayout"))
            if tblLayout is None:
                tblLayout = _lxml_etree.SubElement(tblPr, _q("tblLayout"))
            tblLayout.set(_q("type"), "autofit")

            # ── 2. 每个单元格：tcW type="auto" ──────────────────────────────
            for tc in tbl.iter(_q("tc")):
                tcp = tc.find(_q("tcPr"))
                if tcp is None:
                    tcp = _lxml_etree.SubElement(tc, _q("tcPr"))
                    tc.insert(0, tcp)
                tcW = tcp.find(_q("tcW"))
                if tcW is None:
                    tcW = _lxml_etree.SubElement(tcp, _q("tcW"))
                tcW.set(_q("w"), "0")
                tcW.set(_q("type"), "auto")
            # 表格整体居中
            jc = tblPr.find(_q("jc"))
            if jc is None:
                jc = _lxml_etree.SubElement(tblPr, _q("jc"))
            jc.set(_q("val"), "center")

        except Exception as e:
            warning(f"[docx_engine] set_table_autofit 失败：{e}")

    def set_table_cell_margins(self, table, top=0, bottom=0, left=0, right=0) -> None:
        """将表格所有单元格的内边距统一设置（单位：dxa，默认全 0）。

        通过向每个 ``w:tc`` 下的 ``w:tcPr/w:tcMar`` 写入四个方向的边距实现。
        """
        try:
            tbl = table._tbl if hasattr(table, "_tbl") else table
            for tc in tbl.iter(_q("tc")):
                tcp = tc.find(_q("tcPr"))
                if tcp is None:
                    tcp = _lxml_etree.SubElement(tc, _q("tcPr"))
                    tc.insert(0, tcp)
                mar = tcp.find(_q("tcMar"))
                if mar is not None:
                    tcp.remove(mar)
                mar = _lxml_etree.SubElement(tcp, _q("tcMar"))
                for side, val in (("top", top), ("bottom", bottom),
                                  ("left", left), ("right", right)):
                    el = _lxml_etree.SubElement(mar, _q(side))
                    el.set(_q("w"), str(val))
                    el.set(_q("type"), "dxa")
        except Exception as e:
            warning(f"[docx_engine] set_table_cell_margins 失败：{e}")

    # ───────────────────────── 表格行 / 单元格操作 ─────────────────────────

    def add_table_row(self, table):
        """在表格末尾追加一行，返回新行（python-docx Row）或 None。"""
        try:
            if len(table.rows) == 0:
                return None
            return table.add_row()
        except Exception as e:
            warning(f"[docx_engine] add_table_row 失败：{e}")
            return None

    def remove_table_row(self, table, index: int) -> bool:
        """删除表格第 index 行（0-based）。返回 True / False。"""
        try:
            return remove_table_row(table._tbl, index)
        except Exception as e:
            warning(f"[docx_engine] remove_table_row 失败：{e}")
            return False

    def table_row_count(self, table) -> int:
        """返回表格行数。"""
        try:
            return len(table.rows)
        except Exception:
            return 0

    def set_cell_text(self, table, row: int, col: int, text: str,
                      color=None, font_name: str = None,
                      font_size: int = None, align: str = None) -> bool:
        """设置表格单元格 (row, col) 的文本；可选 color 着色首 run、font_name/font_size 设字体、align 设对齐。返回 True / False。"""
        try:
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD
            cell = table.cell(row, col)
            cell.text = text
            para = cell.paragraphs[0] if cell.paragraphs else None
            if para is not None:
                runs = para.runs
                if color is not None and runs:
                    set_text_color(runs[0], color)
                if (font_name or font_size) and runs:
                    for run in runs:
                        if font_name:
                            run.font.name = font_name
                        if font_size:
                            run.font.size = Pt(font_size)
                if align is not None:
                    _align_map = {
                        "center": _WD.CENTER,
                        "left": _WD.LEFT,
                        "right": _WD.RIGHT,
                    }
                    para.alignment = _align_map.get(align.lower(), _WD.LEFT)
            return True
        except Exception as e:
            warning(f"[docx_engine] set_cell_text 失败：{e}")
            return False

    def set_cell_text_preserve(self, table, row: int, col: int, text: str,
                               color=None, font_name: str = None,
                               font_size: int = None, align: str = None) -> bool:
        """设置单元格文本但**保留单元格既有 run 格式**（与 set_cell_text 的区别）。

        ``set_cell_text`` 通过 ``cell.text = ...`` 重建段落，会丢失模板原有的
        rPr（字体 / 颜色 / 加粗等）；本方法只改写首个既有 run 的 ``w:t``，
        删除多余 run 与多余 ``w:t``，其余格式原样保留。

        Args:
            table: python-docx Table 对象。
            row / col: 单元格坐标（0-based）。
            text: 目标文本。
            color: 可选，着色首 run。
            font_name / font_size: 可选，强制设置首 run 字体 / 字号（pt）。
            align: 可选，``"center"`` / ``"left"`` / ``"right"``。

        Returns:
            True 成功，False 失败（不抛异常）。
        """
        try:
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD
            cell = table.cell(row, col)
            para = cell.paragraphs[0] if cell.paragraphs else None
            if para is None:
                cell.text = text
                return True

            runs = para.runs
            if runs:
                # 保留首 run 的 rPr，删除其余 run，避免重建段落丢失格式
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
                first = runs[0]
                ts = first._r.findall(_q("t"))
                if ts:
                    for t in ts[1:]:
                        first._r.remove(t)
                    ts[0].text = text
                else:
                    from lxml import etree as _et
                    t = _et.SubElement(first._r, _q("t"))
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = text
            else:
                para.add_run(text)

            if color is not None and para.runs:
                set_text_color(para.runs[0], color)
            if (font_name or font_size) and para.runs:
                for run in para.runs:
                    if font_name:
                        run.font.name = font_name
                    if font_size:
                        run.font.size = Pt(font_size)
            if align is not None:
                _align_map = {
                    "center": _WD.CENTER,
                    "left": _WD.LEFT,
                    "right": _WD.RIGHT,
                }
                para.alignment = _align_map.get(align.lower(), _WD.LEFT)
            return True
        except Exception as e:
            warning(f"[docx_engine] set_cell_text_preserve 失败：{e}")
            return False

    def table_cell_text(self, table, row: int, col: int) -> Optional[str]:
        """返回表格单元格 (row, col) 的文本；异常时返回 None。"""
        try:
            return table.cell(row, col).text
        except Exception:
            return None

    def table_column_count(self, table) -> int:
        """返回表格列数。"""
        try:
            return len(table.columns)
        except Exception:
            return 0

    def set_text_color(self, run, color) -> None:
        """设置 run 字体颜色（color: #RRGGBB / 整数 RGB）。"""
        try:
            set_text_color(run, color)
        except Exception as e:
            warning(f"[docx_engine] set_text_color 失败：{e}")

    # ───────────────────────── 内部辅助 ─────────────────────────

    @staticmethod
    def _apply_cell_style(cell, st: dict) -> None:
        run = None
        if cell.paragraphs and cell.paragraphs[0].runs:
            run = cell.paragraphs[0].runs[0]
        elif cell.paragraphs:
            run = cell.paragraphs[0].add_run("")
        if run is not None:
            if st.get("color"):
                set_text_color(run, st["color"])
            if st.get("bold"):
                run.font.bold = True
            if st.get("font_name"):
                run.font.name = st["font_name"]
            if st.get("font_size"):
                run.font.size = Pt(float(st["font_size"]))
            # 对齐
            align = st.get("align")
            if align == "center":
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == "right":
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if st.get("bg"):
            set_cell_background(cell, st["bg"])


def _visual_width(text: str) -> int:
    """计算字符串的视觉宽度：中文/全角=2，其他=1。"""
    if not text:
        return 0
    w = 0
    for ch in text:
        w += 2 if ord(ch) > 0x2E7F else 1
    return w