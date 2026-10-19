"""文本与结构 Mixin（v5.0 · 由 docx_engine.py 拆分而来）。

职责：``DocxEngine`` 的**文本写入**与**文档结构**能力——段落文本替换 / 着色 /
run 级精准替换、分页符插入、章节删除与克隆、页眉填充、相邻单元格填充、
过期日期标红。

以 Mixin 形式被 ``DocxEngine`` 继承；通过 ``self`` 调用 ``NavigationMixin``
（``find_section`` / ``find_section_range`` / ``invalidate_section_cache``）
与 ``TableOpsMixin``（``set_cell_text``）。

依赖：``services.docx_helpers``、``services.ooxml_service``、``lxml``。
"""
from copy import deepcopy as _deepcopy
from typing import Optional

from lxml import etree as _lxml_etree

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from common.logger import warning

from services.docx_helpers import (
    _color_expired_in_p,
    _has_page_break_before,
    _p_is_empty,
    _para_is_bold,
)
from services.ooxml_service import (
    _q,
    append_elements,
    collect_range_elements,
    insert_elements_before,
    set_paragraph_text,
    set_text_color,
)


class StructureOpsMixin:
    """文本与结构 Mixin。"""

    # ───────────────────────── 插入 ─────────────────────────

    def insert_text(self, doc: Document, anchor: object, text: str,
                    font_color=None,
                    font_name: str = None,
                    font_size: int = None) -> bool:
        """在锚点段落的**同行冒号后**替换文本（替换模式，非追加）。

        自动适配全角（：）/半角（:）冒号：保留冒号前前缀，覆盖冒号后内容。
        font_name / font_size：可选，若提供则对所有新建 run 设置字体。
        """
        def _apply_font(run):
            if font_name is not None:
                run.font.name = font_name
            if font_size is not None:
                run.font.size = Pt(font_size)

        try:
            para = anchor
            full = para.text or ""
            has_sep = ":" in full or "：" in full
            if has_sep:
                sep = ":" if ":" in full else "："
                prefix, _ = full.split(sep, 1)
            else:
                prefix, sep = "", ""

            if font_color is None:
                # ── 原有行为：整段写入一个 run ──────────────────────────────
                new_text = f"{prefix}{sep} {text}" if has_sep else text
                runs = para.runs
                if runs:
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)
                    first = runs[0]
                    ts = first._r.findall(_q("t"))
                    if ts:
                        for t in ts[1:]:
                            first._r.remove(t)
                        ts[0].text = new_text
                    else:
                        t = _lxml_etree.SubElement(first._r, _q("t"))
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t.text = new_text
                    _apply_font(first)
                else:
                    _apply_font(para.add_run(new_text))
            else:
                # ── 着色模式：标签 run（黑色）+ 值 run（着色）───────────────
                # 清空所有现有 run
                for r in list(para.runs):
                    r._r.getparent().remove(r._r)
                # run1：标签 + 分隔符（黑色）
                label_text = f"{prefix}{sep} " if has_sep else ""
                if label_text:
                    run1 = para.add_run(label_text)
                    # 确保 run1 无颜色（黑色默认）
                    from docx.oxml.ns import qn as _qn
                    rpr1 = run1._r.get_or_add_rPr()
                    color_el = rpr1.find(_qn("w:color"))
                    if color_el is not None:
                        rpr1.remove(color_el)
                    _apply_font(run1)
                # run2：值（着色）
                run2 = para.add_run(text)
                set_text_color(run2, font_color)
                _apply_font(run2)
            return True
        except Exception as e:
            warning(f"[docx_engine] insert_text 失败：{e}")
            return False

    def replace_paragraph_text(self, para: object, text: str) -> bool:
        """整段替换段落文本（清空原有 run，首 run 写入新文本）。返回 True / False。"""
        try:
            runs = para.runs
            if runs:
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
                runs[0].text = text
            else:
                para.add_run(text)
            return True
        except Exception as e:
            warning(f"[docx_engine] replace_paragraph_text 失败：{e}")
            return False

    def set_paragraph_color(self, para: object, color) -> bool:
        """设置段落所有 run 的字体颜色（color: #RRGGBB / 整数 RGB）。"""
        try:
            for r in para.runs:
                set_text_color(r, color)
            return True
        except Exception as e:
            warning(f"[docx_engine] set_paragraph_color 失败：{e}")
            return False

    def set_paragraph_runs(self, para: object, runs: list) -> bool:
        """整段重建为给定的 run 列表（完全可控格式，用于统计段落等分段着色的场景）。

        Args:
            runs: ``[(text, color_or_None), ...]``，``color`` 为整数 RGB（如 0xFF0000）
                  或 ``None``（不显式着色，沿用默认）。首个 run 保留段首 rPr 风格。
        返回 True / False。
        """
        try:
            from docx.oxml import OxmlElement
            p_el = para._p
            old_runs = list(para.runs)
            for r in old_runs:
                p_el.remove(r._r)
            for i, (text, color) in enumerate(runs):
                r_el = OxmlElement("w:r")
                if i == 0 and old_runs:
                    rPr = old_runs[0]._r.find(_q("rPr"))
                    if rPr is not None:
                        r_el.append(_deepcopy(rPr))
                if color is not None:
                    rPr = r_el.find(_q("rPr"))
                    if rPr is None:
                        rPr = OxmlElement("w:rPr")
                        r_el.append(rPr)
                    color_el = OxmlElement("w:color")
                    color_el.set(_q("val"), f"{color:06X}")
                    rPr.append(color_el)
                t = OxmlElement("w:t")
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = text
                r_el.append(t)
                p_el.append(r_el)
            return True
        except Exception as e:
            warning(f"[docx_engine] set_paragraph_runs 失败：{e}")
            return False

    def replace_in_paragraph(self, para: object, pattern, repl,
                            count: int = 0, *, flags: int = 0) -> int:
        """段落内精准文本替换，**保留各 run 的字体/颜色格式**（按 run 局部替换）。

        用于统计段落等需保留「判定词红色 run」的场景：仅替换命中区间所在的 run
        文本（命中跨多 run 时，按区间拆分各 run 文本），不重建整段，从而保留颜色。

        Args:
            pattern: 字符串（子串匹配）或编译过的 ``re`` 正则对象。
            repl: 替换字符串，或 ``(re.Match) -> str`` 的回调。
            count: 最多替换次数（0 = 全部）。
            flags: 正则标志（仅在 pattern 为正则时生效）。

        Returns:
            实际替换次数；异常返回 0。
        """
        try:
            import re as _re
            runs = para.runs
            if not runs:
                return 0

            # 归一为正则对象
            if isinstance(pattern, str):
                pat = _re.compile(pattern, flags)
                is_regex = True
            elif hasattr(pattern, "sub"):
                pat = pattern
                is_regex = True
            else:
                pat = None
                is_regex = False

            total = 0
            for r in list(runs):
                text = r.text or ""
                if not text:
                    continue

                # 收集命中区间 (start, end)
                if is_regex:
                    matches = list(pat.finditer(text))
                else:
                    ms = []
                    start = 0
                    while True:
                        idx = text.find(pattern, start)
                        if idx < 0:
                            break
                        ms.append((idx, idx + len(pattern)))
                        start = idx + len(pattern)
                        if count and len(ms) >= count:
                            break
                    matches = ms

                if not matches:
                    continue

                # 计算替换后文本（从后往前替换，避免位移）
                if is_regex and callable(repl):
                    def _do(m):
                        return repl(m)
                elif is_regex and not callable(repl):
                    def _do(m):
                        return pat.sub(repl, m.group())
                else:  # 子串 + 字符串替换
                    def _do(_m):
                        return repl

                new_text = text
                replaced = 0
                for m in sorted(matches, reverse=True):
                    s, e = (m.start(), m.end()) if is_regex else m
                    new_text = new_text[:s] + _do(m) + new_text[e:]
                    replaced += 1
                    if count and replaced >= count:
                        break

                # 写回该 run（保留 rPr，即颜色/加粗等格式）
                ts = r._r.findall(_q("t"))
                if ts:
                    for t in ts[1:]:
                        r._r.remove(t)
                    ts[0].text = new_text
                else:
                    from lxml import etree as _et
                    t = _et.SubElement(r._r, _q("t"))
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = new_text
                total += replaced
            return total
        except Exception as e:
            warning(f"[docx_engine] replace_in_paragraph 失败：{e}")
            return 0

    def add_page_break(self, doc: Document, para: object) -> None:
        title = (para.text or "").strip() if hasattr(para, "text") else ""
        already = _has_page_break_before(para)
        warning(f"[docx_engine] add_page_break：章节={title!r}  已有分页符={already}")
        if already:
            warning(f"[docx_engine] add_page_break：跳过（已有分页符）")
            # 清理标题前的纯空白段落，防止它在分页符与标题之间占用页面产生空白页
            try:
                p_el = para._p if hasattr(para, "_p") else para
                prev = p_el.getprevious()
                while prev is not None and prev.tag == _q("p") and _p_is_empty(prev):
                    to_del = prev
                    prev = prev.getprevious()
                    to_del.getparent().remove(to_del)
                    warning(f"[docx_engine] add_page_break：删除标题前空白段落（防空白页）")
            except Exception as _e:
                warning(f"[docx_engine] add_page_break：清理空白段落失败（忽略）：{_e}")
            return
        try:
            from docx.oxml import OxmlElement as _OE
            p_el = para._p if hasattr(para, "_p") else para
            new_p = _OE("w:p")
            r_el = _OE("w:r")
            br_el = _OE("w:br")
            br_el.set(qn("w:type"), "page")
            r_el.append(br_el)
            new_p.append(r_el)
            # 插入到标题段落之前，而不是之后
            p_el.addprevious(new_p)
        except Exception as e:
            warning(f"[docx_engine] add_page_break 失败：{e}")

    def delete_section(self, doc: Document, section_name: str) -> bool:
        """删除整个章节（含标题与全部内容），直到下一个章节标题或文档末尾。

        同时删除"本章节标题前面紧邻的前导空段落 / 分页符段落"，
        但**保留分节符段落**（它是页眉/页码的载体，删了会丢页眉）。
        """
        try:
            start, end = self.find_section_range(doc, section_name)
            if start is None:
                return False
            body = doc.element.body
            elems = collect_range_elements(
                body, start._p, end._p if end is not None else None)

            # ── 额外收集：标题前面紧邻的前导空段落 / 分页符段落 ──
            extra = []
            prev = start._p.getprevious()
            count = 0
            while prev is not None and count < 5:
                tag = prev.tag.split("}")[-1] if "}" in prev.tag else prev.tag
                if tag in ("bookmarkStart", "bookmarkEnd", "proofErr"):
                    prev = prev.getprevious()
                    count += 1
                    continue
                if tag != "p":
                    break
                # 遇到分节符段落 → 保留，停止
                _has_sectpr = False
                for ppr in prev.findall(_q("pPr")):
                    if ppr.find(_q("sectPr")) is not None:
                        _has_sectpr = True
                        break
                if _has_sectpr:
                    break
                # 有文本 / 有图片 / 有表格 → 停止
                _has_content = False
                for t in prev.findall(".//" + _q("t")):
                    if (t.text or "").strip():
                        _has_content = True
                        break
                if not _has_content:
                    for tag2 in ("drawing", "pict", "object", "tbl"):
                        if prev.find(".//" + _q(tag2)) is not None:
                            _has_content = True
                            break
                if _has_content:
                    break
                extra.append(prev)
                prev = prev.getprevious()
                count += 1

            for e in elems + extra:
                parent = e.getparent()
                if parent is not None:
                    parent.remove(e)
            self.invalidate_section_cache(doc)
            return True
        except Exception as e:
            warning(f"[docx_engine] delete_section 失败：{e}")
            return False


    def delete_section_content(self, doc: Document, section_name: str) -> bool:
        """删除章节内所有内容（保留标题），直到下一个加粗标题或文档末尾。"""
        try:
            section = self.find_section(doc, section_name)
            if section is None:
                return False
            # 收集待删除节点（section 之后、下一个加粗标题之前的所有 w:p 与 w:tbl）
            body = doc.element.body
            started = False
            to_remove = []
            for child in body.iterchildren():
                tag = child.tag.split('}')[-1]
                if child is section._p:
                    started = True
                    continue
                if not started:
                    continue
                if tag == "p":
                    # 判断是否为下一个加粗标题
                    from docx.text.paragraph import Paragraph
                    p = Paragraph(child, doc)
                    if _para_is_bold(p) and p.text.strip():
                        break
                    to_remove.append(child)
                elif tag == "tbl":
                    to_remove.append(child)
            for node in to_remove:
                body.remove(node)
            return True
        except Exception as e:
            warning(f"[docx_engine] delete_section_content 失败：{e}")
            return False

    def clone_section(self, doc: Document, template_section_name: str,
                      new_title: str, insert_before_section_name: Optional[str] = None,
                      lang: str = "en") -> bool:
        """克隆模板章节结构、重命名标题、插入到指定章节之前（或文末）。

        ``insert_before_section_name`` 决定插入位置（``None`` = 文末）。
        业务层负责解析「插入到 RoHS 前」等规则后传入具体章节名。

        实现说明（v5.3 修复）：
          原 collect_range_elements 仅收集 w:p / w:tbl，会漏掉 w:sdt、
          w:bookmarkStart / w:bookmarkEnd 等直接子元素，导致克隆章节缺少
          锚点文本（Test Standards:、Result description、设备表等）。
          现改为直接迭代 body.iterchildren()，收集范围内**所有**直接子元素
          （仅跳过 w:sectPr），并对每个节点逐一 deepcopy，确保完整复制。
        """
        from copy import deepcopy as _deepcopy
        try:
            start, end = self.find_section_range(doc, template_section_name)
            if start is None:
                warning(f"[docx_engine] clone_section：模板章节不存在 {template_section_name}")
                return False

            body = doc.element.body
            start_el = start._p
            end_el = end._p if end is not None else None

            # 收集范围内所有直接子元素（含 sdt/bookmarkStart/bookmarkEnd 等），
            # 跳过 w:sectPr（节属性，不应被克隆到新章节）。
            clones = []
            title_clone = None   # 标题段落的克隆节点
            collecting = False
            for child in body.iterchildren():
                if child is start_el:
                    collecting = True
                    cloned = _deepcopy(child)
                    title_clone = cloned   # 记录标题副本，与列表索引无关
                    clones.append(cloned)
                    continue
                if not collecting:
                    continue
                if end_el is not None and child is end_el:
                    break
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if tag == "sectPr":
                    continue
                clones.append(_deepcopy(child))

            if not clones:
                warning(f"[docx_engine] clone_section：模板章节 {template_section_name} 范围为空")
                return False

            # 将标题克隆节点的文本替换为新标题。
            # find_section_range 对中英文模板均为倒序收集，统一反转为正序。
            clones = clones[::-1]
            set_paragraph_text(title_clone, new_title)
            # 强制标题加粗：确保标题克隆里所有 run 都有 w:b
            # （set_paragraph_text 只保留首 run 的 rPr，若首 run 恰好未加粗则丢加粗）
            try:
                for _r in title_clone.findall(_q("r")):
                    _rpr = _r.find(_q("rPr"))
                    if _rpr is None:
                        _rpr = _lxml_etree.SubElement(_r, _q("rPr"))
                        _r.insert(0, _rpr)
                    if _rpr.find(_q("b")) is None:
                        _lxml_etree.SubElement(_rpr, _q("b"))
            except Exception as _e:
                warning(f"[clone_section] 强制标题加粗失败（忽略）：{_e}")

            # 确定插入位置
            ref = None
            if insert_before_section_name:
                ref_para = self.find_section(doc, insert_before_section_name)
                if ref_para is not None:
                    ref = ref_para._p

            if ref is not None:
                insert_elements_before(ref, clones)
            else:
                append_elements(body, clones)
            self.invalidate_section_cache(doc)
            return True
        except Exception as e:
            warning(f"[docx_engine] clone_section 失败：{e}")
            return False

    # ───────────────────────── 文档遍历 ─────────────────────────

    def color_expired_in_paragraph(self, para, today=None) -> None:
        """扫描段落中的日期，对过期日期（区间只标截止日期）标红。"""
        try:
            _color_expired_in_p(para._p if hasattr(para, "_p") else para, today)
        except Exception as e:
            warning(f"[docx_engine] color_expired_in_paragraph 失败：{e}")

    def color_expired_in_table(self, table, today=None) -> None:
        """扫描表格所有单元格中的日期并标红过期项。"""
        try:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _color_expired_in_p(p._p, today)
        except Exception as e:
            warning(f"[docx_engine] color_expired_in_table 失败：{e}")

    def color_expired_in_cell(self, table, row: int, col: int, today=None) -> None:
        """扫描指定单元格中的日期并标红过期项。"""
        try:
            cell = table.cell(row, col)
            for p in cell.paragraphs:
                _color_expired_in_p(p._p, today)
        except Exception as e:
            warning(f"[docx_engine] color_expired_in_cell 失败：{e}")

    # ───────────────────────── 页眉填充 ─────────────────────────

    def insert_text_in_header(self, doc: Document, label: str, value: str) -> bool:
        """在页眉（Primary）中搜索 label，保留页码域，在其后填入 value。"""
        warning(f"[insert_text_in_header] DEBUG: 被调用，label={label!r}, value={value!r}")
        try:
            for section in doc.sections:
                header = section.header
                for para in header.paragraphs:
                    if label in (para.text or ""):
                        self._set_label_value_preserve_fields(para, label, value)
                        return True
        except Exception as e:
            warning(f"[docx_engine] insert_text_in_header 失败：{e}")
        warning(f"[docx_engine] insert_text_in_header 未找到标签：{label}")
        return False

    def set_adjacent_cell_text(self, doc: Document, para, value: str,
                               font_name: str = None, font_size: int = None,
                               align: str = None) -> bool:
        """定位 para 所在的表格单元格，将同行右侧相邻单元格文本设为 value。

        Args:
            font_name: 可选，强制设置相邻单元格所有 run 的字体名称。
            font_size: 可选，强制设置相邻单元格所有 run 的字号（pt）。
        """
        try:
            from docx.table import Table as _Table
            from docx.shared import Pt
            p = para._p if hasattr(para, "_p") else para
            parent = p.getparent()
            while parent is not None and parent.tag != _q("tc"):
                parent = parent.getparent()
            if parent is None:
                return False
            tbl_el = parent.getparent().getparent()
            tbl = _Table(tbl_el, doc)
            for r_i, row in enumerate(tbl.rows):
                for c_i, cell in enumerate(row.cells):
                    if cell._tc is parent and c_i + 1 < len(row.cells):
                        ok = self.set_cell_text(tbl, r_i, c_i + 1, value, align=align)
                        if ok and (font_name or font_size):
                            val_cell = tbl.cell(r_i, c_i + 1)
                            for run in (val_cell.paragraphs[0].runs
                                        if val_cell.paragraphs else []):
                                if font_name:
                                    run.font.name = font_name
                                if font_size:
                                    run.font.size = Pt(font_size)
                        return ok
        except Exception as e:
            warning(f"[docx_engine] set_adjacent_cell_text 失败：{e}")
        return False
