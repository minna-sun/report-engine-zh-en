"""python-docx 文档引擎实现（v5.0 新增 · spec 模块 11）。

职责：``DocumentEngine`` 抽象接口的具体实现，基于 **python-docx** + **ooxml_service**。
本模块是系统中**唯一真正操作 Word 文档底层 XML（OOXML）的技术实现者**。

特点：
  - 不依赖 win32com、不依赖 Word 应用、跨平台（Windows / macOS / Linux）。
  - 无 COM 进程残留。
  - 所有方法失败返回 ``False`` / ``None``，不抛异常（编程错误除外）。
  - 不含业务逻辑、不含数据查询。

结构（v5.x 拆分后）：
  本文件只保留**生命周期**（open / save / close / export_pdf）与文档遍历等骨架方法，
  其余能力按职责拆分为四个 Mixin，通过多重继承组合进 ``DocxEngine``：
    - ``docx_navigation.NavigationMixin``     导航 / 定位（章节、锚点、占位符、表格定位）
    - ``docx_image_ops.ImageOpsMixin``        图片插入
    - ``docx_table_ops.TableOpsMixin``        表格操作（插入 / 复制 / 格式 / 单元格读写）
    - ``docx_structure_ops.StructureOpsMixin`` 文本与结构（写入 / 分页 / 章节增删克隆 / 标红）
  拆分只做代码搬运，未修改任何实现，外部调用方（template_engine、report_engine 等）无感知。

依赖：``python-docx``（核心）、``services.ooxml_service``（高级 OOXML 操作）、
``lxml``（底层节点）、``zipfile``（跨文档复制）、``services.docx_helpers``（底层纯函数）。
"""
from copy import deepcopy as _deepcopy
from io import BytesIO
from typing import List, Optional, Union
from lxml import etree as _lxml_etree

from docx import Document

from common.logger import info, warning

from services.ooxml_service import _q

from services.docx_helpers import _disable_image_compression
from services.docx_navigation import NavigationMixin
from services.docx_image_ops import ImageOpsMixin
from services.docx_table_ops import TableOpsMixin
from services.docx_structure_ops import StructureOpsMixin


class DocxEngine(NavigationMixin, ImageOpsMixin, TableOpsMixin, StructureOpsMixin):
    """python-docx 文档引擎（继承四个职责 Mixin 组合而成）。"""

    def __init__(self):
        # 章节定位缓存：key = (doc_id, section_name, require_bold)，value = 段落句柄或 None
        # 文档结构变更（增删章节、克隆等）后需调用 invalidate_section_cache 清除。
        super().__init__()
        self._section_cache = {}

    # ───────────────────────── 生命周期 ─────────────────────────

    def open(self, source: Union[str, "object"]) -> Optional[Document]:
        try:
            doc = Document(source)
            # 禁用自动压缩（保持默认）
            _disable_image_compression(doc)
            return doc
        except Exception as e:
            warning(f"[docx_engine] open 失败：{e}")
            return None

    def save(self, doc: Document, path: str) -> bool:
        try:
            doc.save(path)
            return True
        except Exception as e:
            warning(f"[docx_engine] save 失败：{path} -> {e}")
            return False

    def close(self, doc: Document) -> None:
        # python-docx 无连接概念，空操作（保持接口一致性）。
        return None


        # python-docx 无连接概念，空操作（保持接口一致性）。
        return None

    def export_pdf(self, doc_path: str, pdf_path: str) -> bool:
        """导出 PDF：优先 docx2pdf（Word COM，保真），缺失时回退 win32com.SaveAs2。

        两者均为 **Windows-only** 且需要 Microsoft Word。返回 False 时仅记录警告，
        不阻塞 DOCX 交付（调用方应把 PDF 视为可选产物）。
        """
        try:
            import os
            abs_doc = os.path.abspath(doc_path)
            abs_pdf = os.path.abspath(pdf_path)
            # 方案1：docx2pdf（需 Word）
            try:
                from docx2pdf import convert
                convert(abs_doc, abs_pdf)
                if os.path.exists(abs_pdf):
                    info(f"[docx_engine] PDF 导出完成: {abs_pdf}")
                    return True
            except Exception:
                pass
            # 方案2：win32com
            try:
                import win32com.client  # type: ignore
                import pythoncom  # type: ignore
                pythoncom.CoInitialize()
                word = win32com.client.DispatchEx("Word.Application")
                try:
                    word.Visible = False
                    wdFormatPDF = 17
                    doc_cm = word.Documents.Open(abs_doc, ReadOnly=True)
                    doc_cm.SaveAs2(abs_pdf, FileFormat=wdFormatPDF)
                    doc_cm.Close(False)
                    info(f"[docx_engine] PDF 导出完成: {abs_pdf}")
                    return True
                finally:
                    word.Quit()
                    pythoncom.CoUninitialize()
            except Exception:
                warning("[docx_engine] export_pdf：win32com 不可用，跳过 PDF 导出")
        except Exception as e:
            warning(f"[docx_engine] export_pdf 失败：{e}")
        return False

    # ───────────────────────── 文档遍历 ─────────────────────────

    def iter_paragraphs(self, doc: Document):
        """返回文档所有正文段落句柄列表。"""
        try:
            return list(doc.paragraphs)
        except Exception as e:
            warning(f"[docx_engine] iter_paragraphs 失败：{e}")
            return []

    def iter_tables(self, doc: Document):
        """返回文档所有表格句柄列表。"""
        try:
            return list(doc.tables)
        except Exception as e:
            warning(f"[docx_engine] iter_tables 失败：{e}")
            return []

    # ───────────────────────── 页眉填充辅助 ─────────────────────────

    @staticmethod
    def _set_run_text(p_el, text):
        """清空 p_el 中所有 w:r 的节点内容，只保留首个 w:r 的 rPr，写入 text。"""
        runs = p_el.findall(_q("r"))
        if not runs:
            return
        first = runs[0]
        for _r in runs[1:]:
            p_el.remove(_r)
        for _t in first.findall(_q("t")):
            first.remove(_t)
        _t = _lxml_etree.SubElement(first, _q("t"))
        _t.text = text
        _t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    @staticmethod
    def _set_label_value_preserve_fields(p_el_or_para, label, value):
        """在页眉段落中，将 label 之后的旧占位符替换为 value。

        跨 run 拼接定位：label 可能被拆散在多个 run（如 'Report No' + '.: '），
        因此先拼接所有 run 的 w:t 文本做整体正则匹配，再把结果写回——
        label 起始 run 承载新文本（保留其前缀），匹配段内的其余 run 清空，
        匹配段之后的所有 run（含页码域）原样保留，从而不会丢失页码。
        """
        import re
        p_el = (p_el_or_para._p if hasattr(p_el_or_para, "_p")
                else p_el_or_para)
        runs = p_el.findall(_q("r"))
        if not runs:
            return

        # 1. 收集每个 run 的 w:t 文本，并记录字符区间
        run_wts = []
        run_texts = []
        offsets = []
        pos = 0
        for r in runs:
            ts = r.findall(_q("t"))
            run_wts.append(ts)
            txt = "".join((t.text or "") for t in ts)
            run_texts.append(txt)
            offsets.append(pos)
            pos += len(txt)
        joined = "".join(run_texts)

        # 2. 在拼接串上定位 "label + 占位符"，整体替换
        pattern = re.compile(rf"({re.escape(label)})\s*\S+")
        m = pattern.search(joined)
        if not m:
            # 兜底：所有 run 都找不到 label，在末尾追加新 run
            last_run = runs[-1]
            new_r = _lxml_etree.Element(_q("r"))
            rpr = last_run.find(_q("rPr"))
            if rpr is not None:
                new_r.append(_deepcopy(rpr))
            new_t = _lxml_etree.SubElement(new_r, _q("t"))
            new_t.text = f" {value}"
            new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            last_run.addnext(new_r)
            return

        repl = f"{m.group(1)} {value}"

        # 3. 跨 run 写回
        #    找到 label 起始所在的 run
        first_idx = 0
        for i, off in enumerate(offsets):
            if off <= m.start() < off + len(run_texts[i]):
                first_idx = i
                break
        else:
            for i in range(len(offsets) - 1, -1, -1):
                if offsets[i] <= m.start():
                    first_idx = i
                    break

        new_run_texts = list(run_texts)
        for i in range(len(runs)):
            s, e = offsets[i], offsets[i] + len(run_texts[i])
            if i == first_idx:
                local_start = m.start() - offsets[i]
                local_end = max(local_start, m.end() - offsets[i])
                new_run_texts[i] = (run_texts[i][:local_start]
                                    + repl
                                    + run_texts[i][local_end:])
            elif s < m.end() and e > m.start():
                # 被匹配段吸收的非首 run：清空
                new_run_texts[i] = ""
            # 其余 run（匹配段之后，含页码域）保持不变

        # 4. 写回（无 w:t 的 run 跳过，保留页码域）
        for i, r in enumerate(runs):
            ts = run_wts[i]
            if not ts:
                continue
            if new_run_texts[i] == run_texts[i]:
                continue
            for j, t in enumerate(ts):
                t.text = new_run_texts[i] if j == 0 else ""
