"""python-docx 引擎的底层工具函数（v5.0 · 由 docx_engine.py 拆分而来）。

职责：承载 ``DocxEngine`` 及其各 Mixin 共用的**纯函数**（不依赖 ``self``）：
单位换算、段落判定、分页符检测、OOXML 元素构造、日期过期标红等。

本模块**不依赖** ``DocxEngine``，可被 ``docx_navigation`` / ``docx_image_ops`` /
``docx_table_ops`` / ``docx_structure_ops`` / ``docx_engine`` 安全导入。

依赖：``python-docx``（核心）、``services.ooxml_service``（``_q``）、
``lxml``（底层节点）。
"""
import re as _re
import datetime as _datetime
from copy import deepcopy as _deepcopy

from lxml import etree as _lxml_etree

from docx.oxml.ns import qn

from common.logger import warning

from services.ooxml_service import _q

# 1 cm = 1440 / 2.54 dxa ≈ 566.93
_CM_TO_DXA = 1440.0 / 2.54

# 章节同步黑名单：封面 / 汇总 / 声明 / 样品描述等「非测试章节」标题，
# sync_chapters 不触碰（保留且不得当作多余章节删除）。
# 注意：真实模板标题常带尾部标点（如 "Summary of Results:"），故采用
# **前缀匹配**而非精确相等，否则会因标点差异被误判为测试章节而遭误删。
# 双语支持：中英文模板的固定页标题同时收录（中文模板的「测试报告/样品描述/
# 结果汇总」等同样不得被 sync_chapters 误删）；这是识别用保护清单，
# 非「按语言动态切换的锚点」，与 language_pack 职责不同。
_BLACKLIST_TITLES = {
    "Test Report", "Statement", "Sample Description", "Sample Picture",
    "Summary of Results",  "Conclusion", "The End",
    "测试报告", "检测报告", "声明", "声  明", "样品描述", "样品图片",
    "结果汇总", "结论", "结束",
}


def _cm_to_dxa(cm: float) -> int:
    return int(round(cm * _CM_TO_DXA))


def _normalize_colon(text: str) -> str:
    """将全角冒号「：」统一替换为半角冒号「:」，用于锚点匹配时消除全半角差异。"""
    return (text or "").replace("：", ":")


def _ensure_table_borders(tbl_el) -> None:
    """确保表格及所有单元格均有完整边框（细实线，黑色，4 dxa）。

    copy_table 跨文档 deepcopy 后，tblBorders / tcBorders 有时丢失；
    本函数检测缺失时补全，避免粘贴后表格变为无框线样式。
    异常时仅记录警告，不阻塞调用方。
    """
    try:
        from docx.oxml import OxmlElement as _OE
        from docx.oxml.ns import qn as _qn

        _WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        def _wqn(tag: str) -> str:
            # 将 "w:xxx" 转为 Clark 格式 "{ns}xxx"，与 qn() 等价但不依赖注册表
            local = tag.split(":", 1)[1] if ":" in tag else tag
            return f"{{{_WNS}}}{local}"

        _SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
        _CELL_SIDES = ("top", "left", "bottom", "right")

        def _make_border(side: str):
            el = _OE(f"w:{side}")
            el.set(_wqn("w:val"), "single")
            el.set(_wqn("w:sz"), "4")
            el.set(_wqn("w:space"), "0")
            el.set(_wqn("w:color"), "000000")
            return el

        # ── 表格整体边框 ─────────────────────────────────────────────────────
        tblPr = tbl_el.find(_wqn("w:tblPr"))
        if tblPr is None:
            tblPr = _OE("w:tblPr")
            tbl_el.insert(0, tblPr)
        tblBorders = tblPr.find(_wqn("w:tblBorders"))
        if tblBorders is None:
            tblBorders = _OE("w:tblBorders")
            tblPr.append(tblBorders)
            for side in _SIDES:
                tblBorders.append(_make_border(side))

        # ── 各单元格边框 ──────────────────────────────────────────────────────
        for tc in tbl_el.iter(_wqn("w:tc")):
            tcPr = tc.find(_wqn("w:tcPr"))
            if tcPr is None:
                tcPr = _OE("w:tcPr")
                tc.insert(0, tcPr)
            if tcPr.find(_wqn("w:tcBorders")) is None:
                tcBorders = _OE("w:tcBorders")
                tcPr.append(tcBorders)
                for side in _CELL_SIDES:
                    tcBorders.append(_make_border(side))
    except Exception as _e:
        warning(f"[docx_engine] _ensure_table_borders 失败（忽略，不阻塞复制）：{_e}")


def _is_chapter_title_para(para) -> bool:
    """判定段落是否为测试章节标题：加粗 + 非表格内 + 文本非空 + 非黑名单。"""
    if _is_in_table(para):
        return False
    if not _para_is_bold(para):
        return False
    txt = (para.text or "").strip()
    if not txt:
        return False
    warning(f"[_is_chapter_title_para] 检查段落：{txt!r}")
    # 去空白后前缀匹配：兼容模板标题中的空格变体（如"声  明"与"声明"等价），
    # 同时保留原有的尾部标点兼容性（如"Summary of Results:"）。
    # 先去掉首尾的装饰性字符（横线、破折号、波浪线、点号、空格等），
    # 再做前缀匹配，兼容模板中的 "------The End------" 之类写法。
    _DECOR = r"^[\s\-—–—~·.、，。:：]+|[\s\-—–—~·.、，。:：]+$"
    txt_clean = _re.sub(_DECOR, "", txt)
    txt_nospace = _re.sub(r"\s+", "", txt_clean)
    is_blacklisted = any(txt_nospace.startswith(_re.sub(r"\s+", "", bad))
                         for bad in _BLACKLIST_TITLES)
    if is_blacklisted:
        warning(f"[_is_chapter_title_para] 黑名单命中：{txt!r}，返回 False")
    else:
        warning(f"[_is_chapter_title_para] 未命中黑名单：{txt!r}，返回 True")
    return not is_blacklisted


def _is_blacklist_title(para) -> bool:
    """判断段落是否为黑名单标题（The End / 声明 / 样品描述 等固定页）。
    与 _is_chapter_title_para 的区别：本函数只看文本是否命中黑名单，
    不要求加粗、非表格内等条件，用于识别"章节边界"。
    """
    txt = (para.text or "").strip()
    if not txt:
        return False
    _DECOR = r"^[\s\-—–—~·.、，。:：]+|[\s\-—–—~·.、，。:：]+$"
    txt_clean = _re.sub(_DECOR, "", txt)
    txt_nospace = _re.sub(r"\s+", "", txt_clean)
    return any(txt_nospace.startswith(_re.sub(r"\s+", "", bad))
               for bad in _BLACKLIST_TITLES)


def _p_has_page_br(p_el) -> bool:
    """判断 w:p 元素自身是否含有 w:br type="page" run（不含样式属性检测）。"""
    for r in p_el.findall(_q("r")):
        br = r.find(_q("br"))
        if br is not None and br.get(_q("type")) == "page":
            return True
    return False


def _p_is_empty(p_el) -> bool:
    """判断 w:p 是否为无实际文本且不含分页符的空段落（视觉留白用空行）。"""
    text = "".join(
        (n.text or "")
        for r in p_el.findall(_q("r"))
        for n in r.findall(_q("t"))
    ).strip()
    return not text and not _p_has_page_br(p_el)


def _has_page_break_before(para) -> bool:
    """判断段落前是否已存在显式分页符（w:br type="page"）。

    注意：w:pageBreakBefore 是 Word 段落样式属性（渲染时自动换页），
    它不等于一个真实的 w:br page break 节点，无法阻止标题前的空白段落
    被一起推到新页上而产生空白页，因此不在此处检测。

    检查顺序：
    1. 段落自身 run 内含 w:br type="page"（add_page_break 写入的位置）
    2. 向前最多扫 3 个兄弟节点（跳过最多 2 个非 w:p 节点），在遇到的
       第一个 w:p 兄弟上检查分页符；找到第一个 w:p 即决策，不继续向上。
    """
    try:
        p = para._p if hasattr(para, "_p") else para
        title_txt = "".join(
            (n.text or "")
            for r in p.findall(_q("r"))
            for n in r.findall(_q("t"))
        ).strip()[:30]

        # 1. 段落自身 run 内的显式 page break
        if _p_has_page_br(p):
            warning(f"[_has_page_break_before] 步骤1命中（自身含 w:br page）：{title_txt!r}")
            return True

        # 2. 向前最多扫 3 个兄弟节点，跳过最多 2 个非 w:p 节点，
        #    在遇到的第一个 w:p 上决策（有分页符→True，无→False）。
        skipped = 0
        prev = p.getprevious()
        while prev is not None and skipped <= 2:
            if prev.tag != _q("p"):
                skipped += 1
                warning(f"[_has_page_break_before] 回溯跳过非 w:p 节点 "
                        f"<{prev.tag.split('}')[-1]}>（已跳 {skipped} 个）：{title_txt!r}")
                prev = prev.getprevious()
                continue
            # 遇到第一个 w:p，在此决策，不再继续向上
            if _p_has_page_br(prev):
                warning(f"[_has_page_break_before] 步骤2命中（紧邻 w:p 含 w:br page）："
                        f"{title_txt!r}")
                return True
            warning(f"[_has_page_break_before] 步骤2：紧邻 w:p 无分页符，判定 False："
                    f"{title_txt!r}")
            return False

        warning(f"[_has_page_break_before] 未找到前序 w:p（回溯耗尽），判定 False：{title_txt!r}")
        return False
    except Exception:
        return False

def _has_section_break_before(para) -> bool:
    """判断段落前是否已有"分节符（下一页）"。

    分节符挂在段落的 pPr/sectPr 里；若 sectPr 的 w:type 是 nextPage 或空
    （默认即下一页），视为"已有换页标记"。
    只查段落的直接前序兄弟，不无限回溯。
    """
    try:
        from services.ooxml_service import _q as _q_local
        p_el = para._p if hasattr(para, "_p") else para
        prev = p_el.getprevious()
        count = 0
        while prev is not None and count < 3:
            local = prev.tag.split("}")[-1] if "}" in prev.tag else prev.tag
            if local in ("bookmarkStart", "bookmarkEnd", "proofErr"):
                prev = prev.getprevious()
                count += 1
                continue
            if local != "p":
                break
            # 查该段落的 pPr/sectPr
            for ppr in prev.findall(_q_local("pPr")):
                sectpr = ppr.find(_q_local("sectPr"))
                if sectpr is not None:
                    _type = sectpr.find(_q_local("type"))
                    # type 缺失 / nextPage → 视为"下一页"换页
                    if _type is None or _type.get(_q_local("val")) in (None, "nextPage"):
                        return True
            prev = prev.getprevious()
            count += 1
        return False
    except Exception:
        return False

def _is_in_table(para) -> bool:
    """判断段落是否位于表格单元格内（表格内段落不是章节标题）。"""
    p = para._p if hasattr(para, "_p") else para
    parent = p.getparent()
    return parent is not None and parent.tag == _q("tc")


def _para_is_bold(para) -> bool:
    """判断段落是否整体加粗（python-docx 无 Paragraph.bold，需按 run/样式判定）。

    规则：任一 run 显式加粗，或段落样式字体加粗。章节标题通常为单 run 加粗。
    """
    runs = para.runs
    if runs and any(r.bold is True for r in runs):
        return True
    try:
        style = para.style
        if style is not None and getattr(style.font, "bold", False):
            return True
    except Exception:
        pass
    return False


def _disable_image_compression(doc) -> None:
    """占位函数（当前为空操作）。

    python-docx 通过 doc.save() 直接写 ZIP，完全绕过 Word 应用程序，
    图片不经过 Word 的压缩后处理。Word 压缩仅在用 Word 打开文件后"另存为"时触发。
    若报告被 Word 打开并手动保存过，图片才会被压缩；程序直接生成的文件不受影响。
    """
    pass


def _new_paragraph(doc):
    from docx.oxml import OxmlElement
    p = OxmlElement("w:p")
    return p


def _build_table_element(nrows: int, ncols: int):
    """用 OOXML 直接构造一个 nrows×ncols 的空 w:tbl 元素。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl = OxmlElement("w:tbl")
    # tblPr
    tblPr = OxmlElement("w:tblPr")
    tblStyle = OxmlElement("w:tblStyle")
    tblStyle.set(qn("w:val"), "TableGrid")
    tblPr.append(tblStyle)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "0")
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    tbl.append(tblPr)
    # tblGrid
    grid = OxmlElement("w:tblGrid")
    for _ in range(ncols):
        gc = OxmlElement("w:gridCol")
        grid.append(gc)
    tbl.append(grid)
    # rows
    for _ in range(nrows):
        tr = OxmlElement("w:tr")
        for _ in range(ncols):
            tc = OxmlElement("w:tc")
            tcPr = OxmlElement("w:tcPr")
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), "0")
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)
            tc.append(tcPr)
            p = OxmlElement("w:p")
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl


def _paragraph_from_element(p_element, doc):
    from docx.text.paragraph import Paragraph
    return Paragraph(p_element, doc)


def _element_after(anchor_el, tbl_el) -> bool:
    """判断 tbl_el 在文档顺序上是否位于 anchor_el 之后。

    两者同级时直接比较；若 tbl_el 位于表格单元格内（如表格内段落），
    则向上追溯到 body 直属祖先再比较，从而支持「章节标题 → 表格内段落」的位置判断。
    """
    if anchor_el is tbl_el:
        return False
    body = anchor_el.getparent()
    if body is None:
        return False

    def _top_ancestor(el):
        """返回 el 在 body 下的直属祖先（用于跨层级比较）。"""
        cur = el
        while cur is not None and cur.getparent() is not None \
                and cur.getparent() is not body:
            cur = cur.getparent()
        return cur

    target = tbl_el if tbl_el.getparent() is body else _top_ancestor(tbl_el)
    if target is None:
        return False
    found = False
    for child in body.iterchildren():
        if child is anchor_el:
            found = True
        elif child is target and found:
            return True
    return False


def _iter_all_paragraphs(doc):
    """按文档顺序迭代**所有**段落，含表格单元格内段落。

    python-docx 的 ``doc.paragraphs`` 只返回正文顶层段落，不含表格内段落；
    真实模板常把字段标签放在表格单元格里，故需要本函数做回退搜索。
    """
    from docx.text.paragraph import Paragraph
    for p_el in doc.element.body.iter(qn("w:p")):
        yield Paragraph(p_el, doc)


def _deepclone_tbl(tbl) -> object:
    from copy import deepcopy
    src = tbl._tbl if hasattr(tbl, "_tbl") else tbl
    return deepcopy(src)


# ───────────────────────── 日期过期标红 ─────────────────────────

_DATE_RE = _re.compile(r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}")
_RANGE_RE = _re.compile(
    r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})"          # 开始日期
    r"\s*[~～至\-–—](?:to)?\s*"                 # 分隔符
    r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})",          # 截止日期
    _re.IGNORECASE,
)


def _parse_date(text: str):
    text = (text or "").strip()
    for pat, ext in (
        (r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",
         lambda m: (int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})",
         lambda m: (int(m.group(3)), int(m.group(2)), int(m.group(1)))),
    ):
        m = _re.match(pat, text)
        if m:
            y, mo, d = ext(m)
            if not (1900 <= y <= 2100):   # ← 新增：年份合理性检查
                continue
            try:
                return _datetime.date(y, mo, d)
            except ValueError:
                continue
    return None


def _expired_spans(full: str, today) -> list:
    """返回 full 中需标红的过期日期字符区间 ``[(start, end), ...]``（end 不含）。

    规则：``开始 ~ 截止`` 区间只取**截止日期**与今天比较；独立日期过期则标红。
    """
    spans = []
    start_spans = set()
    for m in _RANGE_RE.finditer(full):
        start_spans.add(m.start(1))
        start_spans.add(m.start(2))
        dt = _parse_date(m.group(2))
        if dt is not None and dt < today:
            spans.append((m.start(2), m.end(2)))
    for m in _DATE_RE.finditer(full):
        if m.start() in start_spans:
            continue
        dt = _parse_date(m.group())
        if dt is not None and dt < today:
            spans.append((m.start(), m.end()))
    return spans


def _make_run(template_r, text: str, expired: bool):
    """克隆 template_r 的 rPr 创建一个新 w:r；expired 时着红色。"""
    nr = _lxml_etree.Element(_q("r"))
    rPr = template_r.find(_q("rPr"))
    if rPr is not None:
        nr.append(_deepcopy(rPr))
    if expired:
        rPr2 = nr.find(_q("rPr"))
        if rPr2 is None:
            rPr2 = _lxml_etree.SubElement(nr, _q("rPr"))
        color = _lxml_etree.SubElement(rPr2, _q("color"))
        color.set(_q("val"), "FF0000")
    t = _lxml_etree.SubElement(nr, _q("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return nr


def _color_expired_in_p(p_el, today=None) -> None:
    """对段落元素内的过期日期做 run 级标红（区间只标截止日期）。"""
    if today is None:
        today = _datetime.date.today()
    runs = p_el.findall(_q("r"))
    if not runs:
        return
    run_texts = []
    full = ""
    for r in runs:
        t = "".join(n.text or "" for n in r.findall(_q("t")))
        run_texts.append(t)
        full += t
    if not full.strip():
        return
    spans = _expired_spans(full, today)
    if not spans:
        return
    is_exp = [False] * len(full)
    for s, e in spans:
        for i in range(s, min(e, len(is_exp))):
            is_exp[i] = True
    new_runs = []
    pos = 0
    for r, rtext in zip(runs, run_texts):
        if not rtext:
            continue
        seg_start = 0
        cur = is_exp[pos]
        for i in range(len(rtext)):
            exp = is_exp[pos + i]
            if exp != cur:
                if rtext[seg_start:i]:
                    new_runs.append(_make_run(r, rtext[seg_start:i], cur))
                seg_start = i
                cur = exp
        if rtext[seg_start:]:
            new_runs.append(_make_run(r, rtext[seg_start:], cur))
        pos += len(rtext)
    for r in runs:
        p_el.remove(r)
    for nr in new_runs:
        p_el.append(nr)
