"""文档服务层（v5.0 重构 · 经 DocumentEngine 操作，底层由 python-docx 实现）。

职责（见 重构.md §6.3）：
  - 业务层（Processor / Strategy）与 DocumentEngine 之间的唯一隔离层。
  - 接收业务指令，根据「物料类别配置表」确定在模板哪里操作，调用 DocumentEngine 执行。

边界（强制约束）：
  - 不导入 python-docx / lxml / win32com；所有 Word 操作经 engine_factory.get_engine()。
  - 不包含业务逻辑判断（由 Processor / Strategy 负责）。
  - 不暴露文档对象给上层（上层只调用本模块函数）。

对外公开 API（doc-first，doc 可省略并回退到 set_active_doc 设置的活跃文档）：
  sync_chapters / add_page_break / insert_text_after / fill_equipment_table /
  insert_image / insert_image_after_text / insert_excel_table_after_text /
  insert_table_cell_value / copy_table_from_doc / find_section_range /
  highlight_expired_calibrations / fill_sample_description / fill_summary
使用模块级活跃文档（无 doc 参数）的 API：
  insert_image_at_placeholder / insert_text_at_placeholder
"""

import os
import re
import shutil
import datetime

from common.logger import info, warning
from services.engine_factory import get_engine
from services import reference_service
from services import image_service
from services import language_pack
import excel_table_builder  # 顶层模块（非 services 包）：读取 Excel 生成表格数据（Electrical 策略使用）
import data_extractor       # 顶层模块：电性 Excel → PNG（ReliabilityCycle 策略使用）
from models.result import ChapterResult

# ── 模块级常量 ────────────────────────────────
_MAX_IMG_H_CM = 8.0   # Excel 图片框：单张图高度上限（等比缩放）

# ───────────────────────── 模块级活跃文档 ─────────────────────────
# process_*_chapter 与占位符 helper 在 §6.3 签名中不含 doc，依赖此全局状态。
# 调用方（report_engine / generic_processor）应在处理文档前调用 set_active_doc(doc)。

_DOC = None
_ENGINE = None

# 记录"原模板的第一个测试章节名"（sync_chapters 执行前设置）
_FIRST_SECTION_NAME = None


def set_first_section_name(name) -> None:
    """由 report_engine 在 sync_chapters 前调用，记录原模板的第一个章节名。"""
    global _FIRST_SECTION_NAME
    _FIRST_SECTION_NAME = (name or "").strip() or None


def set_active_doc(doc) -> None:
    """设置当前活跃文档（供无 doc 参数的公开方法使用）。"""
    global _DOC, _ENGINE
    _DOC = doc
    _ENGINE = get_engine()


def _eng():
    return _ENGINE if _ENGINE is not None else get_engine()


def _doc(doc):
    return doc if doc is not None else _DOC


# ───────────────────────── 纯文本工具（无文档依赖） ─────────────────────────

def _norm_colon(s: str) -> str:
    """全角冒号归一为半角，便于标签匹配。"""
    return (s or "").replace("\uff1a", ":")


def _strip_brackets(s: str) -> str:
    """去除占位符文本外层的成对括号 {} / [] / <>。

    仅当首尾是"匹配的括号对"时才剥：
      {xxx} / [xxx] / <xxx> → 剥
      <br>{1L-1}            → 不剥（首<配>，但尾是}，不成对）
    """
    s = (s or "").strip()
    pairs = {"{": "}", "[": "]", "<": ">"}
    if len(s) >= 2 and s[0] in pairs and s[-1] == pairs[s[0]]:
        return s[1:-1].strip()
    return s


def _parse_date(text: str):
    """解析日期文本为 datetime.date（支持多种格式）。"""
    text = (text or "").strip()
    for pat, ext in (
        (r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",
         lambda m: (int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})",
         lambda m: (int(m.group(3)), int(m.group(2)), int(m.group(1)))),
    ):
        m = re.match(pat, text)
        if m:
            y, mo, d = ext(m)
            try:
                return datetime.date(y, mo, d)
            except ValueError:
                continue
    return None

_RANGE_RE = re.compile(
    r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})"          # 开始日期
    r"\s*[~～至\-–—](?:\s*to\s*)?\s*"            # 分隔符
    r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})",          # 截止日期
    re.IGNORECASE,
)

def _is_expired(date_str: str) -> bool:
    s = (date_str or "").strip()
    # 区间「开始 ~ 截止」→ 取截止日期（右边）；无区间 → 取唯一日期
    m = _RANGE_RE.search(s)
    if m:
        dt = _parse_date(m.group(2))
    else:
        dt = _parse_date(s)
    if dt is None:
        return False
    return dt < datetime.date.today()


def _parse_image_mapping(mapping) -> dict:
    """解析 image_mapping（"1-1→front, 1-2→back"）为 {关键词: 占位符}。"""
    out = {}
    if not mapping:
        return out
    for part in str(mapping).split(","):
        part = part.strip()
        if "→" in part:
            k, v = part.split("→", 1)
        elif "->" in part:
            k, v = part.split("->", 1)
        else:
            continue
        out[_strip_brackets(k.strip())] = _strip_brackets(v.strip())
    return out


def _list_images(folder: str) -> list:
    """递归列出文件夹下所有图片文件（排除系统文件）。"""
    exts = (".jpg", ".jpeg", ".png", ".bmp", ".gif")
    skip = ("_processed", "thumbs.db", "desktop.ini")
    out = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            low = f.lower()
            if low.startswith("~$"):
                continue
            if any(s in low for s in skip):
                continue
            if low.endswith(exts):
                out.append(os.path.join(root, f))
    return out


def _match_one(imgs: list, keyword: str):
    """三轮匹配：精确 → 子串 → 去后缀/同义词模糊。返回图片路径或 None。"""
    names = [(p, os.path.splitext(os.path.basename(p))[0]) for p in imgs]
    # 第一轮：精确匹配（去扩展名）
    for p, n in names:
        if n == keyword:
            return p
    # 第二轮：子串匹配
    for p, n in names:
        if keyword in n or n in keyword:
            return p
    # 第三轮：去后缀 + 同义词模糊
    suffixes = ["照片", "图片", "photo", "image", "pic", "图"]
    kw = keyword
    for s in suffixes:
        kw = kw.replace(s, "")
    synonyms = [("正面", "反面"), ("焊前", "可焊前"), ("前", "焊前")]
    cands = [kw]
    for a, b in synonyms:
        if a in kw:
            cands.append(kw.replace(a, b))
    for p, n in names:
        nn = n
        for s in suffixes:
            nn = nn.replace(s, "")
        for c in cands:
            if c and (c in nn or nn in c):
                return p
    return None


def _read_image_bytes(path: str, width_cm: float = None) -> bytes:
    """裁剪白边后读取图片二进制。
    
    Args:
        width_cm: 传入则按此显示宽度压缩到 200 DPI；不传则不压。
    """
    processed = image_service.process_image(path, target_width_cm=width_cm)
    with open(processed, "rb") as f:
        return f.read()


def _delete_paragraph(para) -> None:
    """从文档 XML 中删除段落元素（无论其位于 body 还是表格单元格内）。

    删除失败时仅记录警告，不向上抛出，保证主流程不受影响。
    """
    try:
        p = para._element
        parent = p.getparent()
        if parent is not None:
            parent.remove(p)
    except Exception as e:
        warning(f"[template_engine] _delete_paragraph 失败：{e}")


def _strip_chinese(text: str) -> str:
    """仅保留英文/数字/符号，删除中日韩表意文字与全角区。"""
    if not text:
        return ""
    return re.sub(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uff00-\uffef]+", "", text).strip()


def _strip_chinese_paren(text: str) -> str:
    """删除中文/半角括号及其内部内容。"""
    if not text:
        return ""
    return re.sub(r"[（(][^）)]*[）)]", "", text).strip()


def _find_image(folder: str, keyword: str):
    """在 folder 下递归查找文件名含 keyword 的图片。"""
    if not folder or not os.path.isdir(folder):
        return None
    return _match_one(_list_images(folder), keyword)


# ───────────────────────── 内部文档辅助 ─────────────────────────

def _fill_field(doc, section_name: str, label: str, value, font_color=None,
                lang: str = "en") -> None:
    """在 section 内查找 label 段落，冒号后替换 value。

    Args:
        lang: 报告语言，仅用于判定「不合格/参考备注」这条特殊字体的标签
            （不同语言标签文本不同，故按 language_pack 比对，不写死英文）。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    section = eng.find_section(doc, section_name)
    if section is None:
        warning(f"[template_engine] _fill_field：未找到章节 {section_name}")
        return
    para = eng.find_paragraph(doc, section, label)
    if para is None:
        warning(f"[template_engine] _fill_field：章节 {section_name} 未找到标签 {label}")
        return
    extra = {}
    L = language_pack.labels(lang)
    if (section_name == L["summary_title"]
            and label == L["summary_notes"]):
        extra = {"font_name": "Times New Roman", "font_size": 14}
    eng.insert_text(doc, para, value or "", font_color, **extra)


def _insert_image_after_label(doc, label: str, img_path: str, width_cm: float) -> bool:
    """在含 label 的段落后插入图片。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    para = eng.find_paragraph_by_text(doc, label)
    if para is None:
        warning(f"[template_engine] _insert_image_after_label：未找到标签 {label}")
        return False
    try:
        data = _read_image_bytes(img_path)
    except Exception as e:
        warning(f"[template_engine] 读取图片失败 {img_path}：{e}")
        return False
    return eng.insert_image(doc, para, data, width_cm)


def _insert_table_after(test_item: str, anchor: str, table_data, width_cm: float) -> bool:
    """在 test_item 章节内 anchor 段落后插入表格。"""
    doc = _doc(None)
    eng = _eng()
    if doc is None:
        return False
    section = eng.find_section(doc, test_item)
    if section is None:
        warning(f"[template_engine] _insert_table_after：未找到章节 {test_item}")
        return False
    para = eng.find_paragraph(doc, section, anchor)
    if para is None:
        warning(f"[template_engine] _insert_table_after：章节 {test_item} 未找到锚点 {anchor}")
        return False
    return eng.insert_table(doc, para, table_data, width_cm)

def _is_before(a_el, b_el) -> bool:
    """判断 a_el 是否在文档中出现在 b_el 之前。"""
    parent = a_el.getparent()
    if parent is None or b_el.getparent() is not parent:
        return False
    try:
        children = list(parent)
        return children.index(a_el) < children.index(b_el)
    except ValueError:
        return False


def _move_section_before(doc, eng, title: str, target_el) -> bool:
    """把章节 title 的全部元素移到 target_el 之前。"""
    body = doc.element.body
    start_para, end_para = eng.find_section_range(doc, title)
    if start_para is None:
        return False
    start_el = start_para._p
    end_el = end_para._p if end_para is not None else None

    section_els = []
    collecting = False
    for child in body.iterchildren():
        if child is start_el:
            collecting = True
        if collecting:
            if end_el is not None and child is end_el:
                break
            section_els.append(child)

    if not section_els:
        return False

    for el in section_els:
        body.remove(el)
    for el in section_els:
        target_el.addprevious(el)
    return True


def _reorder_sections(doc, eng, required) -> None:
    """按 required 顺序重排文档中已存在的测试章节。

    从后往前处理：对每个相邻对 (curr, next)，若 curr 不在 next 之前，
    则把整个 curr 章节移到 next 之前。从后往前保证"next 的位置"已稳定。
    新增/删除已在 sync_chapters 其他步骤处理，本函数只负责"已存在章节重排"。
    """
    for i in range(len(required) - 2, -1, -1):
        curr_title = required[i]
        next_title = required[i + 1]

        curr_para = eng.find_section(doc, curr_title)
        next_para = eng.find_section(doc, next_title)
        if curr_para is None or next_para is None:
            continue

        if _is_before(curr_para._p, next_para._p):
            continue

        if _move_section_before(doc, eng, curr_title, next_para._p):
            info(f"[sync_chapters] 重排: {curr_title} → 移到 {next_title} 之前")

    eng.invalidate_section_cache(doc)
# ───────────────────────── 章节同步 / 分页 ─────────────────────────

def sync_chapters(doc, chapter_contexts, default_chapter_template, lang: str = "en") -> None:
    """同步模板章节：保留已有、删除多余、新增缺失。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        warning("[sync_chapters] 活跃文档为空")
        return
    if not chapter_contexts:
        warning("[sync_chapters] chapter_contexts 为空，不执行同步")
        return

    required = [ctx.config.test_item for ctx in chapter_contexts]
    required_set = set(required)
    sections = eng.list_sections(doc)
    info(f"[sync_chapters] list_sections 返回：{[txt for txt, _ in sections]}")

    def _norm_title(text: str) -> str:
        # 去除开头的数字编号和紧跟的标点/空格，如 "6. 板面检查" -> "板面检查"
        import re as _re2
        return _re2.sub(r"^\d+[.、_\-\s]*\s*", "", text.strip()).strip()

    existing = {_norm_title(txt): h for txt, h in sections}

    to_delete = set(existing) - required_set
    to_add = [it for it in required if it not in existing]

    # 先克隆缺失章节（此时默认模板章节仍存在于文档中，可被 clone_section 复用）。
    # 若先删除再克隆，默认模板会被误删导致克隆失败。
    #
    # v5.13 修复：不再把所有新增章节统一插到 RoHS 前，而是**按期望顺序**
    # 把每个新增章节插到"期望顺序中，它之后第一个已存在章节"的前面，
    # 保证章节页顺序与汇总页（基于 required 顺序）一致。
    current_existing = set(existing.keys())   # 动态维护：克隆一个就加入

    for item in required:                     # ← 按期望顺序遍历
        if item in current_existing:
            continue                          # 已存在，跳过
        if not default_chapter_template:
            warning("[sync_chapters] default_chapter_template 为空，跳过新增")
            break

        # 找"期望顺序中 item 之后第一个已存在章节"
        before_target = None
        idx = required.index(item)
        for nxt in required[idx + 1:]:
            if nxt in current_existing:
                before_target = nxt
                break

        # 兜底：找不到参考 → 插到 RoHS 前（若无 RoHS 则文末）
        if before_target is None:
            before_target = ("RoHS"
                             if eng.find_section(doc, "RoHS") is not None
                             else None)

        ok = eng.clone_section(doc, default_chapter_template, item,
                               before_target, lang=lang)
        if ok:
            current_existing.add(item)
            info(f"[sync_chapters] 新增章节: {item}（插入至 {before_target or '文末'}）")
        else:
            warning(f"[sync_chapters] 新增章节失败: {item}")

    # 后删除多余章节（含未被需要的默认模板章节本身）。
    for title in to_delete:
        eng.delete_section(doc, title)
        info(f"[sync_chapters] 删除章节: {title}")
    # ── v5.13：对已存在的章节按期望顺序重排（修复"章节页顺序与汇总页不一致"）──
    _reorder_sections(doc, eng, required)


def add_page_break(doc, test_item) -> None:
    """在章节标题前插入分页符（已存在则跳过）。

    特例：如果当前章节是【模板原本的第一个测试章节】，则完全跳过——
    模板本身已保证它独立成页，脚本干预反而会误删分节符（丢页眉）
    或插入多余分页符（多空白页）。

    注意：判断依据是"sync_chapters 之前记录的模板首章节名"，
    而不是"当前文档里排第一的"——因为原第一可能已被删除。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    section = eng.find_section(doc, test_item)
    if section is None:
        warning(f"[add_page_break] 未找到章节 {test_item}")
        return

    # ── 是否 = 模板原本的第一个章节？只跳它 ──
    if _FIRST_SECTION_NAME and (test_item or "").strip() == _FIRST_SECTION_NAME:
        info(f"[add_page_break] {test_item} 是模板首个章节，跳过处理")
        return

    # ── 其他章节：原有逻辑 ──
    clean_blank_pages_before(doc, test_item)
    already_break = eng.has_page_break_before(section)
    # 新增：判断"分节符(下一页)"也算"已有换页标记"，避免重复插分页符
    already_section = False
    try:
        from services.docx_helpers import _has_section_break_before
        already_section = _has_section_break_before(section)
    except Exception as _e:
        warning(f"[add_page_break] 检查分节符失败（忽略）：{_e}")
    warning(f"[add_page_break] 章节={test_item!r}  已有分页符={already_break}  已有分节符={already_section}")
    if already_break or already_section:
        return
    eng.add_page_break(doc, section)


def clean_blank_pages_before(doc, section_name: str) -> None:
    """清理指定章节标题前的多余空白页（连续分页符 / 多余空段落）。

    规则（版本 B：激进清理）：
      1. 删除标题段落自身内的分页符；
      2. 从标题向前扫最多 8 个兄弟节点：
         - bookmark / proofErr → 跳过继续扫
         - 表格（w:tbl）→ 停止
         - 有文本 / 有图片的段落 → 停止（不删）
         - 纯空段落（无文本、无图片、无表格）→ 收集
         - 含分页符的纯空段落 → 收集
      3. 删除策略：
         - 如果有分页符段落：保留最靠近标题的 1 个，其余分页符段 + 所有空段落全删
         - 如果只有空段落：全部删除（后续由 add_page_break 负责插入新分页符）

    失败仅警告，不抛异常。
    """
    from services.ooxml_service import _q as _q_local

    eng = _eng()
    if doc is None:
        return
    section = eng.find_section(doc, section_name)
    if section is None:
        return
    p_el = section._p if hasattr(section, "_p") else section._element
    if p_el is None:
        return

    try:
        def _has_text(el) -> bool:
            for t in el.findall(".//" + _q_local("t")):
                if (t.text or "").strip():
                    return True
            return False

        def _has_media(el) -> bool:
            for tag in ("drawing", "pict", "object", "tbl"):
                if el.find(".//" + _q_local(tag)) is not None:
                    return True
            return False

        def _has_sectPr(el) -> bool:
            """段落是否携带分节符（sectPr 在 pPr 里）——有则不能删。"""
            for ppr in el.findall(".//" + _q_local("pPr")):
                if ppr.find(_q_local("sectPr")) is not None:
                    return True
            return False

        def _has_page_break(el) -> bool:
            for br in el.findall(".//" + _q_local("br")):
                if br.get(_q_local("type")) == "page":
                    return True
            return False

        # ── 1. 删除标题段落自身内的分页符 ──
        for br in p_el.findall(".//" + _q_local("br")):
            if br.get(_q_local("type")) == "page":
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)

        # ── 2. 向前扫描收集 ──
        blank_paras = []          # 纯空段落
        break_paras = []          # 含分页符的纯空段落

        prev = p_el.getprevious()
        count = 0
        while prev is not None and count < 8:
            tag = prev.tag
            local = tag.split("}")[-1] if "}" in tag else tag

            if local in ("bookmarkStart", "bookmarkEnd", "proofErr"):
                prev = prev.getprevious()
                count += 1
                continue

            if local == "tbl":
                break

            if local != "p":
                break

            if _has_text(prev) or _has_media(prev) or _has_sectPr(prev):
                break

            if _has_page_break(prev):
                break_paras.append(prev)
            else:
                blank_paras.append(prev)

            prev = prev.getprevious()
            count += 1

        # ── 3. 决定删除哪些（版本 B：激进清理）──
        to_remove = []
        if break_paras:
            # 保留最靠近标题的 1 个分页符段落，其余分页符和所有空段落全删
            to_remove = break_paras[1:] + blank_paras
        else:
            # 没有分页符，空段落全删
            to_remove = blank_paras

        # ── 4. 执行删除 ──
        for el in to_remove:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

        if to_remove:
            info(f"[clean_blank_pages_before] {section_name}：清理 {len(to_remove)} 个多余节点")
    except Exception as e:
        warning(f"[clean_blank_pages_before] {section_name} 清理失败：{e}")


# ───────────────────────── 文本 / 图片 / 表格填充 ─────────────────────────

def insert_text_after(doc, section_name, text_to_find, value,
                      require_bold=True, font_color=None) -> None:
    """在章节内搜索标签段落，冒号后替换 value（页眉走专用逻辑）。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    if section_name == "Header":
        eng.insert_text_in_header(doc, text_to_find, value or "")
        return
    section = eng.find_section(doc, section_name, require_bold=require_bold)
    if section is None:
        warning(f"[insert_text_after] 未找到章节 {section_name}")
        return
    para = eng.find_paragraph(doc, section, text_to_find)
    if para is None:
        warning(f"[insert_text_after] 章节 {section_name} 未找到标签 {text_to_find}")
        return
    eng.insert_text(doc, para, value or "", font_color)


def insert_image(doc, name, img_path, width_cm) -> None:
    """按占位符/名称文本搜索，在其段落后插入图片，并删除占位符段落。

    支持占位符在普通段落或表格单元格内两种情况。
    插入成功后删除占位符段落；失败则保留占位符便于排查。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    target = _strip_brackets(name)
    para = eng.find_paragraph_by_text(doc, target)
    if para is None:
        warning(f"[insert_image] 未找到占位符/名称 {name}")
        return
    try:
        data = _read_image_bytes(img_path, width_cm)
    except Exception as e:
        warning(f"[insert_image] 读取图片失败 {img_path}：{e}")
        return
    ok = eng.insert_image(doc, para, data, width_cm)
    if ok:
        _delete_paragraph(para)


def insert_image_after_text(doc, section_name, text_to_find, img_path, width_cm) -> bool:
    """章节内搜索锚点段落后插入图片。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    section = eng.find_section(doc, section_name)
    if section is None:
        warning(f"[insert_image_after_text] 未找到章节 {section_name}")
        return False
    para = eng.find_paragraph(doc, section, text_to_find)
    if para is None:
        warning(f"[insert_image_after_text] 章节 {section_name} 未找到 {text_to_find}")
        return False
    try:
        data = _read_image_bytes(img_path)
    except Exception as e:
        warning(f"[insert_image_after_text] 读取图片失败 {img_path}：{e}")
        return False
    return eng.insert_image(doc, para, data, width_cm)


def insert_excel_table_after_text(doc, section_name, text_to_find,
                                  source_folder, width_cm=18.0,
                                  skip_columns=None, skip_rows=None) -> bool:
    """从 source_folder（原始记录子文件夹）读取首个 Excel 构建表格，插入到锚点段落后。

    旧电性/可靠性策略使用；新策略请改用 process_electrical_chapter。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    if not os.path.isdir(source_folder):
        warning(f"[insert_excel_table_after_text] 目录不存在 {source_folder}")
        return False
    folder_path = os.path.dirname(os.path.dirname(source_folder))  # 编号文件夹
    src_name = os.path.basename(source_folder)
    table_data = excel_table_builder.build_table_data(
        folder_path, src_name, table_width_cm=width_cm,
        skip_columns=skip_columns, skip_rows=skip_rows)
    if table_data is None:
        warning(f"[insert_excel_table_after_text] 未读取到表格数据 {src_name}")
        return False
    return _insert_table_after(section_name, text_to_find, table_data, width_cm)


def insert_table_cell_value(doc, section_name, text_to_find, value,
                            require_bold=True, font_name=None, font_size=None,
                            align=None) -> None:
    """填充「标签在单元格A、值在同行下一单元格B」的表格型字段。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    section = eng.find_section(doc, section_name, require_bold=require_bold)
    if section is None:
        warning(f"[insert_table_cell_value] 未找到章节 {section_name}")
        return
    para = eng.find_paragraph(doc, section, text_to_find)
    if para is None:
        warning(f"[insert_table_cell_value] 章节 {section_name} 未找到 {text_to_find}")
        return
    ok = eng.set_adjacent_cell_text(doc, para, value or "",
                                     font_name=font_name, font_size=font_size,
                                     align=align)
    if not ok:
        warning(f"[insert_table_cell_value] 未找到相邻值单元格 {text_to_find}")


def read_table_cell_value(doc, section_name, text_to_find,
                           require_bold=True) -> str:
    """读取「标签在单元格A、值在同行下一单元格B」的表格型字段的值。
    返回标签右侧相邻单元格的文本；找不到标签或相邻单元格时返回空字符串。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return ""
    section = eng.find_section(doc, section_name, require_bold=require_bold)
    if section is None:
        return ""
    para = eng.find_paragraph(doc, section, text_to_find)
    if para is None:
        return ""
    try:
        from docx.table import Table as _Table
        from services.ooxml_service import _q
        p = para._p if hasattr(para, "_p") else para
        parent = p.getparent()
        while parent is not None and parent.tag != _q("tc"):
            parent = parent.getparent()
        if parent is None:
            return ""
        tbl_el = parent.getparent().getparent()
        tbl = _Table(tbl_el, doc)
        for row in tbl.rows:
            for c_i, cell in enumerate(row.cells):
                if cell._tc is parent and c_i + 1 < len(row.cells):
                    return (row.cells[c_i + 1].text or "").strip()
    except Exception as e:
        warning(f"[template_engine] read_table_cell_value 失败：{e}")
    return ""


def copy_table_from_doc(doc, section_name, source_doc_path, source_anchor, target_anchor) -> bool:
    """从源报告复制表格到目标文档（跨文档）。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    src_doc = eng.open(source_doc_path)
    if src_doc is None:
        warning(f"[copy_table_from_doc] 源报告打开失败 {source_doc_path}")
        return False
    return eng.copy_table(doc, section_name, src_doc, source_anchor, target_anchor)

def copy_table_by_source_para_from_doc(doc, section_name, source_doc_path,
                                        source_para_text, target_anchor) -> bool:
    """按源文档加粗段落定位表格，插入目标文档锚点后（跨文档）。

    与 copy_table_from_doc 的区别：源文档用"加粗段落锚点"而非"章节名"定位。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    src_doc = eng.open(source_doc_path)
    if src_doc is None:
        warning(f"[copy_table_by_source_para_from_doc] 源报告打开失败 {source_doc_path}")
        return False
    return eng.copy_table_by_source_para(doc, section_name, src_doc,
                                          source_para_text, target_anchor)

def find_section_range(doc, section_name, require_bold=True):
    """返回 (start_para, end_para)。"""
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return None, None
    return eng.find_section_range(doc, section_name)


def fill_equipment_table(doc, section_name, equipment_info, lang: str = "en") -> None:
    """填充章节内设备表（4 列：Name/Model/No/Expiry）。

    Args:
        lang: 报告语言。设备表锚点（"Test Equipment" / "检测设备"）与表头
            （"Equipment Name…" / "设备名称…"）均按语言切换；表头已与目标语言
            一致时不重写，避免无谓改动模板既有格式。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    if isinstance(equipment_info, dict):
        equipment_info = [equipment_info]

    L = language_pack.labels(lang)
    section = eng.find_section(doc, section_name)
    if section is None:
        warning(f"[fill_equipment_table] 未找到章节 {section_name}")
        return
    tbl = eng.find_table(doc, section, L["equip_title"])
    if tbl is None:
        warning(f"[fill_equipment_table] 章节 {section_name} 未找到设备表")
        return

    # 表头切换（仅当现有表头与目标语言不一致时重写，保留模板原有格式）
    headers = [L["equip_name"], L["equip_model"], L["equip_no"], L["equip_cal"]]
    try:
        if eng.table_row_count(tbl) > 0:
            for j, want in enumerate(headers):
                if j >= eng.table_column_count(tbl):
                    break
                cur = (eng.table_cell_text(tbl, 0, j) or "").strip()
                if cur != want:
                    eng.set_cell_text_preserve(tbl, 0, j, want, align="center")
    except Exception as e:
        warning(f"[fill_equipment_table] 设备表表头切换失败：{e}")

    today = datetime.date.today()
    for i, dev in enumerate(equipment_info):
        row_idx = i + 1  # 第 0 行为表头
        if row_idx >= eng.table_row_count(tbl):
            eng.add_table_row(tbl)
        vals = [
            dev.get("name", "/") or "/",
            dev.get("model", "/") or "/",
            dev.get("no", "/") or "/",
            dev.get("expiry", "/") or "/",
        ]
        for j, val in enumerate(vals):
            eng.set_cell_text(tbl, row_idx, j, val, align="center")
        if vals[3] not in ("", "/") and _is_expired(vals[3]):
            eng.color_expired_in_cell(tbl, row_idx, 3, today)
            # 【新增】打警告日志
            warning(f"[fill_equipment_table] {section_name}："
                    f"设备 {vals[0]}({vals[2]}) 校准已过期（{vals[3]}）")

    # 删除模板预留的多余数据行
    needed = len(equipment_info) + 1  # 表头 + 设备行
    while eng.table_row_count(tbl) > needed:
        eng.remove_table_row(tbl, eng.table_row_count(tbl) - 1)

    # 设置总宽 18cm，autofit 模式（列宽由 Word 按内容自动分配），单元格内边距清零
    eng.set_table_autofit(tbl, 18.0)
    eng.set_table_cell_margins(tbl, top=0, bottom=0, left=0, right=0)
    info(f"[fill_equipment_table] 已填充设备表: {section_name}")


# ───────────────────────── 占位符 helper（无 doc 参数） ─────────────────────────

def insert_image_at_placeholder(placeholder, img_path, width_cm,
                                section_name=None) -> None:
    """按占位符文本搜索，在其段落后插入图片，并删除占位符段落。

    支持占位符在普通段落或表格单元格内两种情况。
    插入成功后删除占位符段落；失败则保留占位符便于排查。

    Args:
        section_name: 传入时使用 ``find_placeholder`` 在章节内精确匹配（推荐）；
                      不传时回退到 ``find_paragraph_by_text`` 全文子串搜索（兼容旧调用）。
    """
    doc = _doc(None)
    eng = _eng()
    if doc is None:
        warning("[insert_image_at_placeholder] 活跃文档为空")
        return

    if section_name:
        # 推荐路径：章节内精确匹配 {placeholder}
        section = eng.find_section(doc, section_name)
        # placeholder 可能带花括号也可能不带，统一补全
        ph_with_braces = (placeholder if placeholder.startswith("{")
                          else f"{{{placeholder}}}")
        para = eng.find_placeholder(doc, section, ph_with_braces)
    else:
        # 兼容路径：全文子串搜索（旧逻辑）
        target = _strip_brackets(placeholder)
        para = eng.find_paragraph_by_text(doc, target)

    if para is None:
        warning(f"[insert_image_at_placeholder] 未找到占位符 {placeholder!r}"
                + (f" (section={section_name!r})" if section_name else ""))
        return
    try:
        data = _read_image_bytes(img_path, width_cm)
    except Exception as e:
        warning(f"[insert_image_at_placeholder] 读取图片失败 {img_path}：{e}")
        return
    ok = eng.insert_image(doc, para, data, width_cm)
    if ok:
        _delete_paragraph(para)


def insert_text_at_placeholder(placeholder, text) -> None:
    """按占位符文本搜索，将其段落文本替换为 text。"""
    doc = _doc(None)
    eng = _eng()
    if doc is None:
        warning("[insert_text_at_placeholder] 活跃文档为空")
        return
    target = _strip_brackets(placeholder)
    para = eng.find_paragraph_by_text(doc, target)
    if para is None:
        warning(f"[insert_text_at_placeholder] 未找到占位符 {placeholder}")
        return
    eng.insert_text(doc, para, text or "")


def insert_image_auto_table(doc, section_name, image_list, width_cm,
                             anchor_para=None, blank_if_empty=False) -> bool:
    """在章节的 Result description 段落后插入图片表格（全自动模式）。

    布局规则（v5.9）：
      - 每张图宽度由文件名"最后一个括号"决定（半角/全角都支持）；
        无括号则用传入的 width_cm；都无则默认 7cm。
      - 宽度 > 9cm 的图：占一整行（2 列合并 → 18cm 单元格），图片按自身
        宽度居中显示。
      - 宽度 ≤ 9cm 的图：占半行（1 列 9cm），两张小图两两配对成一行。
      - 大图前如果有未配对的小图，该小图单独占左半行，右半行填 "/"。
      - 最后若剩一个未配对的小图，同上填 "/"。
      - 表格总宽固定 18cm（2 列 × 9cm）。

    image_list 支持两种格式：
      - [(name, path), ...]          → 宽度统一用 width_cm（旧调用方式）
      - [(name, path, width), ...]   → 每张图用各自 width（新调用方式）
    """
    import math
    from io import BytesIO
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.table import Table as _Table
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from services.ooxml_service import set_cell_background
    from services.docx_helpers import _ensure_table_borders

    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False

    # ── 定位插入点 ──
    if anchor_para is None:
        section = eng.find_section(doc, section_name)
        if section is None:
            warning(f"[insert_image_auto_table] 未找到章节 {section_name}")
            return False
        for label in ("Result description", "结果说明"):
            anchor_para = eng.find_paragraph(doc, section, label)
            if anchor_para:
                break
        if anchor_para is None:
            warning(f"[insert_image_auto_table] 章节 {section_name} 未找到 Result description 锚点")
            return False

    # ── 1. 归一化 image_list（兼容 2 元组和 3 元组）──
    normalized = []
    for item in image_list:
        if len(item) == 3:
            name, path, w = item
            normalized.append((name, path, float(w)))
        else:
            name, path = item
            normalized.append((name, path, float(width_cm)))

    if not normalized:
        return False

    # ── 2. 预计算布局 ──
    # layout 元素：("row",[img]) 大图整行 / ("pair",[img1,img2]) 两小图 / ("half",[img]) 半行小图
    layout = []
    pending_small = None
    for img in normalized:
        _, _, w = img
        if w > 9:
            if pending_small is not None:
                layout.append(("half", [pending_small]))
                pending_small = None
            layout.append(("row", [img]))
        else:
            if pending_small is not None:
                layout.append(("pair", [pending_small, img]))
                pending_small = None
            else:
                pending_small = img
    if pending_small is not None:
        layout.append(("half", [pending_small]))

    # ── 3. 建表格（固定 2 列 × 9cm = 18cm）──
    n_rows = len(layout) * 2      # 每个 layout 项占 2 个物理行（名称行、图片行）
    ncols = 2
    _DXA = 567
    total_dxa = int(18.0 * _DXA)
    col_dxa = int(9.0 * _DXA)

    tbl_el = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    style = OxmlElement("w:tblStyle")
    style.set(qn("w:val"), "TableGrid")
    tblPr.append(style)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    tbl_el.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for _ in range(ncols):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(col_dxa))
        grid.append(gc)
    tbl_el.append(grid)
    for _ in range(n_rows):
        tr = OxmlElement("w:tr")
        for _ in range(ncols):
            tc = OxmlElement("w:tc")
            tcPr = OxmlElement("w:tcPr")
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_dxa))
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)
            tc.append(tcPr)
            tc.append(OxmlElement("w:p"))
            tr.append(tc)
        tbl_el.append(tr)

    _ensure_table_borders(tbl_el)
    anchor_para._p.addnext(tbl_el)
    tbl = _Table(tbl_el, doc)

    # ── 4. 填充辅助 ──
    def _clear_para(p):
        for r in list(p.runs):
            r._r.getparent().remove(r._r)

    def _fill_name(row, col, text):
        cell = tbl.cell(row, col)
        p = cell.paragraphs[0]
        _clear_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.size = Pt(10)
        set_cell_background(cell, "D9D9D9")

    def _fill_slash(row, col):
        cell = tbl.cell(row, col)
        p = cell.paragraphs[0]
        _clear_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("/")

    def _fill_image(row, col, path, img_width):
        cell = tbl.cell(row, col)
        p = cell.paragraphs[0]
        _clear_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not path:
            if blank_if_empty:
                return
            p.add_run("/")
            return
        try:
            processed = image_service.process_image(path, target_width_cm=img_width)
            with open(processed, "rb") as f:
                data = f.read()
            run = p.add_run()
            run.add_picture(BytesIO(data), width=Cm(img_width))
            image_service.cleanup_processed(processed)
        except Exception as e:
            warning(f"[insert_image_auto_table] 图片插入失败 {path}：{e}")
            p.add_run("/")

    # ── 5. 按 layout 填充 ──
    for idx, (kind, imgs) in enumerate(layout):
        name_row = idx * 2
        img_row = idx * 2 + 1

        if kind == "row":
            name, path, w = imgs[0]
            # 合并 2 列
            tbl.cell(name_row, 0).merge(tbl.cell(name_row, 1))
            _fill_name(name_row, 0, name)
            tbl.cell(img_row, 0).merge(tbl.cell(img_row, 1))
            _fill_image(img_row, 0, path, w)

        elif kind == "pair":
            name_l, path_l, w_l = imgs[0]
            name_r, path_r, w_r = imgs[1]
            _fill_name(name_row, 0, name_l)
            _fill_name(name_row, 1, name_r)
            _fill_image(img_row, 0, path_l, w_l)
            _fill_image(img_row, 1, path_r, w_r)

        elif kind == "half":
            name, path, w = imgs[0]
            _fill_name(name_row, 0, name)
            _fill_slash(name_row, 1)
            _fill_image(img_row, 0, path, w)
            _fill_slash(img_row, 1)

    return True

def fill_sample_description(doc, task_ctx) -> None:
    """填充样品描述页（步骤3）。

    保真要点（对照真实模板 报告模板.docx）：
      样品描述字段是**「标签 | 值」两列表格**（T1），值必须写入标签右侧的
      **相邻值单元格**，不能写在标签单元格内。故统一走 ``insert_table_cell_value``
      （引擎的 ``set_adjacent_cell_text``），而非段落型的 ``_fill_field``。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    record = task_ctx.record
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    L = language_pack.labels(lang)
    SEC = L["sample_title"]

    _F = {"font_name": "Times New Roman", "font_size": 12}  # 小四号

    insert_table_cell_value(doc, SEC, L["sample_material_number"], record.material_code, **_F)
    insert_table_cell_value(doc, SEC, L["sample_material_category"],
                            record.material_category, **_F)
    subcat = reference_service.get_translation(
        "material", record.material_subclass, lang) \
        if record.material_subclass else ""
    insert_table_cell_value(doc, SEC, L["sample_material_subcategory"], subcat, **_F)
    insert_table_cell_value(doc, SEC, L["sample_model_series"], record.material_model, **_F)
    _desc = (record.description if lang == "zh"
             else _strip_chinese(_strip_chinese_paren(record.description)))
    insert_table_cell_value(doc, SEC, L["sample_material_description"], _desc,
                            align="left", **_F)
    _supplier = ((record.supplier or "").strip() if lang == "zh"
                 else _strip_chinese_paren(record.supplier))
    insert_table_cell_value(doc, SEC, L["sample_supplier"], _supplier, **_F)
    qty = f"{(record.quantity or '').strip()} pcs"
    insert_table_cell_value(doc, SEC, L["sample_quantity"], qty, **_F)

    # 图片目录：优先使用 原始记录/ 下名称含「标签」的子文件夹，
    # 该子文件夹内存放样品标签图（1-1）和样品图（1-2）；
    # 找不到含「标签」的子文件夹时退回 原始记录/ 本身。
    _raw_data = os.path.join(task_ctx.folder_path, "原始记录")
    img_root = _raw_data  # 默认
    if os.path.isdir(_raw_data):
        for _entry in os.listdir(_raw_data):
            if "标签" in _entry and os.path.isdir(os.path.join(_raw_data, _entry)):
                img_root = os.path.join(_raw_data, _entry)
                from common.logger import info as _info
                _info(f"[fill_sample_description] 使用标签子文件夹：{img_root}")
                break

    def _insert_sample_img(placeholder_key: str, match_kw: str) -> None:
        """按占位符子串搜索定位单元格，插入图片；找不到图片或占位符则填 '/'。

        使用 eng.find_paragraph_by_text（全文子串匹配）搜索占位符，
        绕过 find_paragraph 的章节边界限制（next_title 截止可能提前停止）。
        占位符带花括号，不会误匹配表头行「Sample Label」纯文本。
        """
        img_path = _find_image(img_root, match_kw)
        ph_text = f"{{{placeholder_key}}}"
        # 全文搜索：find_paragraph_by_text 不受章节边界限制，含表格单元格
        ph_para = eng.find_paragraph_by_text(doc, ph_text)
        if ph_para is None:
            warning(f"[fill_sample_description] 未找到占位符 {ph_text!r}，跳过图片插入")
            return
        if not img_path:
            warning(f"[fill_sample_description] 未找到图片（关键词={match_kw!r}），跳过处理")
            return
        try:
            data = _read_image_bytes(img_path)
        except Exception as e:
            warning(f"[fill_sample_description] 读取图片失败 {img_path}：{e}，跳过处理")
            return
        # 宽度：括号优先 → 7.0（局部导入，避免与 strategies 循环依赖）
        from strategies.image_auto import _extract_width_from_name
        _stem = os.path.splitext(os.path.basename(img_path))[0]
        _w = _extract_width_from_name(_stem, 7.0)
        ok = eng.insert_image(doc, ph_para, data, _w)
        if ok:
            from services.ooxml_service import delete_paragraph as _del_p
            _del_p(ph_para)
        else:
            warning(f"[fill_sample_description] 图片插入失败，跳过处理")

    _insert_sample_img("samplelabel", "1-1")
    _insert_sample_img("sampleimage", "1-2")


def fill_sample_description_images_only(doc, task_ctx) -> None:
    """独立模式：仅填充 Sample Label / Sample Image 图片，不填任何文本字段。

    图片来源同 fill_sample_description：优先 {folder_path}/原始记录/ 下名称含
    「标签」的子文件夹；占位符 {samplelabel}→"1-1"、{sampleimage}→"1-2"。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    _raw_data = os.path.join(task_ctx.folder_path, "原始记录")
    img_root = _raw_data
    if os.path.isdir(_raw_data):
        for _entry in os.listdir(_raw_data):
            if "标签" in _entry and os.path.isdir(os.path.join(_raw_data, _entry)):
                img_root = os.path.join(_raw_data, _entry)
                info(f"[fill_sample_description_images_only] 使用标签子文件夹：{img_root}")
                break

    def _insert_sample_img(placeholder_key: str, match_kw: str) -> None:
        img_path = _find_image(img_root, match_kw)
        ph_text = f"{{{placeholder_key}}}"
        ph_para = eng.find_paragraph_by_text(doc, ph_text)
        if ph_para is None:
            warning(f"[fill_sample_description_images_only] 未找到占位符 {ph_text!r}，跳过")
            return
        if not img_path:
            warning(f"[fill_sample_description_images_only] 未找到图片（关键词={match_kw!r}），跳过")
            return
        try:
            data = _read_image_bytes(img_path)
        except Exception as e:
            warning(f"[fill_sample_description_images_only] 读取图片失败 {img_path}：{e}，跳过")
            return
        # 宽度：括号优先 → 7.0（局部导入，避免与 strategies 循环依赖）
        from strategies.image_auto import _extract_width_from_name
        _stem = os.path.splitext(os.path.basename(img_path))[0]
        _w = _extract_width_from_name(_stem, 7.0)
        ok = eng.insert_image(doc, ph_para, data, _w)
        if ok:
            _delete_paragraph(ph_para)
        else:
            warning(f"[fill_sample_description_images_only] 图片插入失败，跳过")

    _insert_sample_img("samplelabel", "1-1")
    _insert_sample_img("sampleimage", "1-2")


def fill_summary(doc, chapter_results, lang: str = "en") -> str:
    """填充汇总页（步骤5），返回最终结论 PASS / FAIL。

    兼容 4 列（Test Item/Test Basis/Sample Number/Test Result）与 5 列
    （No./Test Item/Test Basis/Sample Number/Test Result）两种 Test Contents 表：
    通过表头关键词定位各列索引，不依赖固定列序。

    Args:
        lang: 报告语言。汇总页章节标题、表头关键词、统计段落关键词、不合格/参考
            备注标签全部按语言切换。返回值恒为 ``"PASS"`` / ``"FAIL"``（不随
            语言变化，供 file_renamer 重命名）；中文的「合格 / 不合格」只写进
            文档统计段落的判定词位置。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return "PASS"
    L = language_pack.labels(lang)
    SEC = L["summary_title"]

    section = eng.find_section(doc, SEC)
    if section is None:
        warning(f"[fill_summary] 未找到 {SEC} 章节")
        return "PASS"
    tbl = None
    key_item = L["col_test_item"]
    for t in eng.iter_tables(doc):
        try:
            ncols = eng.table_column_count(t)
            first = [(eng.table_cell_text(t, 0, c) or "") for c in range(ncols)]
        except Exception:
            first = []
        if any(key_item in (h or "") for h in first):
            tbl = t
            break
    if tbl is None:
        warning(f"[fill_summary] 未找到汇总表（表头关键词 {key_item!r}）")
        return "PASS"

    # 定位各列索引（按表头关键词，顺序即优先级）
    headers = [(eng.table_cell_text(tbl, 0, c) or "") for c in
               range(eng.table_column_count(tbl))]
    col_keys = language_pack.summary_col_keys()
    _col = {}
    for c, h in enumerate(headers):
        for key in col_keys:
            if L[key].lower() in (h or "").lower():
                _col.setdefault(key, c)
    k_no, k_item = col_keys[0], col_keys[1]
    k_basis, k_sn, k_result = col_keys[2], col_keys[3], col_keys[4]
    ti = _col.get(k_item, 0)
    tb = _col.get(k_basis, ti + 1)
    sn = _col.get(k_sn, ti + 2)
    tr = _col.get(k_result, ti + 3)
    no_col = _col.get(k_no)

    n = len(chapter_results)
    while eng.table_row_count(tbl) - 1 < n:
        eng.add_table_row(tbl)
    while eng.table_row_count(tbl) - 1 > n:
        eng.remove_table_row(tbl, eng.table_row_count(tbl) - 1)

    notes = []
    # 中文报告将 PASS/FAIL/Reference 映射为中文显示文本
    _ZH_RESULT = {"PASS": "合格", "FAIL": "不合格", "Reference": "供参考"}
    _DATA_FONT = {"font_name": "Times New Roman", "font_size": 12, "align": "center"}
    for i, cr in enumerate(chapter_results):
        r = i + 1
        if no_col is not None:
            eng.set_cell_text(tbl, r, no_col, str(i + 1), **_DATA_FONT)
        eng.set_cell_text(tbl, r, ti, cr.test_item, **_DATA_FONT)
        eng.set_cell_text(tbl, r, tb, cr.test_basis, **_DATA_FONT)
        eng.set_cell_text(tbl, r, sn, cr.sample_number, **_DATA_FONT)
        display_result = _ZH_RESULT.get(cr.test_result, cr.test_result) if lang == "zh" else cr.test_result
        eng.set_cell_text(tbl, r, tr, display_result,
                          color=0xFF0000 if cr.test_result == "FAIL" else None,
                          **_DATA_FONT)
        if cr.test_result in ("FAIL", "Reference") and cr.result_description:
            notes.append(f"{cr.test_item}: {cr.result_description}")

    note_text = "; ".join(f"{i + 1}. {nt}" for i, nt in enumerate(notes))
    _fill_field(doc, SEC, L["summary_notes"], note_text,
                font_color=0xFF0000 if note_text else None, lang=lang)

    # 统计段落（英文 "A total of ( N ) ... judged as: Fail"；
    #           中文 "本次共测试 ( N ) 项 … 判定本次试验样品：不合格"）
    _fill_statistics(doc, eng, len(chapter_results),
                     sum(1 for c in chapter_results if c.test_result == "PASS"),
                     sum(1 for c in chapter_results if c.test_result == "FAIL"),
                     sum(1 for c in chapter_results if c.test_result == "Reference"),
                     lang)

    verdict = "FAIL" if any(cr.test_result == "FAIL" for cr in chapter_results) else "PASS"
    return verdict


def _fill_statistics(doc, eng, total, passing, failing, reference,
                     lang: str = "en") -> None:
    """更新汇总页章节内的统计段落（含总数/通过/失败/参考与结论着色）。

    保真要点（对照真实模板 报告模板.docx 的 P027）：
      - 统计段落为**多 run**结构：计数部分散落多个 run，判定词 ``Fail``/``Pass``
        是段末独立红色/黑色 run。故不能整段 replace（会丢红色），也不能依赖跨 run
        正则匹配（python-docx 单 run 内匹配，跨 run 失效）。
      - 采用**段级 run 重建**：拆为「计数部分（黑色）+ 判定词（红/黑）」两段 run，
        既保证计数替换一定生效，又完整保留判定词着色。

    双语：定位关键词、判定结论引导语、判定词均按 ``lang`` 从 ``language_pack`` 取。
      英文 → "total of" / "judged as: " / ``Fail`` ``Pass``
      中文 → "本次共测试" / "根据以下的检验依据，判定本次试验样品：" / ``不合格`` ``合格``
    颜色逻辑不变（有不合格项 → 判定词标红）。
    """
    L = language_pack.labels(lang)
    # 候选关键词：先用语言包指定词，再兜底尝试中/英常见变体
    _kw_candidates = [
        L["summary_total_of"],
        "本次共测试",
        "total of",
        "A total of",
        "a total of",
    ]
    try:
        para = None
        for _kw in _kw_candidates:
            para = eng.find_paragraph_by_text(doc, _kw)
            if para is not None:
                warning(f"[fill_statistics] 找到统计段落（关键词={_kw!r}）：{(para.text or '').strip()!r}")
                break
        if para is None:
            warning("[fill_statistics] 未找到统计段落，跳过统计更新")
            return
        full = para.text or ""

        # 计数部分：取判定引导语之前文本，按出现顺序替换 4 个计数括号
        # （total / passing / failing / reference）。模板括号为全角空格且可能为空，
        # 故用统一正则匹配"任意空白+括号"，仅在括号内填入/替换数字，保留原全角空格。
        lead_in = L["summary_judged_as"]
        split_pat = re.escape(lead_in.rstrip("：:")) + r"[：:]?"
        head = re.split(split_pat, full, maxsplit=1)[0]
        nums = [total, passing, failing, reference]
        it = iter(nums)

        def _fill_one(m):
            # 直接用迭代器中的数字替换整个括号内内容（空括号或含数字均覆盖）。
            # 注意：python-docx 的 para.text 会把全角空格压缩为半角，空括号 "(  )"
            # 读出来可能只剩一个空格，故不能用 \d+ 子替换，必须整体替换 inner。
            # 重建时保留 "( N )" 的半角空格排版，贴合原模板视觉。
            return "( " + str(next(it)) + " )"

        # \s 在 Python3 默认含全角空格(\u3000)；括号内允许为空或含数字
        head = re.sub(r"[\(（]\s*[\d\s]*\s*[\)）]", _fill_one, head)

        verdict_red = failing > 0
        verdict_word = L["verdict_fail"] if verdict_red else L["verdict_pass"]
        verdict_color = 0xFF0000 if verdict_red else 0x000000

        # 重建为两段 run：计数（黑）+ 判定词（红/黑）
        eng.set_paragraph_runs(para, [
            (head.rstrip() + " " + lead_in + " ", None),
            (verdict_word, verdict_color),
        ])
        # 所有 run 强制设置字体（Times New Roman 四号 14pt），判定词保留原有颜色
        from docx.shared import Pt
        for run in para.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
        info(f"[fill_summary] 统计段落已更新（lang={lang}）")
    except Exception as e:
        warning(f"[fill_summary] 统计段落更新失败：{e}")


# 封面日期标签（中英文）——含这些标签的段落/表格不参与"过期标红"，
# 避免收样日期/检测日期/报告日期等历史日期被误判为"过期校准"
_COVER_DATE_LABELS = (
    "Received Date", "Inspection Date", "Report Date",
    "收样日期", "检测日期", "报告日期",
)


def _has_cover_date_label(text: str) -> bool:
    t = text or ""
    return any(lbl in t for lbl in _COVER_DATE_LABELS)


def highlight_expired_calibrations(doc) -> None:
    """扫描全文校准有效期并标红过期项。

    - 跳过含封面日期标签的段落/表格（收样/检测/报告日期不是校准有效期）
    - 其余段落与表格照常扫描
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return
    today = datetime.date.today()

    for para in eng.iter_paragraphs(doc):
        if _has_cover_date_label(para.text):
            continue
        eng.color_expired_in_paragraph(para, today)

    for tbl in eng.iter_tables(doc):
        try:
            has_label = any(_has_cover_date_label(c.text)
                            for row in tbl.rows for c in row.cells)
        except Exception:
            has_label = False
        if has_label:
            continue
        eng.color_expired_in_table(tbl, today)

    info("[highlight_expired_calibrations] 校准有效期扫描完成")


# ───────────────────────── 旧兼容（M7 移除） ─────────────────────────

def update_header_footer_page_numbers(doc) -> None:
    """v5.0 页码由 Word 打开时自动更新域，无需 COM 更新，留空实现兼容旧调用。"""
    return None


def build_section_index(doc) -> None:
    return None


def invalidate_section_index(doc) -> None:
    return None


def _com_call_with_retry(fn, *, retries=6, delay=0.5):
    """win32com 重试原语（v5.0 不再需要，保留签名以便旧代码导入）。"""
    return fn()


def localize_equipment_table(doc, section_name, lang: str = "zh") -> bool:
    """将复制过来的设备表本地化（中文报告专用，RoHS 章节）。
    源报告设备表为英文，复制到中文报告后：
      1. 表头改写为中文（equip_name / equip_model / equip_no / equip_cal）；
      2. 设备名称列按「设备编号」反查设备字典，写入中文名称。
    英文报告不调用本函数。
    标准列序：设备名称 | 型号 | 设备编号 | 校准日期。
    Args:
        doc: Word 文档对象。
        section_name: 章节标题（如 "RoHS"）。
        lang: 目标语言；仅 "zh" 时执行本地化。
    Returns:
        True 处理完成（含跳过），False 定位失败。
    """
    doc = _doc(doc)
    eng = _eng()
    if doc is None:
        return False
    if language_pack.normalize(lang) != "zh":
        return True
    L = language_pack.labels("zh")
    section = eng.find_section(doc, section_name)
    if section is None:
        warning(f"[localize_equipment_table] 未找到章节 {section_name}")
        return False
    tbl = eng.find_table(doc, section, L["equip_title"])
    if tbl is None:
        warning(f"[localize_equipment_table] 章节 {section_name} 未找到设备表")
        return False
    row_count = eng.table_row_count(tbl)
    col_count = eng.table_column_count(tbl)
    if row_count == 0 or col_count == 0:
        warning(f"[localize_equipment_table] 设备表为空：{section_name}")
        return False
    # ── 列定位：优先英文表头，其次中文表头，最后按标准列序兜底 ──
    _EN_TO_FIELD = {
        "equipment name": "name",
        "equipment model": "model",
        "equipment no": "no",
        "calibration expiration date": "cal",
    }
    _ZH_TO_FIELD = {
        (L["equip_name"] or "").strip(): "name",
        (L["equip_model"] or "").strip(): "model",
        (L["equip_no"] or "").strip(): "no",
        (L["equip_cal"] or "").strip(): "cal",
    }
    _ZH_TO_FIELD = {k: v for k, v in _ZH_TO_FIELD.items() if k}
    col_map = {}
    for c in range(col_count):
        cur = (eng.table_cell_text(tbl, 0, c) or "").strip()
        cur_low = cur.lower()
        hit = False
        for en_key, field in _EN_TO_FIELD.items():
            if en_key in cur_low:
                col_map.setdefault(field, c)
                hit = True
                break
        if hit:
            continue
        for zh_key, field in _ZH_TO_FIELD.items():
            if zh_key in cur:
                col_map.setdefault(field, c)
                break
    # 标准列序兜底：名称 | 型号 | 编号 | 校准日期
    col_map.setdefault("name", 0)
    col_map.setdefault("model", min(1, col_count - 1))
    col_map.setdefault("no", min(2, col_count - 1))
    col_map.setdefault("cal", min(3, col_count - 1))
    name_col = col_map["name"]
    no_col = col_map["no"]
    # ── 1. 表头翻译（保留原格式） ──
    _zh_headers = [L["equip_name"], L["equip_model"], L["equip_no"], L["equip_cal"]]
    try:
        for c in range(min(col_count, len(_zh_headers))):
            want = _zh_headers[c]
            if not want:
                continue
            cur = (eng.table_cell_text(tbl, 0, c) or "").strip()
            if cur != want:
                eng.set_cell_text_preserve(tbl, 0, c, want, align="center")
    except Exception as e:
        warning(f"[localize_equipment_table] 表头翻译失败：{e}")
    # ── 2. 设备名称中文化 ──
    changed = 0
    for r in range(1, row_count):
        try:
            asset_id = (eng.table_cell_text(tbl, r, no_col) or "").strip()
        except Exception:
            continue
        if not asset_id or asset_id == "/":
            continue
        detail = reference_service.get_equipment_detail(asset_id, lang="zh")
        zh_name = (detail or {}).get("name", "") or ""
        if zh_name and zh_name != "/":
            cur_name = (eng.table_cell_text(tbl, r, name_col) or "").strip()
            if cur_name != zh_name:
                eng.set_cell_text_preserve(tbl, r, name_col, zh_name, align="center")
                changed += 1
        else:
            warning(f"[localize_equipment_table] 设备名称中文化失败，保留原值：{asset_id}")
    info(f"[localize_equipment_table] {section_name} 设备表本地化完成，改写设备名 {changed} 个")
    return True


def insert_empty_para_after(doc, anchor_para) -> object:
    """在 anchor_para 后插入一个空段落，返回该段落对象。

    EDS 策略用于在子文件夹内容之间插入空行。
    """
    doc = _doc(doc)
    if doc is None or anchor_para is None:
        return None
    try:
        from docx.oxml import OxmlElement as _OE
        new_p = _OE("w:p")
        anchor_para._p.addnext(new_p)
        from docx.text.paragraph import Paragraph
        return Paragraph(new_p, doc)
    except Exception as e:
        warning(f"[insert_empty_para_after] 失败：{e}")
        return None


def insert_image_after_para(anchor_para, img_path: str, width_cm: float) -> bool:
    """在 anchor_para 后直接插入一张图片（不建表格）。

    EDS 策略用于插入 Excel 转出的 PNG 图。
    """
    doc = _doc(None)
    eng = _eng()
    if doc is None or anchor_para is None:
        return False
    try:
        data = _read_image_bytes(img_path)
    except Exception as e:
        warning(f"[insert_image_after_para] 读取图片失败 {img_path}：{e}")
        return False
    return eng.insert_image(doc, anchor_para, data, width_cm)


def insert_word_table_after_para(anchor_para, table_data, width_cm: float = 18.0) -> bool:
    """在指定段落后插入 Word 表格（不按文字搜索锚点）。

    用于 EDS / Electrical 等需要连续插入表格的场景。
    表格整体居中，宽度 width_cm，列宽由 docx_engine.insert_table 分配。

    Args:
        anchor_para: 锚点段落（表格会插到它正后方）。
        table_data: excel_table_builder.build_table_data 返回的结构化数据。
        width_cm: 表格总宽度（厘米）。

    Returns:
        True 成功，False 失败。
    """
    doc = _doc(None)
    eng = _eng()
    if doc is None or anchor_para is None:
        warning("[insert_word_table_after_para] 文档或锚点为空")
        return False
    try:
        return eng.insert_table(doc, anchor_para, table_data, width_cm)
    except Exception as e:
        warning(f"[insert_word_table_after_para] 插入失败：{e}")
        return False

def insert_excel_images_after_table(anchor_para, images: list,
                                     width_cm: float = 7.0) -> bool:
    """在 anchor_para 后的表格下方，插入一个"图片框表格"。

    规则：
      - 表格固定 2 列（各 9cm，总宽 18cm）
      - 每行最多 2 张图
      - 1 张图独占一行（合并 2 列）
      - 每张图宽度 = min(width_cm, 单元格宽 - 0.5cm)
      - 图片居中，边框细实线
    """
    import math
    from io import BytesIO
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.table import Table as _Table
    from docx.shared import Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from services.docx_helpers import _ensure_table_borders

    doc = _doc(None)
    if doc is None or anchor_para is None or not images:
        return False

    # 定位 anchor_para 后的第一个表格
    tbl_el = anchor_para._p.getnext()
    if tbl_el is None or not tbl_el.tag.endswith("}tbl"):
        warning("[insert_excel_images_after_table] 锚点后未找到表格")
        return False

    n = len(images)
    n_rows = math.ceil(n / 2)
    ncols = 2
    _DXA = 567
    total_dxa = int(18.0 * _DXA)
    col_dxa = int(9.0 * _DXA)

    # 建空表格
    new_tbl_el = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    style = OxmlElement("w:tblStyle")
    style.set(qn("w:val"), "TableGrid")
    tblPr.append(style)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    new_tbl_el.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for _ in range(ncols):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(col_dxa))
        grid.append(gc)
    new_tbl_el.append(grid)
    for _ in range(n_rows):
        tr = OxmlElement("w:tr")
        for _ in range(ncols):
            tc = OxmlElement("w:tc")
            tcPr = OxmlElement("w:tcPr")
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_dxa))
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)
            tc.append(tcPr)
            tc.append(OxmlElement("w:p"))
            tr.append(tc)
        new_tbl_el.append(tr)

    _ensure_table_borders(new_tbl_el)
    # 先在数据表格后插一个空段落，再插图片表格——防止 Word 合并两张表
    from docx.oxml import OxmlElement as _OE
    spacer = _OE("w:p")
    tbl_el.addnext(spacer)
    spacer.addnext(new_tbl_el)
    new_tbl = _Table(new_tbl_el, doc)

    def _fill_img(row, col, img_data, cell_width_cm):
        cell = new_tbl.cell(row, col)
        p = cell.paragraphs[0]
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            target_w = min(width_cm, cell_width_cm - 0.5)
            if target_w <= 0:
                target_w = 1.0

            # 限高：先按宽度算高度，若超 8cm 则改按高度封顶等比回推宽度
            target_h = None
            try:
                from PIL import Image as _PILImage
                with _PILImage.open(BytesIO(img_data)) as _im:
                    w_px, h_px = _im.size
                if w_px > 0:
                    ratio = h_px / w_px               # 高/宽
                    h = target_w * ratio
                    if h > _MAX_IMG_H_CM:
                        target_h = _MAX_IMG_H_CM
                        target_w = _MAX_IMG_H_CM / ratio
            except Exception:
                pass   # 读尺寸失败则退回只按宽度

            run = p.add_run()
            if target_h is not None:
                run.add_picture(BytesIO(img_data),
                                width=Cm(target_w), height=Cm(target_h))
            else:
                run.add_picture(BytesIO(img_data), width=Cm(target_w))
        except Exception as e:
            warning(f"[insert_excel_images_after_table] 插图失败：{e}")
            p.add_run("/")

    if n == 1:
        # 独占一行
        try:
            new_tbl.cell(0, 0).merge(new_tbl.cell(0, 1))
        except Exception:
            pass
        data = images[0].get("data")
        if data:
            _fill_img(0, 0, data, 18.0)
    else:
        for i, img in enumerate(images):
            r = i // 2
            c = i % 2
            data = img.get("data")
            if not data:
                continue
            _fill_img(r, c, data, 9.0)

    info(f"[insert_excel_images_after_table] 已插入 {n} 张 Excel 图片")
    return True

def insert_page_break_after(doc, anchor_para) -> object:
    """在 anchor_para 后插入一个含分页符的段落，返回该段落对象。

    EDS 策略用于让每个子文件夹内容单独成页。
    """
    doc = _doc(doc)
    if doc is None or anchor_para is None:
        return None
    try:
        from docx.oxml import OxmlElement as _OE
        from docx.oxml.ns import qn as _qn
        from docx.text.paragraph import Paragraph
        new_p = _OE("w:p")
        new_r = _OE("w:r")
        new_br = _OE("w:br")
        new_br.set(_qn("w:type"), "page")
        new_r.append(new_br)
        new_p.append(new_r)
        anchor_para._p.addnext(new_p)
        return Paragraph(new_p, doc)
    except Exception as e:
        warning(f"[insert_page_break_after] 失败：{e}")
        return None