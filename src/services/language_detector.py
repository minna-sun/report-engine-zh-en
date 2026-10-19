"""模板语言自动识别服务（双语支持 · 优先级 2）。

职责：在物料信息表未指定「报告语言」时，通过 Word 模板封面关键词判定报告语言。

边界：
- 只做识别，不做决策（决策优先级由 ``main`` 编排）。
- 识别失败 / 平局无中文特征 → 默认英文（spec 优先级 3）。
- 不修改文档、不写配置。

技术栈：python-docx（严禁 win32com / COM）。
"""

from docx import Document

from common.logger import warning
from services.language_pack import ZH, EN, DEFAULT

# 英文特征词（封面页）
_EN_KEYWORDS = [
    "Test Report", "Applicant", "Received Date",
    "Sample Description", "Summary of Results",
]
# 中文特征词（封面页）
_ZH_KEYWORDS = [
    "测试报告", "检测报告", "委托人", "收样日期",
    "样品描述", "结果汇总",
]

# 封面页采样范围
_COVER_PARA_LIMIT = 10      # 前 N 个非空段落
_FALLBACK_TABLES = 3        # 正文无特征词时，补扫前 N 张表格
_FALLBACK_ROWS = 12         # 每张表格补扫的前 N 行


def _cover_texts(doc) -> list:
    """取封面页候选文本：正文前 N 个非空段落；全无特征词时补扫表格首行文本。"""
    texts = []
    for para in doc.paragraphs[:_COVER_PARA_LIMIT]:
        text = (para.text or "").strip()
        if text:
            texts.append(text)
    return texts


def _table_texts(doc) -> list:
    """兜底：取前 N 张表格的前 N 行单元格文本（封面信息常置于表格内）。"""
    texts = []
    for tbl in doc.tables[:_FALLBACK_TABLES]:
        for row in tbl.rows[:_FALLBACK_ROWS]:
            for cell in row.cells:
                ct = (cell.text or "").strip()
                if ct:
                    texts.append(ct)
    return texts


def _score(texts) -> tuple:
    """统计中英文特征词命中次数，返回 (en_score, zh_score)。"""
    full_text = " ".join(texts)
    en_score = sum(1 for kw in _EN_KEYWORDS if kw in full_text)
    zh_score = sum(1 for kw in _ZH_KEYWORDS if kw in full_text)
    return en_score, zh_score


def detect_language(doc_path: str) -> str:
    """识别 Word 模板的语言。

    Args:
        doc_path: 模板文件路径（.docx）。

    Returns:
        ``"zh"`` 或 ``"en"``；识别失败或特征不足时返回 ``"en"``。
    """
    try:
        doc = Document(doc_path)

        # 取前 N 个非空段落（封面页）
        texts = _cover_texts(doc)
        en_score, zh_score = _score(texts)

        # 正文无特征词（封面信息在表格内时）→ 补扫表格，避免漏判
        if en_score == 0 and zh_score == 0:
            texts = _table_texts(doc)
            en_score, zh_score = _score(texts)

        if zh_score > en_score:
            return ZH
        if en_score > zh_score:
            return EN
        # 平局时检查是否有明确中文特征
        if any(kw in " ".join(texts) for kw in _ZH_KEYWORDS):
            return ZH
        return DEFAULT

    except Exception as e:
        warning(f"[language_detector] 识别失败，默认英文: {e}")
        return DEFAULT


__all__ = ["detect_language"]
