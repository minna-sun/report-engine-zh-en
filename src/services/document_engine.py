"""文档引擎抽象接口（v5.0 新增 · spec 模块 9）。

职责：定义文档操作的标准化抽象接口，屏蔽底层技术差异（python-docx / 未来其他引擎）。
所有具体实现（``DocxEngine``）必须遵循本接口定义的契约。

设计要点：
  - 本模块**只定义抽象接口**，不包含任何具体实现。
  - 所有方法失败时返回 ``False`` / ``None``，**不抛出异常**（编程错误除外）。
  - 业务层（Processor / Strategy）不直接依赖本接口；仅 ``template_engine`` 与
    ``report_engine`` 通过 ``engine_factory.get_engine()`` 获取其实例。
  - v5.0 已删除 ``find_bookmark()`` 与 ``update_page_numbers()``，改用配置驱动 + 文本搜索。
"""
from abc import ABC, abstractmethod
from typing import List, Optional, Union

try:
    from docx.document import Document as DocxDocument
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except Exception:  # pragma: no cover - python-docx 缺失时在类型注解处容错
    DocxDocument = "Document"
    Table = "Table"
    Paragraph = "Paragraph"


class DocumentEngine(ABC):
    """文档操作抽象接口。

    契约（行为说明，见 重构.md §6.2）：
      - ``open``        读 .docx，返回文档对象；失败返回 None。
      - ``save``        保存到指定路径；失败返回 False，不抛异常。
      - ``close``       释放资源；python-docx 实现为空操作（接口一致性）。
      - ``find_section``   匹配**加粗**且文本含 section_name 的段落，忽略表格内段落。
      - ``find_paragraph`` 章节内部分匹配 anchor_text 的段落。
      - ``find_table``      锚点段落后第一个 Table 对象。
      - ``insert_text``     锚点段落同行冒号后**替换**文本。
      - ``insert_image``    锚点段落后新建段落插入图片（锁定纵横比、居中）。
      - ``insert_table``    锚点段落后新建段落插入表格（总宽 width_cm，按数据分配列宽）。
      - ``copy_table``      跨文档复制表格（委托 ooxml_service）。
      - ``delete_section_content`` 删除章节内容保留标题。
      - ``add_page_break``  段落后插入分页符。
      - ``set_table_format`` 设置表格总宽并禁用自动调整。
    """

    @abstractmethod
    def open(self, source: Union[str, "object"]) -> Optional[DocxDocument]:
        ...

    @abstractmethod
    def save(self, doc: DocxDocument, path: str) -> bool:
        ...

    @abstractmethod
    def close(self, doc: DocxDocument) -> None:
        ...

    @abstractmethod
    def find_section(self, doc: DocxDocument, section_name: str,
                     require_bold: bool = True) -> Optional[Paragraph]:
        ...

    @abstractmethod
    def find_paragraph(self, doc: DocxDocument, section: Paragraph,
                       anchor_text: str) -> Optional[Paragraph]:
        ...

    @abstractmethod
    def find_table(self, doc: DocxDocument, section: Paragraph,
                   anchor_text: str) -> Optional[Table]:
        ...

    @abstractmethod
    def insert_text(self, doc: DocxDocument, anchor: Paragraph, text: str,
                    font_color=None) -> bool:
        ...

    @abstractmethod
    def insert_image(self, doc: DocxDocument, anchor: Paragraph,
                     image_data: bytes, width_cm: float) -> bool:
        ...

    @abstractmethod
    def replace_paragraph_text(self, para: Paragraph, text: str) -> bool:
        """整段替换段落文本（清空原有 run，首 run 写入新文本）。"""

    @abstractmethod
    def set_paragraph_color(self, para: Paragraph, color) -> bool:
        """设置段落所有 run 的字体颜色（color: #RRGGBB / 整数 RGB）。"""

    @abstractmethod
    def replace_in_paragraph(self, para: Paragraph, pattern, repl,
                            count: int = 0, *, flags: int = 0) -> int:
        """段落内精准替换，**保留各 run 字体/颜色格式**（不整段重排）。

        用于统计段落等需保留「判定词红色 run」的场景。返回实际替换次数。
        ``pattern`` 可为字符串（子串）或已编译 ``re`` 正则；``repl`` 可为字符串或回调。
        """

    @abstractmethod
    def set_paragraph_runs(self, para: Paragraph, runs: list) -> bool:
        """整段重建为给定 run 列表 ``[(text, color_or_None), ...]``（分段着色场景）。

        用于统计段落等需「计数黑 + 判定词红」分段着色的场景。返回 True / False。
        """

    @abstractmethod
    def insert_table(self, doc: DocxDocument, anchor: Paragraph,
                     table_data, width_cm: float) -> bool:
        ...

    @abstractmethod
    def copy_table(self, doc: DocxDocument, section: str,
                   source_doc: DocxDocument, source_anchor: str,
                   target_anchor: str) -> bool:
        ...

    @abstractmethod
    def delete_section_content(self, doc: DocxDocument, section_name: str) -> bool:
        ...

    @abstractmethod
    def add_page_break(self, doc: DocxDocument, para: Paragraph) -> None:
        ...

    @abstractmethod
    def set_table_format(self, table: Table, width_cm: float) -> None:
        ...

    @abstractmethod
    def find_section_range(self, doc: DocxDocument, section_name: str):
        """返回 ``(start_para, end_para)``：章节标题段落与下一个章节标题段落。

        ``end_para`` 为 ``None`` 表示到文档末尾。找不到章节时返回 ``(None, None)``。
        """
        ...

    @abstractmethod
    def delete_section(self, doc: DocxDocument, section_name: str) -> bool:
        """删除整个章节（含标题与全部内容）。返回 ``True`` / ``False``。"""
        ...

    @abstractmethod
    def clone_section(self, doc: DocxDocument, template_section_name: str,
                      new_title: str, insert_before_section_name: Optional[str] = None) -> bool:
        """克隆模板章节结构、重命名标题、插入到指定章节之前（或文末）。

        插入位置由 ``insert_before_section_name`` 决定（``None`` 表示文末）。
        返回 ``True`` / ``False``。
        """
        ...

    @abstractmethod
    def add_table_row(self, table):
        """在表格末尾追加一行，返回新行（python-docx Row）或 ``None``。"""
        ...

    @abstractmethod
    def remove_table_row(self, table, index: int) -> bool:
        """删除表格第 ``index`` 行（0-based）。返回 ``True`` / ``False``。"""
        ...

    @abstractmethod
    def table_row_count(self, table) -> int:
        """返回表格行数。"""
        ...

    @abstractmethod
    def set_cell_text(self, table, row: int, col: int, text: str,
                      color=None) -> bool:
        """设置表格单元格 ``(row, col)`` 的文本；可选 color 着色首 run。"""
        ...

    @abstractmethod
    def table_cell_text(self, table, row: int, col: int) -> Optional[str]:
        """返回表格单元格 ``(row, col)`` 的文本；异常时返回 None。"""

    @abstractmethod
    def table_column_count(self, table) -> int:
        """返回表格列数。"""

    @abstractmethod
    def set_text_color(self, run, color) -> None:
        """设置 run 字体颜色（color: #RRGGBB / 整数 RGB）。"""
        ...

    @abstractmethod
    def list_sections(self, doc: DocxDocument):
        """返回 ``[(文本, 段落句柄), ...]``：文档中所有测试章节标题（按文档顺序）。"""
        ...

    @abstractmethod
    def find_paragraph_by_text(self, doc: DocxDocument, text: str,
                               exact: bool = False) -> Optional[Paragraph]:
        """全局搜索含 text（或文本等于 text）的段落，返回首个匹配句柄。"""
        ...

    @abstractmethod
    def has_page_break_before(self, para: Paragraph) -> bool:
        """判断段落前是否已存在分页符。"""
        ...

    @abstractmethod
    def iter_paragraphs(self, doc: DocxDocument):
        """返回文档所有正文段落句柄列表。"""
        ...

    @abstractmethod
    def iter_tables(self, doc: DocxDocument):
        """返回文档所有表格句柄列表。"""
        ...

    @abstractmethod
    def color_expired_in_paragraph(self, para, today=None) -> None:
        """扫描段落中的日期，对过期日期（区间只标截止日期）标红。"""
        ...

    @abstractmethod
    def color_expired_in_table(self, table, today=None) -> None:
        """扫描表格所有单元格中的日期并标红过期项。"""
        ...

    @abstractmethod
    def color_expired_in_cell(self, table, row: int, col: int, today=None) -> None:
        """扫描指定单元格中的日期并标红过期项。"""
        ...

    @abstractmethod
    def insert_text_in_header(self, doc: DocxDocument, label: str, value: str) -> bool:
        """在页眉（Primary）中搜索 label，保留页码域，在其后填入 value。"""
        ...
