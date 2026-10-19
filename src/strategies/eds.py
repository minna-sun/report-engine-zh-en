"""EDS 策略（模块新增）。

职责：处理 EDS 章节，按子文件夹顺序插入图片表格 + Excel 渲染图，
      最后在最顶部插入汇总（SUM）Excel 渲染图。

文件夹结构（位于 ``原始数据/{folder_name}/图片输出/``）：
  - 顶层：总汇总.xlsx（文件名含"总汇总"）
  - 子文件夹（每个代表一组测试条件）：
      - 图片文件（.jpg/.jpeg/.png/.bmp/.gif）→ 图片表格
      - Excel 文件（.xlsx）→ 渲染为 PNG 后插入

最终章节内容顺序（紧随 Result description 段落之后）：
  SUM图 | 空行 | 条件1图片表 | 条件1Excel图 | 空行 | 条件2图片表 | 条件2Excel图 | ...
  | 条件N图片表 | 条件NExcel图（最后一组无空行）

插入机制：在 Result description 段落上反复调用 addnext，
倒序处理子文件夹，使最终顺序与子文件夹自然顺序一致。

边界：
- 不导入 win32com。
- 不负责章节分页（由 GenericProcessor 统一处理）。
- 无图片的子文件夹跳过图片表格（只插 Excel 图，若 Excel 也无则跳过整个子文件夹）。
- 图片名称去掉扩展名后作为标签显示。
- "取点图"优先：每个子文件夹里，文件名含"取点图"的图片排第一位；
  没有取点图时，最前面补一个 ("", "") 空占位（第一格留空白）。
"""

import os
import re
from typing import Any, List, Optional, Tuple

from common.logger import info, warning
import excel_table_builder
from models.result import ChapterResult
from services import language_pack
from strategies.base import BaseStrategy
from template_engine import (
    _doc,
    insert_empty_para_after,
    insert_image_after_para,
    insert_image_auto_table,
    insert_page_break_after,
    insert_word_table_after_para,
)
from strategies.image_auto import (
    _extract_width_from_name,
    _clean_name_for_display,
    _split_sort_prefix_and_display,
    _sort_key_with_marker,
)


# ── 常量 ──────────────────────────────────────────────────────────────────────
_IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}
_SKIP_PREFIXES = ("~$",)
_SKIP_KEYWORDS = ("_processed", "thumbs.db", "desktop.ini")
_EXCEL_PNG_WIDTH_CM = 18.0   # Excel 渲染 PNG 固定宽度（同电性策略）
_QD_KEYWORD = "取点图"        # 取点图关键字


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def _natural_sort_key(name: str):
    parts = re.split(r"(\d+)", name)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def _is_valid_file(name: str) -> bool:
    low = name.lower()
    if any(low.startswith(p) for p in _SKIP_PREFIXES):
        return False
    if any(k in low for k in _SKIP_KEYWORDS):
        return False
    return True


def _scan_subfolder(sub_path: str, default_width_cm: float = 7.0):
    """扫描子文件夹，返回 (image_list, xlsx_path)。

    image_list 规则（v5.9 更新）：
      - 每项为 3 元组：(显示名称, 完整路径, 宽度)
      - 宽度提取：文件名最后一个括号 > default_width_cm
      - 显示名称：去掉文件名最后一个括号及内容
      - 含"取点图"的图片排第一位
      - 没有"取点图"时，最前面补一个空占位 ("", "", 0.0)

    xlsx_path: 第一个 .xlsx 文件的完整路径，或 None
    """
    all_imgs = []   # [(display_name, path, width), ...]
    xlsx = None
    try:
        entries = sorted(
            os.listdir(sub_path),
            key=lambda n: _sort_key_with_marker(os.path.splitext(n)[0])
        )
    except OSError:
        return [], None

    for name in entries:
        if not _is_valid_file(name):
            continue
        full = os.path.join(sub_path, name)
        if not os.path.isfile(full):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext in _IMG_EXTS:
            raw_label = os.path.splitext(name)[0]
            _, label_after_marker = _split_sort_prefix_and_display(raw_label)
            display_name = _clean_name_for_display(label_after_marker)
            width = _extract_width_from_name(raw_label, default_width_cm)
            all_imgs.append((display_name, full, width))
        elif ext == ".xlsx" and xlsx is None:
            xlsx = full

    # 取点图优先 + 无取点图时补空占位
    qd = [item for item in all_imgs if _QD_KEYWORD in item[0]]
    rest = [item for item in all_imgs if _QD_KEYWORD not in item[0]]
    if qd:
        imgs = [qd[0]] + rest
    elif rest:
        imgs = [("", "", 0.0)] + rest   # 无取点图：第一格空占位（3 元组）
    else:
        imgs = []
    return imgs, xlsx

def _find_sum_xlsx(folder: str) -> Optional[str]:
    """在目录下查找文件名含「总汇总」的 xlsx。

    匹配示例：总汇总.xlsx、总汇总-EDS.xlsx
    """
    try:
        for name in sorted(os.listdir(folder), key=_natural_sort_key):
            if not _is_valid_file(name):
                continue
            if "总汇总" in name and name.lower().endswith(".xlsx"):
                return os.path.join(folder, name)
    except OSError:
        pass
    return None

def _localize_point_label(name: str, lang: str) -> str:
    """英文报告下把显示名中的"取点图"替换为语言包里的 Test Point。

    - 中文报告 / 名称不含"取点图" / 语言包取不到词 → 原样返回
    - 两侧补空格再压缩，避免 '1-1取点图' 变成 '1-1Test Point'
    """
    if lang != "en" or _QD_KEYWORD not in name:
        return name
    en_label = language_pack.pick(lang, "chapter_test_point")
    if not en_label:
        return name
    return re.sub(r"\s+", " ", name.replace(_QD_KEYWORD, f" {en_label} ")).strip()

# ── 核心处理逻辑 ──────────────────────────────────────────────────────────────

def _process_eds_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    doc = _doc(None)
    config = chapter_ctx.config
    test_item = config.test_item
    source_folder = chapter_ctx.source_folder_name or config.folder_name
    width_cm = config.image_width or 9.0

    base = ChapterResult(
        test_item=test_item, test_basis="", sample_number="",
        test_result="", status="success",
    )

    if not source_folder:
        base.status = "skipped"
        base.error = "source_folder 为空"
        warning(f"[EDS] {test_item}：source_folder 为空，跳过")
        return base

    eds_dir = os.path.join(task_ctx.folder_path, "原始记录", source_folder)
    if not os.path.isdir(eds_dir):
        base.status = "skipped"
        base.error = f"EDS 目录不存在：{eds_dir}"
        warning(f"[EDS] {test_item}：{base.error}")
        return base

    # ── 1. 找章节标题段落，再找 Result description 段落作为 anchor ──
    from template_engine import _eng
    eng = _eng()
    section_para = eng.find_section(doc, test_item)
    if section_para is None:
        base.status = "failed"
        base.error = f"未找到章节标题段落：{test_item}"
        warning(f"[EDS] {test_item}：{base.error}")
        return base

    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    result_desc_label = language_pack.labels(lang)["chapter_result_description"]
    anchor = eng.find_paragraph(doc, section_para, result_desc_label)
    if anchor is None:
        for _label in ("Result description", "结果说明"):
            anchor = eng.find_paragraph(doc, section_para, _label)
            if anchor:
                break
    if anchor is None:
        base.status = "failed"
        base.error = f"未找到 Result description 段落：{test_item}"
        warning(f"[EDS] {test_item}：{base.error}")
        return base

    # ── 2. 定位「图片输出」目录并收集其子文件夹（自然排序） ──
    img_output_dir = os.path.join(eds_dir, "图片输出")
    if not os.path.isdir(img_output_dir):
        base.status = "failed"
        base.error = f"图片输出目录不存在：{img_output_dir}"
        warning(f"[EDS] {test_item}：{base.error}")
        return base

    try:
        all_entries = sorted(os.listdir(img_output_dir), key=_natural_sort_key)
    except OSError as e:
        base.status = "failed"
        base.error = f"读取图片输出目录失败：{e}"
        warning(f"[EDS] {test_item}：{base.error}")
        return base

    sub_dirs = [
        name for name in all_entries
        if os.path.isdir(os.path.join(img_output_dir, name)) and _is_valid_file(name)
    ]

    # ── 3. 查找总汇总 xlsx ──
    sum_xlsx = _find_sum_xlsx(img_output_dir)
    if sum_xlsx is None:
        warning(f"[EDS] {test_item}：未在图片输出目录找到「总汇总」xlsx，跳过 SUM 图插入")

       # ── 4. 倒序插入各子文件夹内容（addnext 语义） ──
    #    最终每个子文件夹内容单独一页：分页符 | 图片表 | 空行 | Excel表
    #    （倒序时：先插分页符 → 再插 Excel → 再插空行 → 最后插图片表，
    #      后插的会挤到前面，最终顺序正确。）
    inserted_count = 0
    for sub_name in reversed(sub_dirs):
        sub_path = os.path.join(img_output_dir, sub_name)
        imgs, xlsx_path = _scan_subfolder(sub_path, default_width_cm=width_cm)

        if imgs and lang == "en":
            imgs = [(_localize_point_label(n, lang), p, w) for (n, p, w) in imgs]

        has_imgs = bool(imgs)
        has_xlsx = bool(xlsx_path)

        if not has_imgs and not has_xlsx:
            warning(f"[EDS] {test_item}/{sub_name}：无图片也无 Excel，跳过")
            continue

        is_last_in_order = (sub_name == sub_dirs[-1])

        # (A) 非正序最后一组：插分页符（最终在本组内容之后）
        if not is_last_in_order:
            insert_page_break_after(doc, anchor)

        # (B) 插入 Excel Word 表格（固定宽度 18cm）
        if has_xlsx:
            table_data = excel_table_builder.build_table_data(
                xlsx_path=xlsx_path,
                table_width_cm=18.0,
                skip_columns=(), skip_rows=(),
                require_header=False,
                force_font_name="Times New Roman",
                force_font_size=10.5,
            )
            if table_data:
                ok = insert_word_table_after_para(anchor, table_data, 18.0)
                if not ok:
                    warning(f"[EDS] {test_item}/{sub_name}：Word 表格插入失败")
            else:
                warning(f"[EDS] {test_item}/{sub_name}：表格数据构建失败，跳过")

        # (C) 图片表和 Excel 表之间插空行（两者都有的情况）
        if has_imgs and has_xlsx:
            insert_empty_para_after(doc, anchor)

        # (D) 插入图片表格（最终排在本组最前面）
        if has_imgs:
            ok = insert_image_auto_table(
                doc, test_item, imgs, width_cm,
                anchor_para=anchor, blank_if_empty=True,
            )
            if not ok:
                warning(f"[EDS] {test_item}/{sub_name}：图片表格插入失败")

        inserted_count += 1
        info(f"[EDS] {test_item}/{sub_name}：已插入（图片{len(imgs)}张，Excel{'有' if has_xlsx else '无'}）")
# ── 5. 在最顶部插入 SUM 表格（固定宽度 18cm，紧跟 Result description，不换页） ──
    if sum_xlsx:
        table_data = excel_table_builder.build_table_data(
            xlsx_path=sum_xlsx,
            table_width_cm=18.0,
            skip_columns=(), skip_rows=(),
            require_header=False,
            force_font_name="Times New Roman",
            force_font_size=10.5,
        )
        if table_data:
            # 有子文件夹时，SUM 之后插分页符（保证第一个子文件夹独立成页）
            if sub_dirs:
                insert_page_break_after(doc, anchor)
            ok = insert_word_table_after_para(anchor, table_data, 18.0)
            if not ok:
                warning(f"[EDS] {test_item}：SUM 表格插入失败")
            else:
                info(f"[EDS] {test_item}：SUM 表格插入成功")
        else:
            warning(f"[EDS] {test_item}：SUM 表格数据构建失败")

    if inserted_count == 0 and sum_xlsx is None:
        base.status = "skipped"
        base.error = "无有效子文件夹数据且无 SUM 文件"
        warning(f"[EDS] {test_item}：{base.error}")

    return base


# ── 策略类 ────────────────────────────────────────────────────────────────────

class EDSStrategy(BaseStrategy):
    """EDS 策略——多子文件夹图片 + Excel 渲染图插入。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        try:
            result = _process_eds_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[EDSStrategy] 执行异常 ({test_item}): {e}")
            return self._finalize(
                ChapterResult(
                    test_item=test_item, test_basis="", sample_number="",
                    test_result="", status="failed", error=str(e),
                ),
                chapter_ctx, task_ctx,
            )