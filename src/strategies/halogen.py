"""卤素策略（新增 · 2026-09-29）。

职责：从卤素源报告（docx）复制"Sample Test Result："下方第一个表格，
      插入到目标报告卤素章节的 "Result description" / "结果说明" 段落下方。

与 ComplianceStrategy 的差异：
- 源报告定位：不依赖章节标题，全局搜加粗段落 "Sample Test Result"。
- 只复制一个表格（不复制设备表）。
- 中文报告需翻译（英文报告直接拷）。
- 本地目录名：原始数据/卤素/；config 字段：halogen_source_base。

边界：
- 不导入 python-docx / lxml / DocumentEngine。
- 不执行设备匹配。
- 不处理 Test Standards / Test Result / sample number（由 GenericProcessor 负责）。
- 不负责章节分页（由 GenericProcessor 统一处理）。
"""

import os
import shutil
from typing import Any

from strategies.base import BaseStrategy
from common.logger import info, warning
from services import reference_service
from services import language_pack
from template_engine import _doc, _eng, copy_table_by_source_para_from_doc
from models.result import ChapterResult


# 源报告锚点：加粗段落文本（不含冒号，匹配时归一化处理）
_SRC_PARA_TEXT = "Sample Test Result"


# ───────────────────────── 源报告查找（本地优先 → 远程 fallback） ─────────────────────────

def _resolve_halogen_source_report(task_ctx, material_code: str):
    """解析卤素源报告路径（本地优先 → 远程兜底）。"""
    folder_path = task_ctx.folder_path
    mc_low = (material_code or "").strip().lower()

    # ── 优先级 1：本地卤素文件夹 ──
    local_dir = reference_service.find_actual_folder(folder_path, "卤素")
    if local_dir:
        local_report = _find_docx_in_dir(local_dir, mc_low)
        if local_report:
            info(f"[Halogen] 优先使用本地卤素文件夹: {local_dir}")
            return local_report

    # ── 优先级 2（fallback）：远程按料号拉取 ──
    cfg = task_ctx.config or {}
    source_base = cfg.get("halogen_source_base") if isinstance(cfg, dict) else None
    if not source_base:
        warning("[Halogen] 未配置 halogen_source_base，且本地无卤素文件夹，无法拉取源报告")
        return None
    src_path = _find_source_report(source_base, mc_low)
    if not src_path:
        warning(f"[Halogen] 在 halogen_source_base 下未找到料号 {material_code} 的卤素报告，跳过")
        return None

    # 拷贝到本地
    target_dir = local_dir if local_dir else os.path.join(folder_path, "原始记录", "卤素")
    try:
        os.makedirs(target_dir, exist_ok=True)
    except Exception as e:
        warning(f"[Halogen] 创建本地目录失败（{target_dir}）: {e}")
        return None
    local_path = os.path.join(target_dir, os.path.basename(src_path))
    try:
        shutil.copy2(src_path, local_path)
    except Exception as e:
        warning(f"[Halogen] 拷贝卤素源报告失败（{src_path}）: {e}")
        return None
    info(f"[Halogen] 已从远程定位并拷贝卤素源报告到本地: {local_path}")
    return local_path


def _find_docx_in_dir(dir_path: str, material_code_low: str):
    """在 dir_path 下查找 .docx：优先文件名含料号，否则取最新修改者。"""
    if not os.path.isdir(dir_path):
        return None
    docs = [f for f in os.listdir(dir_path)
            if f.lower().endswith(".docx") and not f.startswith("~$")]
    if not docs:
        return None
    if material_code_low:
        matched = [f for f in docs if material_code_low in f.lower()]
        if matched:
            return _newest(os.path.join(dir_path, f) for f in matched)
    return _newest(os.path.join(dir_path, f) for f in docs)


def _newest(paths):
    """返回 paths 中修改时间最新的路径；均不可读返回 None。"""
    best, best_mtime = None, None
    for p in paths:
        try:
            m = os.path.getmtime(p)
        except Exception:
            continue
        if best_mtime is None or m > best_mtime:
            best, best_mtime = p, m
    return best


def _find_source_report(base: str, material_code_low: str):
    """在 halogen_source_base 下检索含料号的 .docx（月→日 + 任意两层子目录兜底）。"""
    base = os.path.normpath(base)
    if not os.path.isdir(base):
        warning(f"[Halogen] halogen_source_base 路径不存在: {base}")
        return None
    if not material_code_low:
        return None
    best, best_mtime = None, None

    def _consider(path: str) -> None:
        nonlocal best, best_mtime
        try:
            m = os.path.getmtime(path)
        except Exception:
            return
        if best_mtime is None or m > best_mtime:
            best, best_mtime = path, m

    for m in range(1, 13):
        month_dir = os.path.join(base, f"{m}月")
        if not os.path.isdir(month_dir):
            continue
        for d in range(1, 32):
            day_dir = os.path.join(month_dir, f"{d}号")
            if not os.path.isdir(day_dir):
                continue
            try:
                files = [f for f in os.listdir(day_dir)
                         if f.lower().endswith(".docx") and not f.startswith("~$")]
            except Exception:
                continue
            for f in files:
                if material_code_low in f.lower():
                    _consider(os.path.join(day_dir, f))
    for root, dirs, files in os.walk(base):
        depth = root[len(base):].count(os.sep)
        if depth > 2:
            dirs[:] = []
            continue
        for f in files:
            if (f.lower().endswith(".docx") and not f.startswith("~$")
                    and material_code_low in f.lower()):
                _consider(os.path.join(root, f))
    return best


# ───────────────────────── 中文翻译 ─────────────────────────

def _normalize_quotes(s: str) -> str:
    """把各种单引号/双引号归一成半角，用于翻译精确匹配。

    处理：弯单引号 ''（U+2018/U+2019）、弯双引号 ""（U+201C/U+201D）、
          重音 ´（U+00B4）、反引号 `（U+0060）。
    """
    if not s:
        return ""
    table = {
        "\u2018": "'", "\u2019": "'",   # 弯单引号 → 半角
        "\u201c": '"', "\u201d": '"',   # 弯双引号 → 半角
        "\u00b4": "'", "\u0060": "'",   # 重音、反引号
    }
    for k, v in table.items():
        s = s.replace(k, v)
    return s

def _translate_halogen_table(doc, test_item: str, target_anchor: str, lang: str) -> None:
    """中文报告：把刚插入的卤素表格逐单元格翻译（精确键匹配）。"""
    eng = _eng()
    # find_table 需要 section 段落对象，不是章节名字符串
    section_para = eng.find_section(doc, test_item)
    if section_para is None:
        warning(f"[Halogen] 翻译失败：未找到章节 {test_item!r}")
        return
    tbl = eng.find_table(doc, section_para, target_anchor)
    if tbl is None:
        warning(f"[Halogen] 翻译失败：未找到目标表格（章节={test_item!r}, 锚点={target_anchor!r}）")
        return


    # 上面这行是占位写法，实际用下面的 _EN_MAP：

    from services import language_pack as _lp
    en_labels = _lp.labels("en")
    zh_labels = _lp.labels("zh")
    keys = [k for k in en_labels if k.startswith("halogen_")]

    translated = 0
    for r_i, row in enumerate(tbl.rows):
        seen = set()
        for c_i, cell in enumerate(row.cells):
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            txt = (cell.text or "").strip()
            if not txt:
                continue
            _txt_norm = _normalize_quotes(txt)
            for k in keys:
                if _normalize_quotes(en_labels.get(k, "")) == _txt_norm:
                    eng.set_cell_text_preserve(tbl, r_i, c_i, zh_labels.get(k, txt))
                    translated += 1
                    break
    info(f"[Halogen] 表格翻译完成，命中 {translated} 个单元格")


# ───────────────────────── 主流程 ─────────────────────────

def _process_halogen_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理卤素章节：从源报告复制"Sample Test Result"表格。"""
    doc = _doc(None)
    eng = _eng()
    test_item = chapter_ctx.config.test_item
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")

    record = task_ctx.record
    material_code = getattr(record, "material_code", "") if record else ""
    if not material_code:
        base.status = "skipped"
        base.error = "未获取到 material_code"
        warning(f"[Halogen] {test_item}：{base.error}")
        return base

    src_path = _resolve_halogen_source_report(task_ctx, material_code)
    if src_path is None:
        base.status = "skipped"
        base.error = f"未找到料号 {material_code} 的卤素报告（本地/远程均无）"
        warning(f"[Halogen] {test_item}：{base.error}")
        return base

    info(f"[Halogen] 使用卤素源报告: {src_path}")

    if eng.find_section(doc, test_item) is None:
        base.status = "skipped"
        base.error = f"目标文档中未找到 '{test_item}' 章节"
        warning(f"[Halogen] {test_item}：{base.error}")
        return base

    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    L = language_pack.labels(lang)
    target_anchor = L["chapter_result_description"]  # Result description / 结果说明

    ok = copy_table_by_source_para_from_doc(
        doc, test_item, src_path, _SRC_PARA_TEXT, target_anchor)
    if not ok:
        base.status = "skipped"
        base.error = f"表格复制失败（源锚点={_SRC_PARA_TEXT!r}）"
        warning(f"[Halogen] {test_item}：{base.error}")
        return base

    if lang == "zh":
        _translate_halogen_table(doc, test_item, target_anchor, lang)

    info(f"[Halogen] {test_item}：表格复制完成（lang={lang}）")
    return base


class HalogenStrategy(BaseStrategy):
    """卤素策略——源报告表格复制 + 中文翻译。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        try:
            result = _process_halogen_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            import traceback
            test_item = chapter_ctx.config.test_item
            warning(f"[HalogenStrategy] 执行异常 ({test_item}): {e}\n"
                    f"{traceback.format_exc()}")
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)