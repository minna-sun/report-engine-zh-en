"""合规策略（模块 14 · v5.0）。

职责：在合规章节（Compliance 类型）中，从 RoHS 源报告复制两张表：
- "Result description" 下方的数据表格；
- "Test Equipment" 下方的设备表格；
表格格式保持源报告原样（不改变字体）。

实现方式（v5.0）：原委托 template_engine.process_compliance_chapter 的业务编排逻辑
已内联到本模块（_process_compliance_chapter 及 RoHS 源报告解析辅助函数
_resolve_rohs_source_report / _find_docx_in_dir / _find_source_report / _newest），
execute 直接调用本模块逻辑并补全 ChapterResult 公共字段。

边界（见 重构.md §6.4）：
- 不导入 python-docx / lxml / DocumentEngine。
- 不执行设备匹配（直接复制源设备表）。
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
from template_engine import _doc, _eng, localize_equipment_table
from models.result import ChapterResult


def _resolve_rohs_source_report(task_ctx, material_code: str):
    """解析 RoHS 源报告路径（本地优先 → 远程兜底）。

    1. 本地优先：在 {folder_path}/原始记录/ 下按规范化名反查等于 'rohs' 的子目录
       （委托 reference_service.find_actual_folder），在其中查找 .docx，命中即用。
    2. 兜底：本地无可用报告时，从 config.yaml 的 ``rohs_source_base`` 按料号检索
       （月/号目录 + 任意两层子目录），命中后拷贝到本地 {folder_path}/原始记录/RoHS/。

    返回源报告绝对路径；均未命中返回 None。
    """
    folder_path = task_ctx.folder_path
    mc_low = (material_code or "").strip().lower()

    # ── 优先级 1：本地 RoHS 文件夹 ──
    local_dir = reference_service.find_actual_folder(folder_path, "rohs")
    if local_dir:
        local_report = _find_docx_in_dir(local_dir, mc_low)
        if local_report:
            info(f"[process_compliance_chapter] 优先使用本地 RoHS 文件夹: {local_dir}")
            return local_report

    # ── 优先级 2（fallback）：远程按料号拉取 ──
    cfg = task_ctx.config or {}
    rohs_source_base = cfg.get("rohs_source_base") if isinstance(cfg, dict) else None
    if not rohs_source_base:
        warning("[process_compliance_chapter] 未配置 rohs_source_base，且本地无 RoHS 文件夹，"
                "无法拉取源报告")
        return None
    src_path = _find_source_report(rohs_source_base, mc_low)
    if not src_path:
        warning(f"[process_compliance_chapter] 在 rohs_source_base 下未找到料号 {material_code} 的"
                f" RoHS 报告，跳过")
        return None

    # 拷贝到本地（优先复用已有的 rohs 文件夹，没有才新建 "RoHS"）
    target_dir = local_dir if local_dir else os.path.join(folder_path, "原始记录", "RoHS")
    try:
        os.makedirs(target_dir, exist_ok=True)
    except Exception as e:
        warning(f"[process_compliance_chapter] 创建本地目录失败（{target_dir}）: {e}")
        return None
    local_path = os.path.join(target_dir, os.path.basename(src_path))
    try:
        shutil.copy2(src_path, local_path)
    except Exception as e:
        warning(f"[process_compliance_chapter] 拷贝 RoHS 源报告失败（{src_path}）: {e}")
        return None
    info(f"[process_compliance_chapter] 已从远程定位并拷贝 RoHS 源报告到本地: {local_path}")
    return local_path


def _find_docx_in_dir(dir_path: str, material_code_low: str):
    """在 dir_path 下查找 .docx 源报告：优先文件名包含料号，否则取最新修改者。"""
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
    """在 rohs_source_base 下检索文件名含料号的 .docx（月/号目录 + 任意两层子目录兜底）。

    返回修改时间最新的匹配文件绝对路径；未找到返回 None。
    """
    base = os.path.normpath(base)
    if not os.path.isdir(base):
        warning(f"[process_compliance_chapter] rohs_source_base 路径不存在: {base}")
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

    # 按 spec：1月~12月 → 1号~31号
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
    # 兜底：任意两层子目录下匹配
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


def _process_compliance_chapter(chapter_ctx, task_ctx) -> ChapterResult:
    """处理合规章节：从 RoHS 源报告复制数据表 + 设备表。

    源报告查找采用「本地优先、远程兜底」：
      1. 本地优先：reference_service.find_actual_folder(folder_path, "rohs")
         返回非空且其中存在 .docx → 直接使用；
      2. 远程兜底：按 task_ctx.config["rohs_source_base"] 下的「月/号」目录及
         任意两层子目录检索文件名含料号的 .docx，命中后拷贝到本地再使用；
      3. 容错：均未命中 / 拷贝失败 → 返回 status="skipped"，不影响其他章节。

    复制动作经 DocumentEngine.copy_table 完成（目标侧按标签定位 + 删除旧表 +
    跨文档插入三步合一）。
    """
    doc = _doc(None)
    eng = _eng()
    test_item = chapter_ctx.config.test_item
    folder_path = task_ctx.folder_path
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")

    record = task_ctx.record
    material_code = getattr(record, "material_code", "") if record else ""
    if not material_code:
        base.status = "skipped"
        base.error = "未获取到 material_code"
        warning(f"[process_compliance_chapter] {test_item}：{base.error}")
        return base

    src_path = _resolve_rohs_source_report(task_ctx, material_code)
    if src_path is None:
        base.status = "skipped"
        base.error = f"未找到料号 {material_code} 的 RoHS 报告（本地/远程均无）"
        warning(f"[process_compliance_chapter] {test_item}：{base.error}")
        return base

    info(f"[process_compliance_chapter] 使用 RoHS 源报告: {src_path}")

    src_doc = eng.open(src_path)
    if src_doc is None:
        base.status = "failed"
        base.error = "RoHS 源报告打开失败"
        return base

    # ── 前置检查：目标文档中必须存在配置表指定的章节 ──────────────────────
    if eng.find_section(doc, test_item) is None:
        base.status = "skipped"
        base.error = f"目标文档中未找到 '{test_item}' 章节，跳过 copy_table"
        warning(f"[process_compliance_chapter] {test_item}：{base.error}")
        return base

    # ── 源报告章节名：用候选列表匹配（供应商命名可能不一致）──────────────
    _SRC_SECTION_CANDIDATES = ["RoHS Testing", "RoHS test", "RoHS"]
    src_section = next(
        (s for s in _SRC_SECTION_CANDIDATES if eng.find_section(src_doc, s) is not None),
        None,
    )
    if src_section is None:
        base.status = "skipped"
        base.error = "源报告中未找到 RoHS 相关章节"
        warning(f"[process_compliance_chapter] {test_item}：{base.error}")
        return base
    info(f"[process_compliance_chapter] 源报告章节: '{src_section}'，目标章节: '{test_item}'")

    # 双语：源报告（实验室 Z 盘 / 本地 RoHS）为固定英文格式，源锚点不变；
    # 目标文档锚点按报告语言切换（中文：结果说明 / 检测设备）。
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    L = language_pack.labels(lang)
    dst_result_desc = L["chapter_result_description"]
    dst_equip = L["equip_title"]
    ok1 = eng.copy_table(doc, test_item, src_doc, "Result description", dst_result_desc)
    ok2 = eng.copy_table(doc, test_item, src_doc, "Test Equipment", dst_equip)
    if not ok1:
        base.status = "skipped"
        base.error = "数据表复制失败"
    if not ok2:
        warning(f"[process_compliance_chapter] {test_item}：设备表复制失败（不影响状态）")
    elif lang == "zh":
        # 中文报告：复制过来的设备表还是英文表头和英文设备名，本地化一次
        localize_equipment_table(doc, test_item, lang="zh")
    return base


class ComplianceStrategy(BaseStrategy):
    """合规策略——RoHS 源报告表格复制（逻辑内联）。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        """执行合规策略：本模块内联逻辑完成源报告解析与表格复制。

        Args:
            template_engine: template_engine 模块/实例（兼容形参）。
            chapter_ctx: 章节执行上下文（含 config.test_item 等）。
            task_ctx: 任务执行上下文（含 folder_path、record、config 等）。

        Returns:
            ChapterResult：章节执行结果（status / error 等由本模块填充，公共字段经 _finalize 补全）。
        """
        try:
            result = _process_compliance_chapter(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[ComplianceStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
