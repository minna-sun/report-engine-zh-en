"""离子污染策略（新增 · 2026-09-15）。

职责：针对物料类别 1025（PCB），从离子污染源报告（PDF）中找到匹配的报告，
拷贝到原始记录目录，并提取 Flow Rate 方框图插入到报告模板的对应章节。

四条件命中规则（料号、周期号、AP、供应商）：
  - 料号、周期号：必有，来自 R 流程文件夹名
  - AP、供应商：文件夹名有则满足；没有则去 PDF 正文 Lot No 确认
  - 四条件全中才命中；命中多个时按 PDF 修改时间取最新

适用章节：物料类别配置表中类型为 IonContamination 的章节。
"""

import os
import re
import shutil
from typing import Any

from strategies.base import BaseStrategy
from common.logger import info, warning
from models.result import ChapterResult
from services import reference_service
from services import pdf_extractor
from services import language_pack
from template_engine import _doc, insert_image_after_text

# 统一 Lot No 文本的分隔符
_SEP_PAT = re.compile(r"[\s\-_/;,，、|]+")


def _extract_from_folder_name(folder_name: str) -> dict:
    """从 R 流程文件夹名提取料号、周期号、AP、供应商。

    Args:
        folder_name: 只传文件夹名（不含路径），如
            "R0001234+ 10000000001+（DEMO 示例供应商）+场效应管"

    Returns:
        {"material_code": str, "period": str, "ap": bool, "supplier": str}
    """
    out = {"material_code": "", "period": "", "ap": False, "supplier": ""}
    # 料号：第一个 10 位以上数字
    m = re.search(r"\d{10,}", folder_name)
    if m:
        out["material_code"] = m.group(0)
    # 括号内容（兼容全半角括号）
    m = re.search(r"[（(]([^）)]+)[）)]", folder_name)
    if m:
        inner = m.group(1)
        # 周期号：括号内第一个 4 位数字
        pm = re.search(r"\d{4}", inner)
        if pm:
            out["period"] = pm.group(0)
            # 供应商：去掉周期号及前后分隔符
            rest = inner.replace(pm.group(0), "", 1)
            rest = rest.strip(" \t-_,;/;|、")
            out["supplier"] = rest.strip()
    # AP：整个文件夹名里找（不区分大小写，词边界）
    if re.search(r"\bAP\b", folder_name, re.IGNORECASE):
        out["ap"] = True
    return out


def _split_lot_no(text: str) -> list:
    """把 Lot No 文本按多种分隔符拆分。"""
    return [t for t in _SEP_PAT.split(text) if t]


def _find_lot_no(pdf_text: str) -> str:
    """从 PDF 文本找 Lot No 行，返回冒号后的内容。找不到返回空字符串。"""
    for line in pdf_text.splitlines():
        if "Lot No" in line:
            for sep in (":", "："):
                if sep in line:
                    return line.split(sep, 1)[1].strip()
    return ""


def _match_in_lot_no(lot_no: str, info_dict: dict,
                     folder_supplier_id: str) -> tuple:
    """在 Lot No 文本里验证四条件（顺序无关）。

    Returns:
        (料号命中, 周期号命中, AP命中, 供应商命中)
    """
    tokens = _split_lot_no(lot_no)
    joined = " ".join(tokens)
    code_ok = bool(info_dict["material_code"]) and info_dict["material_code"] in joined
    period_ok = bool(info_dict["period"]) and info_dict["period"] in joined
    # AP：不区分大小写
    ap_ok = any(t.upper() == "AP" for t in tokens)
    # 供应商：
    #   文件夹名有供应商 → PDF 里必须能找到同一个（归一化 ID 匹配）
    #   文件夹名无供应商 → PDF 里只要有任何可识别的供应商即可
    if folder_supplier_id:
        sup_ok = any(
            reference_service.get_pcb_supplier_id(t) == folder_supplier_id
            for t in tokens
        )
    else:
        sup_ok = any(
            reference_service.get_pcb_supplier_id(t)
            for t in tokens
        )
    return code_ok, period_ok, ap_ok, sup_ok


def _extract_from_pdf_filename(pdf_filename: str, material_code: str,
                                period: str) -> dict:
    """从 PDF 文件名提取 AP 和供应商信息。

    文件名示例：
      DEMO-A000000-P-V06(10000000001)-DEMO1-AP.pdf
      DEMO-A000001-PC-V01(10000000002)-DEMO2-示例供应商.pdf
      DEMO-B000002-C-V02(10000000003)-DEMO3-示例供应商.pdf

    提取方式：
      1. 去除扩展名
      2. 去除料号、周期号
      3. 去除括号及括号内内容
      4. 剩下的 token 逐个判断：是否是 AP / 是否在供应商映射表中
    """
    out = {"ap": False, "supplier_id": ""}
    if not pdf_filename:
        return out
    name = os.path.splitext(pdf_filename)[0]
    # 判断 AP（不区分大小写，词边界）
    if re.search(r"\bAP\b", name, re.IGNORECASE):
        out["ap"] = True
    # 去除料号和周期号
    if material_code:
        name = name.replace(material_code, " ")
    if period:
        name = name.replace(period, " ")
    # 去除括号及括号内内容（兼容全半角）
    name = re.sub(r"[（(][^）)]*[）)]", " ", name)
    # 拆分 token，逐个尝试识别为供应商
    tokens = re.split(r"[\s\-_/;,，、|]+", name)
    for t in tokens:
        t = t.strip()
        if not t:
            continue
        sid = reference_service.get_pcb_supplier_id(t)
        if sid:
            out["supplier_id"] = sid
            break
    return out


def _scan_candidates(source_base: str, material_code: str,
                     period: str) -> list:
    """递归扫描源报告目录，返回文件名含料号+周期号的 PDF 列表（按修改时间降序）。"""
    candidates = []
    if not os.path.isdir(source_base):
        return candidates
    for root, _dirs, files in os.walk(source_base):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            if material_code and material_code not in f:
                continue
            if period and period not in f:
                continue
            full = os.path.join(root, f)
            try:
                mtime = os.path.getmtime(full)
            except Exception:
                mtime = 0
            candidates.append((full, mtime))
    candidates.sort(key=lambda x: x[1], reverse=True)
    return [c[0] for c in candidates]


def _process_ion_contamination(chapter_ctx, task_ctx) -> ChapterResult:
    doc = _doc(None)
    test_item = chapter_ctx.config.test_item
    folder_path = task_ctx.folder_path
    base = ChapterResult(test_item=test_item, test_basis="", sample_number="",
                         test_result="", status="success")

    # 源报告根路径
    cfg = task_ctx.config or {}
    source_base = cfg.get("ion_contamination_source_base") if isinstance(cfg, dict) else None
    if not source_base:
        base.status = "skipped"
        base.error = "未配置 ion_contamination_source_base"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    # 从 R 流程文件夹名解析
    folder_name = os.path.basename(folder_path.rstrip("\\/"))
    info_dict = _extract_from_folder_name(folder_name)
    _missing = []
    if not info_dict["material_code"]:
        _missing.append("料号")
    if not info_dict["period"]:
        _missing.append("周期号")
    if _missing:
        base.status = "skipped"
        base.error = f"文件夹名缺少{'和'.join(_missing)}：{folder_name}"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    info(f"[IonContamination] {test_item}：解析结果 料号={info_dict['material_code']}, "
         f"周期号={info_dict['period']}, AP={'是' if info_dict['ap'] else '否'}, "
         f"供应商={info_dict['supplier']!r}")

    # 文件夹名里的供应商归一化 ID
    folder_supplier_id = ""
    if info_dict["supplier"]:
        folder_supplier_id = reference_service.get_pcb_supplier_id(info_dict["supplier"])
        if not folder_supplier_id:
            warning(f"[IonContamination] 供应商 {info_dict['supplier']!r} "
                    f"不在映射表中，将退化为字符串比较")

    # 扫描候选
    candidates = _scan_candidates(source_base, info_dict["material_code"],
                                  info_dict["period"])
    if not candidates:
        base.status = "skipped"
        base.error = f"未找到候选 PDF（料号+周期号），扫描根={source_base}"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    info(f"[IonContamination] {test_item}：找到 {len(candidates)} 个候选 PDF")

    # ── 按优先级筛选候选 ──
    def _has_ap(pdf_name):
        return bool(re.search(r"\bAP\b", pdf_name, re.IGNORECASE))

    # 文件名里除料号、周期号、扩展名、括号内容外，剩下的 token 里有供应商
    def _has_supplier(pdf_name):
        name = os.path.splitext(pdf_name)[0]
        if info_dict["material_code"]:
            name = name.replace(info_dict["material_code"], " ")
        if info_dict["period"]:
            name = name.replace(info_dict["period"], " ")
        name = re.sub(r"[（(][^）)]*[）)]", " ", name)
        tokens = re.split(r"[\s\-_/;,，、|]+", name)
        for t in tokens:
            if t.strip() and reference_service.get_pcb_supplier_id(t):
                return True
        return False

    # 按优先级分类
    tier1 = []  # 4要素
    tier2 = []  # 3要素含 AP
    tier3 = []  # 3要素含供应商
    tier4 = []  # 2要素
    for pdf_path in candidates:
        pdf_name = os.path.basename(pdf_path)
        has_ap = _has_ap(pdf_name)
        has_sup = _has_supplier(pdf_name)
        if has_ap and has_sup:
            tier1.append(pdf_path)
        elif has_ap:
            tier2.append(pdf_path)
        elif has_sup:
            tier3.append(pdf_path)
        else:
            tier4.append(pdf_path)

    # 依次检查各优先级，取第一个非空的，再从中取最新
    hit_pdf = None
    hit_tier = 0
    for tier_num, tier_list in [(1, tier1), (2, tier2), (3, tier3), (4, tier4)]:
        if tier_list:
            hit_pdf = max(tier_list, key=lambda p: os.path.getmtime(p))
            hit_tier = tier_num
            break

    if not hit_pdf:
        base.status = "skipped"
        base.error = "所有候选 PDF 均不满足最低条件（料号+周期号）"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    tier_desc = {1: "4要素(裁剪)", 2: "3要素含AP(全图)",
                 3: "3要素含供应商(全图)", 4: "2要素(全图)"}
    info(f"[IonContamination] 命中候选（优先级{hit_tier}：{tier_desc[hit_tier]}）：{hit_pdf}")

    # ── 拷贝 PDF ──
    target_folder_name = (chapter_ctx.source_folder_name or "").strip()
    if not target_folder_name:
        target_folder_name = (chapter_ctx.config.folder_name or "").strip()
    if not target_folder_name:
        target_folder_name = "离子污染"
    target_dir = os.path.join(folder_path, "原始记录", target_folder_name)
    try:
        os.makedirs(target_dir, exist_ok=True)
        target_pdf = os.path.join(target_dir, os.path.basename(hit_pdf))
        shutil.copy2(hit_pdf, target_pdf)
        info(f"[IonContamination] 已拷贝 PDF 到 {target_pdf}")
    except Exception as e:
        base.status = "failed"
        base.error = f"拷贝 PDF 失败：{e}"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    # ── 提取图片 ──
    if hit_tier == 1:
        # 4要素：裁剪坐标图
        img_path = pdf_extractor.extract_main_image(hit_pdf)
    else:
        # 其他：整页全图
        img_path = pdf_extractor.extract_full_page(hit_pdf)

    if not img_path:
        base.status = "failed"
        base.error = "提取 PDF 图片失败"
        warning(f"[IonContamination] {test_item}：{base.error}")
        return base

    info(f"[IonContamination] 已提取图片到 {img_path}")

    # ── 插入 Word（Result description 下一行）──
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))
    anchor = language_pack.labels(lang)["chapter_result_description"]
    ok = insert_image_after_text(doc, test_item, anchor, img_path, 18.0)

    # 清理临时图片
    try:
        os.remove(img_path)
    except Exception:
        pass

    if not ok:
        base.status = "failed"
        base.error = "图片插入失败"
        warning(f"[IonContamination] {test_item}：{base.error}")
    return base


class IonContaminationStrategy(BaseStrategy):
    """离子污染策略——PDF 提取 + 图片插入。"""

    def execute(self, template_engine: Any, chapter_ctx, task_ctx) -> Any:
        try:
            result = _process_ion_contamination(chapter_ctx, task_ctx)
            return self._finalize(result, chapter_ctx, task_ctx)
        except Exception as e:
            test_item = chapter_ctx.config.test_item
            warning(f"[IonContaminationStrategy] 执行异常 ({test_item}): {e}")
            from models.result import ChapterResult
            return self._finalize(
                ChapterResult(test_item=test_item, test_basis="", sample_number="",
                              test_result="", status="failed", error=str(e)),
                chapter_ctx, task_ctx)
