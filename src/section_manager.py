"""章节管理器（spec 模块 8）。

职责：扫描 ``原始记录/`` 文件夹，匹配物料类别配置，确定"有哪些章节"；
加载目标文件夹配置，组装 ``ChapterContext``。

变更记录：
- 2026-08-28  _strip_suffix 新增 -ref / _ref 后缀支持。
  背景：实测文件夹命名含 -REF（如 01_外观-REF、03_xray-REF），原后缀列表未覆盖，
  导致 normalize 后无法匹配 config_map，对应章节被误删。
  影响范围：仅 _strip_suffix，不影响其他匹配逻辑。
- 2026-09-20  新增 get_chapter_template_name()，从物料类别配置表读「新增章节范本」列。

边界：
- 不负责查询设备详情（由 reference_service 处理）
- 不负责填充章节内容（由 GenericProcessor + Strategy 处理）
- 不直接读取目标文件夹配置.xlsx（统一通过 reference_service.load_folder_config）
"""

import os
import re
from typing import List, Dict, Tuple, Optional

from models.material import MaterialRecord
from models.chapter import ChapterConfig, ChapterRuntime, ChapterContext
from services import reference_service
from common.logger import info, warning

_TOOL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_MAT_CFG_DIR = os.path.join(_TOOL_DIR, "docs", "数据库", "物料类别配置")


def _clean_hdr(header: str) -> str:
    """去除表头中的 HTML 标签并清理空白（含不间断空格）。"""
    return re.sub(r"<[^>]+>", "", header).replace("\xa0", " ").strip()


def _parse_md_tables(path: str) -> List[Tuple[str, List[str], List[Dict[str, str]]]]:
    """解析 Markdown 文件中的所有表格。

    Returns:
        [(caption, headers, rows), ...]
    """
    if not os.path.isfile(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        lines = [ln.rstrip("\n") for ln in f]

    tables: List[Tuple[str, List[str], List[Dict[str, str]]]] = []
    cur: List[str] = []
    caption = ""

    for ln in lines:
        if ln.startswith("|"):
            cur.append(ln)
        else:
            if len(cur) >= 3:
                raw_hdrs = [p.strip().replace("\xa0", " ") for p in cur[0].split("|")[1:-1]]
                headers = [_clean_hdr(h) for h in raw_hdrs]
                rows: List[Dict[str, str]] = []
                for l2 in cur[2:]:
                    pts = [p.strip().replace("\xa0", " ") for p in l2.split("|")[1:-1]]
                    if len(pts) >= len(headers):
                        rows.append(dict(zip(headers, pts)))
                tables.append((caption, headers, rows))
            cur, caption = [], ln.strip().replace("\xa0", " ")

    if len(cur) >= 3:
        raw_hdrs = [p.strip().replace("\xa0", " ") for p in cur[0].split("|")[1:-1]]
        headers = [_clean_hdr(h) for h in raw_hdrs]
        rows = []
        for l2 in cur[2:]:
            pts = [p.strip().replace("\xa0", " ") for p in l2.split("|")[1:-1]]
            if len(pts) >= len(headers):
                rows.append(dict(zip(headers, pts)))
        tables.append((caption, headers, rows))

    return tables


def get_chapter_template_name(material_category: str, lang: str = "en") -> str:
    """按物料类别读取「新增章节范本」对应的中/英文模板章节名。

    从 docs/数据库/物料类别配置/{material_category}.md 读表：
      - 找「新增章节范本」列填「是」的那一行
      - lang="en" → 取该行的「测试项目章节名（英文）」列
      - lang="zh" → 取该行的「章节名（中文）」列
      - 多行填「是」时只取第一行并警告
      - 无任何行填「是」→ 返回空字符串（由调用方兜底）

    Args:
        material_category: 物料类别（如 "DEMO-001"）。
        lang: 报告语言（"en" / "zh"）。

    Returns:
        模板章节名（可能为空字符串）。
    """
    if not material_category:
        return ""
    try:
        path = os.path.join(_MAT_CFG_DIR, f"{material_category}.md")
        if not os.path.isfile(path):
            return ""

        tables = _parse_md_tables(path)

        found_rows = []
        for _caption, headers, rows in tables:
            if not any("新增章节范本" in h for h in headers):
                continue
            for row in rows:
                mark = (row.get("新增章节范本") or "").strip()
                if mark in ("是", "Y", "y", "yes", "Yes"):
                    found_rows.append(row)

        if not found_rows:
            return ""
        if len(found_rows) > 1:
            warning(f"[section_manager] {material_category}.md 中"
                    f"「新增章节范本」有多行填'是'，只取第一行")

        row = found_rows[0]
        is_zh = (lang or "").strip().lower().startswith("zh")
        if is_zh:
            for k in row:
                if "章节名（中文）" in k or "章节名(中文)" in k:
                    return (row[k] or "").strip()
        else:
            for k in row:
                if "章节名（英文）" in k or "章节名(英文)" in k:
                    return (row[k] or "").strip()
        # 找不到对应列时回退到"测试项目"
        return (row.get("测试项目") or "").strip()
    except Exception as e:
        warning(f"[section_manager] get_chapter_template_name 失败：{e}")
        return ""


def _find_col(headers: List[str], keywords: List[str]) -> Optional[str]:
    """在表头中查找包含任意关键词的列名（不区分大小写）。"""
    for h in headers:
        clean = h.lower()
        for kw in keywords:
            if kw.lower() in clean:
                return h
    return None

_ANY_TAG_RE = re.compile(r"<[^>]*>")


def _sanitize_cell(val) -> str:
    """清洗配置单元格：去 HTML 标签、换行、不可见字符。

    处理：<br> / <p> 等 HTML 标签；\n \r \t；全角空格；零宽字符；NBSP。
    """
    if not val:
        return ""
    s = str(val)
    s = _ANY_TAG_RE.sub("", s)                       # 去 HTML 标签
    s = s.replace("\r", " ").replace("\n", " ")      # 换行 → 空格
    s = s.replace("\t", " ")
    s = s.replace("\u3000", " ").replace("\u200b", "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()            # 压缩连续空白


def _parse_placeholders(val: str) -> Optional[List[str]]:
    """将占位符列表字符串解析为列表；'—' 或空值返回 None。

    同时把 Markdown 转义的下划线（``\\_``）还原为普通下划线（``_``），
    并自动剥离模板书写常用的括号（``{}`` / ``[]`` / ``<>``），
    使占位符名与 Word 模板中的真实文本一致（如 ``{front}`` -> ``front``、
    ``[back]`` -> ``back``、``<overview>`` -> ``overview``）。
    """
    val = _sanitize_cell(val)          # ← 新增：先清洗（去 <br>、换行、不可见字符）
    if not val or val == "—":
        return None
    out = []
    for b in val.split(","):
        b = b.strip().replace("\\_", "_")
        # 去除包裹的括号（保留括号内的真实占位符名）
        if len(b) >= 2 and b[0] in "{[<" and b[-1] in "}]>":
            b = b[1:-1].strip()
        if b:
            out.append(b)
    return out or None


def _parse_float(val: str) -> Optional[float]:
    """将字符串解析为浮点数；空值或非法值返回 ``None``。

    用于读取物料类别配置中的 ``图片宽度`` 列（如 ``"7.0"`` -> ``7.0``）。
    """
    if not val or val.strip() == "" or val.strip() == "—":
        return None
    try:
        return float(val.strip())
    except (ValueError, TypeError):
        return None


def _load_material_config(material_category: str, lang: str = "en") -> List[ChapterConfig]:
    """加载物料类别配置，返回所有测试章节配置。

    Args:
        material_category: 物料类别（如 ``"DEMO-001"``）。
        lang: 报告语言（``"en"`` / ``"zh"``），决定取哪列作为 ``test_item``：
            - ``"en"`` → 优先取"测试项目章节名（英文）"列，不存在则退回"测试项目"。
            - ``"zh"`` → 优先取"测试项目章节名（中文）"列，不存在则退回"测试项目"。

    Returns:
        ChapterConfig 列表；文件不存在或解析失败返回空列表。
    """
    cfg_path = os.path.join(_MAT_CFG_DIR, f"{material_category}.md")
    if not os.path.isfile(cfg_path):
        warning(f"物料类别配置不存在: {cfg_path}")
        return []

    try:
        tables = _parse_md_tables(cfg_path)
    except Exception as e:
        warning(f"解析物料类别配置失败: {cfg_path} -> {e}")
        return []

    if not tables:
        warning(f"物料类别配置中未找到表格: {cfg_path}")
        return []

    # 定位包含 "测试项目" 列的配置表格（跳过文档信息表等）
    headers: List[str] = []
    rows: List[Dict[str, str]] = []
    for _, hdrs, rws in tables:
        if _find_col(hdrs, ["测试项目"]):
            headers, rows = hdrs, rws
            break

    if not headers:
        warning(f"物料类别配置中未找到含'测试项目'列的表格: {cfg_path}")
        return []

    col_test_item = (
        _find_col(headers, ["测试项目章节名（中文）", "中文"])
        or _find_col(headers, ["测试项目"])
    ) if lang == "zh" else (
        _find_col(headers, ["测试项目章节名（英文）", "英文"])
        or _find_col(headers, ["测试项目"])
    )
    col_folder = _find_col(headers, ["文件夹名"])
    col_category = _find_col(headers, ["类型"])
    col_strategy = _find_col(headers, ["strategy"])
    col_placeholders = _find_col(headers, ["占位符列表", "书签列表"])
    col_image_map = _find_col(headers, ["图片匹配说明"])
    col_standard = _find_col(headers, ["测试标准"])
    col_image_width = _find_col(headers, ["图片宽度", "image_width"])

    configs: List[ChapterConfig] = []
    for idx, row in enumerate(rows):
        test_item = row.get(col_test_item, "").replace("\xa0", " ") if col_test_item else ""
        folder_name = row.get(col_folder, "").replace("\xa0", " ") if col_folder else ""
        if not test_item or not folder_name:
            continue

        category = row.get(col_category, "").replace("\xa0", " ") if col_category else ""
        strategy = row.get(col_strategy, "").replace("\xa0", " ") if col_strategy else ""
        placeholders = _parse_placeholders(row.get(col_placeholders, "").replace("\xa0", " ")) if col_placeholders else None
        image_mapping = row.get(col_image_map, None)
        if image_mapping:
            image_mapping = image_mapping.replace("\xa0", " ")
        if image_mapping == "—":
            image_mapping = None
        standard = row.get(col_standard, None)
        if standard:
            standard = standard.replace("\xa0", " ")
        image_width = _parse_float(row.get(col_image_width, "")) if col_image_width else None

        configs.append(ChapterConfig(
            test_item=test_item,
            category=category,
            strategy=strategy,
            folder_name=folder_name,
            placeholders=placeholders,
            image_mapping=image_mapping,
            standard=standard,
            image_width=image_width,
        ))

    return configs


def _strip_suffix(name: str) -> str:
    """去掉常见的状态/参考后缀（如 -PASS、_FAIL、-参考），用于匹配配置。"""
    # 第一步：去状态后缀（-PASS/-FAIL/-参考 等）
    lower = name.lower()
    for suffix in sorted(("-不合格", "_不合格", "-合格", "_合格",
                          "-pass", "_pass", "-fail", "_fail", "-参考", "_参考",
                          "-ng", "_ng", "-ok", "_ok", "-partial", "_partial",
                          "-ref", "_ref"),
                         key=len, reverse=True):
        if lower.endswith(suffix):
            name = name[: -len(suffix)]
            break
    # 第二步：去第一个空格及其后的所有内容（人名、备注等，如 "外观 刘" -> "外观"）
    if " " in name:
        name = name.split(" ")[0]
    return name


def _folder_sort_key(item: Tuple[str, ChapterConfig]) -> tuple:
    """按文件夹名排序：优先按数字前缀自然排序；无数字前缀按字母顺序排最后。"""
    name = item[0]
    m = re.match(r"^(\d+)", name)
    if m:
        return (int(m.group(1)), name.lower())
    return (999999, name.lower())


def build_chapters(folder_path: str, record: MaterialRecord,
                   lang: str = "en") -> List[ChapterContext]:
    """扫描原始记录并组装章节上下文列表。

    执行流程：
    1. 读取物料类别配置获取策略。
    2. 扫描 ``folder_path/原始记录/`` 子文件夹。
    3. 规范化文件夹名后匹配策略配置的 folder_name。
    4. 按原始记录顺序排序，RoHS 强制置尾。
    5. 加载目标文件夹配置获取运行时数据。
    6. 组装 ChapterContext。

    Args:
        folder_path: 目标文件夹路径。
        record: 当前物料记录（需含 ``material_category``）。
        lang: 报告语言（``"en"`` / ``"zh"``），传递给 ``_load_material_config``
            以决定 ``test_item`` 取英文列还是中文列。

    Returns:
        ChapterContext 列表；无匹配章节时返回空列表。
    """
    material_category = record.material_category
    info(f"加载物料类别配置: {material_category}")

    configs = _load_material_config(material_category, lang=lang)
    if not configs:
        warning(f"物料类别 {material_category} 无有效策略配置")
        return []

    config_map: Dict[str, ChapterConfig] = {
        cfg.folder_name.lower(): cfg for cfg in configs
    }

    raw_data_dir = os.path.join(folder_path, "原始记录")
    if not os.path.isdir(raw_data_dir):
        warning(f"原始记录目录不存在: {raw_data_dir}")
        return []

    matched: List[Tuple[str, ChapterConfig]] = []
    try:
        entries = os.listdir(raw_data_dir)
    except Exception as e:
        warning(f"读取原始记录目录失败: {raw_data_dir} -> {e}")
        return []

    for entry in entries:
        entry_path = os.path.join(raw_data_dir, entry)
        if not os.path.isdir(entry_path):
            continue

        norm_name = reference_service.normalize_folder_name(entry)
        cfg = config_map.get(norm_name.lower())

        # 精确匹配失败时，先尝试 _strip_suffix（剥状态后缀/空格）
        if not cfg:
            stripped = _strip_suffix(norm_name)
            if stripped != norm_name:
                cfg = config_map.get(stripped.lower())

        # 仍失败时，逐段剥尾部（空格 / - / _ / / 都算分隔符）：
        # 从右往左一段段剥，每剥一段查一次配置表，剥到匹配为止。
        # 例："可焊性-罗-pass" → "可焊性-罗" → "可焊性"（匹配）
        #     "可焊性_罗 张"  → "可焊性_罗"  → "可焊性"（匹配）
        if not cfg:
            candidate = norm_name
            while True:
                _m = re.search(r"[ \-_/][^ \-_/]*$", candidate)
                if not _m:
                    break
                candidate = candidate[: _m.start()]
                if not candidate:
                    break
                cfg = config_map.get(candidate.lower())
                if cfg:
                    break

        if cfg:
            matched.append((entry, cfg))
            info(f"匹配章节: {cfg.test_item} <- {entry}")

    if not matched:
        warning("未匹配到任何章节")
        return []

    # RoHS 强制置尾：用子串匹配兼容英文"ROHS"和中文"RoHS测试"等章节名
    non_rohs = [(name, cfg) for name, cfg in matched
                if "rohs" not in cfg.test_item.lower()]
    rohs = [(name, cfg) for name, cfg in matched
            if "rohs" in cfg.test_item.lower()]
    sorted_matched = sorted(non_rohs, key=_folder_sort_key) + sorted(rohs, key=_folder_sort_key)

    # ── 诊断：第 2 步：确认传入 load_folder_config 的 lang ──────────────────
    info(f"[section_manager] build_chapters lang={lang!r}，即将调用 load_folder_config")
    folder_config = reference_service.load_folder_config(folder_path)
    info(f"[section_manager] load_folder_config 返回 {len(folder_config)} 个 key: {list(folder_config.keys())}")

    contexts: List[ChapterContext] = []
    for source_name, cfg in sorted_matched:
        runtime_data = folder_config.get(cfg.test_item, {})
        # ── 诊断：第 3 步：确认 test_item 与 folder_config 匹配情况 ──────────
        info(f"[section_manager] test_item={cfg.test_item!r}，"
             f"命中={'是' if runtime_data else '否（数据为空）'}，"
             f"runtime_data={runtime_data}")

        runtime = ChapterRuntime(
            device=runtime_data.get("设备编号") or runtime_data.get("测试设备编号") or None,
            condition=runtime_data.get("测试条件") or None,
            result_description=runtime_data.get("Result description") or None,
            sample_number=runtime_data.get("样品编号") or None,
        )

        ctx = ChapterContext(
            config=cfg,
            runtime=runtime,
            material_category=material_category,
            source_folder_name=source_name,
        )
        contexts.append(ctx)

    info(f"共生成 {len(contexts)} 个章节上下文")
    return contexts