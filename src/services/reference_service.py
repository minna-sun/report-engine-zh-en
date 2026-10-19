"""统一数据查询接口（spec 模块 6）。
# 验收版本 — 2026-07-29 14:47

职责：屏蔽数据来源细节，为上层提供 5 个只读查询接口 + 1 个文件夹名规范化函数。
所有 .md 表格数据必须经本模块读取，上层禁止自行解析 Markdown。

边界：
- 不缓存（每次调用重新读文件）
- 不验证数据有效性（缺失时返回占位符）
- 不修改或写入数据
- 不为缺失数据提供默认值（由调用方决定）
"""
import os
import re
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

from common.logger import info, warning

_TOOL_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
_DOC_DB = os.path.join(_TOOL_DIR, "docs", "数据库")
_MAT_CFG = os.path.join(_DOC_DB, "物料类别配置")

# 运行期内存缓存（脚本进程存续期间有效，不持久化）
# key: 文件绝对路径 → value: _md_tables 解析结果
_MD_CACHE: Dict[str, List] = {}
# key: folder_path → value: load_folder_config 返回的配置字典
_FOLDER_CFG_CACHE: Dict[str, Dict] = {}


def _find_md(base_dir: str, *cands: str) -> Optional[str]:
    for c in cands:
        p = os.path.join(base_dir, c)
        if os.path.isfile(p):
            return p
    return None


def _clean_hdr(h: str) -> str:
    return re.sub(r"<[^>]+>", "", h).strip()


def _md_tables(path: str) -> List[Tuple[str, List[str], List[Dict[str, str]]]]:
    # 命中缓存直接返回（同一进程内多次调用同一 .md 文件只解析一次）
    abs_path = os.path.abspath(path)
    if abs_path in _MD_CACHE:
        return _MD_CACHE[abs_path]

    if not os.path.isfile(abs_path):
        _MD_CACHE[abs_path] = []
        return []
    with open(abs_path, "r", encoding="utf-8") as f:
        lines = [ln.rstrip("\n") for ln in f]
    tables, cur, cap = [], [], ""
    for ln in lines:
        if ln.startswith("|"):
            cur.append(ln)
        else:
            if len(cur) >= 3:
                hdr = [_clean_hdr(p) for p in cur[0].split("|")[1:-1]]
                rows = []
                for l2 in cur[2:]:
                    pts = [p.strip() for p in l2.split("|")[1:-1]]
                    if len(pts) >= len(hdr):
                        rows.append(dict(zip(hdr, pts)))
                tables.append((cap, hdr, rows))
            cur, cap = [], ln.strip()
    if len(cur) >= 3:
        hdr = [_clean_hdr(p) for p in cur[0].split("|")[1:-1]]
        rows = []
        for l2 in cur[2:]:
            pts = [p.strip() for p in l2.split("|")[1:-1]]
            if len(pts) >= len(hdr):
                rows.append(dict(zip(hdr, pts)))
        tables.append((cap, hdr, rows))
    # 过滤掉文档信息等非业务表格，避免索引错位
    result = [(c, h, r) for c, h, r in tables
              if not (len(h) == 2 and h[0] == "项目" and h[1] == "内容")]
    _MD_CACHE[abs_path] = result
    return result


def load_folder_config(folder_path: str, lang: str = "en") -> Dict[str, Dict[str, str]]:
    # 命中缓存直接返回（同一编号文件夹只读一次 Excel；结果包含双语 key，lang 不影响缓存）
    abs_folder = os.path.abspath(folder_path)
    if abs_folder in _FOLDER_CFG_CACHE:
        return _FOLDER_CFG_CACHE[abs_folder]

    cfg = os.path.join(folder_path, "目标文件夹配置.xlsx")
    if not os.path.isfile(cfg):
        warning(f"[reference_service] 未找到目标文件夹配置表，跳过运行时字段填充：{cfg}")
        _FOLDER_CFG_CACHE[abs_folder] = {}
        return {}
    try:
        wb = load_workbook(cfg, read_only=True, data_only=True)
    except Exception as e:
        warning(f"load_folder_config failed (open): {cfg} -> {e}")
        _FOLDER_CFG_CACHE[abs_folder] = {}
        return {}
    try:
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        if not data:
            _FOLDER_CFG_CACHE[abs_folder] = {}
            return {}
        lr, lc = len(data), max((len(r) for r in data), default=0)
        key_col, hdr_row = 0, 0  # 0-based
        found = False
        for r in range(min(5, lr)):
            for c in range(lc):
                v = data[r][c]
                if v and "测试项目" in str(v):
                    hdr_row, key_col = r, c
                    found = True
                    break
            if found:
                break

        # 收集所有包含"测试项目"的列（支持"测试项目章节名（英文）"+"测试项目章节名（中文）"双列，
        # 也兼容旧版单列"测试项目"）。每行数据将同时以各列的值作为 key 注册到 out，
        # 使英文章节名和中文章节名都能匹配到同一行数据。
        hdr_len = len(data[hdr_row])
        key_cols = [c for c in range(min(lc, hdr_len))
                    if data[hdr_row][c] is not None
                    and "测试项目" in str(data[hdr_row][c]).strip()]
        if not key_cols:
            key_cols = [key_col]   # 完全兜底：退回初始扫描定位的列
        # ── 诊断：打印识别到的 key 列索引及列名 ──────────────────────────────
        info(f"[load_folder_config] hdr_row={hdr_row}, key_cols={key_cols}, "
             f"col_names={[str(data[hdr_row][c]).strip() for c in key_cols if c < hdr_len]}")

        cmap = {c: str(data[hdr_row][c]).strip()
                for c in range(lc) if c < hdr_len and data[hdr_row][c]}
        out: Dict[str, Dict[str, str]] = {}
        for r in range(hdr_row + 1, lr):
            row = data[r]
            keys_to_register = []
            for kc in key_cols:
                v = row[kc] if kc < len(row) else None
                if v is not None:
                    ks = str(v).strip()
                    if ks:
                        keys_to_register.append(ks)
            # ── 诊断：逐行打印注册的 key（第 4 步）────────────────────────────
            info(f"[load_folder_config] 第{r}行 keys_to_register={keys_to_register}")
            if not keys_to_register:
                continue
            row_data = {n: "" if (row[c] if c < len(row) else None) is None
                        else str(row[c]).strip() for c, n in cmap.items()}
            for k in keys_to_register:
                out[k] = row_data

        _FOLDER_CFG_CACHE[abs_folder] = out
        # ── 诊断：第 1 步：确认返回 key 列表 ───────────────────────────────
        info(f"[load_folder_config] 文件夹: {folder_path}")
        info(f"[load_folder_config] 返回的 key 列表: {list(out.keys())}")
        return out
    except Exception as e:
        warning(f"load_folder_config failed: {cfg} -> {e}")
        _FOLDER_CFG_CACHE[abs_folder] = {}
        return {}
    finally:
        wb.close()


def get_equipment_detail(asset_id: str, lang: str = "en") -> Dict[str, str]:
    """查询设备详情。

    Args:
        asset_id: 资产编号。
        lang: 报告语言。``"zh"`` 时优先取"设备中文名称"列，
            ``"en"`` 时优先取"设备英文名称"列；目标列为空则回退到英文名称。
    """
    p = _find_md(_DOC_DB, "设备字典.md")
    for _, _, rows in _md_tables(p or ""):
        for row in rows:
            if row.get("资产编号", "") == asset_id:
                name_en = row.get("设备英文名称") or row.get("设备名称", "/")
                name_zh = row.get("设备中文名称") or name_en
                name = (name_zh if lang == "zh" else name_en) or "/"
                return {"name": name, "model": row.get("型号", "/"),
                        "no": row.get("资产编号", "/"), "expiry": row.get("有效期区间", "/")}
    return {"name": "/", "model": "/", "no": "/", "expiry": "/"}


def get_standard(material_category: str, test_item: str) -> str:
    p = (_find_md(_MAT_CFG, f"{material_category}.md")
         or _find_md(_MAT_CFG, "索引.md"))
    for _, _, rows in _md_tables(p or ""):
        for row in rows:
            tk = next((k for k in row if "测试项目" in k), None)
            if tk and row[tk] == test_item:
                sk = next((k for k in row if "测试标准" in k), None)
                return row.get(sk, "/") if sk else "/"
    return "/"


def get_translation(field_type: str, value: str, lang: str = "en") -> str:
    """按字段类型查询中→英翻译（双语支持）。

    Args:
        field_type: ``"department"``（部门映射表优先）或 ``"material"``（物料映射表优先）。
        value: 原始中文值。
        lang: 报告语言。``lang == "zh"`` 时跳过翻译，直接返回原始中文值
            （中文报告无需翻译，见双语改造「翻译跳过逻辑」）；其余值执行翻译。

    Returns:
        英文翻译；查不到时返回 ``[缺失: {value}] [TR]`` 占位符。
    """
    from services.language_pack import ZH, normalize as _norm
    if _norm(lang) == ZH or not (value or "").strip():
        return value or ""
    p = _find_md(_DOC_DB, "翻译映射表.md")
    tables = _md_tables(p or "")
    order = (0, 1) if field_type == "department" else (1, 0) if field_type == "material" else range(len(tables))
    for i in order:
        if i >= len(tables):
            continue
        rows = tables[i][2]
        for row in rows:
            if row.get("中文名称", "") == value:
                return row.get("英文翻译", f"[缺失: {value}] [TR]")
    return f"[缺失: {value}] [TR]"


def determine_test_result(folder_name: str) -> str:
    """根据文件夹名后缀判定 Test Result。
    后缀大小写不敏感（-FAIL / -fail / -PASS / -pass / -REF / -ref 均可）。
    后缀判定顺序：先判 FAIL（更具体），再判 PASS，最后判 Reference。
    """
    n = re.sub(r"^\d+[-_ ]", "", folder_name.strip()).lower()
    if n.endswith(("-不合格", "-fail")):
        return "FAIL"
    if n.endswith(("-合格", "-pass")):
        return "PASS"
    if n.endswith(("-参考", "-ref")):
        return "Reference"
    return "PASS"


def normalize_folder_name(raw_name: str) -> str:
    n = raw_name.strip().rstrip("/\\")
    n = re.sub(r"^\d+[.\-_ ]", "", n)
    n = re.sub(r"[_-]\d{6,8}$", "", n)
    n = re.sub(r"[_-]\d+$", "", n)
    n = re.sub(r"[_-](备份|copy|bak|old|new)$", "", n, flags=re.I)
    return n


def find_actual_folder(folder_path: str, norm_name: str) -> Optional[str]:
    """在 folder_path/原始记录/ 下，查找名称中包含 norm_name 的子文件夹（大小写不敏感）。

    只要文件夹名（原始名称）中含有 norm_name 字符串即视为匹配，不做规范化处理。
    例如：查找 "rohs" 可匹配 rohs、RoHS、RoHS Testing、14_rohs、X_rohs 等。

    Args:
        folder_path: 目标文件夹路径（编号文件夹根目录）。
        norm_name: 要匹配的关键词（如 "rohs"），大小写不敏感。

    Returns:
        第一个匹配到的文件夹完整路径，未找到返回 None。
    """
    raw_data_dir = os.path.join(folder_path, "原始记录")
    if not os.path.isdir(raw_data_dir):
        return None
    key = norm_name.lower()
    for entry in os.listdir(raw_data_dir):
        entry_path = os.path.join(raw_data_dir, entry)
        if os.path.isdir(entry_path) and key in entry.lower():
            return entry_path
    return None


def get_pcb_supplier_id(name: str) -> str:
    """在 PCB供应商映射表中查找供应商，返回标准 ID（中文名优先）。

    匹配：中文名、英文名、英文简称、别名（不区分大小写、忽略空格）。

    Args:
        name: 供应商名称（中文/英文/简称/别名任一）

    Returns:
        标准 ID（优先中文名）；匹配不到返回空字符串
    """
    if not name or not str(name).strip():
        return ""
    key = str(name).strip().lower().replace(" ", "")
    p = _find_md(_DOC_DB, "PCB供应商映射.md")
    for _, _, rows in _md_tables(p or ""):
        for row in rows:
            zh = (row.get("供应商中文名") or "").strip()
            en = (row.get("供应商英文名") or "").strip()
            short = (row.get("供应商英文简称") or "").strip()
            alias = (row.get("供应商别名") or "").strip()
            aliases = [a.strip() for a in re.split(r"[,，;；]", alias) if a.strip()]
            candidates = [zh, en, short] + aliases
            for c in candidates:
                if c and c.lower().replace(" ", "") == key:
                    return zh or en or short
    return ""
