# config_loader · 2026-09-10
"""配置加载模块（spec 模块 3）。
职责：加载 config.yaml、读取物料信息表、筛选已审批记录、从变更描述提取料号
并取前 4 位作为物料类别（作为 section_manager 的路由键）。
技术栈：.xlsx/.xls 统一通过 openpyxl 读取（v5.0 起弃用 win32com）。
变更（2026-09-10）：
  - 物料类别不再从物料信息表的「物料类别」列读取，改为从「变更描述」中提取的
    料号前 4 位推导。物料信息表中的「物料类别」列可删除。
"""
import os
import re
from collections import Counter
from typing import Dict, List, Optional, Tuple
import yaml
from openpyxl import load_workbook
try:
    import xlrd
except ImportError:
    xlrd = None
from common.exception import ConfigError, MaterialTableError
from common.logger import apply_config, info, warning
from models.material import MaterialRecord
# 字段 → 候选列名（归一化后匹配；支持中英文、括号、第N页后缀、全角逗号）
# 注意：不再包含 material_category —— 物料类别由变更描述里料号前 4 位推导，
#       物料信息表无需再保留「物料类别」列。
_ALIASES = {
    "number": ["编号", "申请编号", "number", "no"],
    "material_code": ["物料编码", "物料代码", "materialcode", "code"],
    "material_subclass": ["物料子类", "物料小类", "materialsubclass", "subclass"],
    "material_model": ["型号", "物料型号", "materialmodel", "model"],
    "creator": ["创建人", "创建者", "申请人", "creator", "applicant"],
    "department": ["部门", "申请部门", "department", "applicantdepartment"],
    "description": ["变更描述", "描述", "description", "remark"],
    "quantity": ["数量", "样品数量", "quantity", "qty"],
    "receive_date": ["收样日期", "接收日期", "样品接收日期", "receivedate", "date"],
    "sample_count": ["送样次数", "samplecount", "samples"],
    "supplier": ["供应商", "供应商名称", "supplier", "vendor", "manufacturer"],
    "approval_status": ["审批状态", "原始记录审批状态", "审批结果", "approvalstatus", "approved"],
    "report_language": ["报告语言", "语言", "reportlanguage", "language"],
}
# 物料子类从变更描述中动态提取，不再依赖关键词列表
def _norm(s) -> str:
    """归一化列名：去括号、去第X页、去逗号、去空格、转小写。"""
    s = re.sub(r"[()（）]", "", str(s))
    s = re.sub(r"第[一二三四五六七八九十]+页", "", s)
    s = s.replace("，", "").replace(",", "")
    return s.strip().lower().replace(" ", "")
def load_config(tool_dir: Optional[str] = None) -> dict:
    """加载 config.yaml，返回 spec 约定的扁平配置字典。"""
    tool_dir = tool_dir or os.path.dirname(os.path.abspath(__file__))
    cfg_path = os.path.join(tool_dir, "configs", "config.yaml")
    if not os.path.isfile(cfg_path):
        raise ConfigError(f"未找到配置文件: {cfg_path}")
    with open(cfg_path, encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}
    config = {
        "target_base": raw.get("target_base") or "",
        "output_dir": raw.get("output_dir") or "",
        "pdf_output_dir": raw.get("pdf_output_dir") or "",
        "material_table": raw.get("material_table") or "",
        "log_level": raw.get("log_level") or "INFO",
        "log_file": raw.get("log_file") or "",
        "log_root": raw.get("log_root") or "",
        "default_chapter_template": raw.get("default_chapter_template"),
        "default_chapter_template_en": raw.get("default_chapter_template_en") or "",
        "default_chapter_template_zh": raw.get("default_chapter_template_zh") or "",
        "rohs_source_base": raw.get("rohs_source_base") or "",
        "halogen_source_base": raw.get("halogen_source_base") or "",          # ← 新增
        "ion_contamination_source_base": raw.get("ion_contamination_source_base") or "",
    }
    if config["default_chapter_template"] is None:
        raise KeyError("config.yaml 缺少 default_chapter_template 字段（视为配置错误）")
    return config
def _locate_material_table(tool_dir: str, config: Optional[dict] = None) -> str:
    # 定位物料信息表。查找顺序：
    # 1) config["material_table"] 非空且文件存在 -> 直接使用；
    # 2) 报告根目录（.bat 所在目录）下找 物料信息表.xlsx/.xls；
    # 3) src 目录下找同名文件；4) 两目录各自唯一表格文件兜底。
    explicit_dir = None
    if config:
        explicit = (config.get("material_table") or "").strip()
        if explicit and os.path.isfile(explicit):
            return explicit
        if explicit and os.path.isdir(explicit):
            explicit_dir = explicit
    report_root = os.path.dirname(tool_dir)
    search_dirs = ([explicit_dir] if explicit_dir else []) + [report_root, tool_dir]
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for name in ("物料信息表.xlsx", "物料信息表.xls"):
            p = os.path.join(d, name)
            if os.path.isfile(p):
                return p
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        cands = [f for f in os.listdir(d)
                 if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$")]
        if len(cands) == 1:
            return os.path.join(d, cands[0])
    raise MaterialTableError(
        f"未在报告根({report_root})或 src({tool_dir})找到物料信息表；"
        f"请把 物料信息表.xlsx 放到与 .bat 同目录即可自动发现")
def _read_rows(xls_path: str) -> List[Dict[str, str]]:
    """用 openpyxl（.xlsx）或 xlrd（.xls）读取首个工作表，
    返回 [{归一化列名: 单元格}] 列表。"""
    abs_path = os.path.abspath(xls_path)
    ext = os.path.splitext(abs_path)[1].lower()
    if ext == ".xls":
        # 旧格式：用 xlrd 读取
        if xlrd is None:
            raise MaterialTableError(
                f"读取 .xls 文件需要 xlrd 库，请先执行：pip install xlrd")
        try:
            wb = xlrd.open_workbook(abs_path)
        except Exception as e:
            raise MaterialTableError(f"xlrd 无法打开文件: {xls_path} -> {e}")
        try:
            ws = wb.sheet_by_index(0)
            if ws.nrows == 0:
                return []
            headers = [_norm(str(ws.cell_value(0, c))) if ws.cell_value(0, c) != ""
                       else "" for c in range(ws.ncols)]
            rows = []
            for r in range(1, ws.nrows):
                row = {}
                for c in range(ws.ncols):
                    val = ws.cell_value(r, c)
                    row[headers[c]] = "" if val is None or val == "" else str(val)
                rows.append(row)
            return rows
        finally:
            pass  # xlrd 不需要显式关闭
    # 新格式：用 openpyxl 读取
    try:
        wb = load_workbook(abs_path, read_only=True, data_only=True)
    except Exception as e:
        raise MaterialTableError(f"openpyxl 无法打开文件: {xls_path} -> {e}")
    try:
        ws = wb[wb.sheetnames[0]]
        rows_iter = ws.iter_rows(values_only=True)
        data = list(rows_iter)
    finally:
        wb.close()
    if not data:
        return []
    headers = [_norm(h) if h is not None else "" for h in data[0]]
    return [{headers[i]: ("" if r[i] is None else str(r[i])) for i in range(len(headers))}
            for r in data[1:]]
def _to_int(v: str) -> Optional[int]:
    """从单元格取值中提取首个数字（1-已审 -> 1）。"""
    m = re.search(r"\d", v or "")
    return int(m.group()) if m else None
def _build_record(get, approved_only: bool = True) -> Optional[MaterialRecord]:
    """按行构造 MaterialRecord。
    物料类别的确定方式（2026-09-10 起）：
      - 从「变更描述」中提取 10 位以上连续数字作为料号；
      - 取料号前 4 位作为物料类别；
      - 料号提取失败时跳过该记录。
    approved_only=True 时额外要求审批状态==1。
    """
    number = get("number").strip()
    if not number:
        return None
    if approved_only and _to_int(get("approval_status")) != 1:
        return None
    code = get("material_code").strip()
    subclass = get("material_subclass").strip()
    model = get("material_model").strip()
    desc = get("description")
    if not code:
        m = re.search(r"\d{10,}", desc or "")
        code = m.group() if m else ""
    if not subclass:
        # 从变更描述中提取完整子类名。
        # 描述格式：<物料编码> <中文品类[+英文后缀]> <供应商\类型> ...
        # 策略：取料号之后第一个 token（以空格/括号/方括号/·为分隔符），
        # token 必须以中文开头，后面可跟连续的英文/数字/连字符（如 PCB、SMD、SOP-8）。
        # 示例："双面PCB" "贴片双向可控硅" "SMD电阻" 均可正确提取。
        _PAT = re.compile(r"[一-鿿][一-鿿 A-Za-z0-9\-]*")
        # 更精确：从非空白开头，中文起始，连续中文+可选英文后缀，遇空格/括号/特殊符号停止
        _TOKEN = re.compile(r"^([一-鿿][一-鿿A-Za-z0-9\-]*)")
        _extracted = ""
        # 候选来源：料号之后的剩余描述
        remaining = ""
        if code and code in (desc or ""):
            after_code = (desc or "").split(code, 1)[1].lstrip()
            remaining = after_code
        if not remaining:
            remaining = (desc or "").lstrip()
        # 从 remaining 开头匹配
        m = _TOKEN.match(remaining)
        if m:
            _extracted = m.group(1)
        else:
            # 退化：在描述中找第一个以中文开头的 token
            for part in re.split(r"[\s（）()\[\]·]", remaining):
                m2 = _TOKEN.match(part)
                if m2:
                    _extracted = m2.group(1)
                    break
        subclass = _extracted
    if not model:
        cands = re.findall(r"[A-Za-z0-9][A-Za-z0-9\.\-]*", desc or "")
        cands = [c for c in cands if any(ch.isalpha() for ch in c) and c != code]
        model = cands[-1] if cands else ""
    # 物料类别 = 变更描述里料号的前 4 位
    category = code[:4] if len(code) >= 4 else ""
    if not category:
        warning(f"[警告] 无法从变更描述提取料号，已跳过编号: {number}")
        return None
    return MaterialRecord(
        number=number,
        material_category=category,
        material_code=code,
        material_subclass=subclass,
        material_model=model,
        creator=get("creator").strip(),
        department=get("department").strip(),
        description=(desc or "").strip(),
        quantity=get("quantity").strip(),
        receive_date=get("receive_date").strip(),
        sample_count=get("sample_count").strip(),
        supplier=get("supplier").strip(),
        approval_status=get("approval_status").strip(),
        report_language=get("report_language").strip(),
    )
def load_materials(config: dict, tool_dir=None, approved_only=True):
    # approved_only=True: 仅返回已审批且料号可提取的记录（批量模式默认）
    # approved_only=False: 返回所有料号可提取的记录，不限审批状态（拖拽模式）
    tool_dir = tool_dir or os.path.dirname(os.path.abspath(__file__))
    table = _locate_material_table(tool_dir, config)
    info(f"读取物料信息表: {table}")
    rows = _read_rows(table)
    if not rows:
        return []
    key_norm = {_norm(k): k for k in rows[0].keys()}
    norm_aliases = {field: [_norm(a) for a in als] for field, als in _ALIASES.items()}
    field_col = {}
    for field, aliases in _ALIASES.items():
        na = norm_aliases[field]
        hit = next((kh for kh in key_norm if kh in na), None)
        if not hit:
            for nk in sorted(set(na), key=len, reverse=True):
                hit = next((kh for kh in key_norm if nk and nk in kh), None)
                if hit:
                    break
        if hit:
            field_col[field] = key_norm[hit]
    info(f"列名映射: {field_col}")
    records: List[MaterialRecord] = []
    for row in rows:
        def get(field):
            col = field_col.get(field)
            return row.get(col, "") if col else ""
        rec = _build_record(get, approved_only=approved_only)
        if rec is not None:
            records.append(rec)
    for diag_field, sub in (("状态", "状态"), ("审批结果", "审批结果"),
                             ("原始记录审批状态", "原始记录审批状态")):
        hit = next((kh for kh in key_norm if sub in kh), None)
        if hit:
            dist = Counter((row.get(key_norm[hit], "") or "").strip() for row in rows)
            info(f"诊断列[{diag_field}] 取值分布: {dict(dist)}")
    label = "已审批且" if approved_only else ""
    info(f"物料表共 {len(rows)} 行；{label}有效记录（料号可提取）: {len(records)} 条")
    return records
def run(tool_dir: Optional[str] = None) -> Tuple[dict, List[MaterialRecord]]:
    """一次性加载配置 + 物料记录（供入口/临时脚本调用）。"""
    config = load_config(tool_dir)
    apply_config(config)
    records = load_materials(config, tool_dir)
    return config, records
