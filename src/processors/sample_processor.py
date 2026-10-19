"""样品处理器（模块 18）。

职责：步骤1~3（页眉、封面、Sample Description）填充，详见《业务执行规范》第4节
"详细填写步骤"的步骤1/步骤2/步骤3。直接由 ``core/report_engine`` 调用
（不参与 ``processor_registry`` 调度，属于非类型化章节处理器）。

边界：
- 不持有 COM 对象（doc 由 report_engine 打开并传入）
- 不负责翻译映射维护（``get_translation`` 查 ``reference_service``）
- 不负责模板书签设置（``insert_*/insert_image`` 由 template_engine 执行）
- 不负责图片裁剪/缩放（``process_image`` 由 image_service 处理，宽度由
  ``insert_image`` 的 7cm 参数控制）
- 不判断物料信息表数据正确性（由 config_loader 筛选后传入）
- 图片缺失 → 填 "/"，不中断执行（仅记录警告）
- 不硬编码语言标签：所有锚点经 ``services.language_pack`` 按
  ``task_ctx.report_language`` 动态取用（中文报告跳过翻译）

输出 ``ChapterResult``：``test_item="Sample Description"``、``test_basis=""``、
``sample_number=record.sample_count``、``test_result="Reference"``（不参与最终结论）。
"""

import re
from datetime import datetime

import template_engine
from common.logger import warning
from models.result import ChapterResult
from services import reference_service
from services import language_pack
from template_engine import (
    insert_text_after,
    insert_table_cell_value,
    read_table_cell_value,
)


# ---------- 辅助函数 ----------

def _sample_count_letter(n) -> str:
    """送样次数 → 字母（1→A … 26→Z），超界或非法返回原字符串。"""
    try:
        i = int(str(n).strip())
    except (ValueError, TypeError):
        return ""
    if 1 <= i <= 26:
        return chr(ord("A") + i - 1)
    return str(i)


def _format_date(value, lang: str = "en") -> str:
    """格式化为日期字符串。

    - lang="en" → MM/DD/YYYY
    - lang="zh" → YYYY年MM月DD日

    无法解析则原样返回。
    """
    if not value:
        return ""
    s = str(value).strip()
    # 先统一解析为 (年, 月, 日)
    ymd = None
    if re.match(r"^\d{4}/\d{2}/\d{2}$", s):
        p = s.split("/")
        ymd = (int(p[0]), int(p[1]), int(p[2]))
    elif re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        p = s.split("/")
        ymd = (int(p[2]), int(p[0]), int(p[1]))
    elif re.match(r"^\d{8}$", s):
        ymd = (int(s[:4]), int(s[4:6]), int(s[6:]))
    else:
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
                    "%Y年%m月%d日",           # ← 新增
                    "%Y年%m月%d",             # ← 新增（无"日"）
                    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                    "%d/%m/%Y", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(s, fmt)
                ymd = (dt.year, dt.month, dt.day)
                break
            except ValueError:
                continue
    if ymd is None:
        return s
    y, m, d = ymd
    if lang == "zh":
        return f"{y}年{m:02d}月{d:02d}日"
    return f"{m:02d}/{d:02d}/{y}"


def _strip_creator(creator: str) -> str:
    """申请人：保留英文/字母工号部分，去掉末尾中文姓名。

    例：``'WXH 吴新华'`` → ``'WXH'``；``'吴新华 WXH'`` → ``'WXH'``。
    若全部 token 均含中文（无工号），回退为原值。
    """
    toks = [t for t in (creator or "").split()
            if not re.search(r"[一-鿿]", t)]
    return " ".join(toks).strip() or (creator or "").strip()


# ---------- 步骤实现 ----------

def _fill_header(doc, record, lang: str) -> None:
    """步骤1：页眉「Report No.:」/「报告编号：」后填 {编号}{送样次数字母}。"""
    L = language_pack.labels(lang)
    number = (record.number or "").strip()
    report_no = f"{number}{_sample_count_letter(record.sample_count)}"
    insert_text_after(doc, "Header", L["header_report_no"], report_no)


def _fill_cover(doc, record, lang: str) -> None:
    """步骤2：封面（Applicant / Applicant Department / Received Date /
    Inspection Date / Report Date）。

    语言：锚点全部经 ``language_pack`` 切换；中文报告（lang="zh"）的「委托单位」
    跳过翻译、直接使用物料表原始中文（见双语改造「翻译跳过逻辑」）。
    日期格式：英文 → MM/DD/YYYY；中文 → YYYY年MM月DD日。
    """
    L = language_pack.labels(lang)
    SEC = L["cover_title"]
    # 日期按语言输出
    receive = _format_date(record.receive_date, lang=lang)
    report_date = _format_date(datetime.now().strftime("%Y-%m-%d"), lang=lang)
    inspection = f"{receive} ~ {report_date}"
    # 封面为「标签|值」两列表格：标签在左格、值在右格（同行下一单元格），
    # 须用 insert_table_cell_value 写入相邻值单元格（不能用 insert_text_after 追加）。
    # 申请人：保留工号、去掉中文姓名（如 "WXH 吴新华" -> "WXH"）。
    # 申请部门：英文报告经 reference_service 翻译为英文名；中文报告跳过翻译。
    _COVER_FONT = {"font_name": "Times New Roman", "font_size": 14}  # 四号
    if lang == "zh":
        _applicant = (record.creator or "").strip()
    else:
        _applicant = _strip_creator(record.creator)
    insert_table_cell_value(doc, SEC, L["cover_applicant"], _applicant,
                            require_bold=False, **_COVER_FONT)
    insert_table_cell_value(
        doc, SEC, L["cover_department"],
        reference_service.get_translation("department", record.department or "", lang),
        require_bold=False, **_COVER_FONT)
    insert_table_cell_value(doc, SEC, L["cover_received_date"], receive,
                           require_bold=False, **_COVER_FONT)
    insert_table_cell_value(doc, SEC, L["cover_inspection_date"], inspection,
                           require_bold=False, **_COVER_FONT)
    insert_table_cell_value(doc, SEC, L["cover_report_date"], report_date,
                           require_bold=False, **_COVER_FONT)


# ---------- 主入口 ----------

def _fill_cover_dates_only(doc, lang: str) -> None:
    """独立模式：从模板读取 Received Date，据此决定 Inspection Date 是否填充。

    - 读到 Received Date → Inspection Date = "{Received} ~ {当前日期}"
    - 读不到 → Inspection Date 跳过填充
    - Report Date 始终填当前日期
    - 其余封面字段保持模板原样
    - 日期格式：英文 → MM/DD/YYYY；中文 → YYYY年MM月DD日
    """
    L = language_pack.labels(lang)
    SEC = L["cover_title"]
    today = _format_date(datetime.now().strftime("%Y-%m-%d"), lang=lang)
    _COVER_FONT = {"font_name": "Times New Roman", "font_size": 14}  # 四号
    # 从模板读取 Received Date（可能是任意格式，按当前语言格式化）
    received = read_table_cell_value(doc, SEC, L["cover_received_date"],
                                      require_bold=False)
    if received:
        received_fmt = _format_date(received, lang=lang)
        # ← 新增：把格式化后的 Received Date 写回封面（原逻辑漏了这一步）
        insert_table_cell_value(doc, SEC, L["cover_received_date"], received_fmt,
                                require_bold=False, **_COVER_FONT)
        inspection = f"{received_fmt} ~ {today}"
        insert_table_cell_value(doc, SEC, L["cover_inspection_date"], inspection,
                                require_bold=False, **_COVER_FONT)
    else:
        warning("[sample_processor] 模板中未找到 Received Date，跳过 Inspection Date 填充")
    insert_table_cell_value(doc, SEC, L["cover_report_date"], today,
                            require_bold=False, **_COVER_FONT)


def execute(doc, task_ctx) -> ChapterResult:
    """执行步骤1~3，返回 ``ChapterResult``。

    Args:
        doc: 已打开的 Word Document（由 report_engine 持有）。
        task_ctx: 任务执行上下文（含 record / folder_path / config /
            report_language）。

    Returns:
        ChapterResult: ``test_item="Sample Description"``、
        ``test_result="Reference"``（不参与最终结论判定）；
        图片缺失仍视为成功（仅警告）；硬失败 ``status="failed"``。
    """
    record = task_ctx.record
    lang = language_pack.normalize(getattr(task_ctx, "report_language", "en"))

    # 独立模式：填封面日期 + Sample Label / Sample Image 图片，跳过页眉/封面文本/样品信息文字字段
    if getattr(task_ctx, "standalone_mode", False):
        try:
            _fill_cover_dates_only(doc, lang)
            template_engine.fill_sample_description_images_only(doc, task_ctx)
        except Exception as exc:
            warning(f"[sample_processor] 独立模式图片填充失败：{exc}")
            return ChapterResult(
                test_item="Sample Description", test_basis="",
                sample_number=record.sample_count, test_result="Reference",
                status="failed", error=f"独立模式图片填充失败：{exc}")
        return ChapterResult(
            test_item="Sample Description", test_basis="",
            sample_number=record.sample_count, test_result="Reference",
            status="success", error=None)

    try:
        _fill_header(doc, record, lang)
        _fill_cover(doc, record, lang)
        # 步骤3：委托给 template_engine.fill_sample_description，
        # 支持中英文锚点切换及 1-1/1-2 图片匹配。
        template_engine.fill_sample_description(doc, task_ctx)
    except Exception as exc:
        warning(f"[sample_processor] 步骤1~3填充失败：{exc}")
        return ChapterResult(
            test_item="Sample Description", test_basis="",
            sample_number=record.sample_count, test_result="Reference",
            status="failed", error=f"步骤1~3填充失败：{exc}")

    return ChapterResult(
        test_item="Sample Description", test_basis="",
        sample_number=record.sample_count, test_result="Reference",
        status="success", error=None)


__all__ = ["execute"]
