"""主入口（模块 23）。

对外暴露可编程接口 run_report_generation(excel_path, base_path)，
同时保留命令行入口 main() 兼容双击 .bat 运行。

技术栈：python-docx（Word 文档生成）+ openpyxl（Excel 读取/回写），
无 win32com / Office / COM 依赖，跨平台（Windows / macOS / Linux）。

作者信息：孙敏娜 · 质量中心
运行模式（2026-09-10 起）：
  双击模式（无参数）：
    - bat 同级有物料表 + 无 原始记录/ → 批量模式（仅处理审批状态=1 的记录）
    - bat 同级有物料表 + 有 原始记录/ → 完整填充该 R 流程（不限审批状态）
    - bat 同级无物料表 + 有 原始记录/ → 独立模式（只填章节 + 标签/样品图片）
    - bat 同级无物料表 + 无 原始记录/ → 报错退出
  拖拽模式（有参数）：
    - 拖入文件夹有物料表 → 完整填充该流程（不限审批状态）
    - 拖入文件夹无物料表 → 独立模式
"""
import os
import sys
import re

_SRC = os.path.dirname(os.path.abspath(__file__))
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from common.logger import (
    info, warning, error,
    setup as _setup_logger,
    apply_config as _apply_config,
    set_number as _set_number,
    attach_run_log, detach_run_log,
)
from models.result import GenerationSummary, ScanResult
import config_loader
import folder_scanner
from core.report_engine import generate_report
import file_renamer
import table_writer
from services import language_pack
from services.language_detector import detect_language


# ──────────────────────────────────────────────────────────────────────────────
# 内部辅助
# ──────────────────────────────────────────────────────────────────────────────

def _find_template(folder_path: str):
    """在 folder_path 根目录下自动发现唯一的 .docx 文件。
    返回 (template_path, error_msg)：成功时 error_msg=None，失败时 template_path=None。
    """
    candidates = [f for f in os.listdir(folder_path)
                  if f.lower().endswith(".docx")
                  and not f.startswith("~$")
                  and os.path.isfile(os.path.join(folder_path, f))]
    if len(candidates) == 1:
        tpl = os.path.join(folder_path, candidates[0])
        info(f"[main] 发现报告模板: {candidates[0]}")
        return tpl, None
    if len(candidates) == 0:
        return None, "未找到报告模板（根目录下无 .docx 文件）"
    return None, f"发现多个 .docx 文件，无法确定模板：{candidates}"


def _decide_language(record, template_path: str, number: str = "") -> str:
    """决策单份报告的语言（三级优先级）。"""
    specified = (getattr(record, "report_language", "") or "").strip()
    if specified:
        lang = language_pack.normalize(specified)
        info(f"[main] {number}：报告语言取自物料表「报告语言」列="
             f"{specified!r} → {lang}")
        return lang
    lang = detect_language(template_path)
    if lang not in language_pack.SUPPORTED:
        lang = language_pack.DEFAULT
    info(f"[main] {number}：报告语言由模板自动识别 → {lang}"
         f"（物料表未指定）")
    return lang


def _build_output_path(folder_path: str, standalone: bool = False,
                       template_name: str = "") -> str:
    """构建 Word 输出路径：写入 folder_path/output/ 子目录。"""
    out_dir = os.path.join(folder_path, "output")
    os.makedirs(out_dir, exist_ok=True)
    if standalone and template_name:
        from datetime import datetime
        base = os.path.splitext(os.path.basename(template_name))[0]
        name = f"{base}_{datetime.now():%Y%m%d}.docx"
        return os.path.join(out_dir, name)
    return os.path.join(out_dir, "生成的报告.docx")


def _find_material_table_in(folder: str) -> str:
    """在指定文件夹查找物料信息表，返回完整路径或空字符串。"""
    if not folder or not os.path.isdir(folder):
        return ""
    for name in ("物料信息表.xlsx", "物料信息表.xls"):
        p = os.path.join(folder, name)
        if os.path.isfile(p):
            return p
    return ""


def _is_r_flow_folder(folder: str) -> bool:
    """判断是否为 R 流程文件夹（依据：存在 原始记录/ 子目录）。"""
    if not folder or not os.path.isdir(folder):
        return False
    return os.path.isdir(os.path.join(folder, "原始记录"))


def _extract_number_from_folder(folder_path: str) -> str:
    """从文件夹名提取流程编号（仅用于日志，可能为空）。"""
    folder_name = os.path.basename(folder_path.rstrip("\\/"))
    m = re.search(r"R\d+", folder_name, re.IGNORECASE)
    if m:
        return m.group(0)
    parts = folder_name.split()
    return parts[0] if parts else ""


def _extract_material_code_from_folder(folder_path: str) -> str:
    """从文件夹名提取料号（10 位以上连续数字）。"""
    folder_name = os.path.basename(folder_path.rstrip("\\/"))
    m = re.search(r"\d{10,}", folder_name)
    return m.group(0) if m else ""


def _extract_material_code_from_template(template_path: str) -> str:
    """从模板 Sample Description 页提取料号（10 位以上连续数字）。"""
    try:
        from docx import Document
        doc = Document(template_path)
        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                for idx, cell in enumerate(cells):
                    txt = (cell.text or "").strip()
                    if "Material Number" in txt or "物料编号" in txt:
                        if idx + 1 < len(cells):
                            val = (cells[idx + 1].text or "").strip()
                            mm = re.search(r"\d{10,}", val)
                            if mm:
                                return mm.group(0)
        for para in doc.paragraphs:
            mm = re.search(r"\d{10,}", para.text or "")
            if mm:
                return mm.group(0)
    except Exception as e:
        warning(f"[main] 从模板提取料号失败：{e}")
    return ""


def _load_records_with_table(table_path: str, base_cfg: dict,
                              approved_only: bool = True):
    """用指定物料表加载记录，返回 (cfg, records)。"""
    cfg = dict(base_cfg)
    cfg["material_table"] = table_path
    tool_dir = os.path.dirname(table_path)
    records = config_loader.load_materials(cfg, tool_dir,
                                            approved_only=approved_only)
    return cfg, records


def _run_standalone(folder_path: str, config: dict) -> GenerationSummary:
    """独立模式：不依赖物料信息表，只填测试章节 + 标签/样品图片。"""
    from models.material import MaterialRecord
    number = _extract_number_from_folder(folder_path)
    _set_number(number)
    _run_log_handler = attach_run_log(folder_path)
    try:
        template_path, err = _find_template(folder_path)
        if err:
            return GenerationSummary(
                status="failed", total=1, succeeded=0, failed=1,
                details=[{"编号": number, "status": "failed",
                          "report_path": None, "error": err}])
        material_code = _extract_material_code_from_template(template_path)
        material_category = material_code[:4] if len(material_code) >= 4 else ""
        if not material_category:
            return GenerationSummary(
                status="failed", total=1, succeeded=0, failed=1,
                details=[{"编号": number, "status": "failed",
                          "report_path": None,
                          "error": "无法确定物料类别（模板中未找到料号）"}])
        record = MaterialRecord(
            number="", material_category=material_category,
            material_code=material_code, material_subclass="", material_model="",
            creator="", department="", description="",
            quantity="", receive_date="", sample_count="",
            supplier="", approval_status="", report_language="",
        )
        template_name = os.path.basename(template_path)
        output_path = _build_output_path(folder_path, standalone=True,
                                         template_name=template_name)
        lang = detect_language(template_path)
        if lang not in language_pack.SUPPORTED:
            lang = language_pack.DEFAULT
        info(f"[standalone] 独立模式：物料类别={material_category}")
        result = generate_report(template_path, record, folder_path,
                                 output_path, config, report_language=lang,
                                 standalone_mode=True)
        if result.status == "success":
            return GenerationSummary(
                status="success", total=1, succeeded=1, failed=0,
                details=[{"编号": number, "status": "success",
                          "report_path": result.report_path, "error": None}])
        return GenerationSummary(
            status="failed", total=1, succeeded=0, failed=1,
            details=[{"编号": number, "status": "failed",
                      "report_path": None, "error": result.error}])
    finally:
        _set_number(None)
        detach_run_log(_run_log_handler)


def _run_pipeline(config: dict, records, base_path: str,
                   _verbose: bool = False,
                   scan_override=None) -> GenerationSummary:
    """核心流水线：扫描 → 逐文件夹生成报告 → 汇总。供多个入口共用。"""
    details = []
    succeeded = 0
    failed = 0
    _tool_dir = os.path.dirname(os.path.abspath(__file__))
    try:
        _table_path = config_loader._locate_material_table(_tool_dir, config)
    except Exception:
        _table_path = None

    if scan_override is not None:
        scan_results = scan_override
        info("[main] 直接使用指定文件夹，跳过 folder_scanner")
    else:
        scan_results = folder_scanner.scan_folders(records, base_path)
    total_numbers = len(scan_results)

    for scan in scan_results:
        _set_number(scan.number)
        _run_log_handler = None
        if scan.folder_paths:
            _run_log_handler = attach_run_log(scan.folder_paths[0])
        try:
            info(f"[main] ▶ 正在处理编号: {scan.number}")

            if scan.status == "failed":
                warning(f"[main] {scan.number}：{scan.error}")
                details.append({"编号": scan.number, "status": "failed",
                                "report_path": None, "error": scan.error})
                failed += 1
                continue

            for folder_path in scan.folder_paths:
                info(f"[main] 正在处理文件夹: {folder_path}")
                template_path, tpl_err = _find_template(folder_path)
                if tpl_err:
                    warning(f"[main] {scan.number}：{tpl_err}")
                    details.append({"编号": scan.number, "status": "failed",
                                    "report_path": None, "error": tpl_err})
                    failed += 1
                    continue

                record = next((r for r in records if r.number == scan.number), None)
                if record is None:
                    err = f"物料记录丢失（编号 {scan.number}）"
                    details.append({"编号": scan.number, "status": "failed",
                                    "report_path": None, "error": err})
                    failed += 1
                    continue

                output_path = _build_output_path(folder_path)
                lang = _decide_language(record, template_path, scan.number)
                result = generate_report(template_path, record, folder_path,
                                         output_path, config, report_language=lang)

                if result.status == "success":
                    try:
                        new_path = file_renamer.rename_report(
                            result.report_path, record, result.final_verdict, lang=lang)
                        info(f"[main] {scan.number} → {new_path}")

                        pdf_base = (config.get("pdf_output_dir") or "").strip()
                        if not pdf_base:
                            info("[main] 未配置 pdf_output_dir，跳过快捷方式创建")
                        elif new_path and os.path.isfile(new_path):
                            try:
                                from datetime import datetime as _dt
                                now = _dt.now()
                                material_category = (record.material_category or "").strip()
                                if not material_category:
                                    info("[main] material_category 为空，跳过快捷方式创建")
                                else:
                                    dated_dir = os.path.join(
                                        os.path.abspath(pdf_base),
                                        str(now.year),
                                        material_category,
                                        str(now.month),
                                        str(now.day),
                                    )
                                    os.makedirs(dated_dir, exist_ok=True)
                                    stem = os.path.splitext(os.path.basename(new_path))[0]
                                    target_abs = os.path.abspath(new_path)
                                    link_path = os.path.join(dated_dir, stem + ".lnk")
                                    try:
                                        if os.path.lexists(link_path):
                                            os.remove(link_path)
                                        from pylnk3 import for_file as _make_lnk
                                        _make_lnk(target_abs, link_path)
                                        if os.path.isfile(link_path):
                                            info(f"[main] 已创建 .lnk 快捷方式：{link_path}")
                                        else:
                                            warning(f"[main] 快捷方式创建失败：{link_path}")
                                    except ImportError:
                                        warning("[main] pylnk3 未安装，跳过快捷方式创建"
                                                "（在命令行运行：pip install pylnk3）")
                                    except Exception as lnk_exc:
                                        warning(f"[main] 报告链接创建失败：{lnk_exc}")
                            except Exception as lnk_exc:
                                warning(f"[main] 快捷方式目录创建失败：{lnk_exc}")

                        if _table_path:
                            table_writer.write_output_time(_table_path, scan.number)

                        details.append({"编号": scan.number, "status": "success",
                                       "report_path": new_path, "error": None})
                        succeeded += 1
                    except Exception as exc:
                        error(f"[main] {scan.number} 重命名失败：{exc}")
                        details.append({"编号": scan.number, "status": "failed",
                                       "report_path": result.report_path,
                                       "error": f"重命名失败：{exc}"})
                        failed += 1
                else:
                    error(f"[main] {scan.number} 报告生成失败：{result.error}")
                    details.append({"编号": scan.number, "status": "failed",
                                   "report_path": None, "error": result.error})
                    failed += 1
        finally:
            _set_number(None)
            detach_run_log(_run_log_handler)

    total = total_numbers
    if total == 0:
        status = "failed"
    elif failed == 0:
        status = "success"
    else:
        status = "partial" if succeeded > 0 else "failed"
    return GenerationSummary(status=status, total=total,
                             succeeded=succeeded, failed=failed, details=details)


def _run_batch_mode(table_path: str, base_cfg: dict) -> None:
    """批量模式：用指定物料表处理所有已审批记录（仅审批状态=1）。"""
    info(f"[main] 批量模式（物料表：{table_path}，仅审批状态=1）")
    cfg, records = _load_records_with_table(table_path, base_cfg,
                                             approved_only=True)
    base_path = cfg.get("target_base", "")
    if not base_path:
        print("[错误] config.yaml 中 target_base 为空，请检查配置")
        return
    summary = _run_pipeline(cfg, records, base_path, _verbose=True)
    _print_summary(summary)
    _log_summary(summary)


def _run_single_flow(folder_path: str, table_path: str, base_cfg: dict) -> None:
    """完整填充单个 R 流程（用指定物料表，不限审批状态）。"""
    material_code = _extract_material_code_from_folder(folder_path)
    number_hint = _extract_number_from_folder(folder_path)
    cfg, records = _load_records_with_table(table_path, base_cfg,
                                             approved_only=False)
    record = None
    if material_code:
        record = next((r for r in records if r.material_code == material_code), None)
    if record is None and number_hint:
        record = next((r for r in records if r.number == number_hint), None)
    if record is None:
        print(f"[错误] 物料表中未找到匹配记录：")
        print(f"       料号={material_code!r}  流程号提示={number_hint!r}")
        print(f"       文件夹={folder_path}")
        return
    number = record.number or number_hint or material_code
    base_path = cfg.get("target_base", "")
    scan = ScanResult(number=number, folder_paths=[folder_path], status="success")
    info(f"[main] 完整填充模式：编号={number}，料号={record.material_code}，文件夹={folder_path}")
    summary = _run_pipeline(cfg, [record], base_path, scan_override=[scan])
    _print_summary(summary)
    _log_summary(summary)


def _print_summary(summary: GenerationSummary) -> None:
    """按规范格式打印汇总结果到命令行。"""
    sep = "─" * 45
    print(f"\n{sep}")
    print("[处理] 全部处理完成！")
    print(sep)
    print(f"  总计编号：{summary.total} 个")
    print(f"  成功生成：{summary.succeeded} 份")
    print(f"  失败：{summary.failed} 个")
    if summary.failed > 0:
        print("\n  失败明细：")
        for d in summary.details:
            if d["status"] == "failed":
                print(f"  - {d['编号']}：{d['error']}")
    if summary.succeeded > 0:
        print("\n  成功生成的报告：")
        for d in summary.details:
            if d["status"] == "success":
                print(f"  - {d['编号']} → {d['report_path']}")
    print(sep)


def _log_summary(summary: GenerationSummary) -> None:
    """将汇总结论写入日志系统（WARNING 级）。"""
    warning(f"[main] ===== 运行汇总 ===== 总计:{summary.total} 成功:{summary.succeeded} 失败:{summary.failed}")
    for d in summary.details:
        if d["status"] == "failed":
            warning(f"[main] 失败 {d['编号']}：{d['error']}")
    for d in summary.details:
        if d["status"] == "success":
            info(f"[main] 成功 {d['编号']} → {d['report_path']}")


# ──────────────────────────────────────────────────────────────────────────────
# 对外接口
# ──────────────────────────────────────────────────────────────────────────────

def run_report_generation(excel_path: str, base_path: str) -> GenerationSummary:
    """可编程接口：直接传入路径调用，不打印、不交互，仅返回结构化结果。"""
    cfg = config_loader.load_config(_SRC)
    cfg["target_base"] = base_path
    records = config_loader.load_materials(cfg, os.path.dirname(os.path.abspath(excel_path)))
    return _run_pipeline(cfg, records, base_path)


def main() -> None:
    """命令行入口：支持双击模式（无参数）和拖拽模式（有参数）。"""
    print("[信息] 脚本开发：孙敏娜")
    print("[信息] 部门：质量中心")
    print("[信息] 联系电话：见 GitHub Issues")
    _setup_logger(level="INFO", print_to_console=True)

    arg_folder = None
    if len(sys.argv) >= 2:
        candidate = sys.argv[1].strip().rstrip("\\/.")
        if os.path.isdir(candidate):
            arg_folder = candidate

    try:
        base_cfg = config_loader.load_config(_SRC)
        _apply_config(base_cfg)
    except Exception as e:
        print(f"[错误] 加载配置失败：{e}")
        input("按任意键退出...")
        return

    # ── 拖拽模式 ──
    if arg_folder:
        info(f"[main] 拖拽模式：folder_path={arg_folder}")
        drag_table = _find_material_table_in(arg_folder)
        has_data_dir = _is_r_flow_folder(arg_folder)
        if drag_table and not has_data_dir:
            info(f"[main] 拖拽模式 → 批量模式（工具文件夹）")
            _run_batch_mode(drag_table, base_cfg)
        elif drag_table and has_data_dir:
            info(f"[main] 拖拽模式 → 完整填充（R 流程文件夹）")
            _run_single_flow(arg_folder, drag_table, base_cfg)
        else:
            info(f"[main] 拖拽模式 → 独立模式")
            summary = _run_standalone(arg_folder, base_cfg)
            _print_summary(summary)
            _log_summary(summary)
        input("按任意键退出...")
        return

    # ── 双击模式 ──
    bat_dir = os.getcwd().rstrip("\\/.")
    bat_table = _find_material_table_in(bat_dir)
    has_data_dir = _is_r_flow_folder(bat_dir)
    if bat_table and not has_data_dir:
        _run_batch_mode(bat_table, base_cfg)
    elif bat_table and has_data_dir:
        info("[main] 双击模式 → 完整填充该 R 流程（不限审批状态）")
        _run_single_flow(bat_dir, bat_table, base_cfg)
    elif not bat_table and has_data_dir:
        info("[main] 双击模式 → 独立模式")
        summary = _run_standalone(bat_dir, base_cfg)
        _print_summary(summary)
        _log_summary(summary)
    else:
        print("[错误] 无法确定运行模式：当前文件夹既无物料信息表，也非 R 流程文件夹")
        print(f"       当前目录：{bat_dir}")
    input("按任意键退出...")


if __name__ == "__main__":
    main()