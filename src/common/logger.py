"""统一日志工具。

职责（spec 模块 2）：统一日志输出格式与级别控制，提供 ``info()`` / ``warning()`` /
``error()`` 等接口；支持根据 config 中的 ``log_level`` 控制级别；命令行模式与编程
调用模式可分别控制是否打印到控制台。

边界：不负责业务逻辑；不直接操作文件（除非显式配置 ``log_root`` 文件日志）。

实现说明：底层复用标准库 ``logging``，避免重复造轮子；对外暴露极简接口。
双 Handler 设计：控制台打印 INFO+（完整进度），日志文件只记录 WARNING+（仅警告/报错）。

v5.x 日志降噪（2026-09-18）：
  - 新增 ``_NOISE_PREFIXES`` 前缀黑名单，命中即丢弃；
  - 这些前缀都是模块内部调试性 warning，不是真警告；
  - 过滤后，界面/文件只保留真警告、错误、进度，阅读更清晰；
  - 需要临时排查时，调用 ``disable_noise_filter()`` 恢复全部输出。
"""

import logging
import sys
from typing import Optional, Union

_LOGGER_NAME = "report_platform"

_logger = logging.getLogger(_LOGGER_NAME)
_logger.addHandler(logging.NullHandler())  # 避免 "no handler" 警告

# 当前正在处理的流程号（由 main.py 在每个编号处理前注入，处理后清除）
_current_number: Optional[str] = None

# ───────────────────────── 噪音过滤 ─────────────────────────
# 命中以下前缀的 warning 一律丢弃（不是真警告，是模块内部调试信息）。
# 如后续需要排查，可调用 disable_noise_filter() 恢复。
_NOISE_PREFIXES = (
    # ── 上次已加的（保留）──
    "[docx_engine] find_section",
    "[docx_engine] find_paragraph",
    "[docx_engine] find_placeholder",
    "[docx_engine] find_table",
    "[docx_engine] insert_image anchor",
    "[docx_engine] insert_image：",
    "[docx_engine] clone_section",
    "[docx_engine] add_page_break",
    "[_is_chapter_title_para]",
    "[_has_page_break_before]",
    "[find_paragraph] 诊断",
    "[find_paragraph] 候选段落",
    "[insert_text_in_header] DEBUG",
    "[add_page_break] 章节=",
    "[load_folder_config]",
    "[section_manager] build_chapters",
    "[sync_chapters] list_sections",
    "[fill_statistics] 找到统计段落",
    # ── 本次新增 ──
    "[section_manager] test_item=",
    "[section_manager] load_folder_config 返回",
    "[data_extractor] force_header_row",
    "[data_extractor] require_header",
    "[image_service] 已清理临时处理图",
    "[excel_table_builder] require_header",
    "[excel_table_builder] 数据区域",
    "[excel_table_builder] 构建完成",
    "[fill_sample_description] 使用标签子文件夹",
)

_noise_filter_enabled = True


def disable_noise_filter() -> None:
    """关闭噪音过滤（排查问题时用；正常运行时不要调用）。"""
    global _noise_filter_enabled
    _noise_filter_enabled = False


def enable_noise_filter() -> None:
    """开启噪音过滤（默认开启）。"""
    global _noise_filter_enabled
    _noise_filter_enabled = True


def _is_noise(msg: str) -> bool:
    """判断消息是否命中噪音黑名单。"""
    if not _noise_filter_enabled:
        return False
    return any(msg.startswith(p) for p in _NOISE_PREFIXES)


# ───────────────────────── 基础接口 ─────────────────────────

def set_number(number: Optional[str]) -> None:
    """设置当前处理的编号（流程号），后续所有日志自动附带该编号前缀。"""
    global _current_number
    _current_number = number


def _prefix(msg: str) -> str:
    """若当前有流程号，在消息前加 [编号] 前缀。"""
    return f"[{_current_number}] {msg}" if _current_number else msg


def setup(level: Union[str, int] = "INFO", *,
          print_to_console: bool = True,
          log_root: Optional[str] = None) -> None:
    """配置日志输出（双 Handler）。

    Args:
        level: 日志级别，字符串（"DEBUG"/"INFO"/"WARNING"/"ERROR"）或 logging 整数。
               控制台 Handler 使用此级别；文件 Handler 固定 WARNING+，与此参数无关。
        print_to_console: True 时添加控制台 Handler（INFO+）；False 时无控制台输出（API 模式）。
        log_root: 日志根目录；非空时按日期自动创建子目录并写入 report.log（WARNING+）。
                  路径格式：{log_root}/{年份}/运行日志/{月份}/{日期}/report.log。
    """
    _level = logging.getLevelName(level) if isinstance(level, str) else level

    _logger.setLevel(logging.DEBUG)   # 全局最低，由各 Handler 独立控制输出级别
    _logger.propagate = False          # 禁止冒泡到 root logger，避免被外部 handler 重复/错误捕获
    _logger.handlers.clear()

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 1. 控制台 Handler：INFO+（完整进度信息）
    if print_to_console:
        sh = logging.StreamHandler(sys.stdout)
        sh.setLevel(_level)
        sh.setFormatter(fmt)
        _logger.addHandler(sh)

    # 2. 文件 Handler：WARNING+（仅警告和报错，按日期分目录）
    if log_root:
        import os
        from datetime import datetime
        now = datetime.now()
        log_dir = os.path.join(
            log_root,
            str(now.year),
            "运行日志",
            str(now.month),
            str(now.day),
        )
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, now.strftime("%Y%m%d%H%M%S") + ".log")
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.WARNING)   # 文件只记录 WARNING 和 ERROR
        fh.setFormatter(fmt)
        _logger.addHandler(fh)


def apply_config(config: Optional[dict] = None) -> None:
    """从全局配置字典解析 ``log_level`` / ``log_root`` 并配置日志。

    ``log_root`` 为日志根目录，脚本自动在其下创建
    「{年份}/运行日志/{月份}/{日期}/report.log」子路径，文件只记录 WARNING+。

    兼容两种 config 形态：
      - 扁平：``config["log_level"]`` / ``config["log_root"]``
      - 嵌套：``config["logging"]["level"]`` / ``config["logging"]["log_root"]``
    均未提供时回退到默认值（INFO / 无文件）。

    Args:
        config: config_loader 返回的全局配置字典（可为 None）。
    """
    if not config:
        setup()
        return
    level = (
        config.get("log_level")
        or (config.get("logging") or {}).get("level")
        or "INFO"
    )
    log_root = (
        config.get("log_root")
        or (config.get("logging") or {}).get("log_root")
        or None
    )
    setup(level=level, log_root=log_root)


def info(msg: str) -> None:
    """输出 INFO 级日志（命中噪音黑名单的前缀会被丢弃）。"""
    if _is_noise(msg):
        return
    _logger.info(_prefix(msg))


def warning(msg: str) -> None:
    """输出 WARNING 级日志（命中噪音黑名单的前缀会被丢弃）。"""
    if _is_noise(msg):
        return
    _logger.warning(_prefix(msg))


def error(msg: str) -> None:
    """输出 ERROR 级日志。"""
    _logger.error(_prefix(msg))


def debug(msg: str) -> None:
    """输出 DEBUG 级日志。"""
    _logger.debug(_prefix(msg))


def attach_run_log(log_dir: str) -> Optional[logging.Handler]:
    """为指定目录挂一个专属日志文件（只记录 WARNING+）。

    调用方（main.py）在每个流程开始时调用，流程结束时调用 ``detach_run_log()``。
    """
    import os
    if not log_dir:
        return None
    try:
        os.makedirs(log_dir, exist_ok=True)
        from datetime import datetime
        log_file = os.path.join(
            log_dir,
            datetime.now().strftime("%Y%m%d%H%M%S") + "_运行日志.log",
        )
        fmt = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.WARNING)   # 只记录 WARNING+
        fh.setFormatter(fmt)
        _logger.addHandler(fh)
        return fh
    except Exception:
        return None


def detach_run_log(handler) -> None:
    """摘掉 ``attach_run_log`` 挂上的 Handler，关闭文件。"""
    if handler is None:
        return
    try:
        _logger.removeHandler(handler)
        handler.close()
    except Exception:
        pass