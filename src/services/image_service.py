"""图片处理服务（spec 模块 7）。

职责：统一图片处理服务，裁剪空白边距，不缩放，保持原始分辨率。

核心逻辑：
  1. 读取原始图片（PIL）。
  2. 裁剪空白边距（自动检测边缘空白像素）。
  3. 不缩放，保持原始像素尺寸。
  4. 保存为 PNG 格式，保持原始分辨率与 DPI。

特殊规则：
  - 不缩放：只裁剪空白边距，保持原始像素尺寸；Word 显示宽度由
    template_engine.insert_image 通过 shape.Width 控制。
  - 不修改 DPI，保持原始图片分辨率（电性图片的 1200DPI 导出由
    data_extractor 负责）。
  - 所有图片的空白边距裁剪由本模块统一完成（PIL 自动检测），涵盖
    电性数据图片与图片型章节图片，无需在 Word 中额外裁剪。
  - 若裁剪后面积小于原图 5%，怀疑误裁浅色内容，回退原图并打印警告。

边界：
  - 不负责判断图片是否应该被处理（由调用方决定）。
  - 不负责将图片插入 Word（由 template_engine 处理）。
  - 不负责设置 DPI（由 data_extractor 负责）。
  - 不负责缩放图片（由 template_engine 通过 shape.Width 控制显示尺寸）。
  - 处理后的临时文件输出到系统临时目录，由调用方在插入 Word 后通过
    ``cleanup_processed`` 删除（用完即删，避免污染原始记录目录）。
"""

import os
import tempfile
from PIL import Image

# 抑制 DecompressionBombWarning：高 DPI 图片像素数可能超过 PIL 默认上限（约 1.78 亿），
# 设置为 None 完全解除限制，或改为具体数值（如 500_000_000）以保留部分保护。
Image.MAX_IMAGE_PIXELS = None

from common.logger import info, warning

def _resize_for_word(img, target_width_cm, dpi=200):
    """按"显示宽度 × DPI"等比缩放到目标像素（只缩小，不放大）。"""
    target_px = int(target_width_cm / 2.54 * dpi)
    if target_px <= 0 or img.width <= target_px:
        return img
    ratio = target_px / img.width
    new_size = (target_px, max(1, int(img.height * ratio)))
    return img.resize(new_size, Image.LANCZOS)

# 处理后图片统一输出到系统临时目录，避免往「原始记录」目录写入 _processed.png
# 这类中间产物（既污染源数据，又在受限 / 沙箱环境下因无写权限而失败）。
# 该临时文件应在插入 Word 后由 cleanup_processed 删除（用完即删）。
# 临时目录（支持环境变量覆盖，用于多进程隔离）
_TASK_ID = os.environ.get("REPORT_PLATFORM_TASK_ID", "")
_TEMP_BASE = os.environ.get("REPORT_PLATFORM_TEMP", tempfile.gettempdir())
_PROCESSED_DIR = os.path.join(
    _TEMP_BASE,
    "report_platform_processed" + (f"_{_TASK_ID}" if _TASK_ID else "")
)
os.makedirs(_PROCESSED_DIR, exist_ok=True)

def process_image(img_path: str, target_width_cm: float = None) -> str:
    """裁剪空白边距；若传入 target_width_cm 则按显示宽度重采样。

    保存格式：按原图格式
      - 原图 .jpg / .jpeg → 存 JPEG（质量 90）
      - 其他（.png / .bmp 等） → 存 PNG

    Args:
        img_path: 原始图片路径。
        target_width_cm: 图片在 Word 里的显示宽度（厘米）。
                         传入则按 200 DPI 重采样（只缩不放）；
                         不传则不压缩（向后兼容）。

    Returns:
        处理后的图片路径（位于系统临时目录 ``_PROCESSED_DIR``）。
        该临时文件应在插入 Word 后由 ``cleanup_processed`` 删除。
    """
    try:
        img = Image.open(img_path)
    except Exception as exc:
        warning(f"[image_service] 无法打开图片 {img_path}: {exc}")
        raise

    # 保存原始 DPI，后续保存时原样写回
    orig_dpi = img.info.get("dpi")

    # 裁剪白边
    img = _trim_white_borders(img)

    # 按显示宽度重采样（只缩不放）
    if target_width_cm and target_width_cm > 0:
        img = _resize_for_word(img, target_width_cm, dpi=200)

    base_name = os.path.splitext(os.path.basename(img_path))[0]
    if base_name.endswith("_processed"):
        base_name = base_name[:-10]

    # 按原格式保存
    _ext = os.path.splitext(img_path)[1].lower()
    if _ext in (".jpg", ".jpeg"):
        processed = os.path.join(_PROCESSED_DIR, f"{base_name}_processed.jpg")
        save_kwargs = {"format": "JPEG", "quality": 90, "optimize": True}
    else:
        processed = os.path.join(_PROCESSED_DIR, f"{base_name}_processed.png")
        save_kwargs = {"format": "PNG"}

    if orig_dpi is not None:
        save_kwargs["dpi"] = orig_dpi

    try:
        img.save(processed, **save_kwargs)
    except Exception as exc:
        warning(f"[image_service] 保存图片失败 {processed}: {exc}")
        raise
    return processed


def cleanup_processed(img_path: str) -> None:
    """删除 ``process_image`` 产生的临时处理图（用完即删）。

    仅在文件确实位于本模块专用临时目录 ``_PROCESSED_DIR`` 内时才删除，
    避免误删用户原始图片或其它路径的文件（安全护栏）。
    删除失败（文件不存在 / 仍被 Word 占用）时仅告警，不影响主流程。
    """
    if not img_path:
        return
    try:
        abs_path = os.path.abspath(img_path)
        if os.path.dirname(abs_path) != os.path.abspath(_PROCESSED_DIR):
            # 非本模块产生的临时文件，不删除（安全护栏）
            return
        if os.path.exists(abs_path):
            os.remove(abs_path)
            info(f"[image_service] 已清理临时处理图: {abs_path}")
    except Exception as exc:
        warning(f"[image_service] 清理临时处理图失败 {img_path}: {exc}")


def _trim_white_borders(img: Image.Image) -> Image.Image:
    """裁剪图片四周的空白边距（自动检测边缘接近白色的区域）。"""
    try:
        gray = img.convert("L")
    except Exception as exc:
        warning(f"[image_service] 灰度转换失败: {exc}")
        return img

    # 阈值为 245：保留浅灰色内容（如 250 可能被误裁），
    # 同时仍能裁剪掉绝大多数白边（纯白=255）。
    # 反色处理：空白(>245) → 0，深色 → 255
    try:
        bbox = gray.point(lambda p: 0 if p > 245 else 255).getbbox()
    except Exception as exc:
        warning(f"[image_service] 空白边距检测失败: {exc}")
        return img

    if bbox:
        try:
            cropped = img.crop(bbox)
        except Exception as exc:
            warning(f"[image_service] 裁剪失败: {exc}")
            return img
        # 安全回退：若裁剪后面积小于原图 5%，怀疑误裁，回退原图
        orig_w, orig_h = img.size
        cr_w = bbox[2] - bbox[0]
        cr_h = bbox[3] - bbox[1]
        if cr_w * cr_h < orig_w * orig_h * 0.05:
            warning(
                f"[image_service] 裁剪后面积过小 ({cr_w}x{cr_h} vs {orig_w}x{orig_h})，"
                "怀疑误裁浅色内容，回退原图"
            )
            return img
        return cropped
    return img
