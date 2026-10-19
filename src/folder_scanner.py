# 验收备份 · 模块4 folder_scanner · 2026-07-28 16:51（含递归穿透月份目录）
# 该副本为模块验收通过的归档版本，仅供追溯；实际运行以 src/folder_scanner.py 为准。
"""文件夹扫描模块（spec 模块 4）。

职责：在数据目录下查找编号对应的目标文件夹（纯物理路径发现，不执行任何业务
条件筛选）。仅用记录的「编号」字段匹配，不校验审批状态/物料类别。

输入：
    - records    ：list[MaterialRecord]，由 config_loader 返回的已审批物料记录
    - target_base：数据目录路径（如 D:/.../AP）
    - recursive  ：是否穿透子目录（默认 True）
    - max_depth  ：递归最大下钻层数（仅 recursive=True 时生效，默认 4）

输出：list[ScanResult]

边界：
    - recursive=False 时仅扫描 target_base 第一层子文件夹（spec 原行为）；
    - recursive=True 时递归收集所有深度 ≤ max_depth 的子目录作为候选，
      以适配真实数据「AP 下多夹一层月份目录（1月~7月）」的结构；
    - 单个编号失败不影响其他编号；
    - 不负责判断文件夹内数据完整性；
    - 不负责查找文件夹内的具体文件（交给 template_engine）；
    - 不负责多匹配的处理策略（交给 main.py）。
"""
import os
from typing import List

from common.exception import FolderNotFoundError
from common.logger import info, warning
from models.material import MaterialRecord
from models.result import ScanResult


def _collect_candidate_dirs(target_base: str, recursive: bool, max_depth: int) -> List[str]:
    """收集参与匹配的候选目录路径。

    - recursive=False：仅 target_base 的直接子目录（spec 第一层行为）。
    - recursive=True ：BFS 递归收集深度 ≤ max_depth 的所有子目录。
      例：target_base=AP，真实结构 AP/6月/R0367840+...，
          number 文件夹相对 AP 处于第 2 层，max_depth≥2 即可命中。
    """
    if not recursive:
        return [
            os.path.join(target_base, d)
            for d in os.listdir(target_base)
            if os.path.isdir(os.path.join(target_base, d))
        ]

    collected: List[str] = []
    frontier: List[str] = [target_base]
    depth = 0
    while frontier and depth < max_depth:
        nxt: List[str] = []
        for parent in frontier:
            try:
                entries = os.listdir(parent)
            except OSError:
                continue
            for name in entries:
                p = os.path.join(parent, name)
                if os.path.isdir(p):
                    collected.append(p)
                    nxt.append(p)
        frontier = nxt
        depth += 1
    return collected


def scan_folders(
    records: List[MaterialRecord],
    target_base: str,
    recursive: bool = True,
    max_depth: int = 4,
) -> List[ScanResult]:
    """按编号在 target_base 下查找目标文件夹。

    recursive=True（默认）会穿透月份等中间层目录，适配真实数据；
    设 recursive=False 可还原 spec 仅第一层扫描的行为。
    """
    if not os.path.isdir(target_base):
        raise FolderNotFoundError(f"数据目录不存在: {target_base}")

    target_base = os.path.normpath(target_base)  # 统一路径分隔符，避免 / 和 \ 混用

    try:
        candidate_dirs = _collect_candidate_dirs(target_base, recursive, max_depth)
    except OSError as e:
        raise FolderNotFoundError(f"无法读取数据目录 {target_base}: {e}")

    mode = f"递归穿透(最大深度 {max_depth})" if recursive else "仅第一层"
    info(f"扫描模式: {mode}；候选目录数 = {len(candidate_dirs)}")

    results: List[ScanResult] = []
    for rec in records:
        number = rec.number.strip()
        # 仅用编号匹配文件夹名（不关心审批/类别）；命中任意深度的候选目录
        matched = [p for p in candidate_dirs if number and number in os.path.basename(p)]
        if matched:
            results.append(ScanResult(
                number=number, folder_paths=matched, status="ok", error=None))
            info(f"编号 {number} 匹配到 {len(matched)} 个文件夹")
        else:
            results.append(ScanResult(
                number=number, folder_paths=[], status="failed",
                error="目标文件夹不存在"))
            warning(f"编号 {number} 未找到目标文件夹")

    ok_count = sum(1 for r in results if r.status == "ok")
    info(f"扫描完成：共 {len(results)} 个编号，匹配成功 {ok_count} 个，"
         f"未找到 {len(results) - ok_count} 个")
    return results
