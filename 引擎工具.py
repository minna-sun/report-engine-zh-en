"""一键运行入口（替代 一键运行.bat）。

作用：
  1. 把工作目录锁定为「本文件」所在目录（= bat 里的 cd /d "%~dp0"）
  2. 找到并 import src/main.py 的 main()，直接调用

可放位置：
  - 工具目录（批量模式）
  - 任意 R 流程文件夹（完整填充 / 独立模式）
两种场景都靠 os.chdir 把 cwd 锁到本文件所在目录。
"""
import importlib.util
import os
import subprocess
import sys

_PIP_MAP = {
    "yaml": "pyyaml",
    "docx": "python-docx",
    "openpyxl": "openpyxl",
    "PIL": "Pillow",
    "lxml": "lxml",
    "fitz": "PyMuPDF",
    "xlrd": "xlrd",
    "pylnk3": "pylnk3",   # ← 新增
}


def _ensure_deps():
    """启动前自检：缺少依赖时用当前 Python 的 pip 自动安装。"""
    missing = [m for m in _PIP_MAP if importlib.util.find_spec(m) is None]
    if not missing:
        return
    pkgs = [_PIP_MAP[m] for m in missing]
    print(f"[依赖自检] 缺少 {pkgs}，正在安装（约 1-2 分钟）...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *pkgs])
        print("[依赖自检] 安装完成，继续运行。")
    except subprocess.CalledProcessError as e:
        print(f"[依赖自检] 安装失败：{e}")
        print("请手动运行：pip install " + " ".join(pkgs))
        input("按任意键退出...")
        sys.exit(1)


_ensure_deps()

# ① 工作目录 = 本文件所在目录（关键：用「本文件」的 __file__，不是 main.py 的）
_HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(_HERE)

# ② 定位 src/main.py（对应 bat 的 LOCAL_MODE / STANDALONE_MODE）
_LOCAL_SRC = os.path.join(_HERE, "src")
if os.path.isfile(os.path.join(_LOCAL_SRC, "main.py")):
    _SRC = _LOCAL_SRC
else:
    # 本工具同级的 src/ 下未找到 main.py，说明仓库目录结构不完整
    print("[错误] 未在本工具同级找到 src/main.py，请确认仓库目录结构完整。")
    sys.exit(1)

# ③ 把 src 目录加入 import 搜索路径
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

# ④ 直接调用 main()（main.py 内部自己读 sys.argv，拖拽模式自动生效）
from main import main

if __name__ == "__main__":
    main()
