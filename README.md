![License: MIT](https://img.shields.io/github/license/minna-sun/report-engine-zh-en)
![GitHub language count](https://img.shields.io/github/languages/count/minna-sun/report-engine-zh-en)
![Top language](https://img.shields.io/github/languages/top/minna-sun/report-engine-zh-en)
![Repo size](https://img.shields.io/github/repo-size/minna-sun/report-engine-zh-en)
![Last commit](https://img.shields.io/github/last-commit/minna-sun/report-engine-zh-en)
![Stars](https://img.shields.io/github/stars/minna-sun/report-engine-zh-en?style=social)

> **ReportPlatform · 中英双语测试报告自动生成引擎**
> 可覆盖多种器件类型与多种测试项：**0 代码扩展**——新增类型/测试项只需添加数据配置，引擎自动识别；采用**驱动配置 + 策略复用**的可迁移架构，可复用于其他领域的中英双语报告秒级输出。
> 架构同构可迁移至：电子元器件 / 半导体 / 汽车零部件(PPAP) / 医疗器械注册检验 / 第三方检测(CNAS·CMA) / 新能源电池(UN38.3) / 食品·化妆品安全检测 / 化工·材料检测 / 建材工程质检 / 环境监测 等行业的测试·检测·合规报告生成。
> **ReportPlatform** covers a wide range of device types and test items with **zero-code extension** — adding new types/items only requires adding data configs, which the engine auto-detects; built on a **config-driven, strategy-reusable** portable architecture, it retargets to other domains for second-level (instant) bilingual (ZH/EN) report output, and is reusable for test/inspection/compliance reporting across electronics, automotive, medical devices, third-party labs, new-energy batteries, food & cosmetics, chemicals, building materials, and environmental monitoring.
---

## 环境要求 / Requirements
Python 3.8+，跨平台支持 Windows / macOS / Linux。核心引擎仅依赖纯 Python 库（python-docx、openpyxl、lxml、Pillow、PyMuPDF、PyYAML），无需安装 Microsoft Office，无 win32com / COM 依赖。运行：`python src/main.py`（Windows 也可双击 `一键运行.bat`）。

> Note: optional docx→PDF export (`export_pdf`) needs local Microsoft Word + win32com and is Windows-only; it does not affect core report generation or cross-platform deployment. Set the paths in `config.yaml` to your local directories.


## 📂 文档导航 / Documentation

*   [01. 顶层架构设计（为什么这么设计）](docs/01_Architecture.md)
*   [02. 文档引擎抽象层 (DocumentEngine)](docs/02_DocumentEngine.md)
*   [03. 人机协同开发纪律 (AI Development Rules)](docs/03_AI_Development_Rules.md)
*   [深度施工规范（模块边界 / 接口 / 数据流）](skill/%E9%87%8D%E6%9E%84%E7%AC%AC%E5%9B%9B%E7%89%88%EF%BC%88%E6%8B%86%E8%A7%A3%EF%BC%89.md)

---

# 报告生成系统 使用说明

## 简介

本工具用于自动生成器件承认测试报告。将测试原始数据和物料信息表放好后，双击运行即可批量生成 Word 格式的报告文件。

---

## 环境要求

- Windows 操作系统
- Python 3.9 及以上版本（需已安装并加入系统 PATH）
- Microsoft Word（需已安装，工具通过 Word 自动化生成报告）

---

## 第一步：安装依赖

打开命令提示符（Win+R → 输入 `cmd` → 回车），执行以下命令：

```
py -m pip install -r requirements.txt
```

只需安装一次。

---

## 第二步：修改配置文件

用记事本打开 `src/configs/config.yaml`，修改以下路径为你本机的实际路径：

| 配置项 | 说明 |
|--------|------|
| `target_base` | 存放编号文件夹的根目录 |
| `material_table` | 物料信息表所在目录 |
| `pdf_output_dir` | Word 快捷方式输出目录（可留空） |
| `log_root` | 日志保存目录（可留空） |

路径中的斜杠请使用正斜杠 `/` 或双反斜杠 `\\`，**不要使用单个反斜杠 `\`**。

---

## 第三步：运行

运行 `python src/main.py`（Windows 也可双击 `一键运行.bat`），等待处理完成后查看控制台输出的汇总信息。

生成的报告保存在各编号文件夹的 `output/` 子目录下。

---

## 常见问题

**Q：运行后提示"未找到物料信息表"**
A：请确认 `config.yaml` 中的 `material_table` 路径指向的目录下有名为 `物料信息表.xlsx` 的文件，文件名必须完全一致。

**Q：提示"报告模板不存在"或"发现多个 .docx 文件"**
A：每个编号文件夹的根目录下应有且只有一个 `.docx` 模板文件。请检查是否有多余的 Word 文件或临时文件（`~$` 开头的文件会被自动忽略）。

**Q：提示"output/ 已有报告，跳过处理"**
A：对应编号的 `output/` 目录下已存在报告文件，系统不会覆盖。如需重新生成，请先手动删除 `output/` 目录下的旧报告。

**Q：报告生成后部分章节内容为空**
A：请确认 `原始数据/` 子目录下各测试文件夹的命名与物料类别配置表中的文件夹名一致（忽略数字前缀和状态后缀如 -PASS、-REF）。

**Q：运行时 Word 一直弹出恢复窗口**
A：工具会自动清理 Word 恢复注册表项。若仍弹出，请手动关闭 Word 后重新运行。
