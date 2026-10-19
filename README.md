![License: MIT](https://img.shields.io/github/license/minna-sun/report-engine-zh-en)
![GitHub language count](https://img.shields.io/github/languages/count/minna-sun/report-engine-zh-en)
![Top language](https://img.shields.io/github/languages/top/minna-sun/report-engine-zh-en)
![Repo size](https://img.shields.io/github/repo-size/minna-sun/report-engine-zh-en)
![Last commit](https://img.shields.io/github/last-commit/minna-sun/report-engine-zh-en)
![Stars](https://img.shields.io/github/stars/minna-sun/report-engine-zh-en?style=social)

# ReportPlatform · 中英双语测试报告自动生成引擎

## 📖 项目简介 / Introduction

ReportPlatform 是一款基于**配置驱动与策略复用**的自动化测试报告生成引擎。支持多种器件类型与多种测试项的**零代码扩展**——只需添加数据配置，引擎即可自动识别并复用既有策略，核心代码完全无需改动，实现中英双语报告的秒级输出。

得益于高度解耦的五层架构，本引擎具备极强的**同构可迁移性**。除电子元器件外，亦可无缝适配至：半导体、汽车零部件（PPAP）、医疗器械注册检验、第三方检测（CNAS/CMA）、新能源电池（UN38.3）、化工与材料检测、建材工程质检、环境监测等行业的测试与合规报告生成。

---

*ReportPlatform is a config-driven, strategy-reusable automated report generation engine. It supports **zero-code extension** for various device types and test items — adding new items only requires data configurations, and the engine will auto-detect and reuse existing strategies without modifying the core code, delivering instant bilingual (ZH/EN) report output.*

*Built on a highly decoupled five-layer architecture, the engine is **architecturally portable**. Beyond electronic components, it can be seamlessly retargeted for compliance and test reporting across industries such as semiconductors, automotive parts (PPAP), medical device registration testing, third-party labs (CNAS/CMA), new-energy batteries (UN38.3), chemicals & materials, building materials, and environmental monitoring.*

## 📂 文档导航 / Documentation

*   [01. 顶层架构设计（为什么这么设计）](docs/01_Architecture.md)
*   [02. 文档引擎抽象层 (DocumentEngine)](docs/02_DocumentEngine.md)
*   [03. 人机协同开发纪律 (AI Development Rules)](docs/03_AI_Development_Rules.md)
*   [深度施工规范（模块边界 / 接口 / 数据流）](skill/%E9%87%8D%E6%9E%84%E7%AC%AC%E5%9B%9B%E7%89%88%EF%BC%88%E6%8B%86%E8%A7%A3%EF%BC%89.md)

---

# 报告生成系统 使用说明

## 简介

本工具用于自动生成器件承认测试报告。将测试原始数据和物料信息表放好后，运行程序即可批量生成格式统一的 Word 格式报告文件。

## 环境要求

*   **操作系统**：支持 Windows / macOS / Linux（跨平台）。
*   **Python 环境**：Python 3.8 及以上版本（需已安装并加入系统 PATH）。
*   **Office 依赖**：**无需安装 Microsoft Word**（系统已彻底解除 COM 依赖，核心引擎仅依赖纯 Python 库）。

> 注：如后续需要将 Word 报告单独导出为 PDF，才需本地安装 Microsoft Word（依赖 `win32com`，仅 Windows 可用），不影响核心报告生成与跨平台部署。

## 第一步：安装依赖

打开终端（Windows 可用 `cmd`，macOS / Linux 用 `Terminal`），执行以下命令：

```bash
pip install -r requirements.txt
```

只需安装一次。

## 第二步：修改配置文件

用文本编辑器打开 `src/configs/config.yaml`，修改以下路径为你本机的实际路径：

| 配置项 | 说明 |
|--------|------|
| `target_base` | 存放编号文件夹的根目录 |
| `material_table` | 物料信息表所在目录（表名通常为 `物料信息表.xlsx`） |
| `pdf_output_dir` | Word 报告快捷方式 (`.lnk`) 输出目录（可留空） |
| `log_root` | 日志保存目录（可留空） |

注意：路径中的斜杠请使用正斜杠 `/` 或双反斜杠 `\\`，避免转义错误。

## 第三步：运行

运行 `python src/main.py`（Windows 也可直接双击根目录的 `引擎工具.py` 一键运行），等待处理完成后查看控制台输出的汇总信息。

生成的报告将保存在各编号文件夹的 `output/` 子目录下。

## 常见问题 (FAQ)

**Q：运行后提示"未找到物料信息表"？**
A：请确认 `config.yaml` 中的 `material_table` 路径指向的目录下有名为 `物料信息表.xlsx` 的文件，文件名必须完全一致。

**Q：提示"报告模板不存在"或"发现多个 .docx 文件"？**
A：每个编号文件夹的根目录下应有且只有一个 `.docx` 模板文件。系统会自动忽略临时文件（如 `~$` 开头的文件），请检查是否有多余的 Word 文件。

**Q：提示"output/ 已有报告，跳过处理"？**
A：对应编号的 `output/` 目录下已存在报告文件，系统默认防覆盖。如需重新生成，请先手动删除 `output/` 目录下的旧报告。

**Q：报告生成后部分章节内容为空？**
A：请确认 `原始数据/` 子目录下各测试文件夹的命名，与物料类别配置表中的文件夹名一致（忽略数字前缀和状态后缀如 `-PASS`、`-REF`）。
