# ReportPlatform

An automated test report generation tool for electronic component qualification reports. It reads material information from an Excel table, matches raw test data folders, fills a Word template chapter by chapter, and saves the finished report.

## Features

- Batch processing: handles multiple part numbers in one run
- Plugin-style strategy system: image insertion, electrical table paste, reliability cycle, RoHS compliance
- Automatic chapter sync: adds or removes chapters based on available raw data folders
- Calibration expiry check: highlights expired calibration dates in red and logs a warning
- Word shortcut output: creates `.lnk` shortcuts organized by date and material category
- Structured logging: console shows full progress; log file records warnings and errors only

## Requirements

- Windows OS
- Python 3.9+
- Microsoft Word (COM automation via `pywin32`)

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

Edit `src/configs/config.yaml` before first run:

```yaml
target_base:              # Root folder containing numbered report folders
material_table:           # Directory containing 物料信息表.xlsx
default_chapter_template: # Chapter name to use as template when adding new chapters
pdf_output_dir:           # (Optional) Root directory for Word shortcut output
log_root:                 # (Optional) Root directory for log files
```

## Usage

Double-click `一键运行.bat`, or run from the command line:

```bash
cd src
python main.py
```

Results are saved to `output/` inside each numbered folder.

## Project Structure

```
src/
├── main.py                 # Entry point
├── config_loader.py        # Config and material table loading
├── section_manager.py      # Chapter discovery and context assembly
├── template_engine.py      # Word document manipulation
├── strategy_registry.py    # Strategy name → class mapping
├── processor_registry.py   # Category → processor mapping
├── core/report_engine.py   # Per-report orchestration
├── strategies/             # Chapter fill strategies
├── processors/             # Pre/post processors (sample, summary)
├── services/               # Image processing, reference data lookup
└── configs/config.yaml     # User configuration
docs/数据库/               # Material category configs and lookup tables
```

## Adding a New Material Category

1. Create `docs/数据库/物料类别配置/{category_code}.md` following the format in `category-example.md`.
2. No code changes required.

## License

MIT — see [LICENSE](LICENSE).
