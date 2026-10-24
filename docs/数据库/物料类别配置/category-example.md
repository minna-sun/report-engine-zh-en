# DEMO-001

# DEMO-001 物料类别配置

**性质**：DEMO-001 物料类别的完整配置（策略、书签、图片匹配、测试标准）。脚本运行时实际读取的数据源。   **维护规则**：当 DEMO-001 的测试项目、策略、书签、标准变化时更新本文件。   **版本**：1.0   **最后更新**：2026-09-30

---

## 文档信息

> **新增章节范本**：本表**只能有一行**填"是"；该行的章节将作为本物料类别新增章节的结构范本。 若本表无任何一行填"是"，则回退到 config.yaml 的 `default_chapter_template`。

---

| 测试项目<br>章节名（中文） | 测试项目章节名（英文） | 文件夹名 | 占位符列表 | 图片匹配说明 | 图片宽度 | 测试标准 | 类型 | Strategy |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 外观 | Appearance | 外观 | front,back | 1-1→front, 1-2→back | 7 | Datasheet | ImageInspection | ImageBasedStrategy |
| X-RAY | X-RAY | xray | _—_ | _—_ | 7 | GJB 4027B-2021, GJB 548C-2021 | ImageAuto | ImageAutoStrategy |
| 电性 | Electrical test | 电性 | _—_ | _—_ | _—_ | Datasheet | ReliabilityCycle | ReliabilityCycleStrategy |
| 可焊性测试 | Solderability | 可焊性 | solderoverview,solderenlarge | 1-2→solderoverview, 1-3→solderenlarge | 7 | GB/T 2423.28-2005, IEC 60068-2-58:2015, EIA IPC J-STD002 | ImageInspection | ImageBasedStrategy |
| 耐焊接热测试 | Resistance to Soldering Heat | 耐焊接热 | _—_ | _—_ | _—_ | GB/T 2423.28-2005, IEC 60068-2-58:2015 | ReliabilityCycle | ReliabilityCycleStrategy |
| 内部检查 | Internal structure inspection | 开封 | _—_ | _—_ | 7 | GJB 4027B-2021, GJB 548C-2021 | ImageAuto | ImageAutoStrategy |
| 键合强度 | Bonding strength | 键合强度 | wirebreak | 1-5→wirebreak | 8 | GJB 4027B-2021, GJB 548C-2021 | ImageInspection | ImageBasedStrategy |
| 弹坑检测 | Crater test | 弹坑 | crater | 1-1→crater | 12 | GJB 4027B-2021, GJB 548C-2021 | ImageAuto | ImageAutoStrategy |
| 高温反偏 | HTRB | HTRB | _—_ | _—_ | _—_ | IEC 60749-23:2025, JESD22-A108G, AEC-Q-101, MIL-STD-750-1 Method 1038 condition A | ReliabilityCycle | ReliabilityCycleStrategy |
| 高温栅偏 | HTGB | HTGB | _—_ | _—_ | _—_ | IEC 60749-23:2025, JESD22-A108G | ReliabilityCycle | ReliabilityCycleStrategy |
| 高温高湿反偏 | H3TRB | H3TRB | — | — | — | IEC 60749-5:2023, JESD22-A101D, EIAJ ED-4701 B122A | ReliabilityCycle | ReliabilityCycleStrategy |
| 温度冲击/循环 | Thermal shock | TC | — | — | — | IEC 60749-25:2003, JESD22-A104F, JESD22-A106B | ReliabilityCycle | ReliabilityCycleStrategy |
| 高加速应力测试 | HAST | HAST | — | — | — | IEC 60749-4:2017, JESD22-A110E | ReliabilityCycle | ReliabilityCycleStrategy |
| 高加速应力测试（无偏置） | UHAST | UHAST | — | — | — | IEC 60749-42:2014, JESD22-A118B | ReliabilityCycle | ReliabilityCycleStrategy |
| 恒定湿热 | Damp heat | 双85 | — | — | — | IEC 60749-42:2014, EIAJ ED-4701/100 Method 103 | ReliabilityCycle | ReliabilityCycleStrategy |
| 高温存储 | High temperature storage | 高温存储 | — | — | — | IEC 60749-6:2017, JESD22-A103E | ReliabilityCycle | ReliabilityCycleStrategy |
| 低温存储 | Low temperature storage | 低温存储 | — | — | — | JESD22-A119A, JEITA ED4701/202A | ReliabilityCycle | ReliabilityCycleStrategy |
| 引出端强度测试 | Robustness of terminations | 引出端强度 | — | — | — | GB/T 4937.14; IEC 60749-14 | ReliabilityCycle | ReliabilityCycleStrategy |
| RoHS测试 | RoHS Test | ROHS | — | — | — | RoHS 2011/65/EU & (EU) 2015/863 | Compliance | ComplianceStrategy |
| EDS 分析 | EDS analysis | EDS | — | — | 9 | ASTME1588-2016 | EDS | EDSStrategy |
