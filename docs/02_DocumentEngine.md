# 02. 文档引擎抽象层 (DocumentEngine)

## 1. 设计定位

`DocumentEngine` 定义了文档操作的标准化抽象接口，是彻底摆脱 Microsoft Office / COM 依赖的核心层。所有具体实现必须遵循本接口定义的契约。

- **业务层（Processor / Strategy）不直接依赖本接口**，仅 `template_engine` 通过工厂获取其实例。
- **原则**：只说"如何操作文档"，不说"用什么技术实现"。

## 2. 核心技术实现：python-docx + OOXML

对于 `python-docx` 原生支持的操作（如插入段落、图片），直接使用 API；对于不支持的高级操作（如跨文档复制表格、删除段落节点），委托给底层 `ooxml_service` 直接操作底层 `lxml` 节点（解压 docx 包、操作 `document.xml` 后重新压缩）。

**彻底摒弃的功能**：

- **`find_bookmark()`**：`python-docx` 对书签支持极差，系统改用配置驱动 + 文本占位符定位。
- **`update_page_numbers()`**：页码更新依赖 Word 域计算，禁止在代码中实现。页码由 Word 模板和用户打开行为自然完成。

## 3. 接口契约（核心方法节选）

### 3.1 生命周期管理

- `open(source)`：读取 `.docx` 文件，返回文档对象。
- `save(doc, path)`：将文档保存到指定路径。
- `close(doc)`：释放资源（`python-docx` 为空操作，保持接口一致）。

### 3.2 文档导航 / 定位 (Navigation)

- `find_section(doc, section_name)`：在文档中查找加粗且文本包含指定名称的章节标题段落（忽略表格内）。
- `find_paragraph(doc, section, anchor_text)`：在章节内查找包含锚点文本的段落（支持全 / 半角冒号）。
- `find_placeholder(doc, section, placeholder)`：精确匹配占位符段落（如 `{front}`）。

### 3.3 图片与表格操作

- `insert_image(doc, anchor, image_data, width_cm)`：在锚点段落后新建段落插入图片，锁定纵横比、居中；若锚点在表格内，自动追加并调整行高防裁剪。
- `insert_table(doc, anchor, table_data, width_cm)`：插入表格，支持结构化数据与样式；合并单元格后自动清理多余空段落。
- `copy_table(...)`：跨文档深拷贝表格节点，并补全边框。
