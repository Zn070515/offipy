> [English](word.en.md)

# Word API

### `new_doc`

新建空白文档，设为活动，返回 doc_id。

- **参数**: _无参数_
- **返回**: `str`
- **标志**: 普通操作

---

### `open_doc`

打开现有 .docx/.doc，设为活动，返回 doc_id。

- **参数**: `path: str`
- **返回**: `str`
- **标志**: 普通操作

---

### `close_doc`

关闭文档（doc_id 缺省为活动）。save=True 先保存（从未保存过则自动落盘用户数据目录，不弹另存为）并返回保存路径；save=False 不保存不弹窗，返回 null。

- **参数**: `save: bool`、`doc_id: str`
- **返回**: `str|null`
- **标志**: 会改动文档/应用状态

---

### `save`

保存文档（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到该路径；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，不弹另存为）；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `str`
- **标志**: 会改动文档/应用状态

---

### `save_pdf`

把文档（doc_id 缺省为活动）导出为 PDF 到指定路径；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 普通操作

---

### `write`

在文档末尾追加文本（不换行）。

- **参数**: `text: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `write_line`

在文档末尾追加一行文本（自动换行）。

- **参数**: `text: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_heading`

在文档末尾添加标题行并应用 Heading 样式（level 1-3）。

- **参数**: `text: str`、`level: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_table`

在文档末尾添加 rows x cols 表格，返回当前表格数。

- **参数**: `rows: int`、`cols: int`、`doc_id: str`
- **返回**: `int`
- **标志**: 会改动文档/应用状态

---

### `set_table_cell`

设置第 table_idx 个表格的 (row, col) 单元格文本（行列 1 基）。

- **参数**: `table_idx: int`、`row: int`、`col: int`、`text: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `format_text`

设置第 paragraph 段（1 基）的文字格式。bold/italic 传布尔；size 字号；name 字体名；color 传 '#RRGGBB'；underline 取 none/single/words/double/dotted/wavy；highlight 取 none/yellow/green/pink/red/blue/bright_green/turquoise。

- **参数**: `paragraph: int`、`bold: bool`、`italic: bool`、`size: float`、`name: str`、`color: str`、`underline: str`、`highlight: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `format_paragraph`

设置第 paragraph 段（1 基）的段落格式。alignment 取 left/center/right/justify；line_spacing 取 single/1.5/double/at_least/exactly/multiple；space_before/space_after/left_indent/first_line_indent 单位磅。

- **参数**: `paragraph: int`、`alignment: str`、`line_spacing: str`、`space_before: float`、`space_after: float`、`left_indent: float`、`first_line_indent: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_header_text`

设置第 section 节的页眉文本。

- **参数**: `text: str`、`section: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_footer_text`

设置第 section 节的页脚文本。

- **参数**: `text: str`、`section: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_page_number`

在页脚插入页码。alignment 取 left/center/right；可带 color '#RRGGBB' 和 size 字号（会清空既有页脚文本）。

- **参数**: `alignment: str`、`color: str`、`size: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `page_setup`

页面设置。orientation 取 portrait/landscape；paper 取 letter/legal/a3/a4/a5；left/right/top/bottom_margin 与 gutter 单位磅。

- **参数**: `orientation: str`、`paper: str`、`left_margin: float`、`right_margin: float`、`top_margin: float`、`bottom_margin: float`、`gutter: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `insert_toc`

在文档开头插入目录（基于标题样式，levels 控制最深标题级别）。

- **参数**: `levels: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `update_toc`

更新文档中的目录域（新增/删除标题后刷新页码）。

- **参数**: `doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_list`

在文档末尾追加 lines 列表；style 取 bullet（项目符号）/ numbered（编号）。

- **参数**: `lines: list`、`style: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `merge_table_cells`

把第 table_idx 个表格的 (start_row,start_col) 到 (end_row,end_col) 合并为一个单元格。

- **参数**: `table_idx: int`、`start_row: int`、`start_col: int`、`end_row: int`、`end_col: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_table_border`

给第 table_idx 个表格设置边框。style 取 none/single/dot/double；weight 取 0.25pt/0.5pt/0.75pt/1pt/1.5pt/2.25pt/3pt/4.5pt/6pt；color 传 '#RRGGBB'；sides 取 all/outside/inside 或 left/top/bottom/right/inside-h/inside-v。

- **参数**: `table_idx: int`、`style: str`、`weight: str`、`color: str`、`sides: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_table_col_width`

设置第 table_idx 个表格第 col 列的宽度（单位磅）。

- **参数**: `table_idx: int`、`col: int`、`width: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_table_row_height`

设置第 table_idx 个表格第 row 行的高度（单位磅）。rule 取 auto/at_least/exactly。

- **参数**: `table_idx: int`、`row: int`、`height: float`、`rule: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `autofit_table`

自动调整第 table_idx 个表格。behavior 取 content/window/fixed。

- **参数**: `table_idx: int`、`behavior: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `find_replace`

在全文执行查找替换。replace_all 为真则替换全部，否则只替换第一处；match_case/whole_word 可选。

- **参数**: `find: str`、`replace: str`、`match_case: bool`、`whole_word: bool`、`replace_all: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `insert_image`

在文档末尾插入图片。width/height 单位磅（省略则保持原尺寸）。

- **参数**: `path: str`、`width: float`、`height: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `insert_page_break`

在文档末尾插入分页符。

- **参数**: `doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `read_doc_text`

读取文档全文文本（只读，不修改状态）。

- **参数**: `doc_id: str`
- **返回**: `str`
- **标志**: 只读

---

### `activate`

把指定 doc_id 设为活动目标，后续缺省 doc_id 的操作作用在它上面。

- **参数**: `doc_id: str`
- **返回**: `void`
- **标志**: 普通操作

---

### `list_docs`

列出当前打开的文档表：{doc_id: {name, path, active}}（只报已登记句柄）。

- **参数**: _无参数_
- **返回**: `dict`
- **标志**: 只读

---

### `get_target`

当前活动文档身份（app/doc_id/name/path）；无则返回 null。可传 doc_id 查询指定文档。

- **参数**: `doc_id: str`
- **返回**: `dict`
- **标志**: 只读

---

### `quit`

退出 Word 会话（关闭应用窗口）。连接的是既有 Office 实例时默认拒绝（不夺走用户正用的窗口），force=True 强制退出。

- **参数**: `force: bool`
- **返回**: `void`
- **标志**: 会改动文档/应用状态
