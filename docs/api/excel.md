> [English](excel.en.md)

# Excel API

### `new_book`

新建空白工作簿，设为活动，返回 doc_id。

- **参数**: _无参数_
- **返回**: `str`
- **标志**: 普通操作

---

### `open_book`

打开现有 .xlsx/.xls，设为活动，返回 doc_id。

- **参数**: `path: str`
- **返回**: `str`
- **标志**: 普通操作

---

### `close_book`

关闭工作簿（doc_id 缺省为活动）。save=True 先保存（从未保存过则自动落盘同层目录，不弹另存为）并返回保存路径；save=False 不保存不弹窗，返回 null。

- **参数**: `save: bool`、`doc_id: str`
- **返回**: `str|null`
- **标志**: 会改动文档/应用状态

---

### `save`

保存工作簿（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到该路径；未给 path 则存回原路径（从未保存过自动落盘同层目录，不弹另存为）；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `str`
- **标志**: 会改动文档/应用状态

---

### `save_pdf`

把工作簿（doc_id 缺省为活动）导出为 PDF 到指定路径；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 普通操作

---

### `add_sheet`

在工作簿（doc_id 缺省为活动）中新建工作表并命名。

- **参数**: `name: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_cell`

写入单元格值；sheet 传表名或序号，cell 如 'A1'。

- **参数**: `sheet: any`、`cell: str`、`value: any`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `get_cell`

读取单元格的值；sheet 传表名或序号，cell 如 'A1'。

- **参数**: `sheet: any`、`cell: str`、`doc_id: str`
- **返回**: `any`
- **标志**: 只读

---

### `set_range`

把二维值列表一次性写入 range_addr（如 'A1:C3'）。

- **参数**: `sheet: any`、`range_addr: str`、`values: any`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_col_width`

设置列宽；col 传列号（1 基）或列字母。

- **参数**: `sheet: any`、`col: any`、`width: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `format_cell`

格式化单元格。bold/italic 传布尔；size 字号；bg/fg 传 '#RRGGBB'；align 传 Excel 水平对齐常量。

- **参数**: `sheet: any`、`cell: str`、`bold: bool`、`size: float`、`italic: bool`、`bg: str`、`fg: str`、`align: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `merge_cells`

把 range_addr（如 'A1:B2'）合并为一个单元格，值保留在左上角。

- **参数**: `sheet: any`、`range_addr: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `unmerge_cells`

取消 range_addr 的合并。

- **参数**: `sheet: any`、`range_addr: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_border`

给 range_addr 设置边框。side 取 all/outside/inside 或 left/top/bottom/right/inside-h/inside-v；style 取 continuous/dash/dash-dot/dash-dot-dot/dot/double/none/slant-dash-dot；weight 取 hairline/thin/medium/thick；color 传 '#RRGGBB'。

- **参数**: `sheet: any`、`range_addr: str`、`side: str`、`style: str`、`weight: str`、`color: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `freeze_panes`

冻结 rows 行上方 + cols 列左侧；rows=0 且 cols=0 取消冻结。

- **参数**: `sheet: any`、`rows: int`、`cols: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `page_setup`

打印设置。orientation 取 portrait/landscape；paper 取 letter/a3/a4；fit_to_pages_wide/tall 传整数；margins 传字典（单位磅）；print_area 传 'A1:C10'；center_horizontally/center_vertically 传布尔；print_titles_rows 如 '$1:$2'。

- **参数**: `sheet: any`、`orientation: str`、`paper: str`、`fit_to_pages_wide: int`、`fit_to_pages_tall: int`、`margins: any`、`print_area: str`、`center_horizontally: bool`、`center_vertically: bool`、`print_titles_rows: str`、`print_titles_cols: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_conditional_format`

给 range_addr 加条件格式。rule 取 cell/databar/colorscale；cell 需 operator+value，colorscale 需 min_color/max_color。

- **参数**: `sheet: any`、`range_addr: str`、`rule: str`、`operator: str`、`value: any`、`value2: any`、`bg: str`、`fg: str`、`min_color: str`、`max_color: str`、`mid_color: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_row_height`

设置某一行的高度（单位磅）。

- **参数**: `sheet: any`、`row: int`、`height: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_number_format`

给 range_addr 设置数字格式，如 '#,##0.00' / '0.0%' / 'yyyy-mm-dd'。

- **参数**: `sheet: any`、`range_addr: str`、`fmt: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `autofit`

自动调整 range_addr 的列宽/行高；不传 range_addr 则调整已用区域。columns/rows 为布尔开关。

- **参数**: `sheet: any`、`range_addr: str`、`columns: bool`、`rows: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `read_range`

读取工作表 range_addr（如 'A1:C3'）的值，返回二维列表（行→列）。

- **参数**: `sheet: any`、`range_addr: str`、`doc_id: str`
- **返回**: `list`
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

当前活动工作簿身份（app/doc_id/name/path）；无则返回 null。可传 doc_id 查询指定工作簿。

- **参数**: `doc_id: str`
- **返回**: `dict`
- **标志**: 只读

---

### `quit`

退出 Excel 会话（关闭应用窗口）。

- **参数**: _无参数_
- **返回**: `void`
- **标志**: 会改动文档/应用状态
