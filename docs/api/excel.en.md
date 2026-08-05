> [中文](excel.md)

# Excel API

### `new_book`

Create a new blank workbook, set it active, return doc_id.

- **Parameters**: _none_
- **Returns**: `str`
- **Flags**: normal operation

---

### `open_book`

Open an existing .xlsx/.xls file, set it active, return doc_id.

- **Parameters**: `path: str`
- **Returns**: `str`
- **Flags**: normal operation

---

### `close_book`

Close the workbook (doc_id defaults to the active one). With save=True it saves first (a never-saved document auto-saves to the same directory without the Save As dialog) and returns the save path; with save=False nothing is saved, no dialog, returns null.

- **Parameters**: `save: bool`, `doc_id: str`
- **Returns**: `str|null`
- **Flags**: mutates document/app state

---

### `save`

Save the workbook (doc_id defaults to the active one) and return the absolute path. If path is given, save-as to that path; otherwise save back to the original path (a never-saved document auto-saves to the same directory without the Save As dialog); overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `str`
- **Flags**: mutates document/app state

---

### `save_pdf`

Export the workbook (doc_id defaults to the active one) to PDF at the given path; overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_sheet`

Add a new worksheet to the workbook (doc_id defaults to the active one) and name it.

- **Parameters**: `name: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_cell`

Write a cell value; sheet takes a sheet name or index, cell like 'A1'.

- **Parameters**: `sheet: any`, `cell: str`, `value: any`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `get_cell`

Read a cell value; sheet takes a sheet name or index, cell like 'A1'.

- **Parameters**: `sheet: any`, `cell: str`, `doc_id: str`
- **Returns**: `any`
- **Flags**: read-only

---

### `set_range`

Write a 2-D list of values into range_addr (e.g. 'A1:C3') in one call.

- **Parameters**: `sheet: any`, `range_addr: str`, `values: any`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_col_width`

Set column width; col takes a column number (1-based) or a column letter.

- **Parameters**: `sheet: any`, `col: any`, `width: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `format_cell`

Format a cell. bold/italic take booleans; size is the font size; bg/fg take '#RRGGBB'; align takes an Excel horizontal-alignment constant.

- **Parameters**: `sheet: any`, `cell: str`, `bold: bool`, `size: float`, `italic: bool`, `bg: str`, `fg: str`, `align: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `merge_cells`

Merge range_addr (e.g. 'A1:B2') into one cell, keeping the value in the top-left.

- **Parameters**: `sheet: any`, `range_addr: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `unmerge_cells`

Unmerge range_addr.

- **Parameters**: `sheet: any`, `range_addr: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_border`

Set borders for range_addr. side is all/outside/inside or left/top/bottom/right/inside-h/inside-v; style is continuous/dash/dash-dot/dash-dot-dot/dot/double/none/slant-dash-dot; weight is hairline/thin/medium/thick; color takes '#RRGGBB'.

- **Parameters**: `sheet: any`, `range_addr: str`, `side: str`, `style: str`, `weight: str`, `color: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `freeze_panes`

Freeze the rows above row `rows` and the columns left of column `cols`; rows=0 and cols=0 unfreezes.

- **Parameters**: `sheet: any`, `rows: int`, `cols: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `page_setup`

Print settings. orientation is portrait/landscape; paper is letter/a3/a4; fit_to_pages_wide/tall take integers; margins takes a dict (in points); print_area takes 'A1:C10'; center_horizontally/center_vertically take booleans; print_titles_rows like '$1:$2'.

- **Parameters**: `sheet: any`, `orientation: str`, `paper: str`, `fit_to_pages_wide: int`, `fit_to_pages_tall: int`, `margins: any`, `print_area: str`, `center_horizontally: bool`, `center_vertically: bool`, `print_titles_rows: str`, `print_titles_cols: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_conditional_format`

Add conditional formatting to range_addr. rule is cell/databar/colorscale; cell needs operator+value, colorscale needs min_color/max_color.

- **Parameters**: `sheet: any`, `range_addr: str`, `rule: str`, `operator: str`, `value: any`, `value2: any`, `bg: str`, `fg: str`, `min_color: str`, `max_color: str`, `mid_color: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_row_height`

Set the height of a row (in points).

- **Parameters**: `sheet: any`, `row: int`, `height: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_number_format`

Set the number format for range_addr, e.g. '#,##0.00' / '0.0%' / 'yyyy-mm-dd'.

- **Parameters**: `sheet: any`, `range_addr: str`, `fmt: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `autofit`

Auto-fit column widths / row heights for range_addr; without range_addr it fits the used range. columns/rows are boolean toggles.

- **Parameters**: `sheet: any`, `range_addr: str`, `columns: bool`, `rows: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `read_range`

Read the values of worksheet range_addr (e.g. 'A1:C3'), returning a 2-D list (rows → columns).

- **Parameters**: `sheet: any`, `range_addr: str`, `doc_id: str`
- **Returns**: `list`
- **Flags**: read-only

---

### `activate`

Set the given doc_id as the active target; subsequent ops with a default doc_id act on it.

- **Parameters**: `doc_id: str`
- **Returns**: `void`
- **Flags**: normal operation

---

### `list_docs`

List the open-document table: {doc_id: {name, path, active}} (only registered handles).

- **Parameters**: _none_
- **Returns**: `dict`
- **Flags**: read-only

---

### `get_target`

Identity of the active workbook (app/doc_id/name/path); null if none. Pass doc_id to query a specific workbook.

- **Parameters**: `doc_id: str`
- **Returns**: `dict`
- **Flags**: read-only

---

### `quit`

Quit the Excel session (close the application window).

- **Parameters**: _none_
- **Returns**: `void`
- **Flags**: mutates document/app state
