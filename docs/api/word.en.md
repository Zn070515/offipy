> [中文](word.md)

# Word API

### `new_doc`

Create a new blank document, set it active, return doc_id.

- **Parameters**: _none_
- **Returns**: `str`
- **Flags**: normal operation

---

### `open_doc`

Open an existing .docx/.doc file, set it active, return doc_id.

- **Parameters**: `path: str`
- **Returns**: `str`
- **Flags**: normal operation

---

### `close_doc`

Close the document (doc_id must be given explicitly or follow_active=True). With save=True it saves first (a never-saved document auto-saves to the user data directory without the Save As dialog) and returns the save path; with save=False nothing is saved, no dialog, returns null.

- **Parameters**: `save: bool`, `doc_id: str`
- **Returns**: `str|null`
- **Flags**: mutates document/app state

---

### `save`

Save the document (doc_id must be given explicitly or follow_active=True) and return the absolute path. If path is given, save-as to that path; otherwise save back to the original path (a never-saved document auto-saves to the user data directory without the Save As dialog); overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `str`
- **Flags**: mutates document/app state

---

### `save_pdf`

Export the document (doc_id must be given explicitly or follow_active=True) to PDF at the given path; overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: normal operation

---

### `write`

Append text at the end of the document (no newline).

- **Parameters**: `text: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `write_line`

Append a line of text at the end of the document (auto newline).

- **Parameters**: `text: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_heading`

Add a heading line at the end of the document and apply the Heading style (level 1-3).

- **Parameters**: `text: str`, `level: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_table`

Add a rows x cols table at the end of the document, returning the current table count.

- **Parameters**: `rows: int`, `cols: int`, `doc_id: str`
- **Returns**: `int`
- **Flags**: mutates document/app state

---

### `set_table_cell`

Set the (row, col) cell text of the table_idx-th table (rows/cols are 1-based).

- **Parameters**: `table_idx: int`, `row: int`, `col: int`, `text: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `format_text`

Set the text format of the paragraph-th paragraph (1-based). bold/italic take booleans; size is the font size; name is the font name; color takes '#RRGGBB'; underline is none/single/words/double/dotted/wavy; highlight is none/yellow/green/pink/red/blue/bright_green/turquoise.

- **Parameters**: `paragraph: int`, `bold: bool`, `italic: bool`, `size: float`, `name: str`, `color: str`, `underline: str`, `highlight: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `format_paragraph`

Set the paragraph format of the paragraph-th paragraph (1-based). alignment is left/center/right/justify; line_spacing is single/1.5/double/at_least/exactly/multiple or numeric 1/1.5/2; space_before/space_after/left_indent/first_line_indent are in points.

- **Parameters**: `paragraph: int`, `alignment: str`, `line_spacing: str | float`, `space_before: float`, `space_after: float`, `left_indent: float`, `first_line_indent: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_header_text`

Set the header text of the section-th section.

- **Parameters**: `text: str`, `section: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_footer_text`

Set the footer text of the section-th section.

- **Parameters**: `text: str`, `section: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_page_number`

Insert a page number in the footer. alignment is left/center/right; optional color '#RRGGBB' and size (font size) style only the page-number field. mode is replace (default; clears the footer then inserts the PAGE field, legacy behavior) / append (keeps existing footer text and appends the field idempotently) / standalone (keeps the text: left flows directly after it, center/right use a tab zone, clearing any pre-existing tab stops in the footer). mode is keyword-only.

- **Parameters**: `alignment: str`, `color: str`, `size: float`, `doc_id: str`, `mode: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `page_setup`

Page setup. orientation is portrait/landscape; paper is letter/legal/a3/a4/a5; left/right/top/bottom_margin and gutter are in points.

- **Parameters**: `orientation: str`, `paper: str`, `left_margin: float`, `right_margin: float`, `top_margin: float`, `bottom_margin: float`, `gutter: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `insert_toc`

Insert a table of contents at the start of the document (based on heading styles; levels controls the deepest heading level included).

- **Parameters**: `levels: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `update_toc`

Update the table-of-contents fields in the document (refresh page numbers after adding/removing headings).

- **Parameters**: `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_list`

Append a list of lines at the end of the document; style is bullet or numbered.

- **Parameters**: `lines: list`, `style: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `merge_table_cells`

Merge cells from (start_row,start_col) to (end_row,end_col) in the table_idx-th table.

- **Parameters**: `table_idx: int`, `start_row: int`, `start_col: int`, `end_row: int`, `end_col: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_table_border`

Set borders of the table_idx-th table. style is none/single/dot/double; weight is 0.25pt/0.5pt/0.75pt/1pt/1.5pt/2.25pt/3pt/4.5pt/6pt; color takes '#RRGGBB'; sides is all/outside/inside or left/top/bottom/right/inside-h/inside-v.

- **Parameters**: `table_idx: int`, `style: str`, `weight: str`, `color: str`, `sides: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_table_col_width`

Set the width of the col-th column of the table_idx-th table (in points).

- **Parameters**: `table_idx: int`, `col: int`, `width: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_table_row_height`

Set the height of the row-th row of the table_idx-th table (in points). rule is auto/at_least/exactly.

- **Parameters**: `table_idx: int`, `row: int`, `height: float`, `rule: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `autofit_table`

Auto-fit the table_idx-th table. behavior is content/window/fixed.

- **Parameters**: `table_idx: int`, `behavior: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `find_replace`

Find and replace throughout the document. replace_all replaces all occurrences, otherwise only the first; match_case/whole_word are optional.

- **Parameters**: `find: str`, `replace: str`, `match_case: bool`, `whole_word: bool`, `replace_all: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `insert_image`

Insert an image at the end of the document. width/height in points (omitted to keep the original size).

- **Parameters**: `path: str`, `width: float`, `height: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `insert_page_break`

Insert a page break at the end of the document.

- **Parameters**: `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `read_doc_text`

Read the full document text (read-only, does not modify state).

- **Parameters**: `doc_id: str`
- **Returns**: `str`
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

Identity of the active document (app/doc_id/name/path); null if none. Pass doc_id to query a specific document.

- **Parameters**: `doc_id: str`
- **Returns**: `dict`
- **Flags**: read-only

---

### `quit`

Quit the Word session (close the application window). Refuses by default when attached to an existing Office instance; force=True overrides.

- **Parameters**: `force: bool`
- **Returns**: `void`
- **Flags**: mutates document/app state
