> [中文](ppt.md)

# PowerPoint API

### `new_pres`

Create a new blank presentation, set it active, return doc_id.

- **Parameters**: _none_
- **Returns**: `str`
- **Flags**: normal operation

---

### `open_pres`

Open an existing .pptx file, set it active, return doc_id.

- **Parameters**: `path: str`
- **Returns**: `str`
- **Flags**: normal operation

---

### `save`

Save the presentation (doc_id defaults to the active one) and return the absolute path. If path is given, save-as to that path (.pptx); otherwise save back to the original path (a never-saved document auto-saves to the user data directory without the Save As dialog); overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `str`
- **Flags**: mutates document/app state

---

### `save_pdf`

Export the presentation (doc_id defaults to the active one) to PDF at the given path; overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: normal operation

---

### `export_slides`

Export each slide of the presentation (doc_id defaults to the active one) as PNG into out_dir (slide_01.png…), for visual inspection/iteration. Default 1920x1080. Returns the list of file paths.

- **Parameters**: `out_dir: str`, `width: int`, `height: int`, `doc_id: str`
- **Returns**: `list`
- **Flags**: normal operation

---

### `add_slide`

Add a slide at the end. layout is {1: Title, 2: Title and Content, 5: Title Only, 12: Blank}, default 2. Returns the current total slide count.

- **Parameters**: `layout: int`, `doc_id: str`
- **Returns**: `int`
- **Flags**: mutates document/app state

---

### `set_title`

Set the title text of the slide_idx-th slide; auto-adds a text box when no title placeholder exists. Returns the shape ID actually modified.

- **Parameters**: `slide_idx: int`, `text: str`, `doc_id: str`
- **Returns**: `int`
- **Flags**: mutates document/app state

---

### `set_body`

Set the body placeholder text of the slide_idx-th slide; lines is a list of strings, one per line. Auto-adds a text box when no body placeholder exists. Returns the shape ID actually modified.

- **Parameters**: `slide_idx: int`, `lines: any`, `doc_id: str`
- **Returns**: `int`
- **Flags**: mutates document/app state

---

### `set_notes`

Write the speaker notes of the slide_idx-th slide. Returns the shape ID actually modified.

- **Parameters**: `slide_idx: int`, `text: str`, `doc_id: str`
- **Returns**: `int`
- **Flags**: mutates document/app state

---

### `add_textbox`

Add a free text box on the slide_idx-th slide (coordinates in points).

- **Parameters**: `slide_idx: int`, `left: float`, `top: float`, `width: float`, `height: float`, `text: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `add_picture`

Insert an image on the slide_idx-th slide (coordinates in points).

- **Parameters**: `slide_idx: int`, `path: str`, `left: float`, `top: float`, `width: float`, `height: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `read_slide_texts`

Read the title/body/notes text of each slide of the presentation (doc_id defaults to the active one) (read-only), returning [{index, title, body, notes}].

- **Parameters**: `doc_id: str`
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

Identity of the active presentation (app/doc_id/name/path); null if none. Pass doc_id to query a specific presentation.

- **Parameters**: `doc_id: str`
- **Returns**: `dict`
- **Flags**: read-only

---

### `quit`

Quit the PowerPoint session (close the application window). Refuses by default when attached to an existing Office instance; force=True overrides.

- **Parameters**: `force: bool`
- **Returns**: `void`
- **Flags**: mutates document/app state
