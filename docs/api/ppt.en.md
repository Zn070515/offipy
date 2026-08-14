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

### `close_pres`

Close the presentation (doc_id must be given explicitly or follow_active=True) without quitting PowerPoint. save=True saves first (a never-saved document auto-saves to the user data directory without the Save As dialog) and returns the saved path; save=False closes without saving or prompting and returns null.

- **Parameters**: `save: bool`, `doc_id: str`
- **Returns**: `str|null`
- **Flags**: mutates document/app state

---

### `save`

Save the presentation (doc_id must be given explicitly or follow_active=True) and return the absolute path. If path is given, save-as to that path (.pptx); otherwise save back to the original path (a never-saved document auto-saves to the user data directory without the Save As dialog); overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `str`
- **Flags**: mutates document/app state

---

### `save_pdf`

Export the presentation (doc_id must be given explicitly or follow_active=True) to PDF at the given path; overwrite=True allows overwriting an existing file.

- **Parameters**: `path: str`, `overwrite: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: normal operation

---

### `export_slides`

Export each slide of the presentation (doc_id must be given explicitly or follow_active=True) as PNG into out_dir (slide_01.png…), for visual inspection/iteration. Default 1920x1080. Returns the list of file paths.

- **Parameters**: `out_dir: str`, `width: int`, `height: int`, `overwrite: bool`, `doc_id: str`
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

### `set_shape_geometry`

Modify the geometry of shape_id on the slide_idx-th slide: left/top/width/height/rotation (coordinates in points, angle in degrees); only the passed attributes are updated and at least one is required. Group-child left/top are written as absolute slide coordinates; descendants of a rotated group reject left/top changes. width/height must be >0.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `left: float`, `top: float`, `width: float`, `height: float`, `rotation: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_text`

Replace the text of shape_id on the slide_idx-th slide (preserving the original font styles). Shapes without a text frame (pictures, lines, etc.) raise an error.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `text: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_font`

Set the font of shape_id text on the slide_idx-th slide: font_name/size/bold/italic/color ('#RRGGBB'). At least one attribute is required; the whole text range is affected.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `font_name: str`, `size: float`, `bold: bool`, `italic: bool`, `color: str`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_fill`

Set the fill of shape_id on the slide_idx-th slide: color takes '#RRGGBB' for a solid fill, transparency takes 0..1; passing neither clears the fill. Shapes without fill capability raise an error.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `color: str`, `transparency: float`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_outline`

Set the outline of shape_id on the slide_idx-th slide: color '#RRGGBB'/width in points/visible boolean. At least one is required; visible controls the final display state. Shapes without outline capability raise an error.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `color: str`, `width: float`, `visible: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_visible`

Show (true) or hide (false) shape_id on the slide_idx-th slide.

- **Parameters**: `slide_idx: int`, `shape_id: int`, `visible: bool`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `delete_shape`

Delete shape_id on the slide_idx-th slide (top-level or group child, resolved recursively).

- **Parameters**: `slide_idx: int`, `shape_id: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `set_shape_z_order`

Move shape_id on the slide_idx-th slide to the 1-based target position z within its containing collection (1 = bottom). Top-level shapes move within slide.Shapes; group children move within the parent GroupItems. z outside 1..Count raises an error (no clamping).

- **Parameters**: `slide_idx: int`, `shape_id: int`, `z: int`, `doc_id: str`
- **Returns**: `void`
- **Flags**: mutates document/app state

---

### `read_slide_texts`

Read every text-capable shape on the slide_idx-th slide (including text inside groups), returning SlideTextRecord entries (shape_id/name/text/coordinates/placeholder/group path). include_empty=True also returns text shapes with empty text; recursive=False skips groups.

- **Parameters**: `slide_idx: int`, `include_empty: bool`, `recursive: bool`, `doc_id: str`
- **Returns**: `list[SlideTextRecord]`
- **Flags**: read-only

---

### `read_slide_summary`

Read the title/body/notes summary of each slide of the presentation (doc_id defaults to the active one) (read-only), returning [{index, title, body, notes}].

- **Parameters**: `doc_id: str`
- **Returns**: `list`
- **Flags**: read-only

---

### `read_shapes`

Read structured records for every shape on the slide_idx-th slide, returning ShapeInfo entries (shape_id/name/type/geometry/fill/outline/text/font/placeholder/group path/z-order). recursive=False lists only the top level; group descendants (including nested) are expanded only when recursive=True. shape_id is strict: any shape whose Id is unreadable raises an error, never emitting 0.

- **Parameters**: `slide_idx: int`, `recursive: bool`, `doc_id: str`
- **Returns**: `list[ShapeInfo]`
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
