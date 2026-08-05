> [English](ppt.en.md)

# PowerPoint API

### `new_pres`

新建空白演示文稿，设为活动，返回 doc_id。

- **参数**: _无参数_
- **返回**: `str`
- **标志**: 普通操作

---

### `open_pres`

打开现有 .pptx，设为活动，返回 doc_id。

- **参数**: `path: str`
- **返回**: `str`
- **标志**: 普通操作

---

### `save`

保存演示文稿（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到该路径（.pptx）；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，不弹另存为）；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `str`
- **标志**: 会改动文档/应用状态

---

### `save_pdf`

把演示文稿（doc_id 缺省为活动）导出为 PDF 到指定路径；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 普通操作

---

### `export_slides`

把演示文稿（doc_id 缺省为活动）每一页导出为 PNG 到 out_dir（slide_01.png…），供视觉检查/迭代。默认 1920x1080。返回文件路径列表。

- **参数**: `out_dir: str`、`width: int`、`height: int`、`overwrite: bool`、`doc_id: str`
- **返回**: `list`
- **标志**: 普通操作

---

### `add_slide`

在末尾添加一张幻灯片。layout 取 {1:标题, 2:标题+文本, 5:仅标题, 12:空白}，默认 2。返回当前总页数。

- **参数**: `layout: int`、`doc_id: str`
- **返回**: `int`
- **标志**: 会改动文档/应用状态

---

### `set_title`

设置第 slide_idx 张幻灯片的标题；无标题占位符自动建框，返回 shape ID。

- **参数**: `slide_idx: int`、`text: str`、`doc_id: str`
- **返回**: `int`
- **标志**: 会改动文档/应用状态

---

### `set_body`

设置第 slide_idx 张幻灯片的正文；无正文占位符自动建框，返回 shape ID。

- **参数**: `slide_idx: int`、`lines: any`、`doc_id: str`
- **返回**: `int`
- **标志**: 会改动文档/应用状态

---

### `set_notes`

写入第 slide_idx 张幻灯片的演讲者备注，返回 shape ID。

- **参数**: `slide_idx: int`、`text: str`、`doc_id: str`
- **返回**: `int`
- **标志**: 会改动文档/应用状态

---

### `add_textbox`

在 slide_idx 页添加自由文本框（坐标单位为磅）。

- **参数**: `slide_idx: int`、`left: float`、`top: float`、`width: float`、`height: float`、`text: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `add_picture`

在 slide_idx 页插入图片（坐标单位为磅）。

- **参数**: `slide_idx: int`、`path: str`、`left: float`、`top: float`、`width: float`、`height: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `read_slide_texts`

读取第 slide_idx 页全部具有文本能力的 shape（含 group 内文本），返回 SlideTextRecord（shape_id/name/text/坐标/占位符/group 路径）。include_empty=True 连空文本 shape 也返回；recursive=False 不递归 group。

- **参数**: `slide_idx: int`、`include_empty: bool`、`recursive: bool`、`doc_id: str`
- **返回**: `list[SlideTextRecord]`
- **标志**: 只读

---

### `read_slide_summary`

逐页读取演示文稿（doc_id 缺省为活动）的标题/正文/备注摘要（只读），返回 [{index, title, body, notes}]。

- **参数**: `doc_id: str`
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

当前活动演示文稿身份（app/doc_id/name/path）；无则返回 null。可传 doc_id 查询指定演示文稿。

- **参数**: `doc_id: str`
- **返回**: `dict`
- **标志**: 只读

---

### `quit`

退出 PowerPoint 会话（关闭应用窗口）。连接的是既有 Office 实例时默认拒绝（不夺走用户正用的窗口），force=True 强制退出。

- **参数**: `force: bool`
- **返回**: `void`
- **标志**: 会改动文档/应用状态
