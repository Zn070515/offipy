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

### `close_pres`

关闭演示文稿（doc_id 必须显式传入或 follow_active=True），不退出 PowerPoint。save=True 先保存（从未保存过则自动落盘用户数据目录，不弹另存为）并返回保存路径；save=False 不保存不弹窗，返回 null。

- **参数**: `save: bool`、`doc_id: str`
- **返回**: `str|null`
- **标志**: 会改动文档/应用状态

---

### `save`

保存演示文稿（doc_id 必须显式传入或 follow_active=True）并返回绝对路径。给 path 则另存到该路径（.pptx）；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，不弹另存为）；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `str`
- **标志**: 会改动文档/应用状态

---

### `save_pdf`

把演示文稿（doc_id 必须显式传入或 follow_active=True）导出为 PDF 到指定路径；overwrite=True 允许覆盖已存在文件。

- **参数**: `path: str`、`overwrite: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 普通操作

---

### `export_slides`

把演示文稿（doc_id 必须显式传入或 follow_active=True）每一页导出为 PNG 到 out_dir（slide_01.png…），供视觉检查/迭代。默认 1920x1080。返回文件路径列表。

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

### `set_shape_geometry`

修改第 slide_idx 页 shape_id 的几何：left/top/width/height/rotation（坐标单位磅、角度单位度），只更新传入属性，至少传一个。group 子元素 left/top 写幻灯片绝对坐标；旋转 group 内后代不支持改 left/top。width/height 必须 >0。

- **参数**: `slide_idx: int`、`shape_id: int`、`left: float`、`top: float`、`width: float`、`height: float`、`rotation: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_text`

整体替换第 slide_idx 页 shape_id 的文本（保留原字体样式）。无文本能力的 shape（图片/线条等）报错。

- **参数**: `slide_idx: int`、`shape_id: int`、`text: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_font`

设置第 slide_idx 页 shape_id 文本的字体：font_name/size/bold/italic/color（'#RRGGBB'）。至少传一个属性；整段文本统一生效。

- **参数**: `slide_idx: int`、`shape_id: int`、`font_name: str`、`size: float`、`bold: bool`、`italic: bool`、`color: str`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_fill`

设置第 slide_idx 页 shape_id 的填充：color 传 '#RRGGBB' 实色填充，transparency 传 0..1 透明度；都不传则清除填充。无填充能力报错。

- **参数**: `slide_idx: int`、`shape_id: int`、`color: str`、`transparency: float`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_outline`

设置第 slide_idx 页 shape_id 的轮廓：color '#RRGGBB'/width 磅/visible 布尔，至少传一个；visible 控制最终显示态。无轮廓能力报错。

- **参数**: `slide_idx: int`、`shape_id: int`、`color: str`、`width: float`、`visible: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_visible`

显示（true）或隐藏（false）第 slide_idx 页 shape_id。

- **参数**: `slide_idx: int`、`shape_id: int`、`visible: bool`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `delete_shape`

删除第 slide_idx 页 shape_id（顶层或 group 子元素，递归定位）。

- **参数**: `slide_idx: int`、`shape_id: int`、`doc_id: str`
- **返回**: `void`
- **标志**: 会改动文档/应用状态

---

### `set_shape_z_order`

把第 slide_idx 页 shape_id 在所在集合内移到 1-based 目标位 z（1=最底）。顶层在 slide.Shapes 内移动，group 子元素在父 GroupItems 内；z 超出 1..Count 报错，不做截断。

- **参数**: `slide_idx: int`、`shape_id: int`、`z: int`、`doc_id: str`
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

### `read_shapes`

读取第 slide_idx 页全部 shape 的结构化记录，返回 ShapeInfo（shape_id/name/类型/几何/填充/轮廓/文本/字体/占位符/group 路径/z-order）。recursive=False 只列顶层；group 后代（含嵌套）仅在 recursive=True 时展开。shape_id 严格：任何 shape 的 Id 读不到即抛错，绝不出 0。

- **参数**: `slide_idx: int`、`recursive: bool`、`doc_id: str`
- **返回**: `list[ShapeInfo]`
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
