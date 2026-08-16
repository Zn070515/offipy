> [English](diagram.en.md)

# 图表 API

### `build`

Mermaid/drawio 源码文件 → 可编辑 PPTX（16:9 整页）。格式按扩展名+内容自动识别：.mmd/.md/.mermaid → Mermaid，.drawio → draw.io。Mermaid 仅支持 flowchart/graph、sequenceDiagram、stateDiagram-v2、erDiagram 四种 kind，其余（gantt/journey/mindmap/timeline 等）报 unsupported diagram kind，请改用 draw.io 表达。输出已存在时默认拒绝覆盖，需 overwrite=true 才替换。

- **参数**: `source: str`、`out: str`、`direction: str`、`page: int | str`、`overwrite: bool`
- **返回**: `dict`
- **标志**: 普通操作

---

### `install_skill`

把 diagram-design + offipy-diagram skill 安装到宿主 agent 技能目录（默认 ~/.claude/skills/，--target_dir 指定）。幂等：目标已存在则跳过（不覆盖用户编辑）；--force 删除并重建目标目录。

- **参数**: `target_dir: str`、`force: bool`
- **返回**: `dict`
- **标志**: 普通操作
