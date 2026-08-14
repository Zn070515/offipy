---
name: offipy-diagram
description: 用 diagram-design 的设计能力把图转成 offipy 可编辑 PPTX。触发词：画一张架构图/流程图/时序图并转成 PPTX、diagram、mermaid、drawio。
---

# offipy-diagram：LLM 设计 → 可编辑 PPTX

把 diagram-design 的视觉设计能力与 offipy 的「图 → 可编辑 PPTX」转换桥接起来。
本 skill 是**产物契约桥接**，不 fork diagram-design——先加载它的设计指引，再按
offipy 的转换契约落地产物。

## 流程

1. **加载设计指引**：读取已安装的 `diagram-design` skill（`offipy diagram install_skill`
   的 `--target_dir` 目录，默认 `~/.claude/skills/`）下的 `SKILL.md` 与其
   `references/`（视觉类型 / 语义模式 / 配色 / 品牌），按编辑风格设计图。
2. **落地产物契约**：把设计写成 **Mermaid 源码文件（.mmd）或 draw.io XML（.drawio）**，
   **不要输出 HTML/SVG**——offipy 只转换 Mermaid/drawio。
3. **转换**：调用 offipy（三选一，见下）。

## 可转换子集（契约边界，必须遵守）

**Mermaid 只支持以下四种 kind**（vendored `mermaid_extract.SUPPORTED_KINDS`）：

- `flowchart` / `graph`
- `sequenceDiagram`
- `stateDiagram-v2`
- `erDiagram`

**不支持** `gantt` / `journey` / `mindmap` / `timeline` / `gitgraph` / `classDiagram`
等——若按 diagram-design 设计了这些类型，**改用 draw.io 表达**，或重构为上面四种
kind，否则 `mermaid_to_pptx` 会报 `unsupported diagram kind`。

draw.io（.drawio）由 `mxGraphModel` 的节点/边类型决定，表达边界宽。

## 转换 handoff（三选一）

**CLI**：

```bash
offipy diagram build --source design.mmd --out design.pptx
offipy diagram build --source design.drawio --out design.pptx
# 安装/更新 skill 本身（幂等；--force 覆盖用户编辑）
offipy diagram install_skill --target_dir ~/.claude/skills
```

**Python API**：

```python
from offipy.diagrams import mermaid_to_pptx
from offipy.drawio import drawio_to_pptx

mermaid_to_pptx("design.mmd", "design.pptx")
drawio_to_pptx("design.drawio", "design.pptx")
```

**MCP**：`diagram_build(source, out, direction?, page?)` / `diagram_install_skill(target_dir?, force?)`

## 要求

- `source` 必须是磁盘上**已存在**的 `.mmd` / `.drawio` 文件（先把设计落地成文件再转换）
- 输出是 16:9 整页**可编辑形状** PPTX——布局由 Mermaid/drawio 引擎重排，不追求像素级还原
