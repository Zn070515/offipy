# Third-Party Notices

offipy 内嵌或依赖以下第三方组件。许可证原文随产物分发，本文件说明来源与用途。

## Vendored 代码（随 wheel / sdist 分发）

### HTML→可编辑 PPTX 转换器

- **项目**：html-to-editable-pptx（`src/offipy/_vendor/html_to_editable_pptx/`）
- **来源**：https://github.com/Hasasasa/html-to-editable-pptx
- **许可证**：MIT — Copyright (c) 2026 Phil
- **用途**：HTML-first 可编辑 PPTX 管线（`offipy.deck`）。渲染前经 chromium 预检，浏览器缺失时报
  `ConversionError` 并提示安装。
- **说明**：该转换器为「vector-first」实现，产物是原生可编辑的 `.pptx`（占位符/文本/图形均保留
  可编辑性），非扁平化图片页。此文件保留原项目 LICENSE 全文。
- **补丁（offipy 维护的第一处 vendored 补丁，#94）**：`scripts/measure.py` 装饰门新增
  `isDrawioPlaceholder` 判定——空 `<div class="drawio">` 占位（无背景无边框）也要测出
  bbox 供 deck 注入定位。改动保持最小，仅扩展一个判定条件。

### 社论图表 skill（diagram-design）

- **项目**：diagram-design（`src/offipy/_vendor/diagram-design/`）
- **来源**：https://github.com/cathrynlavery/diagram-design
- **上游 commit**：`f3622cf66a3c557cb2ead57b687a3c1ff63f5a2b`
- **许可证**：MIT — Copyright (c) 2025 Cathryn Lavery
- **用途**：27 种「社论级」图表（架构/流程图/时序/甘特等）设计系统与 HTML+SVG 模板，后续
  接入 deck 管线（HTML→可编辑 PPTX）作为图表生成能力；本轮仅 vendoring 骨架 + 版权声明。
- **说明**：只 vendor 运行时相关核心——`skills/diagram-design/`（SKILL.md + references +
  assets 模板 + scripts 解析器）、`LICENSE`、`THIRD_PARTY_LICENSES.md`、`README.md`、
  `docs/adr/`（设计决策）。跳过：插件市场元数据（`.claude-plugin/`/`.codex-plugin/`/
  `.agents/`）、agent 命令模板（`commands/`/`prompts/`）、上游发布/校验脚本（根 `scripts/`）、
  截图（`docs/screenshots/`）与项目治理文档（SECURITY/CONTRIBUTING/CODE_OF_CONDUCT）。
  skill 内嵌图标来自 Tabler(MIT)/Simple Icons(CC0)/log-z(MIT)/Devicon(MIT)，其许可证全文
  见 vendored `THIRD_PARTY_LICENSES.md`；品牌 logo 归各商标持有人，仅文档/说明性用途。
- **补丁（offipy 对 diagram-design 的第一处修改，#97）**：`skills/diagram-design/scripts/
  drawio_extract.py` 的 `Node` 增加 `font_size` 字段、`parse_page` 提取 `fontSize`——
  drawio 节点字号随容器缩放（层级不被拍平）。offipy 不自行解析 draw.io XML，故补丁落在
  vendored 提取器内；安全边界（DTD/ENTITY 拒绝、压缩上限）保持不变。

### 图标资产（Phosphor + Lucide 双集）

- **项目**：Phosphor Icons（`src/offipy/assets/icons/phosphor/`）
- **来源**：https://github.com/phosphor-icons
- **许可证**：MIT — Copyright (c) 2023 Phosphor Icons
- **用途**：deck 管线内建图标（`ph:` 前缀，256 viewBox，填充）

- **项目**：Lucide Icons（`src/offipy/assets/icons/lucide/`）
- **来源**：https://github.com/lucide-icons/lucide
- **许可证**：ISC — Copyright (c) 2026 Lucide Icons and Contributors
- **用途**：deck 管线内建图标（`lu:` 前缀，24 viewBox，线形，round 线帽渲染）

图标经 `scripts/fetch_icons.py` 从上游抓取并记录真实 commit sha（见
`src/offipy/assets/icons/README.md`），更新图标集时请复核上游许可证是否变化。

## Python 运行时依赖（pip 安装）

| 包 | 用途 | 许可 |
|----|------|------|
| pywin32 | Windows COM 自动化（Word/Excel/PowerPoint） | PSF |
| python-pptx | PPTX 读写（deck 管线产物构建） | MIT |
| lxml | XML 处理（python-pptx 依赖） | BSD-3-Clause |
| fonttools | 字体度量（deck 管线） | MIT |
| playwright | Chromium 渲染（HTML→图像测量） | Apache-2.0 |
| Pillow | 图像处理（deck 管线） | HPND |
| mcp | MCP server 协议运行时 | MIT |

完整依赖树见 `pyproject.toml` / `uv.lock`。

## 声明

以上第三方组件的权利归属其各自作者/版权方。offipy 本体（`src/offipy/` 非 vendored 部分）为
MIT 许可，见根目录 `LICENSE`。
