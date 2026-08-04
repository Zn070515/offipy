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
