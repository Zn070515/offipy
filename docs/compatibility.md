> [English](compatibility.en.md)

# 兼容矩阵（P2-1）

offipy 是 Windows-only 的 Office COM 自动化库。核心包零平台依赖，能力按
extra 增量安装；各 extra 有各自的平台/版本要求，见下表。

## 三栏总览：Tested / Expected / Unsupported

| 维度 | ✅ Tested（本机实测） | 🟡 Expected（合理预期） | 🚫 Unsupported（明确不支持） |
|------|------------------------|--------------------------|-------------------------------|
| 操作系统 | **Windows 11 Pro for Workstations**（build 26200，x64） | Windows 10 x64、Windows Server 2019+（桌面会话） | macOS、Linux（Office COM 无实现；`import offipy` 可，触发 Office 抛 `UnsupportedPlatformError`） |
| Office | **Microsoft 365**（16.0.20228.20124，三件套 COM 实测） | Office 2016 / 2019 / 2021（对象模型按版本差异以 M365 行为为准） | Office for Mac / 网页版（无 COM） |
| Python | **3.12**（开发与测试） | 3.10 / 3.11 / 3.13（`requires-python >=3.10`，CI 矩阵覆盖） | < 3.10 |
| deck 渲染 | **chromium（Playwright）** on Windows | 非 Windows 纯渲染可行 | — |

> 「Tested」= 本仓库 CI / 真机冒烟实跑过的组合；「Expected」= 依赖官方语义与
> 保守推断，未逐版本实测；「Unsupported」= 架构上不提供。选型以 Tested 列为准。

## 核心 / 通用

| 维度 | 支持范围 | 说明 |
|------|----------|------|
| Python | 3.10 – 3.13 | `requires-python >=3.10`；开发与测试在 3.12 |
| 操作系统 | Windows | COM 自动化与 MCP server 均要求 Windows |
| 核心依赖 | 仅 `tomli`（<3.11） | `import offipy` 零额外依赖，纯标准库可运行 |

非 Windows 平台上核心模块（server/client/schema/CLI 纯模块）可 import、
可测（`offipy check` 会报告缺失的 extra），但任何触发 Office 的操作不可用。

## Windows 版本

| Windows | COM 自动化 | deck 管线 | 备注 |
|---------|-----------|-----------|------|
| Windows 10 (x64) | 🟡 Expected | 🟡 Expected | 合理预期，未逐版本实测 |
| Windows 11 (x64) | ✅ Tested | ✅ Tested | 开发环境（11 Pro for Workstations，本机实测） |
| Server 2019+ | 🟡 Expected | 🟡 Expected | 合理预期；COM 需桌面会话，Server Core 无 Office GUI 不可用 |

## Office 版本

| Office | Word | Excel | PowerPoint | 备注 |
|--------|------|-------|------------|------|
| Office 2016 | 🟡 Expected | 🟡 Expected | 🟡 Expected | 最低支持版本，合理预期未实测 |
| Office 2019 | 🟡 Expected | 🟡 Expected | 🟡 Expected | 合理预期未实测 |
| Office 2021 / LTSC | 🟡 Expected | 🟡 Expected | 🟡 Expected | 合理预期未实测 |
| Microsoft 365 | ✅ Tested | ✅ Tested | ✅ Tested | 开发与验证主力（本机实测） |

依赖 COM 的对象模型按版本存在细微差异（如常量枚举、个别新属性），
遇差异以 Microsoft 365 行为为准并在 `CHANGELOG.md` 记录。

## Extra 支持矩阵

| Extra | 依赖 | 平台 | 能力 |
|-------|------|------|------|
| （核心） | `tomli` | 任意 | `import offipy`、CLI 元命令、server/client 纯模块 |
| `office` | `pywin32` | 仅 Windows | Word/Excel/PowerPoint 会话驱动（全部 op） |
| `deck` | python-pptx / lxml / fonttools / playwright / Pillow | Windows（渲染需 chromium） | HTML→可编辑 PPTX 管线（`deck make/outline/render`） |
| `mcp` | mcp SDK | 任意（服务消费 Office 时需 Windows + office） | `offipy mcp`，Claude Desktop 等接入 |
| `all` | 以上三合一 | 按需 | `pip install offipy[all]` 一键全装 |

`deck` 管线首次使用需装 chromium：`playwright install chromium`。
playwright 渲染在非 Windows 亦可行，但 deck 产物常回灌 Office 会话，因此
整体仍按 Windows 支持。

## 安装

```bash
pip install "offipy[all]"        # 或按用途：offipy[office] / offipy[deck] / offipy[mcp]
# deck 首次：
playwright install chromium
```

`offipy check` 会逐项探测 extra 可用性并给出缺失提示。
