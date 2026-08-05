> [中文](compatibility.md)

# Compatibility Matrix (P2-1)

offipy is a Windows-only Office COM automation library. The core package has zero platform
dependencies, with capabilities installed incrementally by extra; each extra has its own
platform/version requirements, as shown in the table below.

## Three-Column Overview: Tested / Expected / Unsupported

| Dimension | ✅ Tested (verified on this machine) | 🟡 Expected (reasonably expected) | 🚫 Unsupported (explicitly unsupported) |
|------|------------------------|--------------------------|-------------------------------|
| Operating system | **Windows 11 Pro for Workstations** (build 26200, x64) | Windows 10 x64, Windows Server 2019+ (desktop session) | macOS, Linux (Office COM has no implementation; `import offipy` works, but any operation that triggers Office raises `UnsupportedPlatformError`) |
| Office | **Microsoft 365** (16.0.20228.20124, all three apps verified via COM) | Office 2016 / 2019 / 2021 (object model differences per version follow M365 behavior) | Office for Mac / web edition (no COM) |
| Python | **3.12** (development and testing) | 3.10 / 3.11 / 3.13 (`requires-python >=3.10`, covered by the CI matrix) | < 3.10 |
| deck rendering | **chromium (Playwright)** on Windows | Non-Windows pure rendering is possible | — |

> "Tested" = combinations actually run by this repo's CI / real-machine smoke tests;
> "Expected" = based on official semantics and conservative inference, not individually verified
> per version; "Unsupported" = not provided by the architecture. Choose based on the Tested column.

## Core / Common

| Dimension | Supported scope | Notes |
|------|----------|------|
| Python | 3.10 – 3.13 | `requires-python >=3.10`; development and testing on 3.12 |
| Operating system | Windows | Both COM automation and the MCP server require Windows |
| Core dependency | Only `tomli` (<3.11) | `import offipy` has zero additional dependencies; runs on pure standard library |

On non-Windows platforms, the core modules (the pure server/client/schema/CLI modules) can be
imported and tested (`offipy check` reports missing extras), but any operation that triggers
Office is unavailable.

## Windows Versions

| Windows | COM automation | deck pipeline | Notes |
|---------|-----------|-----------|------|
| Windows 10 (x64) | 🟡 Expected | 🟡 Expected | Reasonably expected, not individually verified |
| Windows 11 (x64) | ✅ Tested | ✅ Tested | Development environment (11 Pro for Workstations, verified on this machine) |
| Server 2019+ | 🟡 Expected | 🟡 Expected | Reasonably expected; COM requires a desktop session, Server Core has no Office GUI so it is unavailable |

## Office Versions

| Office | Word | Excel | PowerPoint | Notes |
|--------|------|-------|------------|------|
| Office 2016 | 🟡 Expected | 🟡 Expected | 🟡 Expected | Minimum supported version; reasonably expected, not individually verified |
| Office 2019 | 🟡 Expected | 🟡 Expected | 🟡 Expected | Reasonably expected, not individually verified |
| Office 2021 / LTSC | 🟡 Expected | 🟡 Expected | 🟡 Expected | Reasonably expected, not individually verified |
| Microsoft 365 | ✅ Tested | ✅ Tested | ✅ Tested | Primary development and verification environment (verified on this machine) |

The COM-based object model has minor differences across versions (e.g., constant enums, a few
new properties); when differences arise, Microsoft 365 behavior is authoritative and is recorded
in `CHANGELOG.md`.

## Extra Support Matrix

| Extra | Dependencies | Platform | Capabilities |
|-------|------|------|------|
| (core) | `tomli` | Any | `import offipy`, CLI meta-commands, pure server/client modules |
| `office` | `pywin32` | Windows only | Session-based Word/Excel/PowerPoint driving (all ops) |
| `deck` | python-pptx / lxml / fonttools / playwright / Pillow | Windows (rendering requires chromium) | HTML→editable PPTX pipeline (`deck make/outline/render`) |
| `mcp` | mcp SDK | Any (needs Windows + office when the service consumes Office) | `offipy mcp`, for Claude Desktop and other integrations |
| `all` | The three above combined | As needed | One-command install of everything via `pip install offipy[all]` |

The `deck` pipeline needs chromium installed for first use: `playwright install chromium`.
Playwright rendering also works on non-Windows, but deck output is often fed back into Office
sessions, so overall support remains Windows-based.

## Installation

```bash
pip install "offipy[all]"        # 或按用途：offipy[office] / offipy[deck] / offipy[mcp]
# deck 首次：
playwright install chromium
```

`offipy check` probes extra availability item by item and reports any missing ones.
