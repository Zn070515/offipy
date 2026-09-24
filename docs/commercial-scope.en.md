# Commercial 1.0 Scope Freeze (CF-6)

This page defines the product commitment boundary for commercial offipy 1.0. It freezes what is
supported and what is accepted for support; it does not delete experimental code from the OSS
snapshot or turn every source-visible capability into a commercial guarantee.

## Initial support matrix

| Dimension | Commercial 1.0 commitment | Notes |
| --- | --- | --- |
| Operating system | **Windows 11 x64** | The only clean-machine acceptance target |
| Office | **Microsoft 365 desktop Word / Excel / PowerPoint** | All three COM applications must be usable |
| Connectivity | Local loopback server + MCP / CLI | Defaults to `127.0.0.1` or `::1` only |
| Runtime | Installer-bundled runtime helpers and Chromium | Customers do not install Python, uv, git, or Playwright |
| Network | Local Office control does not require an external service | Licensing and download services are separate release concerns |

Windows 10, Office 2016/2019/2021, Windows Server, and non-Windows platforms may remain Expected
or Unsupported in the OSS compatibility matrix, but they are outside the commercial 1.0 clean-machine
acceptance promise. Do not advertise them as supported without corresponding real-machine evidence.

## Capability tiers

### Formal: supported at launch

- The existing session-based Word / Excel / PowerPoint COM surface: Excel 25, Word 32, PowerPoint 27,
  **84 operations in total**.
- Session persistence, active-document routing, and `doc_id` / `expected_target` / `follow_active`
  target binding.
- Consistent HTTP, CLI, and MCP error contracts, read-back verification, and idempotent retry
  semantics within one server process lifetime.
- Save, overwrite protection, PDF / slide export, and the existing shape-level PowerPoint editing.
- Stability, security boundaries, and diagnostics covered by the real-Office runner.

Formal work is limited to bug fixes, Office compatibility, security, packaging/licensing/diagnostics,
and product UX. Increasing the number of Office operations is not a 1.0 goal.

### Advanced: usable, but with a bounded promise

- HTML to editable PPTX generation.
- Native charts, icons, Mermaid / draw.io post-processing, and quality auditing.
- The current animation and slide-transition injection, including the same-slide `click` / `after`
  mutual-exclusion rule.
- Deck workflows that require extra resources, longer processing, or human visual review.

Advanced capabilities receive fixes for concrete bugs, but are not guaranteed for every HTML input,
Office version, complex animation combination, or third-party template. New effects and converters
are outside the commercial 1.0 scope.

### Experimental: visible in source, not a commercial promise

- Rules marked `experimental` in `offipy.art`, including `experimental_score` and `quality.score`.
- Feedback MLP `train`, `recommend`, `apply`, and `reschema`, plus feedback-driven severity changes.
- Models, training data formats, and experimental CLI paths used for research or internal evaluation.

These capabilities may remain in the OSS snapshot for developer use, but are not part of the default
commercial Agent tool contract, clean-machine acceptance, compatibility promise, or support SLA. Their
presence in the schema or CLI does not imply commercial availability.
Developers who explicitly need to try Experimental tools through MCP must set
`OFFIPY_MCP_INCLUDE_EXPERIMENTAL=1` before starting the MCP process; the commercial default does not set it.

## Freeze rules

From CF-6 onward, the commercial line accepts only:

1. Formal / Advanced bug fixes, compatibility, security, and performance-regression fixes;
2. frozen-runtime, installer, diagnostics, licensing, signing, and release work;
3. Product UX and documentation corrections that do not change existing semantics.

New Office operations, expanded experimental rules, new animation/chart/diagram types, macOS/Linux
Office support, and a public Python API promise are outside the 1.0 freeze and require a separate
commercial-version review.

## 1.0 acceptance boundary

On a clean Windows 11 x64 + Microsoft 365 desktop machine, with no Python, uv, or git installed,
the product must be able to:

1. Install and start the Offipy runtime;
2. Complete entitlement activation and local diagnostics;
3. Connect to the local server through MCP / CLI;
4. Complete representative read, write, read-back, and save flows in Word, Excel, and PowerPoint;
5. Generate and open an editable PPTX;
6. Export a diagnostic report suitable for support triage when something fails.

Experimental capabilities are not prerequisites for any of these steps.

> 中文版：[商业 1.0 范围冻结](commercial-scope.md)
