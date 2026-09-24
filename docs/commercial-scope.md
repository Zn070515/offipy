# Commercial 1.0 范围冻结（CF-6）

本页是 offipy 商业 1.0 的产品承诺边界。它冻结“支持什么”和“售后验收什么”，不删除
开源快照中的实验代码，也不把所有源码中存在的能力都变成商业保证。

## 首发支持矩阵

| 维度 | 商业 1.0 承诺 | 说明 |
| --- | --- | --- |
| 操作系统 | **Windows 11 x64** | 全新机器验收的唯一目标平台 |
| Office | **Microsoft 365 桌面版 Word / Excel / PowerPoint** | 三件套 COM 会话均须可用 |
| 连接方式 | 本机 loopback server + MCP / CLI | 默认只绑定 `127.0.0.1` 或 `::1` |
| 运行时 | 安装包自带 Python/runtime helper 与 Chromium | 消费者机器不要求 Python、uv、git 或 Playwright |
| 网络 | 本地 Office 控制链路不依赖外部服务 | 授权与下载服务另按发布方案处理 |

Windows 10、Office 2016/2019/2021、Windows Server 和非 Windows 平台仍可在开源兼容矩阵中
列为 Expected 或 Unsupported，但不属于商业 1.0 的 clean-machine 验收承诺。没有对应真机
验收记录前，不在商品页写成“已支持”。

## 功能分层

### Formal：首发正式支持

- Word / Excel / PowerPoint 的现有会话式 COM 操作面：Excel 25、Word 32、PowerPoint 27，合计 **84 个操作**。
- 会话保活、活动文档与 `doc_id` / `expected_target` / `follow_active` 目标绑定。
- 本地 HTTP、CLI、MCP 三入口的一致错误契约、读回验证和同一 server 生命周期内的幂等重试语义。
- 保存、覆盖保护、PDF / slides 导出以及现有 PowerPoint shape 级编辑。
- 现有真 Office runner 覆盖的稳定性、安全边界和诊断输出。

Formal 功能只允许 bugfix、Office 兼容性修复、安全修复、打包/授权/诊断和产品 UX 改进；
不以新增 Office API 数量作为 1.0 目标。

### Advanced：可用但不扩大承诺

- HTML → 可编辑 PPTX 管线。
- 原生图表、图标、Mermaid / draw.io 图形后处理与质量审计。
- 当前动画与页面过渡注入（包括 `click` / `after` 的同页互斥规则）。
- 需要额外资源、较长处理时间或人工视觉复核的 deck 产出能力。

Advanced 能力会继续修复明确 bug，但不承诺覆盖所有 HTML、Office 版本、复杂动画组合或
任意第三方模板；新增效果与新转换器不属于商业 1.0 范围。

### Experimental：源码可见，商业不承诺

- `offipy.art` 中标记为 `experimental` 的规则、`experimental_score` / `quality.score`。
- feedback MLP 的 `train`、`recommend`、`apply`、`reschema` 及基于反馈的严重度调整。
- 任何仅用于研究、内部评估或未来产品决策的模型、训练数据格式和实验 CLI。

这些能力可以留在开源快照中供开发者试用，但不进入商业 1.0 的默认 Agent 工具契约、
clean-machine 验收、兼容性承诺或售后 SLA；其 schema/CLI 是否存在不代表商业可用性保证。
开发者若确实需要在 MCP 中试用 Experimental 工具，必须在启动 MCP 进程前显式设置
`OFFIPY_MCP_INCLUDE_EXPERIMENTAL=1`；商业默认配置不设置该变量。

## 冻结规则

从 CF-6 起，商业线只接受：

1. Formal / Advanced 的 bugfix、兼容性、安全、性能回归修复；
2. frozen runtime、安装器、诊断、授权、签名和发布链工作；
3. 不改变现有语义的产品 UX 与文档修正。

新增 Office 操作、扩展实验规则、增加新的动画/图表/图形类型、macOS/Linux Office 支持和
公共 Python API 承诺，均不在 1.0 冻结范围内，必须另开商业版本评审。

## 1.0 验收边界

全新 Windows 11 x64 + Microsoft 365 桌面版机器，在没有 Python、uv、git 的前提下，能够：

1. 安装并启动 Offipy runtime；
2. 完成授权与本机诊断；
3. 通过 MCP / CLI 连接本地 server；
4. 对 Word、Excel、PowerPoint 完成代表性读写、读回和保存流程；
5. 生成并打开一个可编辑 PPTX；
6. 在失败时导出可供售后定位的诊断报告。

Experimental 能力不作为以上任何一步的必要条件。

> English version: [Commercial 1.0 Scope Freeze](commercial-scope.en.md)
