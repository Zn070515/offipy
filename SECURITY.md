# Security Policy

offipy 通过本机 HTTP server 驱动真实的 Microsoft Office 应用。本文件说明其安全模型，
以及发现漏洞时的处理方式。

## 安全模型

### 常驻 server（端口 8890）

- **scope**：server 只监听 `127.0.0.1`（回环地址），不对外网开放。它持有 Office 的 COM 引用，
  因此**任何能访问该端口的进程都能驱动你当前的 Office 文档**。
- **Bearer token 鉴权**：启动时生成随机 token，优先读环境变量 `OFFIPY_SERVER_TOKEN`，
  否则从用户数据目录 `user_data_dir()/token` 读取/生成并持久化。所有请求必须携带
  `Authorization: Bearer <token>` 头，否则返回 401。
- **白名单**：只暴露各 app 类的公开方法（`dir(cls)` 中非下划线开头的成员），白名单外一律 4xx。
- **防滥用**：请求体上限 16MB（超限 413）、`Content-Type` 必须为 `application/json`（否则 415）。
- **健康端点**：`/ping` 免鉴权（仅握手用），`/status` 需鉴权，返回 `{version, protocol, pid, python, started_at}`。

### 会话语义

server 保持 Office 窗口与文档存活，op 作用在用户当前激活的文档上。**别把 token 泄漏给不信任的进程**——
拿到 token 即等于拿到你当前 Office 会话的读写权限。

### 数据落盘

- token：`user_data_dir()/token`（Windows 为 `%LOCALAPPDATA%\offipy`，其它平台 `~/.local/share/offipy`）
- 转换器数据：`OFFIPY_CONVERTER_DATA_DIR` 可覆盖，默认同上
- 反馈学习：`~/.offipy/feedback.jsonl`

## 报告漏洞

如果你发现了安全问题（不只是普通 bug），请不要在公开 issue 里披露细节，避免在修复前被利用。
请私信仓库维护者（见 `pyproject.toml` authors）或在 GitHub 上发起 **Private vulnerability report**
（仓库 → Security → Report a vulnerability）。

请包含：影响的版本、可复现步骤、影响评估（例如「未授权进程可读写 Office 文档」）、
如可行给出修复建议。
