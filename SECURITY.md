# Security Policy

offipy 通过本机 HTTP server 驱动真实的 Microsoft Office 应用。本文件说明其安全模型，
以及发现漏洞时的处理方式。

## 安全模型

### 常驻 server（端口 8890）

- **scope**：server 只监听 `127.0.0.1`（回环地址），不对外网开放。它持有 Office 的 COM 引用，
  因此**任何能访问该端口的进程都能驱动你当前的 Office 文档**。
- **Bearer token 鉴权**：启动时生成随机 token，优先读环境变量 `OFFIPY_SERVER_TOKEN`（**不落盘**），
  否则从用户数据目录 `user_data_dir()/token` 读取/生成并持久化。所有请求必须携带
  `Authorization: Bearer <token>` 头，否则返回 401。**token 写失败即启动失败**——杜绝
  「server 假活、client 必 401」。
- **白名单**：显式注册表（`server._OPS` 手写清单），只暴露各 app 的公开 RPC；会话内部方法
  （`active_pres` / `active_doc` / `active_book`）与私有方法一律不暴露。新增 RPC 必须手动登记。
- **防滥用**：请求体上限 16MB（超限 413）、`Content-Type` 必须为 `application/json`（否则 415）、
  负 `Content-Length` 拒绝（400）、响应体上限 64MB（超限回 500 错误，不向客户端写超大 payload）、
  `POST` 只接受 `/call` 与 `/shutdown` 两个路径，其余一律 404。
- **健康端点**：`/ping` 免鉴权（仅握手用），`/status` 需鉴权，返回 `{version, protocol, pid, python, started_at}`。
  `/shutdown` 需鉴权——身份由 token 证明，走优雅停机；不依赖 pid 强杀。
- **回环绑定**：默认只允许 `127.0.0.1` / `localhost` / `::1`；`--host 0.0.0.0` 需显式
  `--unsafe-allow-remote`，否则启动拒绝（`ServerStartError`）。
- **进程管理（所有权纪律）**：`offipy server status|stop|restart` 用 PID 文件 + netstat 探测管理常驻进程。
  - `status` **只读**：未运行时只报告「未在运行」，**不隐式拉起** server。
  - `stop`：身份可鉴证（token 匹配）→ 走鉴权 `/shutdown` 优雅停机；token 失配（`auth_fail`）→
    **绝不杀**，只提示修正 token；端口被非 offipy 进程占用（`mismatch`）→ 仅当 `server.pid`
    证明该进程是「我们的」server 才强杀，否则拒绝并提示手动处理。
  - 旧 client 连新 server（token 失配）只报错，**不自杀进程**。

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
