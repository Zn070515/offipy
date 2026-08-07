"""offipy 高层 API facade（P1）：Excel() / Word() / Ppt() 上下文管理器。

两种会话模型（P0-4）：
- Excel()/Word()/Ppt()（= offipy.direct.*）：本地直连 COM，doc_id/线程/
  会话状态与 CLI/MCP/Remote* 完全隔离。
- RemoteExcel()/RemoteWord()/RemotePpt()：经 client→server 的远程会话，
  与 CLI/MCP 共享同一常驻 server（同 doc_id/线程/安全），驱动用户当前
  可见的 Office 窗口。

常用操作代理到 app 类的原子操作（与 CLI/server 同源），惰性 COM 与
offipy 异常体系由底层透传。会话式语义：__exit__ 不退出 Office 应用，
窗口与文档跨调用保持存活（要关窗口显式调 .quit()）。
"""

import inspect

from . import client, schema
from .excel import ExcelApp
from .exceptions import InvalidArgumentError
from .ppt import PptApp
from .word import WordApp


class _Facade:
    """通用 facade：未显式定义的属性/操作代理到底层 app 实例。"""

    def __init__(self, app, app_name):
        self._app = app
        self._app_name = app_name

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        # 会话式驱动：退出上下文不关 Office 窗口；异常照常传播
        return False

    def __getattr__(self, name):
        return getattr(self._app, name)

    def __dir__(self):
        # P1-4：dir() 补充显示 schema 正式 op（与 CLI/MCP/文档同批）；
        # super().__dir__() 仍含 _app/_app_name 等内部成员，此处只做并集不隐藏
        return sorted(set(super().__dir__()) | set(schema.ops(self._app_name)) | {"quit"})

    def quit(self, force: bool = False):
        # 实例方法优先于 __getattr__ 代理：必须显式透传 force，否则遮蔽
        # PptApp/WordApp/ExcelApp 的 quit(force)（此前被硬编码无参版本拦住，
        # 导致 quit(force=True) 抛 TypeError，而错误消息又引导用户传 force）。
        return self._app.quit(force=force)


class Excel(_Facade):
    """Excel 高层 facade（代理 ExcelApp 全部原子操作）。"""

    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        super().__init__(
            ExcelApp(visible, modify_existing_visibility=modify_existing_visibility), "excel"
        )


class Word(_Facade):
    """Word 高层 facade（代理 WordApp 全部原子操作）。"""

    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        super().__init__(
            WordApp(visible, modify_existing_visibility=modify_existing_visibility), "word"
        )


class Ppt(_Facade):
    """PowerPoint 高层 facade（代理 PptApp 全部原子操作）。"""

    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        super().__init__(
            PptApp(visible, modify_existing_visibility=modify_existing_visibility), "ppt"
        )


class _RemoteFacade:
    """经 client→server 的远程会话 facade：与 CLI/MCP 同 doc_id/线程/安全。

    每个方法调用发一次 HTTP 到常驻 server（缺省本地 8890），驱动用户当前
    可见的 Office 会话；与本地直连的 Excel()/Word()/Ppt() 互不相通。未显式
    定义的 op 按 App 方法签名绑定位置参数后代理到 client.call（同直接 facade
    的调用手感），异常仍走 offipy 异常体系透传。
    """

    def __init__(self, app: str, base_url: str | None = None):
        self._app_name = app
        self._base_url = base_url

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False  # 会话式驱动：退出上下文不关 Office 窗口

    def __dir__(self):
        # P1-4：远程 facade 在 super().__dir__() 基础上补充 schema 正式 op + quit
        return sorted(set(super().__dir__()) | set(schema.ops(self._app_name)) | {"quit"})

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        method = getattr(_APP_CLASSES.get(self._app_name), name, None)
        if method is None:
            return lambda **kw: client.call(self._app_name, name, base_url=self._base_url, **kw)
        sig = inspect.signature(method)
        params = [p for p in sig.parameters.values() if p.name != "self"]
        # 传输层参数（P0-1/P0-3/#25）：expected_target 仅破坏性/导出 op 暴露；
        # follow_active 额外放行只读 op（accepts_follow_active），对齐 api.op()。
        # App 方法签名不声明它们，经 schema 显式补上（与 MCP tool 同策略）。
        if schema.supports_expected_target(self._app_name, name):
            params = params + [
                inspect.Parameter(
                    "expected_target",
                    inspect.Parameter.KEYWORD_ONLY,
                    annotation=dict,
                    default=None,
                ),
            ]
        if schema.supports_follow_active(self._app_name, name):
            params = params + [
                inspect.Parameter(
                    "follow_active",
                    inspect.Parameter.KEYWORD_ONLY,
                    annotation=bool,
                    default=False,
                ),
            ]
        # P1-4：request_id 作为传输层参数暴露（幂等标识），不进 server 的 op args
        params = params + [
            inspect.Parameter(
                "request_id",
                inspect.Parameter.KEYWORD_ONLY,
                annotation=str,
                default=None,
            ),
        ]
        call_sig = sig.replace(parameters=params)

        def _call(*args, **kw):
            bound = call_sig.bind(*args, **kw)
            args = dict(bound.arguments)
            rid = args.pop("request_id", None)
            # 传输层参数缺省不下发（payload 干净；与 MCP tool_fn 同策略）
            if not args.get("expected_target"):
                args.pop("expected_target", None)
            if not args.get("follow_active"):
                args.pop("follow_active", None)
            return client.call(
                self._app_name, name, base_url=self._base_url, request_id=rid, **args
            )

        _call.__signature__ = call_sig
        return _call

    def quit(self, request_id: str | None = None, **kw):
        return client.call(
            self._app_name, "quit", base_url=self._base_url, request_id=request_id, **kw
        )


_APP_CLASSES = {"excel": ExcelApp, "word": WordApp, "ppt": PptApp}


class RemoteExcel(_RemoteFacade):
    """Excel 远程会话 facade（client→server，与 CLI/MCP 共享会话）。"""

    def __init__(self, base_url: str | None = None):
        if base_url is None:
            client.ensure_server()  # 缺省连本地 8890：先保证 server 存活
        super().__init__("excel", base_url)


class RemoteWord(_RemoteFacade):
    """Word 远程会话 facade（client→server，与 CLI/MCP 共享会话）。"""

    def __init__(self, base_url: str | None = None):
        if base_url is None:
            client.ensure_server()
        super().__init__("word", base_url)


class RemotePpt(_RemoteFacade):
    """PowerPoint 远程会话 facade（client→server，与 CLI/MCP 共享会话）。"""

    def __init__(self, base_url: str | None = None):
        if base_url is None:
            client.ensure_server()
        super().__init__("ppt", base_url)


_APP_FACTORIES = {"excel": Excel, "word": Word, "ppt": Ppt}


def op(app: str, op_name: str, **kw):
    """统一分发：按应用名构造 facade 并调用原子操作（与 CLI/server 同源）。

    例：op("excel", "set_cell", sheet=1, cell="A1", value=42)。
    会话式语义：每次调用复用同一 Office 实例（core.ensure_app 会话管理），
    操作作用在当前激活文档上。返回底层 app 方法的结果。
    """
    factory = _APP_FACTORIES.get(app)
    if factory is None:
        raise InvalidArgumentError(f"未知应用: {app}，可选 {list(_APP_FACTORIES)}")
    with factory() as f:
        return getattr(f, op_name)(**kw)
