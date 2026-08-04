"""offipy 高层 API facade（P1）：Excel() / Word() / Ppt() 上下文管理器。

常用操作代理到 app 类的原子操作（与 CLI/server 同源），惰性 COM 与
offipy 异常体系由底层透传。会话式语义：__exit__ 不退出 Office 应用，
窗口与文档跨调用保持存活（要关窗口显式调 .quit()）。

完整 session id 体系（多开识别 / 精确目标文档）留待 2.0；本轮 op 始终
作用在用户当前激活的实时文档上（P1.2 ActiveDocument 语义）。
"""

from .excel import ExcelApp
from .ppt import PptApp
from .word import WordApp


class _Facade:
    """通用 facade：未显式定义的属性/操作代理到底层 app 实例。"""

    def __init__(self, app):
        self._app = app

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        # 会话式驱动：退出上下文不关 Office 窗口；异常照常传播
        return False

    def __getattr__(self, name):
        return getattr(self._app, name)

    def quit(self):
        return self._app.quit()


class Excel(_Facade):
    """Excel 高层 facade（代理 ExcelApp 全部原子操作）。"""

    def __init__(self, visible: bool = True):
        super().__init__(ExcelApp(visible))


class Word(_Facade):
    """Word 高层 facade（代理 WordApp 全部原子操作）。"""

    def __init__(self, visible: bool = True):
        super().__init__(WordApp(visible))


class Ppt(_Facade):
    """PowerPoint 高层 facade（代理 PptApp 全部原子操作）。"""

    def __init__(self, visible: bool = True):
        super().__init__(PptApp(visible))


_APP_FACTORIES = {"excel": Excel, "word": Word, "ppt": Ppt}


def op(app: str, op_name: str, **kw):
    """统一分发：按应用名构造 facade 并调用原子操作（与 CLI/server 同源）。

    例：op("excel", "set_cell", sheet=1, cell="A1", value=42)。
    会话式语义：每次调用复用同一 Office 实例（core.ensure_app 会话管理），
    操作作用在当前激活文档上。返回底层 app 方法的结果。
    """
    factory = _APP_FACTORIES.get(app)
    if factory is None:
        raise ValueError(f"未知应用: {app}，可选 {list(_APP_FACTORIES)}")
    with factory() as f:
        return getattr(f, op_name)(**kw)
