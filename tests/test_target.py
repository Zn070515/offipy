"""目标身份原语 + 只读不创建（P0-7/P0-8）纯逻辑测试。

用 `__new__` 绕过 `__init__`（避免真连 COM），再 monkeypatch core.active_doc /
core.doc_alive 注入假文档：断言 active_* 无文档返回 None 且不调 Add()；
依赖文档的 op 无文档抛 TargetNotFoundError；get_target 返回身份 dict。
"""

import pytest

from offipy import excel, server, word
from offipy.exceptions import InvalidArgumentError, TargetNotFoundError


class _NoAdd:
    """记录 Add 是否被调用；真被调用即失败。"""

    def __init__(self):
        self.add_called = False

    def Add(self, *a, **k):
        self.add_called = True
        raise AssertionError("P0-8: 读操作不得隐式 Add()")


class _FakeApp:
    def __init__(self, active, no_add):
        self.ActiveWorkbook = active  # excel
        self.ActiveDocument = active  # word
        self.ActivePresentation = active  # ppt
        self.Workbooks = no_add
        self.Documents = no_add
        self.Presentations = no_add
        self.DisplayAlerts = True  # _alerts_scope 读写 DisplayAlerts


class _Book:
    def __init__(self, name, path):
        self.Name = name
        self.FullName = path


def _no_doc_env(monkeypatch):
    monkeypatch.setattr("offipy.core.active_doc", lambda *a: None)
    monkeypatch.setattr("offipy.core.doc_alive", lambda *a: False)


# --- active_* 只读不创建（P0-8） ---


def test_active_book_none_without_doc(monkeypatch):
    no_add = _NoAdd()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_book() is None
    assert no_add.add_called is False


def test_active_doc_none_without_doc(monkeypatch):
    no_add = _NoAdd()
    app = word.WordApp.__new__(word.WordApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_doc() is None
    assert no_add.add_called is False


def test_active_pres_none_without_doc(monkeypatch):
    from offipy import ppt

    no_add = _NoAdd()
    app = ppt.PptApp.__new__(ppt.PptApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_pres() is None
    assert no_add.add_called is False


def test_active_book_falls_back_to_real_active(monkeypatch):
    no_add = _NoAdd()
    book = _Book("Book1", r"C:\x\Book1.xlsx")
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(book, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_book() is book  # 落到 self.app.ActiveWorkbook，仍不 Add
    assert no_add.add_called is False


# --- get_target 身份（P0-7） ---


def test_get_target_excel(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(_Book("Book1", r"C:\x\Book1.xlsx"), _NoAdd())
    _no_doc_env(monkeypatch)
    assert app.get_target() == {
        "app": "excel",
        "doc_id": "book1",  # 实时解析的 ActiveWorkbook 并入文档表后分配 doc_id
        "name": "Book1",
        "path": r"C:\x\Book1.xlsx",
    }


def test_get_target_explicit_doc_id(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {"b1": _Book("Book1", r"C:\x\Book1.xlsx")}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    # 显式 doc_id 路由要求句柄存活（与 _no_doc_env 的 doc_alive=False 相反）
    monkeypatch.setattr("offipy.core.active_doc", lambda *a: None)
    monkeypatch.setattr("offipy.core.doc_alive", lambda *a: True)
    assert app.get_target(doc_id="b1")["doc_id"] == "b1"
    assert app.get_target(doc_id="b1")["name"] == "Book1"


def test_get_target_unknown_doc_id_raises(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.get_target(doc_id="nope")


def test_get_target_none_without_doc(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    assert app.get_target() is None


# --- 依赖文档的 op 无文档 → TargetNotFoundError ---


def test_read_range_no_doc_raises(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_range(1, "A1")  # 只读 op：无目标 → TargetNotFoundError
    # save 是破坏性 op：@destructive 守卫先于 _require_book 拦截（无 doc_id 且
    # 未 follow_active → InvalidArgumentError，而非 TargetNotFoundError）
    with pytest.raises(InvalidArgumentError):
        app.save()


def test_word_read_doc_text_no_doc_raises(monkeypatch):
    app = word.WordApp.__new__(word.WordApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_doc_text()


def test_ppt_read_slide_summary_no_doc_raises(monkeypatch):
    from offipy import ppt

    app = ppt.PptApp.__new__(ppt.PptApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_slide_summary()


# --- expected_target 绑定（P0-7，server dispatch 层） ---


class _Visible:
    Visible = True


class _TargetApp:
    def __init__(self, target):
        self.app = _Visible()
        self._target = target
        self.close_calls = []

    def get_target(self, doc_id=None):
        return self._target

    def close_book(self, doc_id=None):
        self.close_calls.append(doc_id)
        return "closed"

    def read_range(self, sheet, range_addr):
        return [[1]]


def test_dispatch_expected_target_match():
    app = _TargetApp(
        {"app": "excel", "doc_id": "book1", "name": "Book1", "path": r"C:\x\Book1.xlsx"}
    )
    result = server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book1"]  # resolve-once：校验的 doc_id 注入方法调用


def test_dispatch_expected_target_mismatch():
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Other", "path": None})
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert app.close_calls == []  # P0-5：校验失败不得执行（防「校验 A 执行 B」）


def test_dispatch_expected_target_no_target():
    app = _TargetApp(None)
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_empty_dict_rejected():
    # P0-4：旧 _target_matches({}) 恒真绕过——空对象必须拒绝
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    with pytest.raises(InvalidArgumentError):
        server.dispatch(app, "close_book", {"expected_target": {}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_unknown_key_rejected():
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    with pytest.raises(InvalidArgumentError):
        server.dispatch(app, "close_book", {"expected_target": {"bogus": 1}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_doc_id_binding():
    # 显式 doc_id 绑定：校验用 doc_id 解析目标后直接注入方法参数
    app = _TargetApp({"app": "excel", "doc_id": "book2", "name": "Book2", "path": None})
    result = server.dispatch(app, "close_book", {"expected_target": {"doc_id": "book2"}}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book2"]


def test_dispatch_expected_target_on_readonly_rejected():
    # P0-1 严格：只读 op 上的 expected_target 无意义且有害（用户以为绑定了目标，
    # 实际 op 不作用于文档）——直接拒绝，不再静默忽略。
    app = _TargetApp(None)
    with pytest.raises(InvalidArgumentError):
        server.dispatch(
            app,
            "read_range",
            {"sheet": 1, "range_addr": "A1", "expected_target": {"name": "x"}},
            "excel",
        )


# --- follow_active：显式跟随当前活动文档（P0-1/P0-3，server dispatch 层） ---


def test_dispatch_follow_active_injects_active_doc_id():
    # follow_active 显式声明跟随当前活动文档：实时解析并注入其 doc_id 再执行
    app = _TargetApp({"app": "excel", "doc_id": "book9", "name": "Book9", "path": None})
    result = server.dispatch(app, "close_book", {"follow_active": True}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book9"]


def test_dispatch_follow_active_no_target_raises():
    # follow_active 但无活动目标 → TargetNotFoundError，绝不静默落到任何文档
    app = _TargetApp(None)
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"follow_active": True}, "excel")
    assert app.close_calls == []


def test_dispatch_follow_active_ignored_on_readonly():
    # 只读 op 上的 follow_active 无意义：pop 掉，不注入 doc_id，正常执行
    app = _TargetApp(None)
    result = server.dispatch(
        app, "read_range", {"sheet": 1, "range_addr": "A1", "follow_active": True}, "excel"
    )
    assert result == [[1]]


def test_dispatch_follow_active_rejected_on_quit():
    # #46：quit 不接受 follow_active（与 expected_target 对称拒绝），
    # 而非静默消费——调用方传了 follow_active 就显式报错，防误以为生效。
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "B", "path": None})
    app.quit = lambda: "quitting"
    with pytest.raises(InvalidArgumentError, match="quit 不接受 follow_active"):
        server.dispatch(app, "quit", {"follow_active": True}, "excel")


def test_dispatch_no_target_leaf_to_app_guard():
    # 破坏性 op 无 doc_id/expected_target/follow_active → dispatch 层直接拒绝，
    # 错误在碰 COM 前抛出；App 层守卫（@destructive/@requires_target）不再有机会。
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "B", "path": None})
    with pytest.raises(InvalidArgumentError):
        server.dispatch(app, "close_book", {}, "excel")
    assert app.close_calls == []


def test_dispatch_no_target_meta_all_binding_ops_rejected():
    # P0-1 meta-test：schema 里所有绑定目标的 op（destructive ∪ requires_target ∪
    # supports_expected_target，除 quit）无任何目标绑定 → dispatch 统一抛
    # InvalidArgumentError，绝不静默落到「当前活动文档」。
    from offipy import schema

    for app_name in schema.apps():
        for op in schema.OPS[app_name]:
            if not schema.supports_expected_target(app_name, op) or op == "quit":
                continue
            app = _TargetApp(None)
            with pytest.raises(InvalidArgumentError):
                server.dispatch(app, op, {}, app_name)


# --- App 层 @destructive 守卫（P0-3 doc_id 权威） ---


def test_destructive_decorator_requires_doc_id():
    # 破坏性 App 方法无 doc_id 且未 follow_active → InvalidArgumentError，绝不
    # 静默落到「当前活动文档」（防用户看到 B、Agent 改 A）。
    from offipy import core

    calls = {}

    class _Stub:
        @core.destructive
        def mutate(self, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    with pytest.raises(InvalidArgumentError):
        _Stub().mutate()
    assert calls == {}


def test_destructive_decorator_follow_active_injects():
    from offipy import core

    calls = {}

    class _Stub:
        def get_target(self, doc_id=None):
            return {"app": "excel", "doc_id": "b1", "name": "N", "path": None}

        @core.destructive
        def mutate(self, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    assert _Stub().mutate(follow_active=True) == "ok"
    assert calls["doc_id"] == "b1"


def test_destructive_decorator_follow_active_no_target():
    from offipy import core

    class _Stub:
        def get_target(self, doc_id=None):
            return None

        @core.destructive
        def mutate(self, doc_id=None):
            raise AssertionError("不该执行")

    with pytest.raises(TargetNotFoundError):
        _Stub().mutate(follow_active=True)


def test_destructive_decorator_explicit_doc_id_passes():
    # 显式 doc_id：不需要 follow_active，直接放行
    from offipy import core

    calls = {}

    class _Stub:
        @core.destructive
        def mutate(self, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    assert _Stub().mutate(doc_id="b7") == "ok"
    assert calls["doc_id"] == "b7"


# --- App 层 @requires_target 守卫（P0-3 导出/写文件） ---


def test_requires_target_decorator_requires_doc_id():
    # 导出类 op（不破坏源文档但写文件系统）无 doc_id 且未 follow_active →
    # InvalidArgumentError；绝不静默导出「当前活动文档」。
    from offipy import core

    calls = {}

    class _Stub:
        @core.requires_target
        def export(self, path, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    with pytest.raises(InvalidArgumentError):
        _Stub().export("/tmp/x.pdf")
    assert calls == {}


def test_requires_target_decorator_follow_active_injects():
    from offipy import core

    calls = {}

    class _Stub:
        def get_target(self, doc_id=None):
            return {"app": "excel", "doc_id": "b1", "name": "N", "path": None}

        @core.requires_target
        def export(self, path, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    assert _Stub().export("/tmp/x.pdf", follow_active=True) == "ok"
    assert calls["doc_id"] == "b1"


def test_requires_target_decorator_follow_active_no_target():
    from offipy import core

    class _Stub:
        def get_target(self, doc_id=None):
            return None

        @core.requires_target
        def export(self, path, doc_id=None):
            raise AssertionError("不该执行")

    with pytest.raises(TargetNotFoundError):
        _Stub().export("/tmp/x.pdf", follow_active=True)


def test_requires_target_decorator_explicit_doc_id_passes():
    from offipy import core

    calls = {}

    class _Stub:
        @core.requires_target
        def export(self, path, doc_id=None):
            calls["doc_id"] = doc_id
            return "ok"

    assert _Stub().export("/tmp/x.pdf", doc_id="b7") == "ok"
    assert calls["doc_id"] == "b7"


# --- P1-1 expected_target 规范化比较（name casefold / path normcase+abspath） ---


def _norm_paths(monkeypatch):
    # 模拟 Windows 路径语义（normcase 小写化、abspath 恒等），保证测试跨平台可跑
    monkeypatch.setattr(server.os.path, "normcase", lambda p: str(p).lower())
    monkeypatch.setattr(server.os.path, "abspath", lambda p: p)


def test_expected_target_path_normalized_match(monkeypatch):
    # P1-1：path 大小写/写法差异经 normcase+abspath 归一后命中同一文件
    _norm_paths(monkeypatch)
    app = _TargetApp(
        {"app": "excel", "doc_id": "book1", "name": "Book1", "path": r"C:\X\Book1.XLSX"}
    )
    result = server.dispatch(
        app, "close_book", {"expected_target": {"path": r"c:\x\book1.xlsx"}}, "excel"
    )
    assert result == "closed"
    assert app.close_calls == ["book1"]


def test_expected_target_path_mismatch_rejected(monkeypatch):
    _norm_paths(monkeypatch)
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": r"C:\x\a.xlsx"})
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"expected_target": {"path": r"C:\x\b.xlsx"}}, "excel")
    assert app.close_calls == []


def test_expected_target_path_none_target_rejected():
    # 目标无已保存路径 + 期望 path → 视为不匹配（保守方向，不误命中空串/None）
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    with pytest.raises(TargetNotFoundError):
        server.dispatch(
            app, "close_book", {"expected_target": {"path": r"C:\x\Book1.xlsx"}}, "excel"
        )
    assert app.close_calls == []


def test_expected_target_name_casefold_match():
    # P1-1：name 按 casefold 对碰，忽略大小写
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    result = server.dispatch(app, "close_book", {"expected_target": {"name": "book1"}}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book1"]


# --- P0-3 导出 op（requires_target）目标绑定 ---


def test_dispatch_export_op_follow_active_injects():
    # 导出 op 跟随活动文档：实时解析并注入活动 doc_id 再执行
    app = _TargetApp({"app": "ppt", "doc_id": "pres1", "name": "P1", "path": None})
    app.export_slides = lambda **kw: f"exported:{kw.get('doc_id')}"
    result = server.dispatch(app, "export_slides", {"out_dir": "x", "follow_active": True}, "ppt")
    assert result == "exported:pres1"


def test_dispatch_export_op_expected_target_binds():
    # 导出 op 传 expected_target → resolve-once 注入绑定 doc_id
    app = _TargetApp({"app": "ppt", "doc_id": "pres2", "name": "P2", "path": None})
    app.export_slides = lambda **kw: f"exported:{kw.get('doc_id')}"
    result = server.dispatch(
        app, "export_slides", {"out_dir": "x", "expected_target": {"doc_id": "pres2"}}, "ppt"
    )
    assert result == "exported:pres2"


def test_dispatch_export_op_follow_active_no_target_raises():
    # 导出 op follow_active 但无活动目标 → TargetNotFoundError，不静默导出
    app = _TargetApp(None)
    app.export_slides = lambda **kw: "should-not-run"
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "export_slides", {"out_dir": "x", "follow_active": True}, "ppt")
