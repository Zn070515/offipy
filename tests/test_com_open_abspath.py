"""C2（H6）: open_book/open_doc 传给 COM 的路径必须是绝对路径。

COM 服务的 Open 按自身工作目录（通常 System32）解析相对路径，相对路径会
开错文件或报错。与 ppt.open_pres 对齐，先 os.path.abspath。
"""

import os
from types import SimpleNamespace

from offipy import excel, word


def test_open_book_passes_abspath_to_com(monkeypatch, tmp_path):
    opened = []
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app.app = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda p: opened.append(p) or "book1"))
    app._register = lambda obj: obj
    monkeypatch.chdir(tmp_path)
    (tmp_path / "book.xlsx").write_bytes(b"x")
    assert app.open_book("book.xlsx") == "book1"
    assert opened == [os.path.abspath("book.xlsx")]


def test_open_doc_passes_abspath_to_com(monkeypatch, tmp_path):
    opened = []
    app = word.WordApp.__new__(word.WordApp)
    app.app = SimpleNamespace(Documents=SimpleNamespace(Open=lambda p: opened.append(p) or "doc1"))
    app._register = lambda obj: obj
    monkeypatch.chdir(tmp_path)
    (tmp_path / "doc.docx").write_bytes(b"x")
    assert app.open_doc("doc.docx") == "doc1"
    assert opened == [os.path.abspath("doc.docx")]
