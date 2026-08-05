"""P2-2 多实例：client 端口路由 / token·pid·oplog 按端口隔离 / server --port。"""

import json

import pytest

from offipy import client, oplog, server


@pytest.fixture(autouse=True)
def _pin_port(monkeypatch):
    """把 client._PORT 钉回默认值并清掉 env，防止跨测试污染。"""
    monkeypatch.setattr(client, "_PORT", client.PORT)
    monkeypatch.delenv("OFFIPY_SERVER_PORT", raising=False)


def test_port_env_priority(monkeypatch):
    assert client.port() == client.PORT
    client.set_port(8891)
    assert client.port() == 8891
    monkeypatch.setenv("OFFIPY_SERVER_PORT", "9001")
    assert client.port() == 9001  # env 优先于 set_port


def test_base_url_follows_port():
    client.set_port(8891)
    assert client._base_url() == "http://127.0.0.1:8891"
    client.set_port(client.PORT)
    assert client._base_url() == f"http://127.0.0.1:{client.PORT}"


def test_token_and_pid_paths_isolated_by_port(monkeypatch, tmp_path):
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    client.set_port(8891)
    assert client._token_path() == tmp_path / "token-8891"
    assert client._pid_path() == tmp_path / "server-8891.pid"
    client.set_port(client.PORT)
    assert client._token_path() == tmp_path / "token"  # 默认端口沿用旧文件名
    assert client._pid_path() == tmp_path / "server.pid"


def test_ensure_server_passes_port_and_writes_scoped_pid(monkeypatch, tmp_path):
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.setattr(client, "_probe", lambda: "down")
    monkeypatch.setattr(client.time, "sleep", lambda s: None)

    class FakePopen:
        def __init__(self, cmd, **kw):
            self.pid = 4242
            self.cmd = cmd
            self.kw = kw

    popen = {}
    monkeypatch.setattr(
        "offipy.client.subprocess.Popen",
        lambda cmd, **kw: popen.setdefault("p", FakePopen(cmd, **kw)),
    )
    client.set_port(8891)
    with pytest.raises(server.ServerStartError):
        client.ensure_server()  # 握手恒 down → 超时抛错，但进程已按端口拉起
    assert popen["p"].cmd[-2:] == ["--port", "8891"]
    # P0-2：pid 文件始终 JSON；token 未定则 token_sha256 为 null（不误杀）
    data = json.loads((tmp_path / "server-8891.pid").read_text(encoding="utf-8"))
    assert data["pid"] == 4242 and data["port"] == 8891
    assert data["token_sha256"] is None
    assert not (tmp_path / "server.pid").exists()


def test_oplog_configure_scopes_path(monkeypatch, tmp_path):
    monkeypatch.setattr(oplog, "user_data_dir", lambda: tmp_path)
    oplog.configure(8891)
    assert oplog.log_path() == tmp_path / "oplog-8891.jsonl"
    oplog.configure(8890)
    assert oplog.log_path() == tmp_path / "oplog.jsonl"


def test_server_load_token_uses_port_file(monkeypatch, tmp_path):
    monkeypatch.setattr(server, "user_data_dir", lambda: tmp_path)
    (tmp_path / "token-8891").write_text("port-token", encoding="utf-8")
    monkeypatch.delenv("OFFIPY_SERVER_TOKEN", raising=False)
    assert server._load_token(8891) == "port-token"
    assert server._load_token(server.DEFAULT_PORT) != "port-token"


def test_server_token_path_default_and_scoped():
    assert server._token_path(server.DEFAULT_PORT) == server.user_data_dir() / "token"
    assert server._token_path(8891) == server.user_data_dir() / "token-8891"
