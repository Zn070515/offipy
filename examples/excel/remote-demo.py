"""P0-4 会话模型示例：RemoteExcel（共享 server 会话） vs Excel（本地直连 COM）。

RemoteExcel 与 CLI/MCP 共用同一个常驻 server 会话（同 doc_id）；Excel()
直连 COM，doc_id/线程/状态与 server 会话隔离。

运行：uv run python examples/excel/remote-demo.py
（自动拉起 8890 server + Excel；结束时 close_book + quit，不留 Office 进程）
"""

import offipy


def main():
    # Remote：与 CLI `offipy excel list_docs` 共享会话 → 同一个 doc_id
    with offipy.RemoteExcel() as r:
        did = r.new_book()
        print(f"remote doc_id = {did!r}  （CLI `offipy excel list_docs` 可见同一 id）")
        r.set_cell(1, "A1", 42, follow_active=True)
        print("remote A1 =", r.get_cell(1, "A1"))
        r.close_book(save=False, doc_id=did)

    # direct：独立 doc_id 命名空间，不与 server 会话互通
    with offipy.Excel() as d:
        d.new_book()
        print("direct 会话自持 doc_id，server 会话不可见（互不干扰）")
        d.quit()

    offipy.RemoteExcel().quit()


if __name__ == "__main__":
    main()
