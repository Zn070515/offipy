"""从 schema.OPS 生成 docs/api/ 下的 API 参考页面（P2-5 文档站）。

运行：`uv run python scripts/gen_api_ref.py`（纯标准库，无外部依赖）。
生成的 markdown 供 mkdocs 渲染，是 API 参考的唯一事实来源。
"""

from pathlib import Path

from offipy import schema

APP_NAMES = {"excel": "Excel", "word": "Word", "ppt": "PowerPoint"}
DOCS_API = Path(__file__).resolve().parent.parent / "docs" / "api"


def _type_name(t) -> str:
    return {
        str: "str",
        bool: "bool",
        int: "int",
        float: "float",
        list: "list",
        dict: "dict",
        object: "any",
        type(None): "null",
    }.get(t, "any")


def _param_list(params: dict) -> str:
    if not params:
        return "_无参数_"
    return "、".join(f"`{k}: {_type_name(v)}`" for k, v in params.items())


def _flags(spec: schema.OpSpec) -> list[str]:
    out = []
    if spec.readonly:
        out.append("只读")
    if spec.destructive:
        out.append("会改动文档/应用状态")
    if spec.deprecated:
        out.append("已弃用（响应带 warning）")
    return out


def _render_op(app: str, op: str, spec: schema.OpSpec) -> str:
    lines = [f"### `{op}`", "", spec.description or "", "", f"- **参数**: {_param_list(spec.params)}"]
    lines.append(f"- **返回**: `{spec.returns}`")
    flags = _flags(spec)
    lines.append(f"- **标志**: {'，'.join(flags) if flags else '普通操作'}")
    return "\n".join(lines)


def _render_app(app: str) -> str:
    title = APP_NAMES[app]
    ops = schema.OPS[app]
    body = "\n\n---\n\n".join(_render_op(app, op, spec) for op, spec in ops.items())
    return f"# {title} API\n\n{body}\n"


def _render_index() -> str:
    rows = []
    for app in schema.apps():
        count = len(schema.OPS[app])
        read = len(schema.readonly_ops(app))
        destr = len(schema.destructive_ops(app))
        rows.append(f"| [{APP_NAMES[app]}]({app}.md) | {count} | {read} | {destr} |")
    table = "\n".join(rows)
    return (
        "# API 参考\n\n"
        "本参考由 `scripts/gen_api_ref.py` 从 `schema.py` 单一来源生成，"
        "覆盖 server / CLI / MCP 三入口的同一批操作。\n\n"
        "| 应用 | 操作数 | 只读 | 改动状态 |\n"
        "| --- | --- | --- | --- |\n"
        f"{table}\n"
        "\n每个操作：`doc_id` 缺省走当前活动文档（Excel `bookN` / Word `docN` / PPT `presN`）；"
        "`expected_target` 用于破坏性操作的绑定校验。\n"
    )


def main() -> None:
    DOCS_API.mkdir(parents=True, exist_ok=True)
    (DOCS_API / "index.md").write_text(_render_index(), encoding="utf-8")
    for app in schema.apps():
        (DOCS_API / f"{app}.md").write_text(_render_app(app), encoding="utf-8")
    print(f"generated {len(schema.apps()) + 1} pages under docs/api/")


if __name__ == "__main__":
    main()
