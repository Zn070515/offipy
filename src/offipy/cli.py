"""offipy CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    offipy excel new_book
    offipy excel set_cell --sheet 1 --cell A1 --value 100
    offipy word new_doc
    offipy word write_line --text "你好"
    offipy ppt new_pres
    offipy ppt add_slide --layout 2
    offipy deck make --html examples/decks/starter/deck.html --out out/deck.pptx
    offipy quit excel
    offipy mcp                      # 启动 MCP stdio server（Claude Desktop 接入）
    offipy check [--json]           # 环境就绪诊断（Python/依赖/Office/浏览器/server）

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。

参数系统（P1）：
- 普通 `--key value`：值自动转 bool/int/float，无法转时回退字符串
- 重复 `--key`：多次出现聚合为 list
- `--payload '<json>'` / `--json '<json>'`：复杂嵌套参数以 JSON 透传，
  覆盖同名 kwargs
- 未知 `--key`（不在该 op 签名内）→ unrecognized arguments，exit 2
"""

import argparse
import inspect
import json
import re
import sys
import tempfile
import traceback
import types
import typing
from pathlib import Path
from typing import Any, cast

from . import diagram, excel, ppt, schema, word
from .client import call, ensure_server, server_status, set_port, stop_server
from .exceptions import InvalidArgumentError, OffipyError
from .feedback.app import FeedbackApp

_APP_CLASSES = {
    "diagram": diagram.DiagramApp,
    "excel": excel.ExcelApp,
    "feedback": FeedbackApp,
    "word": word.WordApp,
    "ppt": ppt.PptApp,
}

_TRACE_FRAME_LINE_RE = re.compile(r"^\s")  # 帧行以空白开头
# 类型前缀剥离锚定到规范化前缀 "[app::op] 失败: "（client._raise_error 拼消息的
# 固定形态）：只有它紧跟 "{TypeName}: " 才剥离，避免误伤其它以 "失败: " 结尾的
# 标签（如 client.py 的 "[excel::save] 连接失败: ConnectionRefusedError: ..."）。
_TYPE_PREFIX_RE = re.compile(r"^(\[[A-Za-z]+::[A-Za-z_]+\] 失败: )\w+(?:Error|Exception): (.*)$")


def _clean_error_message(msg: str) -> str:
    """去掉 server/client 拼进异常消息里的 trace 帧行与冗余类型名前缀。

    client._raise_error 拼出的消息形如：
        "[word::open_doc] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.docx"
        + 换行 + 若干以空白开头的内部代码帧行
    这里去掉空白开头的帧行（内部文件路径泄漏），并去掉第一行里
    "{TypeName}: "（如 "InvalidArgumentError: "）前缀，保留 [app::op] 失败: 上下文。
    """
    lines = [ln for ln in msg.splitlines() if ln and not _TRACE_FRAME_LINE_RE.match(ln)]
    if not lines:
        return ""
    first = lines[0]
    m = _TYPE_PREFIX_RE.match(first)
    if m:
        first = m.group(1) + m.group(2)
    return "\n".join([first, *lines[1:]])


def _parse_kwargs(tokens: list[str]) -> dict[str, Any]:
    """--key value 解析：token 值保留原始字符串，重复 key 聚合为 list。

    类型转换不再在这里做（弃全局猜测）：值按 schema 声明类型由 _coerce_kwargs
    统一转换，避免 "00123" 丢前导零、"true" 被误当 bool。--payload/--json 仍
    JSON 透传（结构化值覆盖同名 kwargs）。
    """
    kwargs: dict[str, Any] = {}
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok in ("--payload", "--json"):
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                try:
                    payload = json.loads(tokens[i + 1])
                except (ValueError, TypeError) as e:
                    _usage_exit(f"--{tok[2:]} JSON 解析失败: {e}")
                if not isinstance(payload, dict):
                    _usage_exit(
                        f"--{tok[2:]} 必须是 JSON 对象；list 参数用重复 --key "
                        f"或 --payload '{{\"key\": [item1, item2, ...]}}'"
                    )
                kwargs.update(payload)
                i += 2
            else:
                kwargs[tok[2:]] = True
                i += 1
        elif tok in ("--expected-target", "--expected_target"):
            # P0-1/P0-3 传输层参数：目标绑定（JSON 对象），解析后存进 expected_target。
            # 值必须是对象（doc_id/name/path 之一），具体校验在 server 侧统一做。
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                try:
                    value = json.loads(tokens[i + 1])
                except (ValueError, TypeError) as e:
                    _usage_exit(f"{tok} JSON 解析失败: {e}")
                if not isinstance(value, dict):
                    _usage_exit(
                        f'{tok} 必须是 JSON 对象（--expected-target \'{{"doc_id":"book<id>"}}\'）'
                    )
                kwargs["expected_target"] = value
                i += 2
            else:
                _usage_exit(f"{tok} 需要一个 JSON 对象值")
        elif tok in ("--follow-active", "--follow_active"):
            # P0-1/P0-3 传输层参数：显式声明跟随当前活动文档（裸用为 True，
            # 也接受 `--follow-active false`）。
            if (
                i + 1 < len(tokens)
                and not tokens[i + 1].startswith("--")
                and tokens[i + 1].strip().lower() in _BOOL_TOKENS
            ):
                kwargs["follow_active"] = _BOOL_TOKENS[tokens[i + 1].strip().lower()]
                i += 2
            else:
                kwargs["follow_active"] = True
                i += 1
        elif tok in ("--request-id", "--request_id"):
            # P1-4 传输层参数：显式幂等 id（重试复用同 id 防重复执行）。
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                kwargs["request_id"] = tokens[i + 1]
                i += 2
            else:
                _usage_exit(f"{tok} 需要一个值")
        elif tok.startswith("--"):
            # 全局归一化 - → _：--doc-id 与 --doc_id 等价，README 双写法不再打架。
            # 传输层参数（--expected-target/--follow-active/--request-id）在上方
            # elif 已双写法特判，到不了这里，不会二次改名。
            key = tok[2:].replace("-", "_")
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                value = tokens[i + 1]
                if key in kwargs:
                    existing = kwargs[key]
                    if isinstance(existing, list):
                        existing.append(value)
                    else:
                        kwargs[key] = [existing, value]
                else:
                    kwargs[key] = value
                i += 2
            else:
                kwargs[key] = True
                i += 1
        else:
            kwargs[tok] = True
            i += 1
    return kwargs


_BOOL_TOKENS = {
    "true": True,
    "false": False,
    "1": True,
    "0": False,
    "on": True,
    "off": False,
    "yes": True,
    "no": False,
}


class _BoolAction(argparse.Action):
    """--flag / --flag <true|false> 两用布尔：裸用为 True，显式值按 token 解析。

    根除 bool("false") is True 陷阱：显式传 false 必须是 False。
    """

    def __init__(self, option_strings: list[str], dest: str, **kwargs: Any) -> None:
        kwargs.setdefault("nargs", "?")
        kwargs.setdefault("const", True)
        kwargs.setdefault("default", False)
        super().__init__(option_strings, dest, **kwargs)

    def __call__(
        self,
        _parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: object,
        _option_string: str | None = None,
    ) -> None:
        if values is None:
            setattr(namespace, self.dest, True)
            return
        token = str(values).strip().lower()
        if token not in _BOOL_TOKENS:
            raise argparse.ArgumentError(
                self, f"布尔值（true/false/1/0/on/off/yes/no），收到: {values!r}"
            )
        setattr(namespace, self.dest, _BOOL_TOKENS[token])


def _unwrap_optional(ann: Any) -> Any:
    """Optional[X] / X | None → X；其余注解原样返回。"""
    origin = typing.get_origin(ann)
    if origin is not typing.Union and origin is not getattr(types, "UnionType", None):
        return ann
    rest = [a for a in typing.get_args(ann) if a is not type(None)]
    if len(rest) == 1:
        return rest[0]
    return ann


def _coerce_fail(key: str, expected: str, value: object) -> None:
    print(f"offipy: error: --{key} 需要{expected}，收到: {value!r}", file=sys.stderr)
    raise SystemExit(2)


def _coerce_value(key: str, ann: Any, value: object) -> object:
    """按单个参数注解转换；失败 stderr + exit 2（argparse 语义）。"""
    base = _unwrap_optional(ann)
    if base is str or base is inspect.Parameter.empty:
        return value
    if base is int:
        try:
            return int(cast("str", value))
        except (TypeError, ValueError):
            _coerce_fail(key, "整数", value)
    if base is float:
        try:
            return float(cast("str", value))
        except (TypeError, ValueError):
            _coerce_fail(key, "数字", value)
    if base is bool:
        # bool("false") 是 True 的陷阱必须避开：只认显式布尔 token
        if isinstance(value, bool):
            return value
        token = str(value).strip().lower()
        if token in _BOOL_TOKENS:
            return _BOOL_TOKENS[token]
        _coerce_fail(key, "布尔值（true/false/1/0/on/off/yes/no）", value)
    if base is list or typing.get_origin(base) is list:
        # 裸 list 与 list[T] 都按列表处理：标量包一层，列表原样
        return value if isinstance(value, list) else [value]
    return value


def _param_hints(app: str, op: str) -> dict[str, object]:
    """参数名→类型：优先 schema 声明（P1-2 单一来源），缺失时回退方法签名。

    未知 op / 无法检视签名 → 空表（参数原样透传，交 server 侧校验）。
    """
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        return dict(sp.params)
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return {}
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return {}
    return {name: param.annotation for name, param in sig.parameters.items() if name != "self"}


def _coerce_kwargs(app: str, op: str, kwargs: dict[str, object]) -> dict[str, Any]:
    """按 schema 声明类型转换参数（P1-2，取代逐方法签名猜测）。

    无注解/Any/str 保持字符串；int/float/bool/list 按声明类型转换；未知
    op 跳过（交给 server 侧处理）。
    """
    hints = _param_hints(app, op)
    coerced = dict(kwargs)
    for key, value in kwargs.items():
        ann = hints.get(key)
        if ann is None:
            continue
        coerced[key] = _coerce_value(key, ann, value)
    return coerced


def _usage_exit(message: str) -> None:
    """参数用法错误 → stderr 报错 + exit 2（对齐 argparse 的 exit 2 语义）。

    #29：CLI 退出码不统一——deck 缺必填参数走 `raise SystemExit("...")`（exit 1），
    而 argparse/参数校验走 exit 2。统一归到 2：1 留给运行时失败（COM/IO/转换错误）。
    """
    print(message, file=sys.stderr)
    raise SystemExit(2)


def _validate_kwargs(app: str, op: str, kwargs: dict[str, object]) -> None:
    """未知 --key（不在 schema/方法参数内）→ stderr 报错并 exit 2（argparse 语义）。"""
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        known = set(sp.params)
    else:
        cls = _APP_CLASSES.get(app)
        method = getattr(cls, op, None) if cls else None
        if method is None:
            return  # 未知 op 交由 server 侧报错
        try:
            params = inspect.signature(method).parameters.values()
        except (TypeError, ValueError):
            return
        if any(p.kind is inspect.Parameter.VAR_KEYWORD for p in params):
            return
        known = {
            p.name
            for p in params
            if p.name != "self"
            and p.kind
            in (
                inspect.Parameter.POSITIONAL_OR_KEYWORD,
                inspect.Parameter.KEYWORD_ONLY,
                inspect.Parameter.POSITIONAL_ONLY,
            )
        }
    # 传输层参数（P0-1/P0-3/#25）：expected_target 仅破坏性/导出 op 放行；
    # follow_active 额外放行只读 op（accepts_follow_active）。server 侧对
    # expected_target 严格拒绝、follow_active 静默忽略，语义一致。
    if schema.supports_expected_target(app, op):
        known |= {"expected_target"}
    if schema.supports_follow_active(app, op):
        known |= {"follow_active"}
    known |= {"request_id"}  # P1-4：所有 op 可用，透明传 server（幂等标识）
    for key in kwargs:
        if key not in known:
            print(f"offipy: error: {app} {op}: unrecognized arguments: --{key}", file=sys.stderr)
            raise SystemExit(2)


_TYPE_LABEL = {str: "str", int: "int", float: "num", bool: "bool"}


def _required_params(app: str, op: str) -> frozenset[str]:
    """App 方法签名中无默认值的参数（必填）——schema 未独立声明必填，以此派生。"""
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return frozenset()
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return frozenset()
    return frozenset(
        p.name
        for p in sig.parameters.values()
        if p.name != "self"
        and p.default is inspect.Parameter.empty
        and p.kind
        in (
            inspect.Parameter.POSITIONAL_OR_KEYWORD,
            inspect.Parameter.KEYWORD_ONLY,
        )
    )


def _validate_required(app: str, op: str, kwargs: dict[str, object]) -> None:
    """必填参数缺失 → 调用前预校验：stderr 报错 + 用法，exit 2。

    不再等拉起 server / 碰 COM 后才炸；--payload 注入的键也算已提供。
    """
    missing = _required_params(app, op) - set(kwargs)
    if not missing:
        return
    sp = schema.spec(app, op)
    hints = dict(sp.params) if sp is not None and sp.params else {}
    shown = ", ".join(
        f"--{m}" + (f" <{_TYPE_LABEL[hints[m]]}>" if hints.get(m) in _TYPE_LABEL else "")
        for m in sorted(missing)
    )
    print(
        f"offipy: error: {app} {op}: 缺少必填参数 {shown}\n"
        f"  用法: offipy {app} {op} --<参数> <值> ...",
        file=sys.stderr,
    )
    raise SystemExit(2)


def _has_doc_id_param(app: str, op: str) -> bool:
    """op 是否接受 doc_id 参数（schema 声明或 App 方法签名）。quit 等无目标 op 返回 False。"""
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        return "doc_id" in sp.params
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return False
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return False
    return "doc_id" in sig.parameters


def _validate_destructive_target(app: str, op: str, kwargs: dict[str, object]) -> None:
    """破坏性 op 缺目标（doc_id/expected_target/follow_active 均无）→ 提前友好报错。

    P0-3 doc_id 权威：与 server 的 InvalidArgumentError 同语义，但 CLI 先给
    usage 提示（exit 2），不等到拉起 server/碰 COM 后才炸。
    """
    if op == "quit" or not schema.supports_expected_target(app, op):
        return
    if not _has_doc_id_param(app, op):
        return  # 无 doc_id 参数的 op（理论上只有 quit）不需要目标
    if kwargs.get("doc_id") not in (None, ""):
        return
    if "expected_target" in kwargs:
        return
    if kwargs.get("follow_active"):
        return
    print(
        f"offipy: error: {app} {op}: 该操作必须显式指定目标文档\n"
        f"  用法: offipy {app} {op} --doc_id <id> ...\n"
        f"        或 --expected-target '<json>' 绑定目标（如 '{{\"doc_id\":\"book<hex>\"}}'）\n"
        f"        或 --follow-active 跟随当前活动文档",
        file=sys.stderr,
    )
    raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="offipy", description="offipy CLI")
    # P2-2 多实例：--port 指向指定端口的 server（env OFFIPY_SERVER_PORT 亦可）
    p.add_argument("--port", type=int, help="连指定端口的 server 实例（默认 8890）")
    sub = p.add_subparsers(dest="app", required=True)
    for app in schema.apps():
        if app == "feedback":  # #135：feedback 有专属子解析器（下方单独定义）
            continue
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    # #135：feedback 专属子解析器——generic REMAINDER 无 help/无 op 校验，改这里
    # 显式声明 5 个 op 与参数（op 白名单 + 必填由 argparse 兜底，缺参 exit 2）。
    # 参数名双写法（--rule-id/--rule_id、--feedback-dir/--feedback_dir）保留
    # 「全局归一化 - → _」承诺：underscore 写法继续可用（既有测试/文档不破）。
    fb = sub.add_parser(
        "feedback", help="反馈学习系统（train/status/append/recommend/apply/reschema）"
    )
    fb_op = fb.add_subparsers(dest="op", required=True)
    fb_train = fb_op.add_parser("train", help="离线训练反馈学习模型（numpy MLP ensemble）")
    fb_train.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        help="反馈库目录（默认 ~/.offipy）",
    )
    fb_train.add_argument("--seed", type=int, default=42, help="训练随机种子（默认 42）")
    fb_status = fb_op.add_parser("status", help="反馈学习状态：样本数/配对潜力/模型状态")
    fb_status.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        help="反馈库目录（默认 ~/.offipy）",
    )
    fb_append = fb_op.add_parser("append", help="追加一条反馈标签（fixed/accepted/ignored）")
    fb_append.add_argument("--profile", required=True, help="艺术分析 profile 名")
    fb_append.add_argument(
        "--rule-id", "--rule_id", dest="rule_id", required=True, help="finding 的 rule_id"
    )
    fb_append.add_argument(
        "--action", required=True, choices=["fixed", "accepted", "ignored"], help="处置"
    )
    fb_append.add_argument(
        "--severity", required=True, choices=["LOW", "MID", "HIGH"], help="严重度"
    )
    fb_append.add_argument(
        "--slide-index",
        "--slide_index",
        dest="slide_index",
        type=int,
        help="finding 所在 slide（0 基，可省）",
    )
    fb_append.add_argument("--message", default="", help="备注")
    fb_append.add_argument("--source", default="", help="来源标注")
    fb_append.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        help="反馈库目录（默认 ~/.offipy）",
    )
    fb_append.add_argument("--ts", help="ISO 时间戳（默认现在）")
    fb_append.add_argument(
        "--features", help="扁平特征 JSON 字符串（如 '{\"finding.confidence\": 0.5}'）"
    )
    fb_recommend = fb_op.add_parser(
        "recommend", help="对 .pptx 跑学习推理，返回调整 finding 与建议（只读）"
    )
    fb_recommend.add_argument("--pptx", required=True, help="目标 .pptx 路径")
    fb_recommend.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        required=True,
        help="反馈库目录（只读建议必须绑定）",
    )
    fb_recommend.add_argument(
        "--profile", default="balanced", help="艺术分析 profile（默认 balanced）"
    )
    fb_recommend.add_argument("--json", action="store_true", help="输出 JSON（默认即 JSON）")
    fb_apply = fb_op.add_parser("apply", help="把学习到的 rule.delta 持久化到 profile 存储")
    fb_apply.add_argument("--profile", required=True, help="目标 profile 名")
    fb_apply.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        help="反馈库目录（默认 ~/.offipy）",
    )
    fb_reschema = fb_op.add_parser(
        "reschema", help="把 schema 过期的历史反馈记录重写为当前 schema 版本（bump 迁移）"
    )
    fb_reschema.add_argument(
        "--feedback-dir",
        "--feedback_dir",
        dest="feedback_dir",
        help="反馈库目录（默认 ~/.offipy）",
    )
    deck = sub.add_parser("deck")
    deck.add_argument("action", choices=["make", "outline", "audit", "add-anim"])
    # audit 的 HTML 源走位置参数（task 5）：make/outline 传了会在分支内拦截。
    deck.add_argument("source", nargs="?", help="audit 的 HTML 源文件（位置参数）")
    # 专用选项（P0-4）：布尔用 _BoolAction，`--overwrite false` 不再是
    # bool("false")→True 的坑；未知 --key 由 argparse 直接 exit 2。
    deck.add_argument("--html", help="HTML 幻灯片源文件（make 必填）")
    deck.add_argument("--out", help="输出路径（make: .pptx；outline: .html）")
    deck.add_argument("--no-open", action="store_true", help="渲染后不打开实况演示（make）")
    deck.add_argument(
        "--export-png",
        "--feedback",
        dest="feedback",
        help="导出 PNG 反馈目录（make）；--feedback 为旧别名（弃用）",
    )
    deck.add_argument("--theme", help="注入内置主题名（make/outline）")
    deck.add_argument("--layouts", action=_BoolAction, help="注入 data-layout 布局 CSS（make）")
    deck.add_argument("--overwrite", action=_BoolAction, help="覆盖已存在的 .pptx（make）")
    deck.add_argument(
        "--animations",
        action="store_true",
        help="渲染 PPTX 注入入场动画/过渡（make；默认关）",
    )
    deck.add_argument(
        "--spec",
        help='动画/过渡 spec.json（add-anim）：{"animations": [...], "transitions": [...]}',
    )
    deck.add_argument(
        "--audit-mode",
        choices=["report", "strict"],
        help="make 渲染后跑质量审计：report 产出报告并替换；strict 达 --fail-on 拒绝替换（make）",
    )
    deck.add_argument(
        "--fail-on",
        choices=["HIGH", "MID", "LOW"],
        default="HIGH",
        help="strict 门禁严重度阈值，默认 HIGH（make，需 --audit-mode）",
    )
    deck.add_argument(
        "--audit-report",
        help="审计报告输出路径（扩展名定格式：.md/.json/.html，否则 text；make，需 --audit-mode）",
    )
    deck.add_argument("--input", help="大纲 markdown 源文件（outline）")
    deck.add_argument("--md", help="大纲 markdown 源文件别名（outline）")
    deck.add_argument("--pptx", help="audit：直接审计现成 .pptx（与位置 source 互斥）")
    deck.add_argument("--json", action="store_true", help="audit：输出结构化 JSON 建议")
    deck.add_argument(
        "--profile", default="balanced", help="audit：艺术分析 profile（默认 balanced）"
    )
    deck.add_argument(
        "--feedback-dir",
        help="audit：应用反馈学习（加载指定目录的记录/模型；feedback=True 需显式目录）",
    )
    # 参数名避开顶层 subparsers 的 dest "app"，否则 argparse 会用子解析器的
    # 值覆盖 args.app（例如 quit ppt → app 被覆盖成 "ppt"）。
    q = sub.add_parser("quit")
    q.add_argument("target", choices=["excel", "word", "ppt"])
    q.add_argument(
        "--force",
        action="store_true",
        help="即使连接的是既有 Office 实例也强制退出（own 实例缺省即可退）",
    )
    sub.add_parser("mcp", help="启动 MCP stdio server（Claude Desktop 等接入）")
    ck = sub.add_parser("check", help="检查环境就绪（Python/依赖/Office/浏览器/server）")
    ck.add_argument("--json", action="store_true", help="输出 JSON")
    ck.add_argument(
        "--profile",
        choices=["core", "office", "deck", "mcp", "all"],
        default=None,
        help="只检查指定 profile（core 为基线，office/deck/mcp/all 各叠加对应分组）",
    )
    srv = sub.add_parser("server", help="管理常驻 server（status/stop/restart）")
    srv.add_argument(
        "action",
        nargs="?",
        default="status",
        choices=["status", "stop", "restart"],
    )
    # 子解析器的 --port 用 SUPPRESS：未给时不覆盖顶层值（否则默认 None 会把
    # `offipy --port 8891 server status` 的端口冲掉）。
    srv.add_argument(
        "--port", type=int, default=argparse.SUPPRESS, help="目标 server 端口（默认 8890）"
    )
    lg = sub.add_parser("log", help="读取操作日志（oplog.jsonl，P2-3）")
    lg.add_argument("--tail", type=int, help="只显示末尾 N 条")
    lg.add_argument(
        "--port", type=int, default=argparse.SUPPRESS, help="目标 server 端口（默认 8890）"
    )
    au = sub.add_parser("audit", help="PPTX 质量审计 / 基线回归对比（不依赖 Office）")
    au.add_argument("file", help="目标 .pptx（对比模式下为候选）")
    au.add_argument(
        "--format",
        choices=["text", "json", "markdown", "html"],
        default="text",
        help="报告格式（默认 text）",
    )
    au.add_argument(
        "--out",
        help="输出文件；html 缺省写 <stem>.audit.html，其余缺省打 stdout",
    )
    au.add_argument(
        "--fail-on",
        choices=["HIGH", "MID", "LOW"],
        default="HIGH",
        help="审计达该严重度 → exit 1（默认 HIGH）",
    )
    au.add_argument("--baseline", help="基线 .pptx；给出则走回归对比（用 --fail-on-new 门槛）")
    au.add_argument(
        "--fail-on-new",
        choices=["HIGH", "MID", "LOW"],
        help="对比模式：候选新增/恶化问题达该严重度 → exit 1",
    )
    au.add_argument("--safe-margin", type=float, default=0.2, help="安全边距（英寸，默认 0.2）")
    au.add_argument(
        "--bounds-tolerance", type=float, default=0.01, help="越界容差（英寸，默认 0.01）"
    )
    au.add_argument("--no-full-bleed-ignore", action="store_true", help="关闭全页背景豁免")
    au.add_argument("--no-repeated-decoration-ignore", action="store_true", help="关闭重复装饰豁免")
    au.add_argument("--no-page-number-ignore", action="store_true", help="关闭页码豁免")
    au.add_argument("--no-header-footer-ignore", action="store_true", help="关闭页眉页脚豁免")
    au.add_argument(
        "--show-suppressed",
        action="store_true",
        help="豁免项默认已在报告中列出；本标志保留兼容",
    )
    au.add_argument(
        "--slides-dir",
        help="html 专用：PNG 页面背景目录（slide-<n>.png）",
    )
    au.add_argument("--debug", action="store_true", help="失败时打印完整 traceback")
    return p


def _main(argv: list[str] | None = None) -> int | None:
    args = build_parser().parse_args(argv)
    if getattr(args, "port", None):
        set_port(args.port)  # P2-2 多实例：后续调用指向该端口
    if args.app == "mcp":
        try:
            from .mcp_server import main as mcp_main
        except ImportError as exc:
            print(
                "offipy: error: offipy mcp 需要 mcp 扩展：pip install offipy[mcp]\n"
                f"  缺失依赖: {exc}",
                file=sys.stderr,
            )
            return 2
        mcp_main()
        return None
    if args.app == "check":
        from .envcheck import main as check_main

        return check_main(json_output=args.json, profile=getattr(args, "profile", None))
    if args.app == "audit":
        return _audit_main(args)
    if args.app == "quit":
        ensure_server()
        call(args.target, "quit", force=args.force)
        print("quit ok")
        return None
    if args.app == "server":
        if args.action == "status":
            # 只读探测（P0-3）：未运行不拉起，直接报状态
            st = server_status()
            if st is None:
                print("server 未在运行")
            else:
                print(json.dumps(st, ensure_ascii=False))
        elif args.action == "stop":
            print(stop_server())
        else:  # restart
            print(stop_server())
            ensure_server()
            print("server 已重启")
        return None
    if args.app == "log":
        from . import oplog

        oplog.configure(getattr(args, "port", None) or 8890)
        entries = oplog.read(tail=args.tail)
        if not entries:
            print("（暂无操作日志）")
            return None
        for e in entries:
            print(json.dumps(e, ensure_ascii=False))
        return None
    if args.app == "deck":
        if args.action == "make":
            if args.source:
                _usage_exit("offipy deck make 不接受位置参数，请用 --html <deck.html>")
            if args.audit_mode:
                return _deck_make_with_audit(args)
            from .deck import make as deck_make

            if not args.html:
                _usage_exit(
                    "用法: offipy deck make --html <deck.html> "
                    "[--out <x.pptx>] [--no-open] [--export-png <dir>] "
                    "[--theme <name>] [--layouts] [--overwrite]"
                )
            pptx = deck_make(
                args.html,
                out=args.out,
                open_live_flag=not args.no_open,
                feedback_dir=args.feedback,
                theme=args.theme,
                apply_layouts=args.layouts,
                overwrite=args.overwrite,
                animations=args.animations,
            )
            print(json.dumps({"pptx": pptx}, ensure_ascii=False))
        elif args.action == "outline":
            if args.source:
                _usage_exit("offipy deck outline 不接受位置参数，请用 --input <outline.md>")
            from .outline import parse_outline, to_deck_html

            md_path = args.input or args.md
            if not md_path:
                _usage_exit(
                    "用法: offipy deck outline --input <outline.md> "
                    "[--theme <name>] [--out <deck.html>]"
                )
            try:
                with Path(md_path).open(encoding="utf-8") as f:
                    outline = parse_outline(f.read())
            except FileNotFoundError:
                _usage_exit(f"找不到文件: {md_path}")
            except ValueError as e:
                _usage_exit(f"大纲格式错误: {e}")
            html = to_deck_html(outline, theme=args.theme)
            if args.out:
                try:
                    with Path(args.out).open("w", encoding="utf-8") as f:
                        f.write(html)
                except OSError as e:
                    _usage_exit(f"无法写入输出文件: {args.out}: {e}")
                print(json.dumps({"html": str(Path(args.out).resolve())}, ensure_ascii=False))
            else:
                print(outline.to_json())
        elif args.action == "add-anim":
            return _deck_add_anim(args)
        else:  # audit
            return _deck_audit(args)
        return None
    if args.app == "feedback":
        return _feedback_main(args)
    kw = _parse_kwargs(args.kwargs)
    _validate_kwargs(args.app, args.op, kw)
    _validate_required(args.app, args.op, kw)
    _validate_destructive_target(args.app, args.op, kw)
    kw = _coerce_kwargs(args.app, args.op, kw)
    result = call(args.app, args.op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))
    return None


def _feedback_main(args: argparse.Namespace) -> int | None:
    """feedback 专属分发：按子命令构造 kwargs，经 cli.call 转发（与 generic 同语义）。

    只在用户显式给出时传可选参（feedback_dir 等缺省 None → FeedbackApp 走 ~/.offipy）；
    必填参（profile/rule_id/action/severity、recommend 的 pptx/feedback_dir）由 argparse
    required 兜底，缺省 exit 2。recommend 的 --json 透传 True（既有捕获测试断言）。
    """
    op = args.op
    if op == "train":
        kw = {}
        if args.feedback_dir:
            kw["feedback_dir"] = args.feedback_dir
        if args.seed != 42:
            kw["seed"] = args.seed
    elif op == "status":
        kw = {}
        if args.feedback_dir:
            kw["feedback_dir"] = args.feedback_dir
    elif op == "append":
        kw = {
            "profile": args.profile,
            "rule_id": args.rule_id,
            "action": args.action,
            "severity": args.severity,
        }
        if args.slide_index is not None:
            kw["slide_index"] = args.slide_index
        if args.message:
            kw["message"] = args.message
        if args.source:
            kw["source"] = args.source
        if args.feedback_dir:
            kw["feedback_dir"] = args.feedback_dir
        if args.ts:
            kw["ts"] = args.ts
        if args.features:
            kw["features"] = args.features
    elif op == "recommend":
        kw = {"pptx": args.pptx, "feedback_dir": args.feedback_dir}
        if args.profile:
            kw["profile"] = args.profile
        if args.json:
            kw["json"] = True
    elif op == "reschema":
        kw = {}
        if args.feedback_dir:
            kw["feedback_dir"] = args.feedback_dir
    else:  # apply
        kw = {"profile": args.profile}
        if args.feedback_dir:
            kw["feedback_dir"] = args.feedback_dir
    result = call("feedback", op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))
    return None


def _force_utf8() -> None:
    """Windows 终端默认 GBK：统一把 stdout/stderr 切到 UTF-8，避免中文乱码（#149）。

    用 getattr 探测 reconfigure：被替换的流（测试捕获/重定向）没有该能力则跳过，
    不强制改写调用方流对象。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int | None:
    """offipy CLI 入口：InvalidArgumentError→exit 2；其余 OffipyError→exit 1。

    SystemExit/argparse 原样放行（非 OffipyError 内建异常从 _main 逃逸时
    保持裸 traceback + Python 默认 exit 1）。
    """
    _force_utf8()  # #149：先切 UTF-8，再进任何打印/错误路径
    try:
        return _main(argv)
    except InvalidArgumentError as e:
        print(f"offipy: {_clean_error_message(str(e))}", file=sys.stderr)
        return 2
    except OffipyError as e:
        print(f"offipy: {_clean_error_message(str(e))}", file=sys.stderr)
        return 1


# ---------------------------------------------------------------- audit 子命令


def _build_audit_config(args: argparse.Namespace) -> Any:
    from .audit import AuditConfig

    return AuditConfig(
        safe_margin_in=args.safe_margin,
        bounds_tolerance_in=args.bounds_tolerance,
        ignore_full_bleed_shapes=not args.no_full_bleed_ignore,
        ignore_repeated_decorations=not args.no_repeated_decoration_ignore,
        ignore_page_numbers=not args.no_page_number_ignore,
        ignore_headers_footers=not args.no_header_footer_ignore,
    )


def _audit_fail(args: argparse.Namespace, exc: BaseException) -> None:
    if getattr(args, "debug", False):
        traceback.print_exc()
    else:
        print(f"offipy: error: {exc}", file=sys.stderr)


def _audit_render(report: Any, args: argparse.Namespace) -> str:
    from .audit import render_html, render_markdown, render_text

    fmt = args.format
    if fmt == "text":
        return render_text(report)
    if fmt == "markdown":
        return render_markdown(report)
    if fmt == "json":
        return str(report.to_json())
    return render_html(report, slides_dir=args.slides_dir)


def _write_audit_report(path: str, report: Any) -> None:
    """按扩展名定格式落盘审计报告（.md/.json/.html，否则 text）。"""
    from .audit import render_html, render_markdown, render_text

    ext = Path(path).suffix.lower()
    if ext == ".md":
        text = render_markdown(report)
    elif ext == ".json":
        text = report.to_json()
    elif ext == ".html":
        text = render_html(report)
    else:
        text = render_text(report)
    with Path(path).open("w", encoding="utf-8") as f:
        f.write(text)
    print(f"offipy: 审计报告已写入 {path}")


def _deck_make_with_audit(args: argparse.Namespace) -> int | None:
    """offipy deck make --audit-mode：渲染 + 静态审计门禁。

    report：审计后替换，报告（若有 --audit-report）落盘；
    strict：达 --fail-on → 先写报告再 exit 1（旧目标不动），未达 → 替换。
    """
    from .audit import Severity
    from .deck import AuditGateError, export_slides, open_live, render_with_report

    if not args.html:
        _usage_exit(
            "用法: offipy deck make --html <deck.html> "
            "[--audit-mode report|strict] [--fail-on HIGH|MID|LOW] "
            "[--audit-report <path>] [--out <x.pptx>] [--no-open] "
            "[--export-png <dir>] [--theme <name>] [--layouts] [--overwrite] "
            "[--animations]"
        )
    fail_on = {"HIGH": Severity.HIGH, "MID": Severity.MID, "LOW": Severity.LOW}[args.fail_on]
    try:
        result = render_with_report(
            args.html,
            out=args.out,
            theme=args.theme,
            apply_layouts=args.layouts,
            overwrite=args.overwrite,
            audit_mode=args.audit_mode,
            fail_on=fail_on,
            animations=args.animations,
        )
    except AuditGateError as e:
        # strict 未通过：先落盘报告（若指定路径），再以 exit 1 收场（旧目标未动）。
        if args.audit_report:
            _write_audit_report(args.audit_report, e.report)
        sev = e.report.max_severity
        sev_name = sev.name if sev is not None else "?"
        print(
            f"offipy: 审计门槛未通过（最高 {sev_name} ≥ {args.fail_on}），未替换输出",
            file=sys.stderr,
        )
        return 1
    pptx = result.output_path
    if args.audit_report:
        _write_audit_report(args.audit_report, result.audit_report)
    if args.feedback:
        doc_id = open_live(pptx)
        export_slides(args.feedback, doc_id=doc_id, overwrite=args.overwrite)
    elif not args.no_open:
        open_live(pptx)
    print(json.dumps({"pptx": pptx}, ensure_ascii=False))
    return None


def _deck_add_anim(args: argparse.Namespace) -> int | None:
    """offipy deck add-anim --pptx X --spec spec.json：给现成 .pptx 注入动画/过渡。"""
    from .animations.apply import apply_animations
    from .animations.spec import AnimationSpec, TransitionSpec

    if not args.pptx:
        _usage_exit("用法: offipy deck add-anim --pptx <deck.pptx> --spec <spec.json>")
    if not args.spec:
        _usage_exit("用法: offipy deck add-anim --pptx <deck.pptx> --spec <spec.json>")
    try:
        with Path(args.spec).open(encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        _usage_exit(f"找不到 spec 文件: {args.spec}")
    except json.JSONDecodeError as e:
        _usage_exit(f"spec 不是合法 JSON: {args.spec}（{e}）")
    if not isinstance(data, dict):
        _usage_exit('spec.json 必须是对象：{"animations": [...], "transitions": [...]}')
    for key in ("animations", "transitions"):
        if key in data and not isinstance(data[key], list):
            _usage_exit(f"spec.json 的 {key} 必须是数组")
    animations = []
    for item in data.get("animations") or []:
        if not isinstance(item, dict):
            _usage_exit(f"animations 元素必须是对象（实际 {type(item).__name__}）")
        # 字段非法 → InvalidArgumentError 冒泡（main 兜底）
        animations.append(AnimationSpec(**item))
    transitions = []
    for item in data.get("transitions") or []:
        if not isinstance(item, dict):
            _usage_exit(f"transitions 元素必须是对象（实际 {type(item).__name__}）")
        transitions.append(TransitionSpec(**item))
    report = apply_animations(args.pptx, animations=animations, transitions=transitions)
    print(json.dumps({"pptx": args.pptx, **report}, ensure_ascii=False))
    return None


def _deck_audit(args: argparse.Namespace) -> int | None:
    """offipy deck audit：一次性分析（HTML 临时渲染 / PPTX 直读）+ 建议投影。

    HTML 流：render_with_quality_report 已内含一次几何审计 + 一次艺术分析，
    产物只发布到 TemporaryDirectory（命令退出即删），绝不二次 audit/build_scene；
    chromium 不可用 → ConversionError 上抛（main 转 offipy: <msg> + exit 1）。
    PPTX 流：analyze_deck(pptx=...) 一次。缺失文件 / 非法 --profile 属预运行无效输入
    → 友好 offipy: error + exit 2（与 offipy audit 参数错误一致，#39）。
    """
    from .exceptions import InvalidArgumentError

    if args.source and args.pptx:
        _usage_exit("请给出且只给一个源：位置 <html> 或 --pptx")
    if not args.source and not args.pptx:
        _usage_exit(
            "用法: offipy deck audit <deck.html> [--theme <name>] [--layouts] "
            "[--profile <p>] [--json] [--feedback-dir <dir>]\n"
            "      或 offipy deck audit --pptx <file.pptx> [--profile <p>] "
            "[--json] [--feedback-dir <dir>]\n"
            "      --feedback-dir <dir> 应用反馈学习（加载指定目录的记录/模型）"
        )

    if args.source:
        # 惰性 import：让测试能 patch offipy.deck.render_with_quality_report
        from .deck import render_with_quality_report

        try:
            with tempfile.TemporaryDirectory(prefix="offipy-deck-audit-") as td:
                result = render_with_quality_report(
                    args.source,
                    out=str(Path(td) / "audit.pptx"),
                    theme=args.theme,
                    apply_layouts=args.layouts,
                    overwrite=True,
                    profile=args.profile,
                    pixel_analysis="off",
                    feedback_dir=args.feedback_dir,
                    include_experimental_score=args.feedback_dir is not None,
                )
                report = result.deck_quality
        except KeyError as e:
            # 未知 --profile：get_profile 抛 KeyError，转友好 offipy: error + exit 2
            print(f"offipy: error: {e.args[0] if e.args else e}", file=sys.stderr)
            return 2
        if report is None:
            # 契约上不会发生：render_with_quality_report 总是产出 DeckQualityReport
            return 1
    else:
        from .art.analyze import analyze_deck

        try:
            report = analyze_deck(
                pptx=args.pptx,
                profile=args.profile,
                include_experimental_score=args.feedback_dir is not None,
                feedback=args.feedback_dir is not None,
                feedback_dir=args.feedback_dir,
            )
        except FileNotFoundError:
            print(f"offipy: error: 找不到文件: {args.pptx}", file=sys.stderr)
            return 2
        except InvalidArgumentError as e:
            if "不存在" in str(e):
                print(f"offipy: error: 找不到文件: {args.pptx}", file=sys.stderr)
                return 2
            raise
        except KeyError as e:
            # 未知 --profile：analyze_deck → get_profile 抛 KeyError，转友好报错
            print(f"offipy: error: {e.args[0] if e.args else e}", file=sys.stderr)
            return 2
    return _emit_deck_audit(report, args)


def _emit_deck_audit(report: Any, args: argparse.Namespace) -> int:
    """输出 deck audit 结果：--json 结构化；否则按维度分组文本。"""
    from .art.suggest import project_adjusted_findings, project_suggestions

    source = args.source or args.pptx
    suggestions = project_suggestions(report, source=source)
    adjusted = project_adjusted_findings(report)
    warnings = [{"code": w.code, "message": w.message} for w in report.warnings]
    art = report.art
    score = art.experimental_score if art else None
    score_mode = art.experimental_score_mode if art else None
    feedback_adjustments = dict(art.feedback_adjustments) if art else {}
    quality_score_coverage = art.quality_score_coverage if art else None
    if args.json:
        payload = {
            "source": source,
            "profile": args.profile,
            "experimental_score": score,
            "experimental_score_mode": score_mode,
            "quality_score_coverage": quality_score_coverage,
            "feedback_adjustments": feedback_adjustments,
            "adjusted_findings": adjusted,
            "warnings": warnings,
            "suggestions": suggestions,
        }
        print(json.dumps(payload, ensure_ascii=False))
    else:
        _print_deck_audit_text(
            args,
            warnings,
            suggestions,
            score=score,
            score_mode=score_mode,
            feedback_adjustments=feedback_adjustments,
            quality_score_coverage=quality_score_coverage,
            adjusted_findings=adjusted,
        )
    return 0


def _print_deck_audit_text(
    args: argparse.Namespace,
    warnings: list[dict[str, object]],
    suggestions: list[dict[str, object]],
    *,
    score: float | None = None,
    score_mode: str | None = None,
    feedback_adjustments: dict[str, int] | None = None,
    quality_score_coverage: dict[str, object] | None = None,
    adjusted_findings: list[dict[str, object]] | None = None,
) -> None:
    """按维度分组打印文本建议（确定性顺序：逐条记录，维度变化时出标题）。"""
    from .art.render import score_label

    lines = [f"offipy deck audit（profile={args.profile}）"]
    if score is not None:
        lines.append(f"{score_label(score_mode)}: {score}")
    if quality_score_coverage:
        cov = quality_score_coverage
        lines.append(
            "分数覆盖: "
            f"covered={cov['covered_findings']}/{cov['total_findings']} "
            f"(abstain {cov['abstained_count']}, ood {cov['ood_count']})"
        )
    for rule_id, delta in sorted((feedback_adjustments or {}).items()):
        lines.append(f"模型调整 (rule.delta) {rule_id}: {delta:+d}")
    lines.extend(
        f"  调整 {rec['rule_id']}: {rec['severity_before']}→{rec['severity_after']} "
        f"(worth={rec['worth']}, shift={rec['shift']})"
        for rec in adjusted_findings or []
    )
    if warnings:
        lines.append("警告:")
        lines.extend(f"- [{w['code']}] {w['message']}" for w in warnings)
    if not suggestions:
        lines.append("（无建议记录）")
    else:
        last_dim = None
        for rec in suggestions:
            dim = rec["dimension"]
            if dim != last_dim:
                lines.append("")
                lines.append(f"维度: {dim}")
                last_dim = dim
            slide_label = rec["slide_index"] if rec["slide_index"] is not None else "（全篇）"
            lines.append(
                f"  页 {slide_label} [{rec['severity']}] {rec['rule_id']}: {rec['message']}"
            )
            lines.append(f"    建议: {rec['suggestion']}")
    print("\n".join(lines))


def _audit_main(args: argparse.Namespace) -> int:
    """offipy audit 入口：自捕全部预期异常转 exit 码，绝不让 OffipyError 逃逸。

    退出码：0=未达门槛 / 1=成功但达 --fail-on 或 --fail-on-new / 2=参数或输入错 /
    3=依赖或解析错。--debug 时保留完整 traceback。
    """
    from .audit import PptxAuditReport, PptxDiffReport, Severity, audit_pptx, compare_pptx
    from .exceptions import ConversionError, InvalidArgumentError

    threshold_map = {"HIGH": Severity.HIGH, "MID": Severity.MID, "LOW": Severity.LOW}
    cfg = _build_audit_config(args)
    report: PptxAuditReport | PptxDiffReport
    try:
        if args.baseline:
            if args.fail_on_new is None:
                print(
                    "offipy: error: 对比模式需要 --fail-on-new （对候选新增/恶化问题设门槛）",
                    file=sys.stderr,
                )
                return 2
            report = compare_pptx(args.baseline, args.file, audit_config=cfg)
            threshold = threshold_map[args.fail_on_new]
            gate = report.gate_severity()
            triggered = gate is not None and gate >= threshold
        else:
            if args.fail_on_new is not None:
                print(
                    "offipy: error: --fail-on-new 只用于 --baseline 对比模式",
                    file=sys.stderr,
                )
                return 2
            report = audit_pptx(args.file, cfg)
            threshold = threshold_map[args.fail_on]
            gate = report.max_severity
            triggered = gate is not None and gate >= threshold
    except InvalidArgumentError as e:
        _audit_fail(args, e)
        return 2
    except (ConversionError, ImportError) as e:
        _audit_fail(args, e)
        return 3

    out_path = args.out
    if not out_path and args.format == "html":
        p = Path(args.file)
        out_path = str(p.with_name(p.stem + ".audit.html"))
    text = _audit_render(report, args)
    if out_path:
        with Path(out_path).open("w", encoding="utf-8") as f:
            f.write(text)
        print(f"offipy: 报告已写入 {out_path}")
    else:
        sys.stdout.write(text)
    return 1 if triggered else 0


if __name__ == "__main__":
    sys.exit(main())
