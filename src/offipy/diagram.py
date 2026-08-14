"""diagram app：把宿主 agent 用 diagram-design skill 设计的 Mermaid/drawio 图
转为可编辑 PPTX，并安装 skill 到宿主 agent 技能目录。

Agent 原生模式：offipy 自身不调用 LLM、不 spawn agent。skill 注册给宿主 agent
（Claude Code/Codex），agent 按产物契约把设计落地为 .mmd/.drawio 文件，
本模块只做「产物 → 可编辑 PPTX」。顶层 import 仅标准库，绝不顶层 import
pptx（惰性 import 红线，与 diagrams.py/drawio.py 一致）。
"""

import re
from pathlib import Path

_MMD_EXTS = {".mmd", ".mermaid", ".md"}
_DRAWIO_EXTS = {".drawio"}

_MMD_MARKER_RE = re.compile(
    r"\b(graph|flowchart|subgraph|sequenceDiagram|classDiagram|stateDiagram|"
    r"erDiagram|gantt|journey|mindmap|timeline)\b"
)


def _detect_format(path: Path) -> str:
    """扩展名优先（.mmd/.md→mermaid，.drawio→drawio），内容嗅探兜底。"""
    ext = path.suffix.lower()
    if ext in _MMD_EXTS:
        return "mermaid"
    if ext in _DRAWIO_EXTS:
        return "drawio"
    text = path.read_text(encoding="utf-8", errors="replace")[:8192]
    if "mxGraphModel" in text:
        return "drawio"
    if _MMD_MARKER_RE.search(text):
        return "mermaid"
    raise ValueError(
        "无法识别图格式（支持 Mermaid .mmd/.md/.mermaid 与 draw.io .drawio）："
        f"{path}（扩展名未命中，内容无 Mermaid 标记或 mxGraphModel）"
    )


class DiagramApp:
    """纯函数 app：无 COM 根（server._alive 对它恒 True），不绑定 Office 文档目标。"""

    def build(self, source, out, *, direction=None, page=None):
        """Mermaid/drawio 源码文件 → 可编辑 PPTX（16:9 整页）。返回 {"pptx": out}。

        source 必须是已存在文件路径（不接受内联文本）。格式按扩展名+内容自动识别。
        direction 仅 Mermaid 用、page 仅 draw.io 用——不适用时不透传。
        """
        path = Path(source)
        if not path.is_file():
            raise FileNotFoundError(f"源文件不存在: {path}")
        if _detect_format(path) == "mermaid":
            from .diagrams import mermaid_to_pptx

            return {"pptx": mermaid_to_pptx(str(path), out, direction=direction)}
        from .drawio import drawio_to_pptx

        return {"pptx": drawio_to_pptx(str(path), out, page=page)}
