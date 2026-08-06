"""offipy 艺术层（v0.12）：演示文稿艺术分析与设计一致性系统。

只做建议（带 confidence 的可解释 Finding），默认不阻断；
strict 门禁仍归 0.11 几何审计层。

注意：Task 18 会补全完整公开导出集；目前只导出 deck.py 惰性 import 所需子集。
"""

from .adapters import build_scene
from .analyze import analyze_scene
from .models import ArtWarning, DeckQualityReport

__all__ = ["ArtWarning", "DeckQualityReport", "analyze_scene", "build_scene"]
