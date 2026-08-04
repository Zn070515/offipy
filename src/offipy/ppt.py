"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import os
from contextlib import contextmanager

from . import core
from ._comguard import guard_com
from .exceptions import TargetNotFoundError
from .paths import ensure_writable

PP_ALERTS_NONE = 1  # ppAlertsNone（=0 是 ppAlertsAll）
PP_FIXED_FORMAT_TYPE_PDF = 2  # ppFixedFormatTypePDF（ExportAsFixedFormat 的 OutputType）
PP_LAYOUT_TITLE = 1
PP_LAYOUT_TEXT = 2
PP_LAYOUT_TITLE_ONLY = 5
PP_LAYOUT_BLANK = 12


@guard_com
class PptApp:
    def __init__(self, visible: bool = True):
        self.app, _ = core.ensure_app("ppt", visible=visible)
        # DisplayAlerts 不再永久静音（P0-5）：按需用 _alerts_scope 临时抑制
        self._saved_alerts = self.app.DisplayAlerts  # quit() 兜底还原
        self._pres = None

    @contextmanager
    def _alerts_scope(self, value: int = PP_ALERTS_NONE):
        """临时抑制模态对话框；退出时（含异常路径）还原 DisplayAlerts 原值。"""
        prev = self.app.DisplayAlerts
        self.app.DisplayAlerts = value
        try:
            yield
        finally:
            self.app.DisplayAlerts = prev

    # --- 演示文稿 ---
    def new_pres(self):
        self._pres = self.app.Presentations.Add()
        return self._pres

    def open_pres(self, path: str):
        self._pres = self.app.Presentations.Open(os.path.abspath(path))
        return self._pres

    def active_pres(self):
        # 会话语义（P1.2）：优先解析实时 ActivePresentation（用户当前激活的
        # 文稿），仅当无活动文稿时回退缓存句柄 + liveness probe。
        # P0-8：无活动文稿时返回 None（纯探测），绝不隐式 Presentations.Add()。
        pres = core.active_doc("ppt", "ActivePresentation")
        if pres is not None:
            self._pres = pres
            return pres
        if self._pres is not None and core.doc_alive(self._pres):
            return self._pres
        pres = self.app.ActivePresentation
        if pres is None:
            return None
        self._pres = pres
        return pres

    def _require_pres(self):
        """操作前置：无活动演示文稿则抛 TargetNotFoundError，不隐式创建。"""
        pres = self.active_pres()
        if pres is None:
            raise TargetNotFoundError("没有打开的演示文稿，请先 new_pres/open_pres")
        return pres

    def get_target(self):
        """当前活动演示文稿身份（app/name/path）；无则返回 None。只读探测。"""
        pres = self.active_pres()
        if pres is None:
            return None
        try:
            name = pres.Name
        except Exception:
            name = None
        try:
            path = pres.FullName
        except Exception:
            path = None
        return {"app": "ppt", "name": name, "path": path}

    def save(self, path: str | None = None, overwrite: bool = False):
        dest = ensure_writable(path, overwrite) if path else None
        with self._alerts_scope():
            pres = self._require_pres()
            if dest:
                pres.SaveAs(dest)
            else:
                pres.Save()

    def save_pdf(self, path: str, overwrite: bool = False):
        dest = ensure_writable(path, overwrite)
        # ExportAsFixedFormat 第二位置参数是 Intent（打印=2），OutputType 才是
        # 输出格式——必须显式指定 PDF，不能只传一个 2 了事。
        with self._alerts_scope():
            self._require_pres().ExportAsFixedFormat(
                dest, Intent=2, OutputType=PP_FIXED_FORMAT_TYPE_PDF
            )

    def export_slides(self, out_dir: str, width: int = 1920, height: int = 1080):
        """把当前演示文稿每一页导出为 PNG，供 Claude 视觉迭代。"""
        out_dir = os.path.abspath(out_dir)
        pres = self._require_pres()
        os.makedirs(out_dir, exist_ok=True)
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            out = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out, "PNG", width, height)
            paths.append(out)
        return paths

    # --- 幻灯片 ---
    def add_slide(self, layout: int = PP_LAYOUT_TEXT):
        pres = self._require_pres()
        pres.Slides.Add(pres.Slides.Count + 1, layout)
        return pres.Slides.Count

    def set_title(self, slide_idx: int, text: str):
        slide = self._require_pres().Slides(slide_idx)
        if slide.Shapes.HasTitle:
            slide.Shapes.Title.TextFrame.TextRange.Text = text

    def set_body(self, slide_idx: int, lines):
        slide = self._require_pres().Slides(slide_idx)
        if isinstance(lines, str):
            lines = [lines]
        ph = slide.Shapes.Placeholders(2)
        ph.TextFrame.TextRange.Text = "\r".join(lines)

    def set_notes(self, slide_idx: int, text: str):
        slide = self._require_pres().Slides(slide_idx)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text

    def add_textbox(
        self, slide_idx: int, left: float, top: float, width: float, height: float, text: str
    ):
        slide = self._require_pres().Slides(slide_idx)
        tb = slide.Shapes.AddTextbox(1, left, top, width, height)
        tb.TextFrame.TextRange.Text = text

    def add_picture(
        self, slide_idx: int, path: str, left: float, top: float, width: float, height: float
    ):
        slide = self._require_pres().Slides(slide_idx)
        slide.Shapes.AddPicture(os.path.abspath(path), 0, 0, left, top, width, height)

    def read_slide_texts(self):
        """逐页读取幻灯片文本：标题/正文占位符/备注，缺字段用空串。返回 list[dict]。"""
        pres = self._require_pres()
        result = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            title = ""
            if slide.Shapes.HasTitle:
                try:
                    title = slide.Shapes.Title.TextFrame.TextRange.Text
                except Exception:
                    title = ""
            body = ""
            try:
                body = slide.Shapes.Placeholders(2).TextFrame.TextRange.Text
            except Exception:
                body = ""
            notes = ""
            try:
                notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            except Exception:
                notes = ""
            result.append({"index": i, "title": title, "body": body, "notes": notes})
        return result

    def quit(self):
        # 库改全局状态（DisplayAlerts），释放前还原原值
        self.app.DisplayAlerts = self._saved_alerts
        core.quit_app("ppt")
