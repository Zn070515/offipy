"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import os

from . import core
from .paths import ensure_writable

PP_ALERTS_NONE = 1  # ppAlertsNone（=0 是 ppAlertsAll）
PP_FIXED_FORMAT_TYPE_PDF = 2  # ppFixedFormatTypePDF（ExportAsFixedFormat 的 OutputType）
PP_LAYOUT_TITLE = 1
PP_LAYOUT_TEXT = 2
PP_LAYOUT_TITLE_ONLY = 5
PP_LAYOUT_BLANK = 12


class PptApp:
    def __init__(self, visible: bool = True):
        self.app, _ = core.ensure_app("ppt", visible=visible)
        self._saved_alerts = self.app.DisplayAlerts  # 库改全局状态，释放时还原
        self.app.DisplayAlerts = PP_ALERTS_NONE
        self._pres = None

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
        pres = core.active_doc("ppt", "ActivePresentation")
        if pres is not None:
            self._pres = pres
            return pres
        if self._pres is not None and core.doc_alive(self._pres):
            return self._pres
        pres = self.app.ActivePresentation
        if pres is None:
            pres = self.app.Presentations.Add()
        self._pres = pres
        return pres

    def save(self, path: str | None = None, overwrite: bool = False):
        dest = ensure_writable(path, overwrite) if path else None
        pres = self.active_pres()
        if dest:
            pres.SaveAs(dest)
        else:
            pres.Save()

    def save_pdf(self, path: str, overwrite: bool = False):
        dest = ensure_writable(path, overwrite)
        # ExportAsFixedFormat 第二位置参数是 Intent（打印=2），OutputType 才是
        # 输出格式——必须显式指定 PDF，不能只传一个 2 了事。
        self.active_pres().ExportAsFixedFormat(dest, Intent=2, OutputType=PP_FIXED_FORMAT_TYPE_PDF)

    def export_slides(self, out_dir: str, width: int = 1920, height: int = 1080):
        """把当前演示文稿每一页导出为 PNG，供 Claude 视觉迭代。"""
        out_dir = os.path.abspath(out_dir)
        pres = self.active_pres()
        os.makedirs(out_dir, exist_ok=True)
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            out = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out, "PNG", width, height)
            paths.append(out)
        return paths

    # --- 幻灯片 ---
    def add_slide(self, layout: int = PP_LAYOUT_TEXT):
        pres = self.active_pres()
        pres.Slides.Add(pres.Slides.Count + 1, layout)
        return pres.Slides.Count

    def set_title(self, slide_idx: int, text: str):
        slide = self.active_pres().Slides(slide_idx)
        if slide.Shapes.HasTitle:
            slide.Shapes.Title.TextFrame.TextRange.Text = text

    def set_body(self, slide_idx: int, lines):
        slide = self.active_pres().Slides(slide_idx)
        if isinstance(lines, str):
            lines = [lines]
        ph = slide.Shapes.Placeholders(2)
        ph.TextFrame.TextRange.Text = "\r".join(lines)

    def set_notes(self, slide_idx: int, text: str):
        slide = self.active_pres().Slides(slide_idx)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text

    def add_textbox(
        self, slide_idx: int, left: float, top: float, width: float, height: float, text: str
    ):
        slide = self.active_pres().Slides(slide_idx)
        tb = slide.Shapes.AddTextbox(1, left, top, width, height)
        tb.TextFrame.TextRange.Text = text

    def add_picture(
        self, slide_idx: int, path: str, left: float, top: float, width: float, height: float
    ):
        slide = self.active_pres().Slides(slide_idx)
        slide.Shapes.AddPicture(os.path.abspath(path), 0, 0, left, top, width, height)

    def read_slide_texts(self):
        """逐页读取幻灯片文本：标题/正文占位符/备注，缺字段用空串。返回 list[dict]。"""
        pres = self.active_pres()
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
