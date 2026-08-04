"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import os

from . import core

PP_SAVE_PDF = 32  # ppSaveAsPDF
PP_LAYOUT_TITLE = 1
PP_LAYOUT_TEXT = 2
PP_LAYOUT_TITLE_ONLY = 5
PP_LAYOUT_BLANK = 12


class PptApp:
    def __init__(self, visible: bool = True):
        self.app, _ = core.ensure_app("ppt", visible=visible)
        self._pres = None

    # --- 演示文稿 ---
    def new_pres(self):
        self._pres = self.app.Presentations.Add()
        return self._pres

    def open_pres(self, path: str):
        self._pres = self.app.Presentations.Open(os.path.abspath(path))
        return self._pres

    def active_pres(self):
        if self._pres is not None:
            return self._pres
        pres = self.app.ActivePresentation
        if pres is None:
            pres = self.app.Presentations.Add()
            self._pres = pres
        return pres

    def save(self, path: str | None = None):
        pres = self.active_pres()
        if path:
            pres.SaveAs(os.path.abspath(path))
        else:
            pres.Save()

    def save_pdf(self, path: str):
        self.active_pres().SaveAs(os.path.abspath(path), PP_SAVE_PDF)

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
        slide.Shapes.AddPicture(path, 0, 0, left, top, width, height)

    def quit(self):
        core.quit_app("ppt")
