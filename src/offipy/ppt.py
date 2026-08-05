"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import os
from contextlib import contextmanager
from typing import Any

from . import core
from ._comguard import guard_com
from .exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError
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
        self._docs: dict[str, Any] = {}  # doc_id → 演示文稿句柄（P2-2 多文档）
        self._active_id: str | None = None
        self._seq = 0

    @contextmanager
    def _alerts_scope(self, value: int = PP_ALERTS_NONE):
        """临时抑制模态对话框；退出时（含异常路径）还原 DisplayAlerts 原值。"""
        prev = self.app.DisplayAlerts
        self.app.DisplayAlerts = value
        try:
            yield
        finally:
            self.app.DisplayAlerts = prev

    def _register(self, obj) -> str:
        """登记一个新文档句柄，分配 doc_id 并设为活动。返回 doc_id。"""
        self._seq += 1
        did = f"pres{self._seq}"
        self._docs[did] = obj
        self._active_id = did
        return did

    def _sync_registered(self, obj) -> str:
        """把实时解析到的句柄并入文档表：已登记则复用并置活动，否则登记为新文档。"""
        for did, pres in self._docs.items():
            if pres is obj:
                self._active_id = did
                return did
        return self._register(obj)

    # --- 演示文稿（P2-2 多文档：doc_id 显式路由，缺省走活动） ---
    def new_pres(self) -> str:
        """新建空白演示文稿，登记进文档表并设为活动。返回 doc_id。"""
        return self._register(self.app.Presentations.Add())

    def open_pres(self, path: str) -> str:
        """打开现有演示文稿并设为活动。返回 doc_id。"""
        return self._register(self.app.Presentations.Open(os.path.abspath(path)))

    def active_pres(self, doc_id: str | None = None):
        # 显式 doc_id：绑定目标路由，只查文档表；未知/失效句柄抛 TargetNotFoundError。
        # 缺省 active：优先 app 登记的活动句柄（new_pres/open_pres/activate 切换）；
        # 无登记或失效时实时解析 ActivePresentation 并并入文档表（重连既有会话场景）。
        # P0-8：全程纯探测，绝不隐式 Presentations.Add()。
        if doc_id is not None:
            pres = self._docs.get(doc_id)
            if pres is None or not core.doc_alive(pres):
                raise TargetNotFoundError(
                    f"未知演示文稿句柄: {doc_id!r}（用 list_docs 查看当前打开的）"
                )
            return pres
        if self._active_id is not None:
            pres = self._docs.get(self._active_id)
            if pres is not None and core.doc_alive(pres):
                return pres
        pres = core.active_doc("ppt", "ActivePresentation")
        if pres is not None:
            self._sync_registered(pres)
            return pres
        pres = self.app.ActivePresentation
        if pres is None:
            return None
        self._sync_registered(pres)
        return pres

    def _require_pres(self, doc_id: str | None = None):
        """操作前置：目标演示文稿不存在则抛 TargetNotFoundError，不隐式创建。"""
        pres = self.active_pres(doc_id)
        if pres is None:
            raise TargetNotFoundError("没有打开的演示文稿，请先 new_pres/open_pres")
        return pres

    def activate(self, doc_id: str) -> str:
        """把指定文档设为活动目标并同步真实 UI；未知句柄抛 TargetNotFoundError。"""
        pres = self._docs.get(doc_id)
        if pres is None or not core.doc_alive(pres):
            raise TargetNotFoundError(
                f"未知演示文稿句柄: {doc_id!r}（用 list_docs 查看当前打开的）"
            )
        old = self._active_id
        self._active_id = doc_id
        try:
            # PowerPoint 的 Presentation 无 Activate：激活其文档窗口；失败兜底
            try:
                pres.Windows.Item(1).Activate()
            except Exception:
                pres.Activate()
        except Exception as e:
            self._active_id = old  # 同步不上则回滚，不静默假活
            raise ComOperationError(f"激活演示文稿 {doc_id} 失败: {e}") from e
        return doc_id

    def list_docs(self) -> dict:
        """当前打开的文档表：{doc_id: {"name", "path", "active"}}。只报已登记句柄，不隐式枚举。"""
        out = {}
        for did, pres in self._docs.items():
            if not core.doc_alive(pres):
                continue
            try:
                name = pres.Name
            except Exception:
                name = None
            try:
                path = pres.FullName
            except Exception:
                path = None
            out[did] = {"name": name, "path": path, "active": did == self._active_id}
        return out

    def get_target(self, doc_id: str | None = None):
        """目标身份 {app, doc_id, name, path}；无目标返回 None。只读探测。

        显式 doc_id：只查文档表，未注册/失效抛 TargetNotFoundError；
        缺省：当前活动目标。
        """
        if doc_id is not None:
            pres = self.active_pres(doc_id)
            resolved = doc_id
        else:
            pres = self.active_pres()
            if pres is None:
                return None
            active_id = self._active_id
            assert active_id is not None  # active_pres 非 None 时活动 id 必已同步
            resolved = active_id
        try:
            name = pres.Name
        except Exception:
            name = None
        try:
            path = pres.FullName
        except Exception:
            path = None
        return {"app": "ppt", "doc_id": resolved, "name": name, "path": path}

    def save(self, path: str | None = None, overwrite: bool = False, doc_id: str | None = None):
        dest = ensure_writable(path, overwrite) if path else None
        with self._alerts_scope():
            pres = self._require_pres(doc_id)
            if dest:
                pres.SaveAs(dest)
            else:
                pres.Save()

    def save_pdf(self, path: str, overwrite: bool = False, doc_id: str | None = None):
        dest = ensure_writable(path, overwrite)
        # ExportAsFixedFormat 第二位置参数是 Intent（打印=2），OutputType 才是
        # 输出格式——必须显式指定 PDF，不能只传一个 2 了事。
        with self._alerts_scope():
            self._require_pres(doc_id).ExportAsFixedFormat(
                dest, Intent=2, OutputType=PP_FIXED_FORMAT_TYPE_PDF
            )

    def export_slides(
        self, out_dir: str, width: int = 1920, height: int = 1080, doc_id: str | None = None
    ):
        """把当前演示文稿每一页导出为 PNG，供 Claude 视觉迭代。"""
        out_dir = os.path.abspath(out_dir)
        pres = self._require_pres(doc_id)
        os.makedirs(out_dir, exist_ok=True)
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            out = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out, "PNG", width, height)
            paths.append(out)
        return paths

    # --- 幻灯片 ---
    def add_slide(self, layout: int = PP_LAYOUT_TEXT, doc_id: str | None = None):
        pres = self._require_pres(doc_id)
        pres.Slides.Add(pres.Slides.Count + 1, layout)
        return pres.Slides.Count

    def set_title(self, slide_idx: int, text: str, doc_id: str | None = None):
        if not text:
            raise InvalidArgumentError("set_title: text 不能为空")
        slide = self._require_pres(doc_id).Slides(slide_idx)
        if slide.Shapes.HasTitle:
            slide.Shapes.Title.TextFrame.TextRange.Text = text

    def set_body(self, slide_idx: int, lines, doc_id: str | None = None):
        if isinstance(lines, str):
            lines = [lines]
        if not lines:
            raise InvalidArgumentError("set_body: lines 不能为空")
        slide = self._require_pres(doc_id).Slides(slide_idx)
        ph = slide.Shapes.Placeholders(2)
        ph.TextFrame.TextRange.Text = "\r".join(lines)

    def set_notes(self, slide_idx: int, text: str, doc_id: str | None = None):
        slide = self._require_pres(doc_id).Slides(slide_idx)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text

    def add_textbox(
        self,
        slide_idx: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str,
        doc_id: str | None = None,
    ):
        slide = self._require_pres(doc_id).Slides(slide_idx)
        tb = slide.Shapes.AddTextbox(1, left, top, width, height)
        tb.TextFrame.TextRange.Text = text

    def add_picture(
        self,
        slide_idx: int,
        path: str,
        left: float,
        top: float,
        width: float,
        height: float,
        doc_id: str | None = None,
    ):
        slide = self._require_pres(doc_id).Slides(slide_idx)
        slide.Shapes.AddPicture(os.path.abspath(path), 0, 0, left, top, width, height)

    def read_slide_texts(self, doc_id: str | None = None):
        """逐页读取幻灯片文本：标题/正文占位符/备注，缺字段用空串。返回 list[dict]。"""
        pres = self._require_pres(doc_id)
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
