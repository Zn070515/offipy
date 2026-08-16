"""Capture PowerPoint's native <p:timing> XML for the 6 curated entrance effects.

Dev-only (docs/development/, gitignored). Requires a live PowerPoint instance.
Usage: uv run python scripts/capture_animation_templates.py
"""

import json
import sys
import zipfile
from pathlib import Path

# PowerPoint SaveAs 按它自己的进程 cwd 解析相对路径，必须用绝对路径，
# 否则 PowerPoint 会在它自己的 cwd 下找不到 build\anim_capture.pptx。
OUT_PPTX = Path(__file__).resolve().parent.parent / "build" / "anim_capture.pptx"


def add_effect(shape, effect_id, subtype=0):
    """给 shape 加一个入场效果，返回 (ok, info)。

    PowerPoint 对象模型里 Shape 没有 .Slide 属性（那是 Excel/Word 的），
    所在 Slide 通过 Shape.Parent 取；AddEffect 的第一个参数就是 Shape 本身。
    """
    try:
        slide = shape.Parent
        seq = slide.TimeLine.MainSequence
        # 只传必需的 Shape/effectId 两个参数：其余可选参（Level/trigger/Index）
        # 交给 PowerPoint 默认值（trigger 默认=1 即 msoAnimTriggerOnPageClick）。
        # 显式传 None 会让 win32com 对 VT_I4 参数做 int(None) 强转而抛错。
        seq.AddEffect(shape, effect_id)
        return True, {"preset": int(effect_id), "subtype": subtype}
    except Exception as e:
        return False, {"error": str(e)}


def main() -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Add()
    try:
        slide = pres.Slides.Add(1, 1)  # ppLayoutTitle
        # msoAnimEffect 常量（PowerPoint 对象模型）：
        # 1=Appear 2=FlyIn 10=Fade 11=FloatUp 13=GrowAndTurn 19=Wipe 20=ZoomIn
        # （grow 取 GrowAndTurn=13；若 PowerPoint 用 <p:animScale> 表达，Task 3
        #  按捕获结果改走 scale 模板。）
        effects = {
            "fade": 10,
            "float_up": 11,
            "fly_in": 2,
            "wipe": 19,
            "zoom_in": 20,
            "grow": 13,
        }
        results = {}
        # 每个效果一个独立 textbox：保证每个效果在 timing 里是一个干净的
        # <p:par>，方便逐效果提取它的原生 <p:anim*>/<p:set> XML。
        for i, (name, eff_id) in enumerate(effects.items()):
            shape = slide.Shapes.AddTextbox(1, 50, 100 + i * 60, 400, 40)
            ok, info = add_effect(shape, eff_id)
            results[name] = {"ok": ok, **info}
        OUT_PPTX.parent.mkdir(exist_ok=True)
        pres.SaveAs(str(OUT_PPTX))
    finally:
        pres.Close()
        app.Quit()

    # 解 zip 读 slide1.xml 的 <p:timing>
    with zipfile.ZipFile(OUT_PPTX) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    start = xml.find("<p:timing>")
    end = xml.find("</p:timing>") + len("</p:timing>")
    timing = xml[start:end] if start >= 0 else "<p:timing>（无——未产生动画）</p:timing>"
    doc = {
        "results": results,
        "timing_xml": timing,
        "note": "PPT 把这些效果存成 presetID 模板或 animEffect/animMotion。"
        "float_up/zoom 等可能在 AddEffect 里用扩展常量，见 output 校准。",
    }
    out = Path("docs/development/animation-templates.md")
    out.parent.mkdir(exist_ok=True)
    out.write_text(
        "```json\n" + json.dumps(doc, indent=2, ensure_ascii=False) + "\n```\n", encoding="utf-8"
    )
    print(json.dumps(doc, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
