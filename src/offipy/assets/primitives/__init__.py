"""offipy.assets.primitives — native editable presentation primitives (A5).

Each primitive module exposes a single ``render(slide, params, context)``
function that returns created python-pptx shapes in visual stacking order
bottom → top. Renderers are imported lazily so importing this package never
pulls in python-pptx until a native primitive is actually rendered.
"""

from __future__ import annotations

import importlib

from offipy.exceptions import InvalidArgumentError

_RENDERER_MODULES: dict[str, str] = {
    "quote-mark": "quote_mark",
    "section-number": "section_number",
    "label-pill": "label_pill",
    "metric-badge": "metric_badge",
    "timeline-node": "timeline_node",
    "process-arrow": "process_arrow",
    "device-frame": "device_frame",
    "browser-mockup": "browser_mockup",
}


def get_native_renderer(primitive: str):
    """Return the ``render(slide, params, ctx)`` callable for a primitive name.

    Unknown primitive reaching the renderer is a contract violation and raises.
    """
    module_name = _RENDERER_MODULES.get(primitive)
    if module_name is None:
        raise InvalidArgumentError(f"unknown native primitive {primitive!r}")
    module = importlib.import_module(f"offipy.assets.primitives.{module_name}")
    render = getattr(module, "render", None)
    if render is None:
        raise InvalidArgumentError(f"native primitive {primitive!r} has no render()")
    return render
