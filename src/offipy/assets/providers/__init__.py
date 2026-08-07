"""offipy.assets.providers — built-in asset providers.

Providers register lazily into the default registry; nothing here scans the
vendored asset directories at import time.
"""

from offipy.assets.providers.icons import IconProvider, icon_render_mode

__all__ = ["IconProvider", "icon_render_mode"]
