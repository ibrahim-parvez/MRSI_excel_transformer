"""Locating and loading the files in ``assets/``.

Asset lookup has to work in two very different layouts:

* a source checkout, where the assets sit next to ``src/``;
* a PyInstaller build, where they are unpacked into a temporary directory
  that PyInstaller advertises as ``sys._MEIPASS``.

Everything in the app goes through :func:`asset_path` so neither case needs
to be special-cased at the call site. The specs in ``packaging/`` are what
put ``assets/images`` inside the build; if that ever stops happening the
loaders below degrade to a null pixmap rather than raising, and callers
fall back to a text label.
"""

from pathlib import Path
import sys

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPixmap

__all__ = ["ASSETS_DIR", "asset_path", "logo_pixmap", "app_icon"]

_IMAGES = ("assets", "images")


def _assets_root() -> Path:
    """Return the directory that holds ``assets/``."""
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled:
        return Path(bundled)

    # Walk up from this file until we find the checkout that owns the assets.
    # Doing it by search rather than by a fixed number of ``.parent`` hops
    # keeps this module free to move around inside src/.
    for candidate in Path(__file__).resolve().parents:
        if (candidate / Path(*_IMAGES)).is_dir():
            return candidate
    return Path(__file__).resolve().parents[2]


ASSETS_DIR = _assets_root() / Path(*_IMAGES)

#: The window/taskbar icon, in the format each platform expects.
_ICON_FILE = "logo.icns" if sys.platform == "darwin" else "logo.ico"


def asset_path(name: str) -> Path:
    """Absolute path to ``name`` inside ``assets/images``."""
    return ASSETS_DIR / name


def logo_pixmap(size: int | None = None) -> QPixmap:
    """The MRSI logo, optionally scaled to fit a ``size`` x ``size`` box.

    Returns a null pixmap if the asset is missing, so callers can test
    ``pixmap.isNull()`` and substitute a text label.
    """
    pixmap = QPixmap(str(asset_path("mrsi_logo.png")))
    if pixmap.isNull() or size is None:
        return pixmap
    return pixmap.scaled(
        size,
        size,
        Qt.AspectRatioMode.KeepAspectRatio,
        Qt.TransformationMode.SmoothTransformation,
    )


def app_icon() -> QIcon:
    """The application icon, for window decorations and the taskbar."""
    icon = QIcon(str(asset_path(_ICON_FILE)))
    if icon.isNull():
        icon = QIcon(str(asset_path("logo.png")))
    return icon
