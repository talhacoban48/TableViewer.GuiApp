from __future__ import annotations

import os
from PyQt5.QtGui import QIcon

from .constants import ASSETS_DIR


def load_icon(name: str) -> QIcon:
    return QIcon(os.path.join(ASSETS_DIR, name))


def load_pixmap(name: str, w: int = 16, h: int = 16):
    return load_icon(name).pixmap(w, h)


def _try_float(val: str):
    """Return float(val) or None if not parseable."""
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _column_is_numeric(values: list) -> bool:
    """Return True when >50 % of non-empty values in `values` are numeric."""
    non_empty = [v for v in values if v != ""]
    if not non_empty:
        return False
    n_num = sum(1 for v in non_empty if _try_float(v) is not None)
    return n_num / len(non_empty) > 0.5
