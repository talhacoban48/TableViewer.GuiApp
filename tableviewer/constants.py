from __future__ import annotations

import os
from PyQt5.QtCore import Qt

SUPPORTED_EXTENSIONS = ('.xlsx', '.xls', '.csv')
CSV_ENCODINGS = ['utf-8-sig', 'utf-8', 'cp1254', 'latin-1', 'iso-8859-9']

# assets/ sits next to the tableviewer/ package (one level up)
ASSETS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "assets",
)

# Custom item data role: stores a set of property names explicitly set on a cell
# (e.g. {'bold', 'size', 'fg'}).  Used by _fmt_from_item to avoid inheriting app defaults.
_FMT_KEYS_ROLE = Qt.UserRole + 100

# Operators shown in the "Number Filter" tab
NUMBER_OPS = ['=', '≠', '>', '>=', '<', '<=', 'between']
