from __future__ import annotations

from PyQt5.QtCore import Qt, QSortFilterProxyModel

from .utils import _try_float


class MultiColumnFilterProxyModel(QSortFilterProxyModel):
    """
    Supports two filter types per column, stored as a spec dict:
      {'type': 'values',  'values': set[str]}
      {'type': 'number',  'op': str, 'val': float}          # single-value ops
      {'type': 'number',  'op': 'between', 'val1': float, 'val2': float}
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._filters: dict = {}   # col_index -> spec dict
        self._global_filter: str = ""

    def set_global_filter(self, text: str):
        self._global_filter = text.lower().strip()
        self.invalidateFilter()

    def set_column_filter(self, col: int, spec):
        """Pass spec=None to clear the filter for that column."""
        if spec is None:
            self._filters.pop(col, None)
        else:
            self._filters[col] = spec
        self.invalidateFilter()

    def get_column_filter(self, col: int):
        return self._filters.get(col)

    def has_filter(self, col: int) -> bool:
        return col in self._filters

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:
        model = self.sourceModel()
        n_cols = model.columnCount()

        if self._global_filter:
            found = any(
                self._global_filter in (model.data(
                    model.index(source_row, c, source_parent), Qt.DisplayRole
                ) or "").lower()
                for c in range(n_cols)
            )
            if not found:
                return False

        for col, spec in self._filters.items():
            idx = model.index(source_row, col, source_parent)
            val = model.data(idx, Qt.DisplayRole) or ""
            if spec['type'] == 'values':
                if val not in spec['values']:
                    return False
            elif spec['type'] == 'number':
                if not self._matches_number_filter(val, spec):
                    return False
        return True

    @staticmethod
    def _matches_number_filter(val_str: str, spec: dict) -> bool:
        num = _try_float(val_str)
        if num is None:
            return False
        op = spec['op']
        if op == 'between':
            return spec['val1'] <= num <= spec['val2']
        v = spec['val']
        if op == '=':  return num == v
        if op == '≠':  return num != v
        if op == '>':  return num > v
        if op == '>=': return num >= v
        if op == '<':  return num < v
        if op == '<=': return num <= v
        return True
