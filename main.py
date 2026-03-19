from __future__ import annotations

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableView, QFileDialog, QAction,
    QStatusBar, QLabel, QMessageBox, QHeaderView, QFrame, QVBoxLayout,
    QHBoxLayout, QLineEdit, QCheckBox, QListWidget, QListWidgetItem,
    QPushButton, QTabWidget, QWidget, QComboBox,
)
from PyQt5.QtGui import QKeySequence, QStandardItemModel, QStandardItem, QIcon, QPainter, QPalette
from PyQt5.QtCore import Qt, pyqtSignal, QSortFilterProxyModel, QRect, QRectF, QPoint

SUPPORTED_EXTENSIONS = ('.xlsx', '.xls', '.csv')
CSV_ENCODINGS = ['utf-8-sig', 'utf-8', 'cp1254', 'latin-1', 'iso-8859-9']
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

# Operators shown in the "Number Filter" tab
NUMBER_OPS = ['=', '≠', '>', '>=', '<', '<=', 'between']


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


# ---------------------------------------------------------------------------
# Multi-column filter proxy model
# ---------------------------------------------------------------------------

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

    # -- filtering logic --

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:
        for col, spec in self._filters.items():
            idx = self.sourceModel().index(source_row, col, source_parent)
            val = self.sourceModel().data(idx, Qt.DisplayRole) or ""
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
            return False          # non-numeric values are excluded by a number filter
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


# ---------------------------------------------------------------------------
# Column filter popup  (Excel-style)
# ---------------------------------------------------------------------------

class FilterPopup(QFrame):

    filter_changed = pyqtSignal(int, object)   # col_index, spec | None

    def __init__(self, col_index: int, all_values: list,
                 current_filter, is_numeric: bool, parent=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.col_index = col_index
        self.is_numeric = is_numeric
        self.current_filter = current_filter   # the full spec dict or None

        # Sorted value list: blanks last, then case-insensitive
        self.all_values = sorted(all_values, key=lambda v: (v == "", v.lower()))

        # current_selection mirrors which values are checked
        if current_filter is None or current_filter.get('type') == 'number':
            self.current_selection = set(self.all_values)
        else:
            self.current_selection = set(current_filter['values'])

        self.setFrameShape(QFrame.StyledPanel)
        self.setFixedWidth(290)
        self._build_ui()

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        if self.is_numeric:
            self._tabs = QTabWidget()

            # ---- Values tab ----
            val_widget = QWidget()
            val_layout = QVBoxLayout(val_widget)
            val_layout.setContentsMargins(4, 4, 4, 4)
            val_layout.setSpacing(4)
            self._build_values_widgets(val_layout)
            self._tabs.addTab(val_widget, load_icon("filter.ico"), "Values")

            # ---- Number filter tab ----
            num_widget = QWidget()
            num_layout = QVBoxLayout(num_widget)
            num_layout.setContentsMargins(4, 8, 4, 4)
            num_layout.setSpacing(8)
            self._build_number_widgets(num_layout)
            self._tabs.addTab(num_widget, load_icon("sort.ico"), "Number Filter")

            # Open the number tab if an active number filter exists
            if self.current_filter and self.current_filter.get('type') == 'number':
                self._tabs.setCurrentIndex(1)

            outer.addWidget(self._tabs)
        else:
            self._tabs = None
            self._build_values_widgets(outer)

        # ---- OK / Cancel ----
        btn_row = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self._apply)
        cancel_btn = QPushButton(load_icon("cancel.ico"), "Cancel")
        cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        outer.addLayout(btn_row)

    def _build_values_widgets(self, layout: QVBoxLayout):
        # Search row
        search_row = QHBoxLayout()
        lbl = QLabel()
        lbl.setPixmap(load_pixmap("search.ico"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search values…")
        self.search_input.textChanged.connect(self._on_search_changed)
        search_row.addWidget(lbl)
        search_row.addWidget(self.search_input)
        layout.addLayout(search_row)

        # Select All (tristate)
        self.select_all_cb = QCheckBox("(Select All)")
        self.select_all_cb.setTristate(True)
        self._refresh_select_all()
        self.select_all_cb.stateChanged.connect(self._on_select_all_changed)
        layout.addWidget(self.select_all_cb)

        # Value list
        self.list_widget = QListWidget()
        self.list_widget.setMaximumHeight(200)
        self._populate_list(self.all_values)
        layout.addWidget(self.list_widget)

    def _build_number_widgets(self, layout: QVBoxLayout):
        # Operator
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Condition:"))
        self.op_combo = QComboBox()
        self.op_combo.addItems(NUMBER_OPS)
        self.op_combo.currentTextChanged.connect(self._on_op_changed)
        row1.addWidget(self.op_combo)
        layout.addLayout(row1)

        # Value 1
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Value:"))
        self.val1_input = QLineEdit()
        self.val1_input.setPlaceholderText("e.g.  50")
        row2.addWidget(self.val1_input)
        layout.addLayout(row2)

        # Value 2  (only for "between")
        row3 = QHBoxLayout()
        self.val2_label = QLabel("And:")
        self.val2_input = QLineEdit()
        self.val2_input.setPlaceholderText("e.g.  100")
        row3.addWidget(self.val2_label)
        row3.addWidget(self.val2_input)
        layout.addLayout(row3)

        layout.addStretch()

        # Pre-fill if an existing number filter is active
        nf = self.current_filter
        if nf and nf.get('type') == 'number':
            if nf['op'] in NUMBER_OPS:
                self.op_combo.setCurrentIndex(NUMBER_OPS.index(nf['op']))
            if nf['op'] == 'between':
                self.val1_input.setText(str(nf.get('val1', '')))
                self.val2_input.setText(str(nf.get('val2', '')))
            else:
                self.val1_input.setText(str(nf.get('val', '')))

        # Trigger initial visibility of val2 row
        self._on_op_changed(self.op_combo.currentText())

    # -----------------------------------------------------------------------
    # Helpers
    # -----------------------------------------------------------------------

    def _populate_list(self, values: list):
        try:
            self.list_widget.itemChanged.disconnect(self._on_item_changed)
        except TypeError:
            pass
        self.list_widget.clear()
        for val in values:
            label = val if val != "" else "(Blanks)"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, val)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(
                Qt.Checked if val in self.current_selection else Qt.Unchecked
            )
            self.list_widget.addItem(item)
        self.list_widget.itemChanged.connect(self._on_item_changed)

    def _refresh_select_all(self):
        self.select_all_cb.blockSignals(True)
        n_sel = sum(1 for v in self.all_values if v in self.current_selection)
        n_all = len(self.all_values)
        if n_sel == n_all:
            self.select_all_cb.setCheckState(Qt.Checked)
        elif n_sel == 0:
            self.select_all_cb.setCheckState(Qt.Unchecked)
        else:
            self.select_all_cb.setCheckState(Qt.PartiallyChecked)
        self.select_all_cb.blockSignals(False)

    # -----------------------------------------------------------------------
    # Slots
    # -----------------------------------------------------------------------

    def _on_search_changed(self, text: str):
        filtered = [v for v in self.all_values if text.lower() in v.lower()]
        self._populate_list(filtered)

    def _on_select_all_changed(self, state: int):
        checked = (state == Qt.Checked)
        try:
            self.list_widget.itemChanged.disconnect(self._on_item_changed)
        except TypeError:
            pass
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            val = item.data(Qt.UserRole)
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            if checked:
                self.current_selection.add(val)
            else:
                self.current_selection.discard(val)
        self.list_widget.itemChanged.connect(self._on_item_changed)

    def _on_item_changed(self, item: QListWidgetItem):
        val = item.data(Qt.UserRole)
        if item.checkState() == Qt.Checked:
            self.current_selection.add(val)
        else:
            self.current_selection.discard(val)
        self._refresh_select_all()

    def _on_op_changed(self, op: str):
        is_between = (op == 'between')
        self.val2_label.setVisible(is_between)
        self.val2_input.setVisible(is_between)

    # -----------------------------------------------------------------------
    # Apply
    # -----------------------------------------------------------------------

    def _apply(self):
        if self._tabs is not None and self._tabs.currentIndex() == 1:
            self._apply_number_filter()
        else:
            self._apply_value_filter()

    def _apply_value_filter(self):
        if self.current_selection == set(self.all_values):
            self.filter_changed.emit(self.col_index, None)
        else:
            self.filter_changed.emit(self.col_index, {
                'type': 'values',
                'values': set(self.current_selection),
            })
        self.close()

    def _apply_number_filter(self):
        op       = self.op_combo.currentText()
        val1_str = self.val1_input.text().strip()

        if not val1_str:                     # empty input → clear filter
            self.filter_changed.emit(self.col_index, None)
            self.close()
            return

        val1 = _try_float(val1_str)
        if val1 is None:
            QMessageBox.warning(self, "Invalid value",
                                f"'{val1_str}' is not a valid number.")
            return

        if op == 'between':
            val2_str = self.val2_input.text().strip()
            val2 = _try_float(val2_str)
            if val2 is None:
                QMessageBox.warning(self, "Invalid value",
                                    f"'{val2_str}' is not a valid number.")
                return
            spec = {
                'type': 'number', 'op': 'between',
                'val1': min(val1, val2), 'val2': max(val1, val2),
            }
        else:
            spec = {'type': 'number', 'op': op, 'val': val1}

        self.filter_changed.emit(self.col_index, spec)
        self.close()


# ---------------------------------------------------------------------------
# Custom header view – sort & filter icons, no label-click sort
# ---------------------------------------------------------------------------

class SortFilterHeaderView(QHeaderView):

    sort_requested   = pyqtSignal(int, int)   # col, Qt.SortOrder
    filter_requested = pyqtSignal(int)        # col

    _ICON_SIZE = 16
    _ICON_GAP  = 6    # px gap between icons and from right edge
    _BTN_PAD   = 3    # extra padding around icon for the button background
    # Extra horizontal space reserved for the two icons inside each section
    _ICON_RESERVE = _ICON_SIZE * 2 + _ICON_GAP * 3 + 16  # + left text padding

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.setSectionsClickable(False)   # no label-click sort
        self.setSortIndicatorShown(False)
        self.setHighlightSections(False)

        self._sort_px   = load_pixmap("sort.ico",   self._ICON_SIZE, self._ICON_SIZE)
        self._filter_px = load_pixmap("filter.ico", self._ICON_SIZE, self._ICON_SIZE)

        self._sort_col:    int = -1
        self._sort_order:  int = Qt.AscendingOrder
        self._filtered_cols: set = set()

        # Hover tracking
        self._hover_col:   int = -1
        self._hover_which: str = ''   # 'sort' | 'filter' | ''
        self.viewport().setMouseTracking(True)

    # -- public API --

    def mark_filter_active(self, col: int, active: bool):
        if active:
            self._filtered_cols.add(col)
        else:
            self._filtered_cols.discard(col)
        self.viewport().update()

    def reset_state(self):
        self._sort_col = -1
        self._filtered_cols.clear()
        self.viewport().update()

    # -- geometry --

    def _section_rect(self, logical: int) -> QRect:
        x = self.sectionViewportPosition(logical)
        return QRect(x, 0, self.sectionSize(logical), self.height())

    def _icon_rects(self, sec: QRect):
        """Return (sort_rect, filter_rect) aligned to the right of the section."""
        sz, gap = self._ICON_SIZE, self._ICON_GAP
        cy = sec.center().y()
        filter_r = QRect(sec.right() - gap - sz,              cy - sz // 2, sz, sz)
        sort_r   = QRect(filter_r.left() - gap - sz,          cy - sz // 2, sz, sz)
        return sort_r, filter_r

    # -- painting --

    def paintSection(self, painter: QPainter, rect: QRect, logical: int):
        if not rect.isValid():
            return

        # ── 1. Let the base class draw background + left-aligned text ──
        # (TextAlignmentRole = AlignLeft is set in the model by _load_dataframe)
        painter.save()
        super().paintSection(painter, rect, logical)
        painter.restore()

        # ── 3. Icon buttons ──
        if rect.width() < self._ICON_RESERVE:
            return

        sort_r, filter_r = self._icon_rects(rect)
        pad = self._BTN_PAD

        sort_active    = (self._sort_col == logical)
        sort_hovered   = (self._hover_col == logical and self._hover_which == 'sort')
        filter_active  = (logical in self._filtered_cols)
        filter_hovered = (self._hover_col == logical and self._hover_which == 'filter')

        pal        = self.palette()
        btn_base   = pal.color(QPalette.Button)
        border_col = pal.color(QPalette.Dark)
        hi_col     = pal.color(QPalette.Highlight)

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        def _draw_btn(icon_r: QRect, hovered: bool, active: bool, use_accent: bool = False):
            if not (hovered or active):
                return
            r = QRectF(icon_r.adjusted(-pad, -pad, pad, pad))
            if hovered:
                bg = btn_base.darker(130)
            elif use_accent:
                bg = hi_col.lighter(165)
            else:
                bg = btn_base.darker(115)
            painter.setPen(border_col)
            painter.setBrush(bg)
            painter.drawRoundedRect(r, 4.0, 4.0)

        _draw_btn(sort_r,   sort_hovered,   sort_active,   use_accent=False)
        _draw_btn(filter_r, filter_hovered, filter_active, use_accent=True)

        painter.restore()

        painter.setOpacity(1.0 if (sort_active or sort_hovered) else 0.35)
        painter.drawPixmap(sort_r, self._sort_px)

        painter.setOpacity(1.0 if (filter_active or filter_hovered) else 0.35)
        painter.drawPixmap(filter_r, self._filter_px)

        painter.setOpacity(1.0)

    # -- interaction --

    def mouseMoveEvent(self, event):
        pos     = event.pos()
        logical = self.logicalIndexAt(pos)
        old = (self._hover_col, self._hover_which)

        if logical >= 0:
            sort_r, filter_r = self._icon_rects(self._section_rect(logical))
            if filter_r.contains(pos):
                self._hover_col, self._hover_which = logical, 'filter'
            elif sort_r.contains(pos):
                self._hover_col, self._hover_which = logical, 'sort'
            else:
                self._hover_col, self._hover_which = -1, ''
        else:
            self._hover_col, self._hover_which = -1, ''

        if old != (self._hover_col, self._hover_which):
            self.viewport().update()

        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        self._hover_col, self._hover_which = -1, ''
        self.viewport().update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            pos     = event.pos()
            logical = self.logicalIndexAt(pos)
            if logical >= 0:
                sort_r, filter_r = self._icon_rects(self._section_rect(logical))

                if filter_r.contains(pos):
                    self.filter_requested.emit(logical)
                    return

                if sort_r.contains(pos):
                    if self._sort_col == logical:
                        new_order = (Qt.DescendingOrder
                                     if self._sort_order == Qt.AscendingOrder
                                     else Qt.AscendingOrder)
                    else:
                        new_order = Qt.AscendingOrder
                    self._sort_col   = logical
                    self._sort_order = new_order
                    self.sort_requested.emit(logical, new_order)
                    self.viewport().update()
                    return

        super().mousePressEvent(event)   # resize / drag pass-through


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------

class TableViewerApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.df                = None
        self.current_file_path = None
        self._source_model     = None
        self._proxy_model      = None
        self._filter_popup     = None   # keep reference – prevents GC
        self._init_ui()

    def _init_ui(self):
        self.setGeometry(200, 150, 1100, 680)
        self.setWindowTitle("Table Viewer")
        self.setWindowIcon(load_icon("favicon.ico"))
        self.setAcceptDrops(True)

        self._header = SortFilterHeaderView()
        self._header.sort_requested.connect(self._on_sort_requested)
        self._header.filter_requested.connect(self._show_filter_popup)

        self.table_view = QTableView(self)
        self.table_view.setHorizontalHeader(self._header)
        self.table_view.horizontalHeader().setSectionsMovable(True)
        self.table_view.setSortingEnabled(False)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.setCentralWidget(self.table_view)

        self.status_label = QLabel("No file loaded")
        status_bar = QStatusBar(self)
        status_bar.addWidget(self.status_label)
        self.setStatusBar(status_bar)

        self._build_menu()

    def _build_menu(self):
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("File")

        open_action = QAction(load_icon("search.ico"), "Open…", self)
        open_action.setShortcut(QKeySequence("Ctrl+O"))
        open_action.triggered.connect(self.show_open_dialog)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        save_excel = QAction("Save as Excel (.xlsx)", self)
        save_excel.setShortcut(QKeySequence("Ctrl+S"))
        save_excel.triggered.connect(self.save_as_excel)
        file_menu.addAction(save_excel)

        save_csv = QAction("Save as CSV (.csv)", self)
        save_csv.setShortcut(QKeySequence("Ctrl+Shift+S"))
        save_csv.triggered.connect(self.save_as_csv)
        file_menu.addAction(save_csv)

        tools_menu = menu_bar.addMenu("Tools")
        reg_action = QAction(
            load_icon("apps.ico"),
            "Register File Associations (.xlsx .xls .csv)", self
        )
        reg_action.triggered.connect(self.register_file_associations)
        tools_menu.addAction(reg_action)

    # ------------------------------------------------------------------
    # Sort / Filter
    # ------------------------------------------------------------------

    def _on_sort_requested(self, col: int, order: int):
        if self._proxy_model:
            self._proxy_model.sort(col, order)

    def _show_filter_popup(self, col_index: int):
        if self._source_model is None:
            return

        # Collect every distinct value for this column from the source
        all_values: set = set()
        for row in range(self._source_model.rowCount()):
            idx = self._source_model.index(row, col_index)
            all_values.add(self._source_model.data(idx, Qt.DisplayRole) or "")

        numeric        = _column_is_numeric(list(all_values))
        current_filter = self._proxy_model.get_column_filter(col_index)

        # Position popup just below the clicked column header
        sec_x      = self._header.sectionViewportPosition(col_index)
        global_pos = self._header.mapToGlobal(QPoint(sec_x, self._header.height()))

        self._filter_popup = FilterPopup(
            col_index, list(all_values),
            current_filter, is_numeric=numeric
        )
        self._filter_popup.filter_changed.connect(self._on_filter_changed)
        self._filter_popup.move(global_pos)
        self._filter_popup.show()

    def _on_filter_changed(self, col: int, spec):
        self._proxy_model.set_column_filter(col, spec)
        self._header.mark_filter_active(col, spec is not None)
        self._update_status()

    # ------------------------------------------------------------------
    # File loading
    # ------------------------------------------------------------------

    def show_open_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "",
            "All Supported (*.xlsx *.xls *.csv);;"
            "Excel Files (*.xlsx *.xls);;"
            "CSV Files (*.csv);;"
            "All Files (*)"
        )
        if file_path:
            self.open_file(file_path)

    def open_file(self, file_path: str):
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext in ('.xlsx', '.xls'):
                df = pd.read_excel(file_path)
            elif ext == '.csv':
                df = self._read_csv_with_auto_encoding(file_path)
            else:
                QMessageBox.warning(
                    self, "Unsupported File",
                    f"Unsupported file type: {ext}\n\nSupported: .xlsx, .xls, .csv"
                )
                return
            self._load_dataframe(df, file_path)
        except Exception as e:
            QMessageBox.critical(
                self, "Error Opening File",
                f"Could not open file:\n{file_path}\n\n{e}"
            )

    def _read_csv_with_auto_encoding(self, file_path: str) -> pd.DataFrame:
        for enc in CSV_ENCODINGS:
            try:
                return pd.read_csv(file_path, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError(
            "Could not decode the CSV file.\n"
            f"Tried encodings: {', '.join(CSV_ENCODINGS)}"
        )

    def _load_dataframe(self, df: pd.DataFrame, file_path: str):
        self.df                = df
        self.current_file_path = file_path

        source = QStandardItemModel(len(df), len(df.columns))
        source.setHorizontalHeaderLabels(df.columns.astype(str).tolist())

        # Left-align all column headers so text never drifts into the icon zone
        left_vcenter = int(Qt.AlignLeft | Qt.AlignVCenter)
        for col in range(len(df.columns)):
            source.setHeaderData(col, Qt.Horizontal, left_vcenter, Qt.TextAlignmentRole)

        for row_idx, (_, row_data) in enumerate(df.iterrows()):
            for col_idx, value in enumerate(row_data):
                display = "" if pd.isna(value) else str(value)
                item = QStandardItem(display)
                item.setEditable(False)
                source.setItem(row_idx, col_idx, item)

        self._source_model = source
        self._proxy_model  = MultiColumnFilterProxyModel()
        self._proxy_model.setSourceModel(source)

        self.table_view.setModel(self._proxy_model)
        self.table_view.resizeColumnsToContents()
        self._ensure_header_column_widths()   # make room for icons

        self._header.reset_state()

        file_name = os.path.basename(file_path)
        self.setWindowTitle(f"Table Viewer — {file_name}")
        self._update_status()

    def _ensure_header_column_widths(self):
        """
        After resizeColumnsToContents(), verify each column is wide enough
        to display its header label *plus* the two icon buttons without overlap.
        """
        fm         = self._header.fontMetrics()
        reserve    = self._header._ICON_RESERVE
        # add section-internal horizontal padding (Qt typically uses 8 px each side)
        h_padding  = 16

        for col in range(self._source_model.columnCount()):
            label = self._source_model.headerData(col, Qt.Horizontal, Qt.DisplayRole) or ""
            min_w = fm.horizontalAdvance(str(label)) + reserve + h_padding
            if self.table_view.columnWidth(col) < min_w:
                self.table_view.setColumnWidth(col, min_w)

    def _update_status(self):
        if self.df is None:
            return
        file_name = os.path.basename(self.current_file_path)
        total   = len(self.df)
        visible = self._proxy_model.rowCount() if self._proxy_model else total
        cols    = len(self.df.columns)
        if visible == total:
            self.status_label.setText(
                f"{file_name}   |   {total:,} rows   |   {cols} columns"
            )
        else:
            self.status_label.setText(
                f"{file_name}   |   {visible:,} / {total:,} rows (filtered)   |   {cols} columns"
            )

    # ------------------------------------------------------------------
    # File saving
    # ------------------------------------------------------------------

    def save_as_excel(self):
        if self.df is None:
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save as Excel", self._default_save_name('.xlsx'),
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            try:
                self.df.to_excel(file_path, index=False)
                self.statusBar().showMessage("Saved successfully.", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error Saving File", str(e))

    def save_as_csv(self):
        if self.df is None:
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save as CSV", self._default_save_name('.csv'),
            "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            try:
                self.df.to_csv(file_path, index=False, encoding='utf-8-sig')
                self.statusBar().showMessage("Saved successfully.", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error Saving File", str(e))

    def _default_save_name(self, new_ext: str) -> str:
        if self.current_file_path:
            return os.path.splitext(self.current_file_path)[0] + new_ext
        return ""

    # ------------------------------------------------------------------
    # Drag & Drop
    # ------------------------------------------------------------------

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0].toLocalFile()
            if url.lower().endswith(SUPPORTED_EXTENSIONS):
                event.acceptProposedAction()

    def dropEvent(self, event):
        self.open_file(event.mimeData().urls()[0].toLocalFile())

    # ------------------------------------------------------------------
    # File association  (Windows Registry)
    # ------------------------------------------------------------------

    def register_file_associations(self):
        try:
            import winreg
        except ImportError:
            QMessageBox.warning(self, "Not Supported",
                                "File association registration is only supported on Windows.")
            return

        app_path = os.path.abspath(sys.argv[0])
        command  = (f'"{sys.executable}" "{app_path}" "%1"'
                    if app_path.endswith('.py')
                    else f'"{app_path}" "%1"')

        try:
            for ext in SUPPORTED_EXTENSIONS:
                prog_id = f"TableViewer.{ext[1:].upper()}"

                with winreg.CreateKey(
                    winreg.HKEY_CURRENT_USER,
                    rf"Software\Classes\{prog_id}\shell\open\command"
                ) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command)

                icon_src = app_path if not app_path.endswith('.py') else sys.executable
                with winreg.CreateKey(
                    winreg.HKEY_CURRENT_USER,
                    rf"Software\Classes\{prog_id}\DefaultIcon"
                ) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, f'"{icon_src}",0')

                with winreg.CreateKey(
                    winreg.HKEY_CURRENT_USER,
                    rf"Software\Classes\{ext}\OpenWithProgids"
                ) as k:
                    winreg.SetValueEx(k, prog_id, 0, winreg.REG_NONE, b"")

            QMessageBox.information(
                self, "File Associations Registered",
                "Done! .xlsx, .xls and .csv files are now associated with Table Viewer.\n\n"
                "To set as default:\n"
                "  Right-click any .xlsx/.xls/.csv file\n"
                "  → Open with → Choose another app\n"
                "  → Table Viewer → Always use this app"
            )
        except Exception as e:
            QMessageBox.critical(self, "Registration Failed",
                                 f"Could not register file associations:\n\n{e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TableViewerApp()
    window.show()

    if len(sys.argv) > 1:
        window.open_file(sys.argv[1])

    sys.exit(app.exec_())
