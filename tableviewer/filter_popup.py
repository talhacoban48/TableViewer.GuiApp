from __future__ import annotations

from PyQt5.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QCheckBox, QListWidget, QListWidgetItem, QPushButton,
    QTabWidget, QWidget, QComboBox, QMessageBox,
)
from PyQt5.QtCore import Qt, pyqtSignal

from .constants import NUMBER_OPS
from .utils import load_icon, load_pixmap, _try_float


class FilterPopup(QFrame):

    filter_changed = pyqtSignal(int, object)   # col_index, spec | None

    def __init__(self, col_index: int, all_values: list,
                 current_filter, is_numeric: bool, parent=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.col_index = col_index
        self.is_numeric = is_numeric
        self.current_filter = current_filter

        self.all_values = sorted(all_values, key=lambda v: (v == "", v.lower()))

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

            val_widget = QWidget()
            val_layout = QVBoxLayout(val_widget)
            val_layout.setContentsMargins(4, 4, 4, 4)
            val_layout.setSpacing(4)
            self._build_values_widgets(val_layout)
            self._tabs.addTab(val_widget, load_icon("filter.ico"), "Values")

            num_widget = QWidget()
            num_layout = QVBoxLayout(num_widget)
            num_layout.setContentsMargins(4, 8, 4, 4)
            num_layout.setSpacing(8)
            self._build_number_widgets(num_layout)
            self._tabs.addTab(num_widget, load_icon("sort.ico"), "Number Filter")

            if self.current_filter and self.current_filter.get('type') == 'number':
                self._tabs.setCurrentIndex(1)

            outer.addWidget(self._tabs)
        else:
            self._tabs = None
            self._build_values_widgets(outer)

        btn_row = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self._apply)
        cancel_btn = QPushButton(load_icon("cancel.ico"), "Cancel")
        cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        outer.addLayout(btn_row)

    def _build_values_widgets(self, layout: QVBoxLayout):
        search_row = QHBoxLayout()
        lbl = QLabel()
        lbl.setPixmap(load_pixmap("search.ico"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search values…")
        self.search_input.textChanged.connect(self._on_search_changed)
        search_row.addWidget(lbl)
        search_row.addWidget(self.search_input)
        layout.addLayout(search_row)

        self.select_all_cb = QCheckBox("(Select All)")
        self.select_all_cb.setTristate(True)
        self._refresh_select_all()
        self.select_all_cb.stateChanged.connect(self._on_select_all_changed)
        layout.addWidget(self.select_all_cb)

        self.list_widget = QListWidget()
        self.list_widget.setMaximumHeight(200)
        self._populate_list(self.all_values)
        layout.addWidget(self.list_widget)

    def _build_number_widgets(self, layout: QVBoxLayout):
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Condition:"))
        self.op_combo = QComboBox()
        self.op_combo.addItems(NUMBER_OPS)
        self.op_combo.currentTextChanged.connect(self._on_op_changed)
        row1.addWidget(self.op_combo)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Value:"))
        self.val1_input = QLineEdit()
        self.val1_input.setPlaceholderText("e.g.  50")
        row2.addWidget(self.val1_input)
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.val2_label = QLabel("And:")
        self.val2_input = QLineEdit()
        self.val2_input.setPlaceholderText("e.g.  100")
        row3.addWidget(self.val2_label)
        row3.addWidget(self.val2_input)
        layout.addLayout(row3)

        layout.addStretch()

        nf = self.current_filter
        if nf and nf.get('type') == 'number':
            if nf['op'] in NUMBER_OPS:
                self.op_combo.setCurrentIndex(NUMBER_OPS.index(nf['op']))
            if nf['op'] == 'between':
                self.val1_input.setText(str(nf.get('val1', '')))
                self.val2_input.setText(str(nf.get('val2', '')))
            else:
                self.val1_input.setText(str(nf.get('val', '')))

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

        if not val1_str:
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
