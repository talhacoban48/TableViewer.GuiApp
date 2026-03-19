from __future__ import annotations

from PyQt5.QtWidgets import QHeaderView
from PyQt5.QtGui import QPainter, QPalette, QColor
from PyQt5.QtCore import Qt, pyqtSignal, QRect, QRectF

from .utils import load_pixmap


class SortFilterHeaderView(QHeaderView):

    sort_requested         = pyqtSignal(int, int)   # col, Qt.SortOrder
    filter_requested       = pyqtSignal(int)        # col
    clear_filter_requested = pyqtSignal(int)        # col

    _ICON_SIZE     = 16
    _ICON_GAP      = 6    # px gap between icons and from right edge
    _BTN_PAD       = 3    # extra padding around icon for the button background
    _ICON_RESERVE  = _ICON_SIZE * 3 + _ICON_GAP * 4 + 16
    _HEADER_HEIGHT = 36   # px

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.setSectionsClickable(True)
        self.setSortIndicatorShown(False)
        self.setHighlightSections(False)
        self.setMinimumHeight(self._HEADER_HEIGHT)
        self.setMaximumHeight(self._HEADER_HEIGHT)

        self._sort_px   = load_pixmap("sort.ico",   self._ICON_SIZE, self._ICON_SIZE)
        self._filter_px = load_pixmap("filter.ico", self._ICON_SIZE, self._ICON_SIZE)
        self._clear_px  = load_pixmap("cancel.ico", self._ICON_SIZE, self._ICON_SIZE)

        self._sort_col:      int = -1
        self._sort_order:    int = Qt.AscendingOrder
        self._filtered_cols: set = set()

        self._hover_col:   int = -1
        self._hover_which: str = ''   # 'sort' | 'filter' | 'clear' | ''
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
        """Return (sort_rect, filter_rect, clear_rect) right-aligned in the section."""
        sz, gap = self._ICON_SIZE, self._ICON_GAP
        cy = sec.center().y()
        clear_r  = QRect(sec.right()     - gap - sz, cy - sz // 2, sz, sz)
        filter_r = QRect(clear_r.left()  - gap - sz, cy - sz // 2, sz, sz)
        sort_r   = QRect(filter_r.left() - gap - sz, cy - sz // 2, sz, sz)
        return sort_r, filter_r, clear_r

    # -- painting --

    def paintSection(self, painter: QPainter, rect: QRect, logical: int):
        if not rect.isValid():
            return

        painter.save()
        super().paintSection(painter, rect, logical)
        painter.restore()

        if rect.width() < self._ICON_RESERVE:
            return

        sort_r, filter_r, clear_r = self._icon_rects(rect)
        pad = self._BTN_PAD

        sort_active    = (self._sort_col == logical)
        sort_hovered   = (self._hover_col == logical and self._hover_which == 'sort')
        filter_active  = (logical in self._filtered_cols)
        filter_hovered = (self._hover_col == logical and self._hover_which == 'filter')
        clear_hovered  = (self._hover_col == logical and self._hover_which == 'clear')

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
            bg = (btn_base.darker(130) if hovered
                  else hi_col.lighter(165) if use_accent
                  else btn_base.darker(115))
            painter.setPen(border_col)
            painter.setBrush(bg)
            painter.drawRoundedRect(r, 4.0, 4.0)

        _draw_btn(sort_r,   sort_hovered,   sort_active,   use_accent=False)
        _draw_btn(filter_r, filter_hovered, filter_active, use_accent=True)
        if filter_active:
            _draw_btn(clear_r, clear_hovered, active=True, use_accent=False)

        painter.restore()

        painter.setOpacity(1.0)
        painter.drawPixmap(sort_r, self._sort_px)

        painter.setOpacity(1.0 if (filter_active or filter_hovered) else 0.75)
        painter.drawPixmap(filter_r, self._filter_px)

        if filter_active:
            painter.setOpacity(1.0)
            painter.drawPixmap(clear_r, self._clear_px)

        painter.setOpacity(1.0)

    # -- interaction --

    def mouseMoveEvent(self, event):
        pos     = event.pos()
        logical = self.logicalIndexAt(pos)
        old = (self._hover_col, self._hover_which)

        if logical >= 0:
            sort_r, filter_r, clear_r = self._icon_rects(self._section_rect(logical))
            if logical in self._filtered_cols and clear_r.contains(pos):
                self._hover_col, self._hover_which = logical, 'clear'
            elif filter_r.contains(pos):
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
                sort_r, filter_r, clear_r = self._icon_rects(self._section_rect(logical))

                if logical in self._filtered_cols and clear_r.contains(pos):
                    self.clear_filter_requested.emit(logical)
                    return

                if filter_r.contains(pos):
                    self.filter_requested.emit(logical)
                    return

                if sort_r.contains(pos):
                    new_order = (Qt.DescendingOrder
                                 if self._sort_col == logical
                                 and self._sort_order == Qt.AscendingOrder
                                 else Qt.AscendingOrder)
                    self._sort_col   = logical
                    self._sort_order = new_order
                    self.sort_requested.emit(logical, new_order)
                    self.viewport().update()
                    return

        super().mousePressEvent(event)
