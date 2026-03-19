from __future__ import annotations

from PyQt5.QtWidgets import QWidget
from PyQt5.QtGui import QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QRect


class MarchingAntsOverlay(QWidget):
    """Transparent widget drawn on top of the table viewport to show the
    animated dashed border around the copied / cut region."""

    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self._rect   = QRect()
        self._offset = 0.0
        self.hide()

    def set_rect(self, rect: QRect, offset: float):
        self._rect   = rect
        self._offset = offset
        self.resize(self.parent().size())
        self.raise_()
        self.show()
        self.update()

    def clear_rect(self):
        self._rect = QRect()
        self.hide()

    def paintEvent(self, _):
        if not self._rect.isValid():
            return
        p = QPainter(self)
        r = self._rect.adjusted(1, 1, -1, -1)
        for color, off_extra in (("#000000", 0.0), ("#ffffff", 4.0)):
            pen = QPen(QColor(color), 1.5)
            pen.setStyle(Qt.CustomDashLine)
            pen.setDashPattern([4.0, 4.0])
            pen.setDashOffset(self._offset + off_extra)
            p.setPen(pen)
            p.setBrush(Qt.NoBrush)
            p.drawRect(r)
