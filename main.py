from __future__ import annotations

import sys
from PyQt5.QtWidgets import QApplication

from tableviewer.app import TableViewerApp


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TableViewerApp()
    window.show()

    if len(sys.argv) > 1:
        window.open_file(sys.argv[1])

    sys.exit(app.exec_())
