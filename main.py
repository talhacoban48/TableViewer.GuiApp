import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QFileDialog, QAction, QStatusBar, QLabel, QMessageBox
)
from PyQt5.QtGui import QKeySequence


SUPPORTED_EXTENSIONS = ('.xlsx', '.xls', '.csv')
CSV_ENCODINGS = ['utf-8-sig', 'utf-8', 'cp1254', 'latin-1', 'iso-8859-9']


class TableViewerApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.df = None
        self.current_file_path = None
        self._init_ui()

    def _init_ui(self):
        self.setGeometry(200, 150, 1000, 650)
        self.setWindowTitle("Table Viewer")
        self.setAcceptDrops(True)

        self.table = QTableWidget(self)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.setCentralWidget(self.table)

        self.status_label = QLabel("No file loaded")
        status_bar = QStatusBar(self)
        status_bar.addWidget(self.status_label)
        self.setStatusBar(status_bar)

        self._build_menu()

    def _build_menu(self):
        menu_bar = self.menuBar()

        # --- File Menu ---
        file_menu = menu_bar.addMenu("File")

        open_action = QAction("Open...", self)
        open_action.setShortcut(QKeySequence("Ctrl+O"))
        open_action.triggered.connect(self.show_open_dialog)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        save_excel_action = QAction("Save as Excel (.xlsx)", self)
        save_excel_action.setShortcut(QKeySequence("Ctrl+S"))
        save_excel_action.triggered.connect(self.save_as_excel)
        file_menu.addAction(save_excel_action)

        save_csv_action = QAction("Save as CSV (.csv)", self)
        save_csv_action.setShortcut(QKeySequence("Ctrl+Shift+S"))
        save_csv_action.triggered.connect(self.save_as_csv)
        file_menu.addAction(save_csv_action)

        # --- Tools Menu ---
        tools_menu = menu_bar.addMenu("Tools")

        register_action = QAction("Register File Associations (.xlsx .xls .csv)", self)
        register_action.triggered.connect(self.register_file_associations)
        tools_menu.addAction(register_action)

    # ------------------------------------------------------------------
    # File Loading
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

    def open_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext in ('.xlsx', '.xls'):
                df = pd.read_excel(file_path)
            elif ext == '.csv':
                df = self._read_csv_with_auto_encoding(file_path)
            else:
                QMessageBox.warning(self, "Unsupported File",
                                    f"Unsupported file type: {ext}\n\nSupported: .xlsx, .xls, .csv")
                return
            self._load_dataframe(df, file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error Opening File",
                                 f"Could not open file:\n{file_path}\n\n{e}")

    def _read_csv_with_auto_encoding(self, file_path):
        for encoding in CSV_ENCODINGS:
            try:
                return pd.read_csv(file_path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError(
            "Could not decode the CSV file.\n"
            f"Tried encodings: {', '.join(CSV_ENCODINGS)}"
        )

    def _load_dataframe(self, df, file_path):
        self.df = df
        self.current_file_path = file_path

        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str).tolist())

        for row_index, row_data in df.iterrows():
            self.table.insertRow(self.table.rowCount())
            for col_index, value in enumerate(row_data):
                display = "" if pd.isna(value) else str(value)
                self.table.setItem(row_index, col_index, QTableWidgetItem(display))

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

        file_name = os.path.basename(file_path)
        self.setWindowTitle(f"Table Viewer — {file_name}")
        self.status_label.setText(
            f"{file_name}   |   {len(df):,} rows   |   {len(df.columns)} columns"
        )

    # ------------------------------------------------------------------
    # File Saving
    # ------------------------------------------------------------------

    def save_as_excel(self):
        if self.df is None:
            return
        default_name = self._default_save_name('.xlsx')
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save as Excel", default_name, "Excel Files (*.xlsx);;All Files (*)"
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
        default_name = self._default_save_name('.csv')
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save as CSV", default_name, "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            try:
                self.df.to_csv(file_path, index=False, encoding='utf-8-sig')
                self.statusBar().showMessage("Saved successfully.", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error Saving File", str(e))

    def _default_save_name(self, new_ext):
        if self.current_file_path:
            base = os.path.splitext(self.current_file_path)[0]
            return base + new_ext
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
        file_path = event.mimeData().urls()[0].toLocalFile()
        self.open_file(file_path)

    # ------------------------------------------------------------------
    # File Association (Windows Registry)
    # ------------------------------------------------------------------

    def register_file_associations(self):
        try:
            import winreg
        except ImportError:
            QMessageBox.warning(self, "Not Supported",
                                "File association registration is only supported on Windows.")
            return

        app_path = os.path.abspath(sys.argv[0])
        if app_path.endswith('.py'):
            command = f'"{sys.executable}" "{app_path}" "%1"'
        else:
            command = f'"{app_path}" "%1"'

        try:
            for ext in SUPPORTED_EXTENSIONS:
                prog_id = f"TableViewer.{ext[1:].upper()}"

                # Register the ProgID with its open command
                key_path = rf"Software\Classes\{prog_id}\shell\open\command"
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)

                # Register icon (uses the Python interpreter icon as fallback)
                icon_key_path = rf"Software\Classes\{prog_id}\DefaultIcon"
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, icon_key_path) as key:
                    icon_src = app_path if not app_path.endswith('.py') else sys.executable
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{icon_src}",0')

                # Associate the extension with this ProgID
                ext_key_path = rf"Software\Classes\{ext}\OpenWithProgids"
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, ext_key_path) as key:
                    winreg.SetValueEx(key, prog_id, 0, winreg.REG_NONE, b"")

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

    # Opened via double-click or file association
    if len(sys.argv) > 1:
        window.open_file(sys.argv[1])

    sys.exit(app.exec_())
