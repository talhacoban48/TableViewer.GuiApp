import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget,QTableWidgetItem, QFileDialog, QAction, QMenuBar



class CSVTableApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):

        self.setGeometry(550, 250, 500, 300)
        self.setWindowTitle("CSV Table")

        self.control = False
        self.table = QTableWidget(self)
        self.setCentralWidget(self.table)

        self.menu_bar = QMenuBar(self)
        self.setMenuBar(self.menu_bar)

        self.file_menu = self.menu_bar.addMenu("File")
        self.import_menu = self.menu_bar.addMenu("import")

        self.save_excel_action = QAction("Save as Excel", self)
        self.save_excel_action.triggered.connect(self.save_excel)
        self.file_menu.addAction(self.save_excel_action)
        
        self.save_csv_action = QAction("Save as CSV", self)
        self.save_csv_action.triggered.connect(self.save_csv)
        self.file_menu.addAction(self.save_csv_action)
        
        self.open_excel_action = QAction("import Excel", self)
        self.open_excel_action.triggered.connect(self.open_excel)
        self.import_menu.addAction(self.open_excel_action)
        
        self.open_csv_action = QAction("import CSV", self)
        self.open_csv_action.triggered.connect(self.open_csv)
        self.import_menu.addAction(self.open_csv_action)
        

    def open_table(self, df):
        
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

        self.table.setColumnCount(len(df.columns))

        self.table.setHorizontalHeaderLabels(df.columns)

        for i, row in df.iterrows():
            self.table.insertRow(self.table.rowCount())
            for j, cell in enumerate(row):
                self.table.setItem(self.table.rowCount()-1, j, QTableWidgetItem(str(cell)))

        self.df = df
        self.control = True


    def open_excel(self):
        
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_name:
            df = pd.read_excel(file_name)
            self.open_table(df)


    def open_csv(self):

        file_name, _ = QFileDialog.getOpenFileName(self, "Open CSV", "", "CSV Files (*.csv);;All Files (*)")

        if file_name:
            df = pd.read_csv(file_name)
            self.open_table(df)

        
    def save_excel(self):

        if self.control:
            file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx);;All Files (*)")
            if file_name:
                
                self.df.to_excel(file_name, index=False)
            
            
    def save_csv(self):

        if self.control:
            file_name, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv);;All Files (*)")
            if file_name:
                
                self.df.to_csv(file_name, index=False)
            
            
            


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CSVTableApp()
    window.show()
    sys.exit(app.exec_())






        
