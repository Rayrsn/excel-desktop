# This Python file uses the following encoding: utf-8
import sys

from PySide6.QtWidgets import QApplication, QMainWindow

# Important:
# You need to run the following command to generate the ui_form.py file
#     pyside6-uic form.ui -o ui_form.py, or
#     pyside2-uic form.ui -o ui_form.py

from .ui_form import Ui_MainWindow
# from ui_form import Ui_MainWindow

from PySide6.QtWidgets import QTableWidgetItem
import openpyxl

def load_excel_data(self):
    wb = openpyxl.load_workbook("list.xlsx")
    sheet = wb.active

    self.ui.tableWidget.setRowCount(sheet.max_row)
    self.ui.tableWidget.setColumnCount(sheet.max_column)

    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        for j, value in enumerate(row):
            self.ui.tableWidget.setItem(i, j, QTableWidgetItem(str(value)))



from PySide6.QtWidgets import QMainWindow, QTableWidgetItem
import openpyxl

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.load_excel_data()
        self.ui.tableWidget.cellChanged.connect(self.save_excel_data)

    def load_excel_data(self):
        self.wb = openpyxl.load_workbook("list.xlsx")
        sheet = self.wb.active

        self.ui.tableWidget.setRowCount(sheet.max_row)
        self.ui.tableWidget.setColumnCount(sheet.max_column)

        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            for j, value in enumerate(row):
                self.ui.tableWidget.setItem(i, j, QTableWidgetItem(str(value)))

    def save_excel_data(self, row, column):
        sheet = self.wb.active
        sheet.cell(row=row+1, column=column+1, value=self.ui.tableWidget.item(row, column).text())
        self.wb.save("list.xlsx")


def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.showMaximized()
    sys.exit(app.exec())

if __name__ == "__main__":
    run()