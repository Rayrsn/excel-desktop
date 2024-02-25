# This Python file uses the following encoding: utf-8

# Important:
# You need to run the following command to generate the ui_form.py file
#     pyside6-uic form.ui -o ui_form.py, or
#     pyside2-uic form.ui -o ui_form.py

import sys

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTableWidgetItem,
    QWidget,
)

from ui.ui_form import Ui_MainWindow

import openpyxl

# excel_file = "../docs/Law Clients Excel Sheet Shared_MainV3.xlsm"


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.load_excel_data()
        self.ui.tableWidget.cellChanged.connect(self.save_excel_data)

    def load_excel_data(self):
        self.wb = openpyxl.load_workbook(exel_file)

        self.sheet_number = len(self.wb.sheetnames)
        # create tabs
        if self.sheet_number > 1:
            for _ in range(self.sheet_number):
                self.tab = QWidget()
                self.tab.setObjectName("tab")
                self.ui.tabWidget.addTab(self.tab, "")

        for sh_num in range(self.sheet_number):
            sheet_name = self.wb.sheetnames[sh_num]
            self.ui.tableWidget.setRowCount(self.wb[sheet_name].max_row)
            self.ui.tableWidget.setColumnCount(self.wb[sheet_name].max_column)

            # Change tabs name
            self.ui.tabWidget.setTabText(sh_num, self.wb.sheetnames[sh_num])

            # set value of table from exel file
            for i, row in enumerate(self.wb[sheet_name].iter_rows(values_only=True)):
                for j, value in enumerate(row):
                    self.ui.tableWidget.setItem(i, j, QTableWidgetItem(str(value)))

    def save_excel_data(self, row, column):
        sheet = self.wb.active
        sheet.cell(
            row=row + 1,
            column=column + 1,
            value=self.ui.tableWidget.item(row, column).text(),
        )
        self.wb.save(exel_file)


def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
