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
    QFileDialog,
    QMessageBox,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget
)

from ui.ui_form import Ui_MainWindow

import openpyxl

# excel_file = "../docs/Law Clients Excel Sheet Shared_MainV3.xlsm"


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # self.load_excel_data()
        self.ui.button21.clicked.connect(self.closeApplication)
        self.ui.button1.clicked.connect(self.openFile)
        # clear existing tabs
        self.ui.tabWidget.clear()

    def load_excel_data(self, exel_file):
        try:
            self.wb = openpyxl.load_workbook(exel_file)
        except:
            self.showAlarm("Format error", "File format is not supported!")
            return

        self.sheet_number = len(self.wb.sheetnames)

        # create tabs
        if self.sheet_number > 1:
            for _ in range(self.sheet_number):
                tab = QWidget()
                tab.setObjectName("tab")
                self.ui.tabWidget.addTab(tab, "")

        for sh_num in range(self.sheet_number):
            sheet_name = self.wb.sheetnames[sh_num]

            # Create a new QTableWidget for this tab
            tableWidget = QTableWidget()
            tableWidget.setRowCount(self.wb[sheet_name].max_row)
            tableWidget.setColumnCount(self.wb[sheet_name].max_column)

            # Add the QTableWidget to a QHBoxLayout inside a QVBoxLayout
            hboxLayout = QHBoxLayout()
            hboxLayout.addWidget(tableWidget)
            vboxLayout = QVBoxLayout()
            vboxLayout.addLayout(hboxLayout)
            self.ui.tabWidget.widget(sh_num).setLayout(vboxLayout)

            # Change tabs name
            self.ui.tabWidget.setTabText(sh_num, self.wb.sheetnames[sh_num])

            # set value of table from exel file
            for i, row in enumerate(self.wb[sheet_name].iter_rows(values_only=True)):
                for j, value in enumerate(row):
                    tableWidget.setItem(i, j, QTableWidgetItem(str(value)))

    def showAlarm(self, header, mes):
        QMessageBox.warning(self, header, mes)

    def openFile(self):
        try:
            if self.exel_file:
                self.ui.tableWidget.cellChanged.disconnect(self.save_excel_data)
        except:
            pass

        filePath, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "All Files (*)"
        )

        if filePath:
            self.exel_file = filePath
            self.load_excel_data(self.exel_file)
            self.ui.tableWidget.cellChanged.connect(self.save_excel_data)

    def save_excel_data(self, row, column):
        if not self.exel_file:
            self.showAlarm("Error", "file does not exist !")

        sheet = self.wb.active
        sheet.cell(
            row=row + 1,
            column=column + 1,
            value=self.ui.tableWidget.item(row, column).text(),
        )
        self.wb.save(self.exel_file)

    def closeApplication(self):
        self.close()

def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
