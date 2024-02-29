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
    QTableWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QDialog,
    QComboBox
)

from PySide6.QtGui import QPixmap, QFont

from ui.ui_form import Ui_MainWindow

import openpyxl

# excel_file = "../docs/Law Clients Excel Sheet Shared_MainV3.xlsm"


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # self.load_excel_data()
        self.ui.exitbutton.clicked.connect(self.closeApplication)
        self.ui.importbutton.clicked.connect(self.openFile)
        self.ui.newentrybutton.clicked.connect(self.ask_for_new_entry)
        # clear existing tabs
        self.ui.tabWidget.clear()

        # create a new tab for the logo
        tab = QWidget()
        tab.setObjectName("logoTab")
        self.ui.tabWidget.addTab(tab, "Start")

        # create a QLabel for the logo
        logoLabel = QLabel(tab)
        logoPixmap = QPixmap("bkp_logo.jpg")
        logoLabel.setPixmap(logoPixmap)
        textLabel = QLabel("BKP Solicitors Client Data", tab)
        font = QFont("Calibri", 72)
        textLabel.setFont(font)

        # create a layout for the tab and add the logoLabel
        hboxLayout = QHBoxLayout()
        hboxLayout.addWidget(logoLabel)
        hboxLayout.addWidget(textLabel)
        tab.setLayout(hboxLayout)

    def load_excel_data(self, excel_file):
        try:
            self.wb = openpyxl.load_workbook(excel_file)
        except:
            self.showAlarm("Format error", "File format is not supported!")
            return

        self.sheet_number = len(self.wb.sheetnames)

        # clear existing tabs
        self.ui.tabWidget.clear()

        # create tabs
        if self.sheet_number > 1:
            for _ in range(self.sheet_number):
                tab = QWidget()
                tab.setObjectName("tab")
                self.ui.tabWidget.addTab(tab, "")

        for sh_num in range(self.sheet_number):
            sheet_name = self.wb.sheetnames[sh_num]

            # Create a new QTableWidget for this tab
            self.tableWidget = QTableWidget()
            self.tableWidget.setRowCount(self.wb[sheet_name].max_row)
            self.tableWidget.setColumnCount(self.wb[sheet_name].max_column)
            # Enable sorting
            self.tableWidget.setSortingEnabled(True)

            # Add the Qself.tableWidget to a QHBoxLayout inside a QVBoxLayout
            hboxLayout = QHBoxLayout()
            hboxLayout.addWidget(self.tableWidget)
            vboxLayout = QVBoxLayout()
            vboxLayout.addLayout(hboxLayout)
            self.ui.tabWidget.widget(sh_num).setLayout(vboxLayout)

            # Change tabs name
            self.ui.tabWidget.setTabText(sh_num, self.wb.sheetnames[sh_num])

            # set value of table from excel file
            headers = None
            for i, row in enumerate(self.wb[sheet_name].iter_rows(values_only=True)):
                # if all entries in the row are None or empty, skip the row
                if all(
                    value is None
                    or value == "BKP Solicitors Client Data"
                    or str(value).strip() == ""
                    for value in row
                ):
                    continue  # skip this row
                if headers is None:  # if headers haven't been set yet
                    headers = [
                        str(value) for value in row
                    ]  # use this row as the headers
                    self.tableWidget.setHorizontalHeaderLabels(headers)  # set the headers
                    continue  # skip the rest of this iteration
                for j, value in enumerate(row):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(value)))

            # remove empty rows
            rows_to_remove = []
            for i in range(self.tableWidget.rowCount()):
                if all(
                    self.tableWidget.item(i, j) is None
                    or self.tableWidget.item(i, j).text() == ""
                    for j in range(self.tableWidget.columnCount())
                ):
                    rows_to_remove.append(i)
            for i in reversed(rows_to_remove):
                self.tableWidget.removeRow(i)

    def showAlarm(self, header, mes):
        QMessageBox.warning(self, header, mes)

    def openFile(self):
        try:
            if self.excel_file:
                self.tableWidget.cellChanged.disconnect(self.save_excel_data)
        except:
            pass

        filePath, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "Excel Files (*.xlsx *.xlsm)"
        )

        if filePath:
            self.excel_file = filePath
            self.load_excel_data(self.excel_file)
            self.tableWidget.cellChanged.connect(self.save_excel_data)

    def save_excel_data(self, row, column):
        if not self.excel_file:
            self.showAlarm("Error", "file does not exist !")

        sheet = self.wb.active
        sheet.cell(
            row=row + 1,
            column=column + 1,
            value=self.tableWidget.item(row, column).text(),
        )
        self.wb.save(self.excel_file)

    
    def ask_for_new_entry(self):
        # Check if wb has been defined
        if not hasattr(self, 'wb'):
            self.showAlarm("Error", "Please load an Excel file first!")
            return

        # open a popup window for new entry
        dialog = NewEntryDialog(self.wb, self)
        if dialog.exec():
            selected_sheet = dialog.comboBox.currentText()
            new_entry = {column: lineEdit.text() for column, lineEdit in zip([cell.value for cell in self.wb[selected_sheet][1]], dialog.lineEdits)}
            self.wb[selected_sheet].append(list(new_entry.values()))
            self.wb.save(self.excel_file)

            # Get the tableWidget of the currently selected tab
            current_tab = self.ui.tabWidget.currentWidget()
            tableWidget = current_tab.findChild(QTableWidget)

            # update the table without reloading the file
            tableWidget.setRowCount(tableWidget.rowCount() + 1)
            for i, value in enumerate(new_entry.values()):
                tableWidget.setItem(tableWidget.rowCount() - 1, i, QTableWidgetItem(str(value)))

    def closeApplication(self):
        self.close()

class NewEntryDialog(QDialog):
    def __init__(self, wb, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Entry")

        self.layout = QVBoxLayout(self)

        self.comboBox = QComboBox(self)
        self.comboBox.addItems(wb.sheetnames)
        self.comboBox.setCurrentText(parent.ui.tabWidget.tabText(parent.ui.tabWidget.currentIndex()))
        
        self.layout.addWidget(self.comboBox)

        self.lineEditsLayout = QVBoxLayout()
        self.layout.addLayout(self.lineEditsLayout)

        self.lineEdits = []
        self.updateLineEdits([cell.value for cell in wb[self.comboBox.currentText()][1]])

        self.comboBox.currentIndexChanged.connect(lambda: self.updateLineEdits([cell.value for cell in wb[self.comboBox.currentText()][1]]))

        self.button = QPushButton("Submit", self)
        self.button.clicked.connect(self.accept)
        self.layout.addWidget(self.button)

    def updateLineEdits(self, columns):
        # Remove existing QLineEdit widgets
        for lineEdit in self.lineEdits:
            self.lineEditsLayout.removeWidget(lineEdit)
            lineEdit.deleteLater()
        self.lineEdits.clear()

        # Add new QLineEdit widgets
        for column in columns:
            lineEdit = QLineEdit(self)
            lineEdit.setPlaceholderText(column)
            self.lineEditsLayout.addWidget(lineEdit)
            self.lineEdits.append(lineEdit)

def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
