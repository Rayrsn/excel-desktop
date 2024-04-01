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
    QComboBox,
    QGridLayout,
    QSpacerItem,
    QProgressBar,
)

from PySide6.QtGui import QPixmap, QFont
from PySide6.QtCore import Qt, QThread, Signal

from ui.ui_form import Ui_MainWindow
from update_doc_file import gen_docs
import utils.pw_query as pw_query

import openpyxl
from utils.btn import (
    generate_monthly_cases_report,
    generate_weekly_cases_report,
    generate_legal_aid_report,
    generate_bail_refused_report,
    generate_empty_counsel_report,
    generate_non_zero_balance_report,
    generate_stage_reports,
)

excel_file = "../docs/Law Clients Excel Sheet Shared_MainV3.xlsm"


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.QueryWorker = None
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # self.loadExcelData()
        self.ui.exitbutton.clicked.connect(self.closeApplication)
        self.ui.importbutton.clicked.connect(self.openFile)
        self.ui.exportbutton.clicked.connect(self.genDocsBtn)
        self.ui.newentrybutton.clicked.connect(self.showEnewEntryDialog)
        self.ui.operationsbutton.clicked.connect(self.showOprationDialog)

        self.ui.exitbutton.setCursor(Qt.PointingHandCursor)
        self.ui.importbutton.setCursor(Qt.PointingHandCursor)
        self.ui.exportbutton.setCursor(Qt.PointingHandCursor)
        self.ui.newentrybutton.setCursor(Qt.PointingHandCursor)
        self.ui.operationsbutton.setCursor(Qt.PointingHandCursor)

        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #333;
            }

            QTableWidget {
                gridline-color: #999;
                font-size: 14px;
            }

            QTableWidget QHeaderView::section {
                background-color: #666;
                color: #fff;
                padding: 5px;
                border: 1px solid #999;
            }

            QTableWidget QTableCornerButton::section {
                background-color: #666;
                border: 1px solid #999;
            }

            QWidget > QPushButton {
                background-color: #007bff; /* Green */
                border: none;
                color: white;
                padding: 15px 32px;
                text-align: center;
                text-decoration: none;
                font-size: 16px;
                margin: 4px 2px;
            }

            QWidget > QPushButton:hover {
                background-color: #3094fd;
            }
        """
        )

        # clear existing tabs
        self.ui.tabWidget.clear()

        # create a new tab for the logo
        tab = QWidget()
        tab.setObjectName("logoTab")
        self.ui.tabWidget.addTab(tab, "Start")

        # create a QLabel for the logo
        logoLabel = QLabel(tab)
        if hasattr(sys, "_MEIPASS"):
            logoPixmap = QPixmap(sys._MEIPASS + "/bkp_logo.jpg")
        else:
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

    def runQueryWithLoding(self):
        loading_dialog = LoadingDialog(self)
        loading_dialog.show()
        self.QueryWorker = QueryWorker(self.excel_file)
        # close loading dialog after finished query work
        self.QueryWorker.finished.connect(loading_dialog.close)
        # run power query
        self.QueryWorker.start()

    def loadExcelData(self, excel_file):
        try:
            # run power query
            self.runQueryWithLoding()
        except Exception as e:
            self.showAlarm("Error", "File does not exist!\n" + str(e))
            return

        try:
            self.wb = openpyxl.load_workbook(excel_file)
        except:
            self.showAlarm("Format error", "File format is not supported!")
            return

        self.sheet_number = len(self.wb.sheetnames)

        # Remove empty columns from all sheets
        for sheet in self.wb:
            self.removeEmptyColumns(sheet)

        # clear existing tabs
        self.ui.tabWidget.clear()

        # create tabs
        if self.sheet_number > 1:
            for _ in range(self.sheet_number):
                tab = QWidget()
                tab.setObjectName("tab")
                self.ui.tabWidget.addTab(tab, "")

        # show data of excel in qt table
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
                    self.tableWidget.setHorizontalHeaderLabels(
                        headers
                    )  # set the headers
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

            # resize columns to fit the contents
            self.tableWidget.resizeColumnsToContents()

    def showAlarm(self, header, mes):
        QMessageBox.warning(self, header, mes)

    def openFile(self, auto_load_file=False):
        try:
            if self.excel_file:
                self.tableWidget.cellChanged.disconnect(self.saveExcelData)
        except:
            pass

        # auto loadfile for debug
        if auto_load_file:
            filePath = "Law Clients.xlsm"
        else:
            filePath, _ = QFileDialog.getOpenFileName(
                self, "Open File", "", "Excel Files (*.xlsx *.xlsm)"
            )

        if filePath:
            self.excel_file = filePath
            self.loadExcelData(self.excel_file)
            self.tableWidget.cellChanged.connect(self.saveExcelData)

    def removeEmptyColumns(self, sheet):
        columns_to_remove = []
        for i, column in enumerate(sheet.iter_cols(values_only=True), start=1):
            if all(cell is None for cell in column):
                columns_to_remove.append(i)

        for i in reversed(columns_to_remove):
            sheet.delete_cols(i)

    # BUG: save data just into first sheet
    def saveExcelData(self, row, column):
        if not self.excel_file:
            self.showAlarm("Error", "file does not exist !")

        # BUG: get first sheet
        sheet = self.wb.active
        sheet.cell(
            row=row + 1,
            column=column + 1,
            value=self.tableWidget.item(row, column).text(),
        )

        self.wb.save(self.excel_file)

    # BUG: write into rows that have logo
    def askForNewEntry(self) -> bool:
        # Check if wb has been defined
        if not hasattr(self, "wb"):
            self.showAlarm("Error", "Please load an Excel file first!")
            return False

        # open a popup window for new entry
        dialog = NewEntryDialog(self.wb, self)
        if dialog.exec():
            selected_sheet = dialog.comboBox.currentText()
            new_entry = {}
            for i, lineEdit in enumerate(dialog.lineEdits):
                new_entry[
                    dialog.lineEditsLayout.itemAt(i).widget().placeholderText()
                ] = lineEdit.text()
            for entry in new_entry:
                # fix showing None value cell
                if new_entry[entry] == "":
                    new_entry[entry] = None
            rows = list(self.wb[selected_sheet].iter_rows(values_only=True))
            for i in reversed(range(len(rows))):
                # skip deleting logo
                # NOTE: index 16 is header
                if selected_sheet == "Opening File" and i == 16:
                    break
                if all(cell is None for cell in rows[i]):
                    self.wb[selected_sheet].delete_rows(i + 1)
                else:
                    print(rows[i])

            # save entry data into selected sheet
            self.wb[selected_sheet].append(list(new_entry.values()))
            self.wb.save(self.excel_file)

            # Get the tableWidget of the tab that corresponds to the selected sheet
            selected_tab_index = self.wb.sheetnames.index(selected_sheet)
            selected_tab = self.ui.tabWidget.widget(selected_tab_index)
            tableWidget = selected_tab.findChild(QTableWidget)

            # update the table without reloading the file
            tableWidget.setRowCount(tableWidget.rowCount() + 1)
            for i, value in enumerate(new_entry.values()):
                tableWidget.setItem(
                    tableWidget.rowCount() - 1, i, QTableWidgetItem(str(value))
                )
            return True
        return False

    def genDocsBtn(self):
        if gen_docs(self.excel_file):
            QMessageBox.information(
                self,
                "Success",
                "Word documents successfully generated! Files are placed in the Docs folder.",
            )
        else:
            self.showAlarm("Error", "Word documents generation failed!")

    def showEnewEntryDialog(self):
        if not self.askForNewEntry():
            return
        # run power query
        self.runQueryWithLoding()
        print("New entry added")

    def showOprationDialog(self):
        dialog = OperationsDialog(self.excel_file)
        dialog.exec()
        self.loadExcelData(self.excel_file)

    def closeApplication(self):
        self.close()


class NewEntryDialog(QDialog):
    def __init__(self, wb, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Entry")

        self.setStyleSheet(
            """
            QDialog {
                background-color: #f0f0f0;
            }

            QLabel {
                font-size: 14px;
            }

            QLineEdit, QComboBox, QDateEdit {
                background-color: #fff;
                border: 1px solid #999;
                padding: 5px;
            }

            QPushButton {
                background-color: #007BFF;
                color: #fff;
                border: none;
                padding: 5px 10px;
                margin: 10px;
            }

            QPushButton:hover {
                background-color: #0056b3;
            }
            
            QComboBox:!editable, QComboBox::drop-down:editable {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                            stop: 0 #E1E1E1, stop: 0.4 #DDDDDD,
                                            stop: 0.5 #D8D8D8, stop: 1.0 #D3D3D3);
            }
        """
        )

        # Set the size of the dialog to be 3/4 of the size of the parent
        if parent is not None:
            self.resize(parent.size() * 0.5)

        self.layout = QVBoxLayout(self)

        self.comboBox = QComboBox(self)
        self.comboBox.addItems(wb.sheetnames)
        self.comboBox.setCurrentText(
            parent.ui.tabWidget.tabText(parent.ui.tabWidget.currentIndex())
        )

        self.layout.addWidget(self.comboBox)

        self.lineEditsLayout = QGridLayout()
        self.layout.addLayout(self.lineEditsLayout)

        self.lineEdits = []
        self.updateLineEdits(wb[self.comboBox.currentText()])

        self.comboBox.currentIndexChanged.connect(
            lambda: self.updateLineEdits(wb[self.comboBox.currentText()])
        )

        # Add a stretchable space
        self.layout.addStretch(1)

        self.button = QPushButton("Submit", self)
        # change size of button
        self.button.setMinimumSize(100, 60)
        self.button.clicked.connect(self.accept)
        self.layout.addWidget(self.button)

    def updateLineEdits(self, sheet):
        # Remove existing QLineEdit widgets
        for lineEdit in self.lineEdits:
            self.lineEditsLayout.removeWidget(lineEdit)
            lineEdit.deleteLater()
        self.lineEdits.clear()

        # Find the first non-empty row
        for row in sheet.iter_rows(values_only=True):
            if all(cell is not None and str(cell).strip() != "" for cell in row):
                columns = row
                break
        else:
            columns = []

        # Add new QLineEdit widgets
        for i, column in enumerate(columns):
            lineEdit = QLineEdit(self)
            lineEdit.setPlaceholderText(column)
            self.lineEditsLayout.addWidget(lineEdit, i // 3, i % 3)
            self.lineEdits.append(lineEdit)


class OperationsDialog(QDialog):
    def __init__(self, filepath):
        super().__init__()

        self.setStyleSheet(
            """
            QPushButton {
                background-color: #007cff; /* Green */
                border: none;
                color: white;
                padding: 15px 32px;
                text-align: center;
                text-decoration: none;
                font-size: 16px;
                margin: 4px 2px;
            }
            
            QPushButton:hover {
                background-color: #3094fd;
            }
        """
        )

        self.excel_file = filepath

        self.setWindowTitle("Operations")

        self.layout = QGridLayout(self)

        self.monthly_cases_report_button = QPushButton("Cases this Month")
        self.monthly_cases_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.monthly_cases_report_button, 0, 0)
        self.monthly_cases_report_button.clicked.connect(
            lambda: (generate_monthly_cases_report(self.excel_file), self.accept())
        )

        self.weekly_cases_report_button = QPushButton("Cases this Week")
        self.weekly_cases_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.weekly_cases_report_button, 0, 1)
        self.weekly_cases_report_button.clicked.connect(
            lambda: (generate_weekly_cases_report(self.excel_file), self.accept())
        )

        self.legal_aid_report_button = QPushButton("Legal Aid")
        self.legal_aid_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.legal_aid_report_button, 1, 0)
        self.legal_aid_report_button.clicked.connect(
            lambda: (generate_legal_aid_report(self.excel_file), self.accept())
        )

        self.bail_refused_report_button = QPushButton("Bail Refused")
        self.bail_refused_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.bail_refused_report_button, 1, 1)
        self.bail_refused_report_button.clicked.connect(
            lambda: (generate_bail_refused_report(self.excel_file), self.accept())
        )

        self.empty_counsel_report_button = QPushButton("Empty Counsel")
        self.empty_counsel_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.empty_counsel_report_button, 2, 0)
        self.empty_counsel_report_button.clicked.connect(
            lambda: (generate_empty_counsel_report(self.excel_file), self.accept())
        )

        self.non_zero_balance_report_button = QPushButton("Non Zero Balance")
        self.non_zero_balance_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.non_zero_balance_report_button, 2, 1)
        self.non_zero_balance_report_button.clicked.connect(
            lambda: (generate_non_zero_balance_report(self.excel_file), self.accept())
        )

        self.stage_reports_button = QPushButton("Stage Reports")
        self.stage_reports_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.stage_reports_button, 3, 0)
        self.stage_reports_button.clicked.connect(
            lambda: (generate_stage_reports(self.excel_file), self.accept())
        )


class LoadingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setModal(True)
        self.setWindowTitle("Loading...")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowCloseButtonHint)
        # self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.Tool)
        self.setStyleSheet(
            """
            QDialog {
                background-color: #fff;
                color: #000;
                border: 1px solid #999;
            }
        """
        )

        layout = QVBoxLayout()
        label = QLabel("Please wait...")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        progress = QProgressBar(self)
        progress.setRange(
            0, 0
        )  # Set range to 0,0 to create an indeterminate progress bar
        layout.addWidget(progress)

        self.setLayout(layout)


class QueryWorker(QThread):
    finished = Signal()

    def __init__(self, excel_file):
        super().__init__()
        self.excel_file = excel_file

    def run(self):
        pw_query.main(self.excel_file)
        self.finished.emit()


def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    # NOTE: auto load excel file for debug
    widget.openFile(auto_load_file=True)
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
