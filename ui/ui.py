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

import ui.network as network

URL = "http://localhost:8000"

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # self.loadExcelData()
        self.ui.exitbutton.clicked.connect(self.closeApplication)
        self.ui.importbutton.clicked.connect(self.openFile)
        self.ui.exportbutton.clicked.connect(self.genDocsBtn)
        self.ui.newentrybutton.clicked.connect(self.showNewEntryDialog)
        self.ui.operationsbutton.clicked.connect(self.showOperationDialog)

        self.ui.exitbutton.setCursor(Qt.PointingHandCursor)
        self.ui.importbutton.setCursor(Qt.PointingHandCursor)
        self.ui.exportbutton.setCursor(Qt.PointingHandCursor)
        self.ui.newentrybutton.setCursor(Qt.PointingHandCursor)
        self.ui.refreshbutton.setCursor(Qt.PointingHandCursor)
        self.ui.operationsbutton.setCursor(Qt.PointingHandCursor)
        self.first_run_entry = True

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

    def updateUI(self):
        self.loadExcelData()
        print("Data loaded")

    def loadExcelData(self):
        """
        load excel data into qt tables
        """

        try:
            self.wb = openpyxl.load_workbook(self.excel_file)
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
    
    """ Example JSON response from the server
    {
        "sheet1": {
            "data": {
                "header A": [
                    "row 1": "value 1",
                    "row 2": "value 2",
                    "row 3": "value 3"
                ],
                "header B": [
                    "row 1": "value 4",
                    "row 2": "value 5",
                    "row 3": "value 6"
                ],
            }
        }
    }
    """

    def loadJsonData(self, url):
        """
        load JSON data into qt tables
        """

        try:
            json_data = network.get_data(url)
        except Exception as e:
            print(f"Error: {e}")
            self.showAlarm("Network error", "Failed to fetch data from the server!")
            return

        sheets = network.get_sheets(json_data)

        # clear existing tabs
        self.ui.tabWidget.clear()

        # create tabs
        for _ in range(len(sheets)):
            tab = QWidget()
            tab.setObjectName("tab")
            self.ui.tabWidget.addTab(tab, "")

        # show data of JSON in qt table
        for sh_num, sheet in enumerate(sheets):
            # Create a new QTableWidget for this tab
            self.tableWidget = QTableWidget()
            self.tableWidget.setRowCount(network.get_row_count(json_data, sheet))
            self.tableWidget.setColumnCount(network.get_column_count(json_data, sheet)-1)
            # Enable sorting
            self.tableWidget.setSortingEnabled(True)
            # set default sorting by first column
            # self.tableWidget.sortItems(0)

            # Add the Qself.tableWidget to a QHBoxLayout inside a QVBoxLayout
            hboxLayout = QHBoxLayout()
            hboxLayout.addWidget(self.tableWidget)
            vboxLayout = QVBoxLayout()
            vboxLayout.addLayout(hboxLayout)
            self.ui.tabWidget.widget(sh_num).setLayout(vboxLayout)

            # Change tabs name
            self.ui.tabWidget.setTabText(sh_num, sheet)

            # set value of table from JSON data
            headers = network.get_headers(json_data, sheet)
            if "row" in headers:
                headers.remove("row")
            self.tableWidget.setHorizontalHeaderLabels(headers)
            
            # sort the data in each sheet by the Sr_No column

            for i in range(self.tableWidget.rowCount()):
                for j in range(self.tableWidget.columnCount()):
                    headers = list(headers)
                    cell_data = network.get_data_from_cell(json_data, sheet, i, headers[j])
                    # if cell_data is float then convert it to int
                    if cell_data == "__null__":
                        cell_data = ""
                    if cell_data is not None and cell_data != "" and cell_data != "__null__":
                        if isinstance(cell_data, float) or (isinstance(cell_data, str) and cell_data.replace('.', '', 1).isdigit()):
                            cell_data = int(float(cell_data))
                        # if in column sr no, then convert it to int
                        if headers[j] == "Sr_No":
                            cell_data = int(cell_data)
                    # Skip setting the item if the header is "row"
                    if headers[j] != "row":
                        self.tableWidget.setItem(i, j, QTableWidgetItem(str(cell_data)))
            
            # resize columns to fit the contents
            self.tableWidget.resizeColumnsToContents()
    
    def addSheetToTabs(self, sheet_name, data):
        # Create a new tab
        tab = QWidget()
        tab.setObjectName("tab")
        self.ui.tabWidget.addTab(tab, "")
        # Create a new QTableWidget for this tab
        self.tableWidget = QTableWidget()
        self.tableWidget.setRowCount(len(data[next(iter(data))]))
        self.tableWidget.setColumnCount(len(data))
        # Enable sorting
        self.tableWidget.setSortingEnabled(True)
        # set default sorting by first column
        self.tableWidget.sortItems(0)

        # Add the Qself.tableWidget to a QHBoxLayout inside a QVBoxLayout
        hboxLayout = QHBoxLayout()
        hboxLayout.addWidget(self.tableWidget)
        vboxLayout = QVBoxLayout()
        vboxLayout.addLayout(hboxLayout)
        self.ui.tabWidget.widget(self.ui.tabWidget.count() - 1).setLayout(vboxLayout)

        # Change tabs name
        self.ui.tabWidget.setTabText(self.ui.tabWidget.count() - 1, sheet_name)

        # set value of table from data
        headers = list(data.keys())
        self.tableWidget.setHorizontalHeaderLabels(headers)

        for i, row in enumerate(data):
            for j, header in enumerate(headers):
                row_data = data[header][i]  # Get the dictionary for the current row
                cell_data = list(row_data.values())[0]
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(cell_data)))
        
        # resize columns to fit the contents
        self.tableWidget.resizeColumnsToContents()
    
    def loadReport(self, url, name):
        # Fetch the data
        url = f"{url}/operations/{name}"
        try:
            data = network.get_data(url)
        except Exception as e:
            print(f"Error: {e}")
            self.showAlarm("Network error", "Failed to fetch data from the server!")
            return
        
        if data is None:
            return

        # Display the data in a new table popup
        dialog = QDialog(self)
        dialog.setWindowTitle(name)
        dialog.resize(800, 600)
        layout = QVBoxLayout(dialog)
        table = TableViewer(dialog)
        if name == "legal-aid":
            table.set_data(data, name="legal-aid")
        else:
            table.set_data(data)
        layout.addWidget(table)
        dialog.exec()
        
        
    
    def showAlarm(self, header, mes):
        QMessageBox.warning(self, header, mes)

    def openFile(self, auto_load_file=False):
        # open file after first time
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

            # show excel data into tables
            # self.loadExcelData()
            self.loadJsonData(URL)

            # connect tables to saveExcelData function
            self.tableWidgetCellChange(is_connect=True)

    def removeEmptyColumns(self, sheet):
        columns_to_remove = []
        for i, column in enumerate(sheet.iter_cols(values_only=True), start=1):
            if all(cell is None for cell in column):
                columns_to_remove.append(i)

        for i in reversed(columns_to_remove):
            sheet.delete_cols(i)

    def saveExcelData(self):
        # Disconnect the cellChanged signal
        self.tableWidgetCellChange(is_connect=False)

        # Save data into excel file
        for sheet_index in range(self.sheet_number):
            tableWidget = self.ui.tabWidget.widget(sheet_index).findChild(QTableWidget)
            # Write the data from the tableWidget back to the sheet
            for i in range(tableWidget.rowCount()):
                for j in range(tableWidget.columnCount()):
                    if tableWidget.item(i, j) is not None:
                        # seprate logic of first sheet for Writing
                        row_offset = 17 if sheet_index == 0 else 1
                        self.wb[self.wb.sheetnames[sheet_index]].cell(
                            row=i + 1 + row_offset,
                            column=j + 1,
                            value=tableWidget.item(i, j).text(),
                        )
        self.wb.save(self.excel_file)

        # Reconnect the cellChanged signal
        self.tableWidgetCellChange(is_connect=True)

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
                if new_entry[entry].strip() == "":
                    new_entry[entry] = None
            rows = list(self.wb[selected_sheet].iter_rows(values_only=True))
            for i in reversed(range(len(rows))):
                # skip deleting logo
                # NOTE: index 16 is header
                if selected_sheet == "Opening File" and i == 16:
                    break
                if all(cell is None or str(cell).strip() == "" for cell in rows[i]):
                    self.wb[selected_sheet].delete_rows(i + 1)
                # else:
                # print(rows[i])

            if not self.first_run_entry:
                # Disconnect the cellChanged signal
                self.tableWidgetCellChange(is_connect=False)

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

            # Reconnect the cellChanged signal
            self.tableWidgetCellChange(is_connect=True)

            self.first_run_entry = False
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

    def showNewEntryDialog(self):
        if not self.askForNewEntry():
            return
        print("New entry added")

    def showOperationDialog(self):
        dialog = OperationsDialog(self, URL)
        dialog.exec()

    def tableWidgetCellChange(self, is_connect: bool):
        """
        connect and disconnect trigger for change cell
        """
        try:
            for i in range(self.sheet_number):
                self.tableWidget = self.ui.tabWidget.widget(i).findChild(QTableWidget)
                if is_connect:
                    self.tableWidget.cellChanged.connect(self.saveExcelData)
                else:
                    self.tableWidget.cellChanged.disconnect(self.saveExcelData)
        except Exception as e:
            print(f"have error when is_connect is {is_connect} in :{e} ")

    def exportToJson(self):
        json_data = {}
        for i in range(self.ui.tabWidget.count()):
            tableWidget = self.ui.tabWidget.widget(i).findChild(QTableWidget)
            sheet_name = self.ui.tabWidget.tabText(i)
            sheet_data = {}
            for j in range(tableWidget.columnCount()):
                header = tableWidget.horizontalHeaderItem(j).text()
                column_data = []
                for k in range(tableWidget.rowCount()):
                    cell_data = tableWidget.item(k, j).text()
                    column_data.append({f"row {k + 1}": cell_data})
                sheet_data[header] = column_data
            json_data[sheet_name] = {"data": sheet_data}
        return json_data

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
    def __init__(self, parent, url):
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

        self.parent = parent

        self.setWindowTitle("Operations")

        self.layout = QGridLayout(self)

        self.monthly_cases_report_button = QPushButton("Cases this Month")
        self.monthly_cases_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.monthly_cases_report_button, 0, 0)
        self.monthly_cases_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "monthly"), self.accept())
        )

        self.weekly_cases_report_button = QPushButton("Cases this Week")
        self.weekly_cases_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.weekly_cases_report_button, 0, 1)
        self.weekly_cases_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "weekly"), self.accept())
        )

        self.legal_aid_report_button = QPushButton("Legal Aid")
        self.legal_aid_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.legal_aid_report_button, 1, 0)
        self.legal_aid_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "legal-aid"), self.accept())
        )

        self.bail_refused_report_button = QPushButton("Bail Refused")
        self.bail_refused_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.bail_refused_report_button, 1, 1)
        self.bail_refused_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "bail-refused"), self.accept())
        )

        self.empty_counsel_report_button = QPushButton("Empty Counsel")
        self.empty_counsel_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.empty_counsel_report_button, 2, 0)
        self.empty_counsel_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "empty-counsel"), self.accept())
        )

        self.non_zero_balance_report_button = QPushButton("Non Zero Balance")
        self.non_zero_balance_report_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.non_zero_balance_report_button, 2, 1)
        self.non_zero_balance_report_button.clicked.connect(
            lambda: (parent.loadReport(url, "non-zero"), self.accept())
        )

        self.stage_reports_button = QPushButton("Stage Reports")
        self.stage_reports_button.setMinimumSize(100, 40)
        self.layout.addWidget(self.stage_reports_button, 3, 0)
        self.stage_reports_button.clicked.connect(
            lambda: (parent.loadReport(url, "stage"), self.accept())
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


"""EXAMPLE DATA for legal-aid

{
    "Submitted": 4,
    "Approved": 2,
    "Appealed": 1,
    "Refused": 2,
    "Date Stamped": 2,
    "__null__": 4,
    "Total Number of Clients": 15
}
"""

class TableViewer(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSortingEnabled(True)
        self.sortItems(0)

    def set_data(self, data, name=None):
        # Get the headers from the first row (the keys of the dictionary)
        if name == "legal-aid":
            headers = list(data.keys())
            self.setColumnCount(2)
            self.setHorizontalHeaderLabels(["Legal Aid Category", "Number of Clients"])
        else:
            headers = list(data[0].keys())
            self.setColumnCount(len(headers))
            # Set the headers
            self.setHorizontalHeaderLabels(headers)
        self.setRowCount(len(data))
        
        
        # Set the data
        if name == "legal-aid":
            for i, header in enumerate(headers):
                cell_data = data[header]
                if header == "__null__":
                    header = ""
                    cell_data = ""
                if isinstance(cell_data, float):
                    cell_data = int(cell_data)
                if cell_data == "__null__":
                    cell_data = ""
                self.setItem(i, 0, QTableWidgetItem(header))
                self.setItem(i, 1, QTableWidgetItem(str(cell_data)))
        else:
            for i, row in enumerate(data):
                for j, header in enumerate(headers):                    
                    cell_data = row[header]
                    if isinstance(cell_data, float):
                        cell_data = int(cell_data)
                    if cell_data == "__null__":
                        cell_data = ""
                    self.setItem(i, j, QTableWidgetItem(str(cell_data)))
        
        # Resize the columns to fit the contents
        self.resize_columns()
        
    def resize_columns(self):
        self.resizeColumnsToContents()

def run():
    app = QApplication(sys.argv)
    widget = MainWindow()
    # NOTE: auto load excel file for debug
    widget.openFile(auto_load_file=True)
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
