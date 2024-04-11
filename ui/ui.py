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

URL = "https://excel-api.fly.dev"
DATA = {}


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
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

    def loadJsonDataFinished(self, data):
        global DATA
        DATA = data
        json_data = data
        
        sheets = network.get_sheets(json_data)

        # close loading dialog
        self.loading_dialog.close()
        
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
            tableWidget = QTableWidget()
            tableWidget.setRowCount(network.get_row_count(json_data, sheet))
            tableWidget.setColumnCount(network.get_column_count(json_data, sheet)-1)
            # Enable sorting
            tableWidget.setSortingEnabled(True)
            # set default sorting by first column
            # tableWidget.sortItems(0)

            # Add the QtableWidget to a QHBoxLayout inside a QVBoxLayout
            hboxLayout = QHBoxLayout()
            hboxLayout.addWidget(tableWidget)
            vboxLayout = QVBoxLayout()
            vboxLayout.addLayout(hboxLayout)
            self.ui.tabWidget.widget(sh_num).setLayout(vboxLayout)

            # Change tabs name
            self.ui.tabWidget.setTabText(sh_num, sheet)

            # set value of table from JSON data
            headers = network.get_headers(json_data, sheet)
            if "row" in headers:
                headers.remove("row")
            tableWidget.setHorizontalHeaderLabels(headers)
            
            # sort the data in each sheet by the Sr_No column

            for i in range(tableWidget.rowCount()):
                for j in range(tableWidget.columnCount()):
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
                        tableWidget.setItem(i, j, QTableWidgetItem(str(cell_data)))
            
            # resize columns to fit the contents
            tableWidget.resizeColumnsToContents()
            # on cell change print the changed cell
            tableWidget.cellChanged.connect(self.handleCellChanged)

    def loadJsonData(self, url):
        """
        load JSON data into qt tables
        """
        
        # make a loading dialog
        self.loading_dialog = LoadingDialog(self)
    
        self.fetchDataThread = FetchDataThread(url)
        self.fetchDataThread.dataReady.connect(self.loadJsonDataFinished)
        self.fetchDataThread.start()
        
        self.loading_dialog.exec()
        
        # Fetch the data
        global DATA
        try:
            json_data = network.get_data(url)
            DATA = json_data
        except Exception as e:
            print(f"Error: {e}")
            self.showAlarm("Network error", "Failed to fetch data from the server!")
            return

    def addSheetToTabs(self, sheet_name, data):
        # Create a new tab
        tab = QWidget()
        tab.setObjectName("tab")
        self.ui.tabWidget.addTab(tab, "")
        # Create a new QTableWidget for this tab
        tableWidget = QTableWidget()
        tableWidget.setRowCount(len(data[next(iter(data))]))
        tableWidget.setColumnCount(len(data))
        # Enable sorting
        tableWidget.setSortingEnabled(True)
        # set default sorting by first column
        tableWidget.sortItems(0)

        # Add the QtableWidget to a QHBoxLayout inside a QVBoxLayout
        hboxLayout = QHBoxLayout()
        hboxLayout.addWidget(tableWidget)
        vboxLayout = QVBoxLayout()
        vboxLayout.addLayout(hboxLayout)
        self.ui.tabWidget.widget(self.ui.tabWidget.count() - 1).setLayout(vboxLayout)

        # Change tabs name
        self.ui.tabWidget.setTabText(self.ui.tabWidget.count() - 1, sheet_name)

        # set value of table from data
        headers = list(data.keys())
        tableWidget.setHorizontalHeaderLabels(headers)

        for i, row in enumerate(data):
            for j, header in enumerate(headers):
                row_data = data[header][i]  # Get the dictionary for the current row
                cell_data = list(row_data.values())[0]
                tableWidget.setItem(i, j, QTableWidgetItem(str(cell_data)))
        
        # resize columns to fit the contents
        tableWidget.resizeColumnsToContents()
    
    def handleCellChanged(self, row, column):
        tableWidget = self.sender()
        cell_id = tableWidget.item(row, 0).text()
        print(f"Row: {row}, ID: {cell_id}: {tableWidget.item(row, column).text()}")
    
    def loadReport(self, url, name):
        # Fetch the data
        global DATA
        url = f"{url}/operations/{name}"
        try:
            data = network.get_data(url)
            DATA = data
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
        try:
            if name == "legal-aid":
                table.set_data(data, name="legal-aid")
            elif name == "bail-refused":
                table.set_data(data, name="bail-refused")
            else:
                table.set_data(data)
        except IndexError:
            self.showAlarm("Error", "No data to display!")
            return
        layout.addWidget(table)
        dialog.exec()
        
        
    
    def showAlarm(self, header, mes):
        QMessageBox.warning(self, header, mes)

    def openFile(self):
        self.loadJsonData(URL)

        # connect tables to saveExcelData function
        # self.tableWidgetCellChange(is_connect=True)

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

    def askForNewEntry(self, data) -> bool:
        # open a popup window for new entry
        dialog = NewEntryDialog(data, self)
        if dialog.exec():
            selected_sheet = dialog.comboBox.currentText().replace(" ", "_")
            new_entry_data = {}
            for i, lineEdit in enumerate(dialog.lineEdits):
                new_entry_data[
                    dialog.lineEditsLayout.itemAt(i).widget().placeholderText()
                ] = lineEdit.text()
            for entry in new_entry_data:
                # fix showing None value cell
                if new_entry_data[entry].strip() == "":
                    new_entry_data[entry] = None

            print(new_entry_data)

            new_entry = {
                "sheetname": selected_sheet,
                "data": new_entry_data
            }

            # Send a POST request to the server with the new entry data
            response = network.post_data(
                f"{URL}/create", new_entry
            )

            if response.status_code == 200:
                self.showAlarm("Success", "New entry added successfully!")
                return True
            else:
                self.showAlarm("Error", "Failed to add new entry!")
                return False
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
        global DATA
        data = DATA
        if not self.askForNewEntry(data):
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
                tableWidget = self.ui.tabWidget.widget(i).findChild(QTableWidget)
                if is_connect:
                    tableWidget.cellChanged.connect(self.saveExcelData)
                else:
                    tableWidget.cellChanged.disconnect(self.saveExcelData)
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
    def __init__(self, data, parent=None):
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
        
        main_data = data
        headers = list(data.get("headers"))
        data = data.get("data")
    
        # Set the size of the dialog to be 3/4 of the size of the parent
        if parent is not None:
            self.resize(parent.size() * 0.5)

        self.layout = QVBoxLayout(self)

        self.comboBox = QComboBox(self)
        self.comboBox.addItems(headers)
        self.comboBox.setCurrentText(
            parent.ui.tabWidget.tabText(parent.ui.tabWidget.currentIndex())
        )

        self.layout.addWidget(self.comboBox)

        self.lineEditsLayout = QGridLayout()
        self.layout.addLayout(self.lineEditsLayout)

        self.lineEdits = []
        self.updateLineEdits(main_data.get("headers")[self.comboBox.currentText()])
        

        self.comboBox.currentIndexChanged.connect(
            lambda: self.updateLineEdits(main_data.get("headers")[self.comboBox.currentText()])
        )

        # Add a stretchable space
        self.layout.addStretch(1)

        self.button = QPushButton("Submit", self)
        # change size of button
        self.button.setMinimumSize(100, 60)
        self.button.clicked.connect(self.accept)
        self.layout.addWidget(self.button)

    def updateLineEdits(self, columns):
        # Remove existing QLineEdit widgets
        for lineEdit in self.lineEdits:
            self.lineEditsLayout.removeWidget(lineEdit)
            lineEdit.deleteLater()
        self.lineEdits.clear()

        # Filter out "id" and "row" columns
        columns = [column for column in columns if column not in ["id", "row"]]

        # Create a new QGridLayout
        self.lineEditsLayout = QGridLayout()

        # Add new QLineEdit widgets
        for i, column in enumerate(columns):
            lineEdit = QLineEdit(self)
            lineEdit.setPlaceholderText(column)
            self.lineEditsLayout.addWidget(lineEdit, i // 4, i % 4)
            self.lineEdits.append(lineEdit)

        # Add the new QGridLayout to the dialog's layout
        self.layout.insertLayout(1, self.lineEditsLayout)

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

class FetchDataThread(QThread):
    dataReady = Signal(object)

    def __init__(self, url):
        super().__init__()
        self.url = url

    def run(self):
        data = network.get_data(self.url)
        self.dataReady.emit(data)


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


"""EXAMPLE DATA for bail-refused report

{
    "Magistrates Merge": [],
    "Crown Court Merge": [
        {
            "id": 2,
            "row": 1,
            "Office": "London",
            "Matter_Type": "Criminal (Magistrates)",
            "Email": "Haroon@bkpsolicitors.com",
            "File_No": "5",
            "Fee_Earner": "__null__",
            "Clients_Surname": "Umer",
            "Clients_Forenames": "Haroon",
            "Address": "101 Leeds Rd",
            "City": "__null__",
            "Postcode": "BD3 0NG",
            "Mobile_Number": "7858977659",
            "HMP": "__null__",
            "Prison_Number": "A945829",
            "National_Insurance_Number": "JZ636434C",
            "Legal_Aid": "Appealed",
            "Legal_Aid1": null,
            "Court": "Police Station",
            "Name_of_Counsel": "__null__",
            "Counsel_Email_Address": "__null__",
            "Chambers": "__null__",
            "PTPH_Date": "__null__",
            "Type_of_Letter": "__null__",
            "Letter_Sent": "__null__",
            "Stage_1": "2024-03-01T00:00:00",
            "Stage_2": "2024-03-02T00:00:00",
            "Stage_3": "2024-03-24T00:00:00",
            "Stage_4": "2024-03-25T00:00:00",
            "CTR_Date": "__null__",
            "Trial_Date": "__null__",
            "Outcome": "__null__",
            "Type_of_Letter_2": "__null__",
            "Appeal": "__null__",
            "Legal_Aid2": "__null__",
            "Matt_Number_If_Legal_Aid_Granted": "__null__",
            "Offence": null,
            "Type_of_Offence": null,
            "First_Hearing_Date": null,
            "Outcome_First_Hearing": null,
            "Letter_Issued": null,
            "Second_Hearing_Date": null,
            "Outcome_Second_Hearing": null,
            "Type_of_Letter_3": null,
            "Letter_Issued2": null,
            "Sr_No": null,
            "Previous_Number": null,
            "CRIME": null,
            "Date_Opened": null,
            "Clients_Title": null,
            "Marital_Status": null,
            "Letters_to_Home_Address": null,
            "Postal_Address_if_Different": null,
            "Postal_Address_Postcode": null,
            "Home_Telephone": null,
            "Work_Telephone": null,
            "Occupation": null,
            "Date_of_Birth": "01.01.2000",
            "Ethnicity": null,
            "_3rd_Party": null,
            "Initial": null,
            "Conflict": null,
            "Date": null,
            "Costs_Information": null,
            "Cost_Estimate": null,
            "Charge_Basis": null,
            "Next_Date": null,
            "Time": null,
            "Location": null,
            "Has_Result_Been_Diarised": null,
            "Co_Accused": null,
            "Conflict_1": null,
            "Name_of_Co_Accused": null,
            "Represented_by": null,
            "Comments": "__null__",
            "Bail": "Bail Refused",
            "DCS_Uploaded": "__null__",
            "Defense_Case_Statement_Date": "__null__",
            "URN": "__null__"
        },
        {
            "id": 3,
            "row": 2,
            "Office": "London",
            "Matter_Type": "Criminal (Magistrates)",
            "Email": "IK@hotmail.com",
            "File_No": "8",
            "Fee_Earner": "__null__",
            "Clients_Surname": "Khan",
            "Clients_Forenames": "Imran",
            "Address": "43 Hprton Grange Rd",
            "City": "__null__",
            "Postcode": "bd7 3ah",
            "Mobile_Number": "7858977659",
            "HMP": "__null__",
            "Prison_Number": "A945832",
            "National_Insurance_Number": "JZ636434C",
            "Legal_Aid": "Submitted",
            "Legal_Aid1": null,
            "Court": "Police Station",
            "Name_of_Counsel": "__null__",
            "Counsel_Email_Address": "__null__",
            "Chambers": "__null__",
            "PTPH_Date": "__null__",
            "Type_of_Letter": "__null__",
            "Letter_Sent": "__null__",
            "Stage_1": "__null__",
            "Stage_2": "__null__",
            "Stage_3": "__null__",
            "Stage_4": "__null__",
            "CTR_Date": "__null__",
            "Trial_Date": "__null__",
            "Outcome": "__null__",
            "Type_of_Letter_2": "__null__",
            "Appeal": "__null__",
            "Legal_Aid2": "__null__",
            "Matt_Number_If_Legal_Aid_Granted": "__null__",
            "Offence": null,
            "Type_of_Offence": null,
            "First_Hearing_Date": null,
            "Outcome_First_Hearing": null,
            "Letter_Issued": null,
            "Second_Hearing_Date": null,
            "Outcome_Second_Hearing": null,
            "Type_of_Letter_3": null,
            "Letter_Issued2": null,
            "Sr_No": null,
            "Previous_Number": null,
            "CRIME": null,
            "Date_Opened": null,
            "Clients_Title": null,
            "Marital_Status": null,
            "Letters_to_Home_Address": null,
            "Postal_Address_if_Different": null,
            "Postal_Address_Postcode": null,
            "Home_Telephone": null,
            "Work_Telephone": null,
            "Occupation": null,
            "Date_of_Birth": "01.01.92",
            "Ethnicity": null,
            "_3rd_Party": null,
            "Initial": null,
            "Conflict": null,
            "Date": null,
            "Costs_Information": null,
            "Cost_Estimate": null,
            "Charge_Basis": null,
            "Next_Date": null,
            "Time": null,
            "Location": null,
            "Has_Result_Been_Diarised": null,
            "Co_Accused": null,
            "Conflict_1": null,
            "Name_of_Co_Accused": null,
            "Represented_by": null,
            "Comments": "__null__",
            "Bail": "JC Bail Refused",
            "DCS_Uploaded": "__null__",
            "Defense_Case_Statement_Date": "__null__",
            "URN": "__null__"
        }
    ]
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
        elif name == "bail-refused":
            table_headers = ["Office", "Matter_Type", "Email", "File_No", "Fee_Earner", "Clients_Surname",
                            "Clients_Forenames", "Address", "City", "Postcode", "Mobile_Number", "Date_of_Birth",
                            "HMP", "Prison_Number", "National_Insurance_Number", "Legal_Aid", "Court"]
            self.setColumnCount(len(table_headers))
            self.setHorizontalHeaderLabels(table_headers)
        else:
            headers = list(data[0].keys())
            self.setColumnCount(len(headers))
            # Set the headers
            self.setHorizontalHeaderLabels(headers)
        
        if name == "bail-refused":
            row_count = 0
            for value in data.values():
                row_count += len(value)
            self.setRowCount(row_count)
        else:
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
        elif name == "bail-refused":
            row_count = 0
            for value in data.values():
                for dic in value:
                    for j, header in enumerate(table_headers):
                        cell_data = dic[header]
                        if isinstance(cell_data, float):
                            cell_data = int(cell_data)
                        if cell_data == "__null__":
                            cell_data = ""
                        self.setItem(row_count, j, QTableWidgetItem(str(cell_data)))
                    row_count += 1
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
    widget.openFile()
    widget.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
