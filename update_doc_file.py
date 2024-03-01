from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from word_content import get_content
from utils.people_date import *


def add_dict_to_doc(doc, my_list, logo_path):
    # Add a paragraph with right alignment
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # Add logo
    run = paragraph.add_run()
    run.add_picture(logo_path, width=Inches(1.5))  # Adjust size as needed

    for value in my_list:
        if isinstance(value, list):
            # Create a table with one row and a number of columns equal to the number of items in the list
            if all(isinstance(i, dict) for i in value):
                table = doc.add_table(rows=1, cols=len(value) - 1)
            else:
                table = doc.add_table(rows=1, cols=len(value))
            table.style = "Table Grid"  # Optional: set the style of the table
            for i, item in enumerate(value):
                # Set each item in its own cell
                cell = table.cell(0, i)
                cell.text = str(item)
        else:
            # If the value is not a list, just add it as a new paragraph
            doc.add_paragraph(str(value))


def create_doc(doc_name, wd_var, logo_path):
    # Create a new Document
    doc = Document()
    add_dict_to_doc(doc, get_content(wd_var), logo_path)

    # Save the document
    doc.save(doc_name)


def gen_docs():
    logo_path = "./bkp_logo.jpg"

    excel_file = "../docs/Law Clients Excel Sheet Shared_MainV3.xlsx"
    # Load the workbook and select the first sheet
    wb = load_workbook(excel_file)
    # This selects the first sheet
    sheet = wb.worksheets[0]
    datas = get_all_people_data(wb, sheet)
    for data in datas:
        wd_var = {
            "email": data["Email"],
            "file_no": data["File No."],
            "date_opened": data["Date Opened"],
            "fee_earner": data["Fee Earner"],
            "client_username": data["Client's Surname"],
            "client_fornames": data["Client's Forename(s)"],
            "client_title": data["Client's Title"],
            "marital_status": data["Marital Status"],
            "letters_to_home_address": data["Letters to Home Address"],
            "address": data["Address"],
            "city": data["City"],
            "postcode": data["Postcode"],
            "postal_address_if_different": data["Postal Address (if Different)"],
            "home_telephone": data["Home Telephone"],
            "work_telephone": data["Work Telephone"],
            "mobile_telephone": data["Mobile Number"],
            "occupation": data["Occupation"],
            "date_of_birth": data["Date of Birth"],
            "ethnicity": data["Ethnicity"],
            "prison_number": data["Prison Number"],
            "national_insurance_no": data["National Insurance Number"],
            "matter_type": data["Matter Type"],
            "m_3rd_party": data["3rd Party"],
            "initial": data["Initial"],
        }
        name = data["Client's Surname"]
        forname = data["Client's Forename(s)"]
        doc_name = f"./{name}_{forname}.docx"
        create_doc(doc_name, logo_path)
