from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from word_content import get_content


def add_dict_to_doc(doc, my_list):
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


logo_path = "./bkp_logo.jpg"

wd_var = {
    "email": "test@email.com",
    "file_no": "",
    "date_opened": "",
    "fee_earner": "",
    "client_username": "",
    "client_fornames": "",
    "client_title": "",
    "marital_status": "",
    "letters_to_home_address": "",
    "address": "",
    "city": "",
    "postcode": "",
    "postal_address_if_different": "",
    "home_telephone": "",
    "work_telephone": "",
    "mobile_telephone": "",
    "occupation": "",
    "date_of_birth": "",
    "ethnicity": "",
    "prison_number": "",
    "national_insurance_no": "",
    "matter_type": "",
    "m_3rd_party": "",
    "initial": "",
}

# Create a new Document
doc = Document()
add_dict_to_doc(doc, get_content(wd_var))

# Save the document
doc.save("./gen_doc.docx")
