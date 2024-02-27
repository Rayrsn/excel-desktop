from docx import Document


def add_dict_to_doc(doc, my_list):
    for value in my_list:
        if isinstance(value, list):
            # Create a table with one row and a number of columns equal to the number of items in the list
            table = doc.add_table(rows=1, cols=len(value))
            table.style = "Table Grid"  # Optional: set the style of the table
            for i, item in enumerate(value):
                # Set each item in its own cell
                cell = table.cell(0, i)
                cell.text = str(item)
        else:
            # If the value is not a list, just add it as a new paragraph
            doc.add_paragraph(str(value))


email = ""
file_no = ""
date_opened = ""
fee_earner = ""
client_username = ""
client_fornames = ""
client_title = ""
marital_status = ""
letters_to_home_address = ""
address = ""
city = ""
postcode = ""
postal_address_if_different = ""
home_telephone = ""
work_telephone = ""
mobile_telephone = ""
occupation = ""
date_of_birth = ""
ethnicity = ""
prison_number = ""
national_insurance_no = ""
matter_type = ""
m_3rd_party = ""
initial = ""
my_dict = [
    "File Opening Form",
    [f"E-mail: {email}", "LOGO"],
    [f"Previous Number ", "CRIME"],
    [f"File No: {file_no}", f"Date Opened: {date_opened}", f"Fee Earner: {fee_earner}"],
    [
        f"Client's surname: {client_username}",
        f"Forename(s): {client_fornames}",
        f"Title: {client_title}",
    ],
    [
        f"Marital Status: {marital_status}",
        f"Letters to home address: {letters_to_home_address}",
    ],
    [
        f"Address: {address}\n {city}\nPostcode : {postcode}",
        f"Postal Address (if different): {postal_address_if_different}\n Postcode{postcode}",
    ],
    [
        f"Telephone",
        f"Home: {home_telephone}",
        f"Work: {work_telephone}",
        f"Mobile: {mobile_telephone}",
    ],
    [
        f"Occupation: {occupation}",
        f"Date of birth: {date_of_birth}",
    ],
    [
        f"Ethnicity: {ethnicity}",
        f"prison number: {prison_number}",
    ],
    [
        f"National Insurance No: {national_insurance_no}",
        f"Surname At Birth: N/A",
    ],
    "",
    "",
    [
        f"Matter type (full description)",
        f"{matter_type}",
    ],
    [
        f"3nd Party",
        f"{m_3rd_party}",
        "initial",
        f"{initial}",
        f"conflict",
        f"Yes",
        f"No",
    ],
    [
        f"3nd Party",
        f"",
        "Date",
        f"",
        f"conflict",
        f"Yes",
        f"No",
    ],
    ["COSTS INFORMATION"],
    "",
    [
        f"Legal Aid X",
        f"Cost Estimate",
        f"Private",
        f"Cost Estimate ",
    ],
    "",
    [
        f"CHARGE BASIC",
        f"TYPE",
        f"COURT",
    ],
    [
        f"Criminal Investigation: X",
        f"Criminal (Magistrates, Franchise)",
        f"Magistrates",
    ],
    [
        f"Criminal Proceedings:",
        f"Criminal (Crown)",
        f"Crown/County",
    ],
    [
        f"Public Funding Certificate",
        f"Duty Solicitor",
        f"High Court",
    ],
    [
        f"",
        f"Criminal Private",
        f"Non-designated",
    ],
    [
        f"",
        f"Civil",
        f"",
    ],
    [
        f"",
        f"",
        f"",
    ],
]

# Create a new Document
doc = Document()
add_dict_to_doc(doc, my_dict)

# Save the document
doc.save("./gen_doc.docx")
