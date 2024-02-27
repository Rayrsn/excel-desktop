def get_content(wd_var):
    content = [
        "File Opening Form",
        [
            f"E-mail: {wd_var.get("email")}",
        ],
        [f"Previous Number ", "CRIME"],
        [f"File No: {wd_var.get("file_no")}", f"Date Opened: {wd_var.get("date_opened")}", f"Fee Earner: {wd_var.get("fee_earner")}"],
        [
            f"Client's surname: {wd_var.get("client_username")}",
            f"Forename(s): {wd_var.get("client_fornames")}",
            f"Title: {wd_var.get("client_title")}",
        ],
        [
            f"Marital Status: {wd_var.get("marital_status")}",
            f"Letters to home address: {wd_var.get("letters_to_home_address")}",
        ],
        [
            f"Address: {wd_var.get("address")}\n {wd_var.get("city")}\nPostcode : {wd_var.get("postcode")}",
            f"Postal Address (if different): {wd_var.get("postal_address_if_different")}\n Postcode{wd_var.get("postcode")}",
        ],
        [
            f"Telephone",
            f"Home: {wd_var.get("home_telephone")}",
            f"Work: {wd_var.get("work_telephone")}",
            f"Mobile: {wd_var.get("mobile_telephone")}",
        ],
        [
            f"Occupation: {wd_var.get("occupation")}",
            f"Date of birth: {wd_var.get("date_of_birth")}",
        ],
        [
            f"Ethnicity: {wd_var.get("ethnicity")}",
            f"prison number: {wd_var.get("prison_number")}",
        ],
        [
            f"National Insurance No: {wd_var.get("national_insurance_no")}",
            f"Surname At Birth: N/A",
        ],
        "",
        "",
        [
            f"Matter type (full description)",
            f"{wd_var.get("matter_type")}",
        ],
        [
            f"3nd Party",
            f"{wd_var.get("m_3rd_party")}",
            "initial",
            f"{wd_var.get("initial")}",
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
    return content
