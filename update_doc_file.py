from docx import Document


def add_dict_to_doc(doc, my_dict):
    for key, value in my_dict.items():
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


# Your dictionary
my_dict = {
    "Key1": "Value1",
    "Key2": ["List item 1", "List item 2", "List item 3"],
    "Key3": "Value3",
    "Key4": ["List item 4", "List item 5"],
}

# Create a new Document
doc = Document()
add_dict_to_doc(doc, my_dict)

# Save the document
doc.save("./test_create_doc.docx")
