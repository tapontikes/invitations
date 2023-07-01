from docxtpl import DocxTemplate
from dotenv import load_dotenv
import os
from docxcompose.composer import Composer
from docx import Document as Document_compose


def merge_documents(temp_path, output_path):
    master = None
    for index, file in enumerate(os.listdir(temp_path)):
        file = temp_path + "/" + file
        if index == 0:
            master = Document_compose(file)
            continue
        composer = Composer(master)
        # filename_second_docx is the name of the second docx file
        append = Document_compose(file)
        # append the doc2 into the master using composer.append function
        composer.append(append)
        # Save the combined docx with a name
        composer.save(output_path)


def merge_address_info(template_path, output_path, addresses):
    output_path = output_path + "/output_%d.docx"
    for index, address in enumerate(addresses):
        # Load the template
        doc = DocxTemplate(template_path)
        # Replace the placeholders with the provided address information
        doc.render(address)
        # Save the modified document as a new file
        doc.save(output_path.replace("%d", str(index)))


dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path)

# Example usage
template_path = os.environ.get("TEMPLATE_PATH")
output_path = os.environ.get("OUTPUT_PATH")
temp_path = os.environ.get("TEMP_PATH")

addresses = [
    {
        "NAME": "Thomas Pontikes",
        "ADDRESS_1": "4 Phillips Ct.",
        "ADDRESS_2": "",
        "CITY": "Cranston",
        "STATE": "RI",
        "ZIP": "02921"
    },
    {
        "NAME": "Jane Smith",
        "ADDRESS_1": "456 Elm St",
        "ADDRESS_2": "Suite 2C",
        "CITY": "Los Angeles",
        "STATE": "CA",
        "ZIP": "90001"
    }
]

merge_address_info(template_path, temp_path, addresses)
merge_documents(temp_path, output_path)
