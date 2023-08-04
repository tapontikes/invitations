import os, requests, openpyxl
from docxtpl import DocxTemplate
from dotenv import load_dotenv
from docxcompose.composer import Composer
from docx import Document


def merge_documents(temp_path, output_path):
    master = None
    for index, file in enumerate(os.listdir(temp_path)):
        file = temp_path + "/" + file
        if index == 0:
            master = Document(file)
            continue
        composer = Composer(master)
        append = Document(file)
        composer.append(append)
        composer.save(output_path)


def merge_address_info(template_path, output_path, addresses):
    output_path = output_path + "/output_%d.docx"
    for index, address in enumerate(addresses):
        # Load the template
        doc = DocxTemplate(template_path)
        # Replace the placeholders with the provided address information
        doc.render(address, autoescape=True)
        # Save the modified document as a new file
        doc.save(output_path.replace("%d", str(index)))


def get_sheet():
    sheet_id = "13FSw1U-ZZZWU50gShZyZp6nAHHSqZfOaR8u1Q6RETyE"
    url = "https://docs.google.com/spreadsheets/d/{}/export?format=xlsx".format(sheet_id)
    response = requests.get(url)
    with open("guest.xlsx", "wb") as f:
        f.write(response.content)
    return response.content


def iter_rows():
    wb = openpyxl.load_workbook("guest.xlsx")
    # Get the first sheet
    ws = wb.active

    data = []

    for i in range(5, ws.max_row + 1):
        cells = [cell.value for cell in ws[i]]
        if cells[0] == "Print":
            data.append({
            "PREFIX": cells[1],
            "SUFFIX": "",
            "NAME": cells[2] + " " + cells[3],
            "ADDRESS1": "Test ADDR",
            "ADDRESS2": " ",
            "CITY": "Cranston",
            "STATE": "Rhode Island",
            "ZIP": "02920"
        })
    return data


for file in os.listdir("tmp"):
    if file:
        os.remove(os.path.join("tmp", file))


dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path)
template_path = os.environ.get("TEMPLATE_PATH")
output_path = os.environ.get("OUTPUT_PATH")
temp_path = os.environ.get("TEMP_PATH")

get_sheet()
addresses = iter_rows()

merge_address_info(template_path, temp_path, addresses)
merge_documents(temp_path, output_path)
