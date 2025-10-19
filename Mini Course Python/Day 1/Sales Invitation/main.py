# Import package yang akan kita gunakan
import openpyxl
from docxtpl import DocxTemplate

# 1. Cek file excel
file_excel = "Day 1/Sales Invitation/brosur_data.xlsx"

# 2. Cek sheet yang active
load = openpyxl.load_workbook(file_excel)
cek_sheet = load.active

# 3. Get value from sheet
get_value = list(cek_sheet.values)
print(get_value)

# 4. Render document
file_doc = DocxTemplate("Day 1/Sales Invitation/brosur.docx")

for value in get_value[1:]:
    file_doc.render({
        "HEADLINE": value[0],
        "ALAMAT": value[1],
        "NOPE": value[2],
        "EMAIL": value[3],
        "INFO": value[4],
        "NAMA": value[5],
    })

    # Simpan file dalam format baru
    file_doc.name = f"Brosur Undangan untuk {value[5]}.docx"
    file_doc.save(file_doc.name)





