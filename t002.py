from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl
import docx
import pandas as pd
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


output_path="output\\1\\文件1.docx"

def delete_empty_rows(table):
    for row in table.rows:
        first_cell = row.cells[0]
        # print the text in the first cell
        # print(first_cell.text)
        if first_cell.text == '':
            row._element.getparent().remove(row._element)

doc = docx.Document(output_path)

for table in doc.tables:
    delete_empty_rows(table)

doc.save(output_path)