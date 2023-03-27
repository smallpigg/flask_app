import pandas as pd
from docxtpl import DocxTemplate

excel_file = 'input.xlsx'
word_file = 'template.docx'

# 读取 Excel 文件
df = pd.read_excel(excel_file)

output_dir = "output/"
# print("output")

# 渲染 Word 文件
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_file)
    doc.render(record)

    output_path = output_dir + f"{record['filename']}"

    # print(output_path)
    doc.save(output_path)

print("done!")