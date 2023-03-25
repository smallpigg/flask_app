from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from docx import Document
# from docx.exceptions import PendingDeprecationWarning
# import warnings

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/render', methods=['POST'])
def render():
    # 获取上传的文件
    excel_file = request.files['excel_file']
    word_file = request.files['word_file']

    # 读取 Excel 文件
    df = pd.read_excel(excel_file)

    # 渲染 Word 文件
    document = Document(word_file)
    table = document.tables[0]
    for i, row in df.iterrows():
        cells = table.rows[i + 1].cells
        cells[0].text = row['Name']
        cells[1].text = str(row['Age'])
        cells[2].text = row['Gender']

    # 保存渲染后的 Word 文件
    document.save('output.docx')

    # 提供下载链接ggg
    return send_file('output.docx', as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
