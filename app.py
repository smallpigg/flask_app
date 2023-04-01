from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from docx import Document
from docxtpl import DocxTemplate  # pip install docxtpl
import docx
# from docx.exceptions import PendingDeprecationWarning
# import warnings

app = Flask(__name__)

# 定义全局变量并初始化为0
visit_count = 0

@app.route('/')
def index():
    global visit_count
    visit_count += 1
    return render_template('index.html')


@app.route('/render', methods=['POST'])
def render():
    # 获取上传的文件
    excel_file = request.files['excel_file']
    word_file = request.files['word_file']

    # 读取 Excel 文件
    df = pd.read_excel(excel_file)

    output_dir = "output/"
    # 渲染 Word 文件
    for record in df.to_dict(orient="records"):
        doc = DocxTemplate(word_file)
        doc.render(record)

        output_path = output_dir + f"{record['filename']}"
        doc.save(output_path)

        # 保存渲染后的 Word 文件
        # document.save('output.docx')

    # 提供下载链接ggg
    return send_file(output_path, as_attachment=True)

@app.route('/doc')
def doc():
    return render_template('doc.html')

@app.route('/about')
def about():
    return render_template('about.html', visit_count=visit_count)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
