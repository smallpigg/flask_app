from flask import Flask, render_template, request, send_file, send_from_directory
import pandas as pd
from docxtpl import DocxTemplate  # pip install docxtpl
import docx
import os
import zipfile
import shutil


def delete_empty_rows(table):
    for row in table.rows:
        first_cell = row.cells[0]
        if first_cell.text == 'nan':
            row._element.getparent().remove(row._element)

app = Flask(__name__)

# 定义全局变量并初始化为0
visit_count = 0
files_count = 0

@app.route('/')
def index():
    global visit_count
    visit_count += 1
    return render_template('index.html')


@app.route('/render', methods=['POST'])
def render():
    global visit_count

    # 获取上传的文件
    excel_file = request.files['excel_file']
    word_file = request.files['word_file']

    # 读取 Excel 文件
    df = pd.read_excel(excel_file, dtype=str)

    for col in df.columns:
        if "日期" in col:
            df[col] = pd.to_datetime(df[col]).dt.date

    home_path = os.getcwd()
    output_dir = "output\\" + str(visit_count) + "\\"

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    # 渲染 Word 文件
    for record in df.to_dict(orient="records"):
        doc = DocxTemplate(word_file)
        doc.render(record)
        output_path = output_dir + f"{record['文件名']}" + ".docx"
        doc.save(output_path)
        global files_count
        files_count += 1

        checkbox_value = request.form.get('delete_blank_rows')
        if checkbox_value:
            doc = docx.Document(output_path)
            for table in doc.tables:
                delete_empty_rows(table)
            doc.save(output_path)

    filenames = os.listdir(output_dir)

    os.chdir(output_dir)
    # 创建压缩文件
    zip_filename = '结果.zip'

    with zipfile.ZipFile(zip_filename, 'w') as zip:
        for filename in filenames:
            zip.write(filename)
    os.chdir(home_path)

    result_dir = output_dir + '结果.zip'

    # 提供下载链接
    return send_file(result_dir, as_attachment=True)

@app.route('/demo')
def demo():
    demo_dir = '例子.zip'
    return send_file(demo_dir, as_attachment=True)

@app.route('/doc')
def doc():
    return render_template('doc.html')

@app.route('/about')
def about():
    global visit_count
    return render_template('about.html', visit_count=visit_count, files_count=files_count)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
