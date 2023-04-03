import zipfile
import os


output_dir = "output/"

filenames = os.listdir(output_dir)

os.chdir(output_dir)
# 创建压缩文件
zip_filename = '例子.zip'

with zipfile.ZipFile(zip_filename, 'w') as zip:
    for filename in filenames:
        zip.write(filename)


# return send_file(zip_filename, as_attachment=True)
