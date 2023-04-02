import zipfile, os


output_dir = "output/"
filenames = os.listdir(output_dir)

# print(filenames)

# 创建压缩文件
zip_filename = output_dir + 'files.zip'

with zipfile.ZipFile(zip_filename, 'w') as zip:
    for filename in filenames:
        #print(filename)
        zip.write(output_dir+filename)
    # 提供给用户下载


# return send_file(zip_filename, as_attachment=True)
