# split pdfs in to single pages, pdf's name from excel
# 需要excel中名字行数与对应的pdf页数相同
import os
import pandas as pd
from PyPDF2 import PdfFileReader, PdfFileWriter
from pip import main 
# import openpyxl
try:
    import openpyxl
except:
     main(['install', 'openpyxl'])
def pdf_splitter(path, dest):
    '''path:原pdf文件完整路径;dest:另存pdf完整路径'''
    pdf = PdfFileReader(path)  # 初始化对象
    for page in range(pdf.getNumPages()): # get page number
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        filename = str(name_df['name1'][page]) + '.pdf'
        output_filename = os.path.join(dest, filename) 
        print(output_filename)
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
        print('Created: {}'.format(output_filename))
        
# 检查模块是作为程序运行还是被导入另一个程序
if __name__ == '__main__':
    wd = 'E:/GitHub/python/temp/' # working directory
    fileList = os.listdir(wd + "pdf") #获取该目录下所有文件，存入列表中
    fileList = [word.replace('.pdf','') for word in fileList] #文件列表，去掉后缀名.pdf
    for n in fileList:
        name_df = pd.read_excel(wd + "excel/" + n + '.xlsx', sheet_name = "Sheet1", engine='openpyxl')  #读取pdf同名的excel
        path = wd + 'pdf/' + n + '.pdf'  
        if not os.path.exists(wd + 'pdf_files'):
            os.mkdir(wd + 'pdf_files')
        if not os.path.exists(wd + 'pdf_files/' + n):
            os.mkdir(wd + 'pdf_files/' + n)  # 新建pdf命名的文件夹
        dest = wd + 'pdf_files/' + n 
        pdf_splitter(path, dest)
        print(n + ':' + "拆分完成！")
    print("All PDF Files Have Splited Successfully!")