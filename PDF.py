# =========================================================
# PDF文件截取某几页另存新文件
# =========================================================
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
def pdf_split(pdf_in,pdf_out,start,end):
    '''定义函数:将PDF文件截取某几页,另存到新pdf文件'''
    # 初始化一个pdf
    output = PdfFileWriter()
    # 读取pdf
    with open(pdf_in, 'rb') as in_pdf:
        pdf_file = PdfFileReader(in_pdf)
        # 从pdf中取出指定页
        for i in range(start, end):
            output.addPage(pdf_file.getPage(i))
        # 写出pdf
        with open(pdf_out, 'ab') as out_pdf:
            output.write(out_pdf)
root = 'E:/personal_zhy/github/python/666'  # 原pdf文件所在的绝对路径
inpath = root +  '/' +"pdf"  # 原pdf所在路径
outpath = root +  '/' +"pdf_files" # 导出的目标路径
pdfList = os.listdir(inpath)  # 需拆分的pdf列表
for pdfname in pdfList:  # 循环批量将各pdf文件中某几页截取后另存新文件
    pdf_in = inpath + '/' + pdfname
    pdf_out = outpath + '/' + pdfname
    s, e = 0, 1  # 拆分的起始位置和结束位置
    pdf_split(pdf_in, pdf_out, s, e) # 拆分函数

# =========================================================
# PDF 拆分成单页
# split pdfs in to single pages, pdf's name from excel
# 需要excel中名字行数与对应的pdf页数相同
# =========================================================
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

# =========================================================
# PDF 合并成一个文件
# =========================================================
import os
from PyPDF2 import PdfFileMerger
root = 'E:/GitHub/python/temp/'
path = root + 'pdf'  # 需要合并的单个pdf所在路径
out_path = root + 'pdf_files/' # 合并后存放路径
pdfs = os.listdir(path)

pdfs_idx = [i.split('.')[0] for i in pdfs] # pdfs排序
pdfs_sorted = []
for i in sorted(pdfs_idx):
    pdfs_sorted.append(str(i) + '.pdf')

merger = PdfFileMerger()  # 初始化对象
for pdf in pdfs_sorted:
    merger.append(path + '/' + pdf, import_bookmarks=False)
merger.write(out_path + 'append.pdf')
merger.close()
print("完成合并")

