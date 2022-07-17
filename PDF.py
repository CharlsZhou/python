# PDF文件截取某几页另存新文件
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

