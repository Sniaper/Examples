from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
import io
import re
from PyPDF2 import PdfFileWriter, PdfFileReader
import os

"""
Программа позволяет делить разделять большие PDF файлы. (персонально для некоммерческой организации)
"""


def pdfparser(data):
    with open(data, 'rb') as fp:

        rsrcmgr = PDFResourceManager()
        retstr = io.StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        for page in PDFPage.get_pages(fp):
            interpreter.process_page(page)
            data = retstr.getvalue()

    # print(data)
    list_page = data.split('\n')
    for i, fi in enumerate(list_page):
        if fi == "Лицевой счет":
            print(data.split('\n')[i+2])
            return data.split('\n')[i+2]



def cut_fail(name):
    input_PDF = PdfFileReader(open(name, 'rb'))

    for i in range(input_PDF.getNumPages()):
        output = PdfFileWriter()
        new_File_PDF = input_PDF.getPage(i)
        output.addPage(new_File_PDF)
        output_Name_File = 'SomePDF_' + str(i + 1) + ".pdf"
        outputStream = open(output_Name_File, 'wb')
        output.write(outputStream)
        outputStream.close()

        os.rename(output_Name_File, pdfparser(output_Name_File) + '.pdf')


if __name__ == '__main__':
    cut_fail(r'июль 2021.pdf')
