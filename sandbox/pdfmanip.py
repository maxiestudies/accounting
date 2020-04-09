from PyPDF2 import PdfFileReader, PdfFileWriter
import pdb

writer = PdfFileWriter()
srcfile = open('pdftester.pdf', "rb")
srcreader = PdfFileReader(srcfile)

numpages = srcreader.getNumPages()
page = srcreader.getPage(numpages - 1)
writer.addPage(page)

pdb.set_trace()
with open('cropped.pdf', 'wb') as output_pdf:
    writer.write(output_pdf)

srcfile.close()
