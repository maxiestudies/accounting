import os, datetime, pyoo, argparse, subprocess
from PyPDF2 import PdfFileReader, PdfFileWriter

# # a function to initialize libreoffice with an open socket
def start_loffice():
subprocess.run("/Applications/LibreOffice.app/Contents/MacOS/soffice --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\" &", shell=True)

def get_inputs():
    """"gets the command line paramters and handles help to get the comman line parameters"""""

    parser = argparse.ArgumentParser()
    parser.add_argument("template",
                        choices=['klangregie', 'mullermusic', 'musikfabrik',
                                 'veranstaltungstechnik_1', 'veranstaltungstechnik_2',
                                 'otros+mwst'],
                        help="specify the invoice type")

    parser.add_argument("-t", help="specify one of the templates", choices=['decoder','on'])
    args = parser.parse_args()
    return vars(args)

def handle_inputs(args):
    """takes the command line inputs and matches them to open the right file"""
    category = args['template']
    template = args['t']
    
    if category  == 'klangregie':
        file_to_open = 'Maximiliano_Estudies-Rechnung_Klangregie-2020.ods'
    elif category == 'musikfabrik':
        file_to_open = 'Maximiliano_Estudies-Rechnung_Musikfabrik-2020.ods'
    elif category == 'otros+mwst':
        file_to_open = 'Maximiliano_Estudies-Otros+MwSt-2020_copy.ods'
    else:
        file_to_open = "no_file"

    return file_to_open

        
def pdf_cropper(filename):

    writer = PdfFileWriter()
    srcfile = open(filename, "rb")
    srcreader = PdfFileReader(srcfile)

    numpages = srcreader.getNumPages()
    page = srcreader.getPage(numpages - 1)
    writer.addPage(page)

    output_name = invoices_path + saving_name 
    with open(output_name, 'wb') as output_pdf:
        writer.write(output_pdf)

    srcfile.close()

invoices_path = "/Users/maxiestudies/Documents/Trabajo/Facturas/2020/" 
#start_loffice()
arguments = get_inputs()
file_to_open = handle_inputs(arguments)


#open connection to document
desktop = pyoo.Desktop('localhost',2002)

# open the document and store the handle in variable
doc = desktop.open_spreadsheet("/Users/maxiestudies/Documents/Trabajo/ingresos/" + file_to_open)

doc.sheets.copy('Base', 'new sheet')

active_sheet = doc.sheets['new sheet']

# get the filenames in the rechnung dir and count them
files = []
with os.scandir("/Users/maxiestudies/Documents/Trabajo/Facturas/2020") as it:
    for entry in it:
        if not entry.name.startswith('.') and entry.is_file():
            files.append(entry.name)


# extract the number of the last rechnung
lastfilenr = int(files[-1][0:2]) + 1

#get the day of today and format it
today = str(datetime.datetime.now())
year = today[0:4]
month = today[5:7]
write_out = "Rechnung-" + month + "." + year + "-" + str(lastfilenr)

# write it out to the sheet
active_sheet[19,0].formula = write_out

#wait until the file is finished and export it as pdf
go = input("Tell me when to print\n")
if go == 'print':
    saving_name = str(lastfilenr) + ".Maximiliano_Estudies_Rechnung-" + write_out + ".pdf"
    result_saved = doc.save(saving_name, pyoo.FILTER_CALC_PDF_EXPORT)

doc.close()

pdf_cropper(saving_name)
