import PyPDF2
import random
import string
import win32com.client
import os

def generate_random_password():
    characters = list(string.ascii_letters + string.digits + "!@#$%^&*()")
    length = 12
    ## shuffling the characters
    random.shuffle(characters)     # shuffling the characters
    password = []
    for i in range(length):
        password.append(random.choice(characters)) # picking random characters from the list
    # shuffling the resultant password
    random.shuffle(password) # shuffling the resultant password
    return ("".join(password))

def pwd_xlsx(file,new_filename,pwd_str):
    file = os.path.abspath(file)  # convert path to absolute path
    new_filename = os.path.normpath(new_filename)

    xcl = win32com.client.Dispatch("Excel.Application")
    # pw_str Open password for , If there is no Access password , Set to ''
    wb = xcl.Workbooks.Open(file, False, False, None)
    xcl.DisplayAlerts = False
    # When saving, you can set the access password .
    wb.SaveAs(new_filename, 51, pwd_str, '') # XlFileFormat enumeration is 51 to xsls
    xcl.Quit()

# pwd_str = '654321'# New password customization
# pwd_xlsx('f:\\Temp\\1.xlsx','f:\\Temp\\2.xlsx',pwd_str)


def pwd_docx(file,new_filename,pwd_str):
    file = os.path.abspath(file) # convert path to absolute path
    new_filename = os.path.normpath(new_filename)

    word =win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(file, False, True, None)
    doc.SaveAs(new_filename,12,None,pwd_str)
    doc.Close()
    word.Quit()

# file = 'f:\\Temp\\טופס פרטי לקוח .doc'
# new_filename = 'f:\\Temp\\טופס פרטי לקוחי.doc'
# pwd_docx(file,new_filename,generate_random_password())




def encrypt_pdf(file,new_filename,pwd_str):
    file = os.path.abspath(file)  # convert path to absolute path
    new_filename = os.path.normpath(new_filename)

    pdf_in_file = open(file, 'rb')
    inputpdf = PyPDF2.PdfFileReader(pdf_in_file)
    pages_no = inputpdf.numPages
    output = PyPDF2.PdfFileWriter()
    for i in range(pages_no):
        inputpdf = PyPDF2.PdfFileReader(pdf_in_file)
        output.addPage(inputpdf.getPage(i))
        output.encrypt(pwd_str) #set the password
        with open(new_filename, "wb") as outputStream:
            output.write(outputStream)
    pdf_in_file.close()

# pdfile = 'F:\\Temp\\101-000101660_101.pdf'
# encrypt_pdf(pdfile)





