import PyPDF2
import random
import string

import os, sys
import win32com.client
import time
import hashlib

def pwd_xlsx(old_filename,new_filename,pwd_str,pw_str=''):
    xcl = win32com.client.Dispatch("Excel.Application")
    # pw_str Open password for , If there is no Access password , Set to ''
    wb = xcl.Workbooks.Open(old_filename, False, False, None, pw_str)
    xcl.DisplayAlerts = False
    # When saving, you can set the access password .
    wb.SaveAs(new_filename, None, pwd_str, '')
    xcl.Quit()

pwd_str = '654321'# New password customization
# pwd_xlsx('f:\\Temp\\1.xlsx','f:\\Temp\\2.xlsx',pwd_str)


def pwd_docx(file,new_filename,pwd_str):
    word =win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(file, False, True, None)
    doc.SaveAs('f:\\Temp\\3.docx',12,None,'123456')
    doc.Quit()
    word.Quit()

file = 'f:\\Temp\\1.docx'
new_filename = 'f:\\Temp\\2.docx'
pwd_docx(file,new_filename,pwd_str)

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


def encrypt_pdf(pdfile):
    pdf_in_file = open(pdfile, 'rb')
    inputpdf = PyPDF2.PdfFileReader(pdf_in_file)
    pages_no = inputpdf.numPages
    output = PyPDF2.PdfFileWriter()
    Newname = pdfile.replace(pdfile.split('.')[-1] ,'') + '_password_protected.pdf'
    Temppass = generate_random_password() #call generate_random_password
    print(Temppass)

    for i in range(pages_no):
        inputpdf = PyPDF2.PdfFileReader(pdf_in_file)

        output.addPage(inputpdf.getPage(i))
        output.encrypt(Temppass) #set the password

        # with open("simple_password_protected.pdf", "wb") as outputStream:
        with open(Newname, "wb") as outputStream:
            output.write(outputStream)

    pdf_in_file.close()

# pdfile = 'F:\\Temp\\101-000101660_101.pdf'
# encrypt_pdf(pdfile)





