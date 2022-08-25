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
    wb.Close()
    xcl.Quit()
    return pwd_str


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
    return pwd_str

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
    return pwd_str

def send_mail(Subject, Message, address,attachment=None ):
    const = win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = Subject
    # newMail.Body = Message
    newMail.BodyFormat = 2  # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    newMail.HTMLBody = f"<HTML><BODY>{Message}</BODY></HTML>"
    newMail.To = address
    if attachment:
        attachment1 = os.path.normpath(attachment)
        newMail.Attachments.Add(Source=attachment1)
        newMail.display(False)
    newMail.Send()

def mail_send(self, file,password):
    self.label.setText('-- The file is encrypted!')
    # self.lineEdit.setText(nname.replace('/','\\'))
    self.lineEdit_2.setText(password)
    self.lineEdit_2.show()
    send_mail('Encrypted file from the Samelet company',  # send the file only
                   'The file is attached to this email, the password will be sent in a separate email',
                   self.lineEdit.text(), os.path.normpath(file))
    send_mail('Encrypted file from the Samelet company',
                   f'The password is: {password}',
                   self.lineEdit.text())
    self.lineEdit.setText(os.path.normpath(file)) #write the file name



