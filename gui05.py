# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\untitled2.ui'
#
# Created by: PyQt5 UI code generator 5.15.5
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

## pyinstaller --onefile --noconsole gui05.py

from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QFileDialog
import main
import os
import re


class Ui_GroupBox(object):
    def setupUi(self, GroupBox):
        GroupBox.setObjectName("GroupBox")
        GroupBox.resize(490, 125)
        GroupBox.setAutoFillBackground(True)
        self.lineEdit = QtWidgets.QLineEdit(GroupBox)
        self.lineEdit.setGeometry(QtCore.QRect(10, 10, 281, 41))
        self.lineEdit.setAutoFillBackground(False)
        self.lineEdit.setText("")
        self.lineEdit.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(GroupBox)
        self.label.setGeometry(QtCore.QRect(310, 20, 171, 31))
        self.label.setObjectName("label")
        self.pushButton = QtWidgets.QPushButton(GroupBox)
        self.pushButton.setGeometry(QtCore.QRect(260, 80, 101, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(GroupBox)
        self.pushButton_2.setGeometry(QtCore.QRect(370, 80, 101, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(GroupBox)
        self.lineEdit_2.setGeometry(QtCore.QRect(10, 80, 145, 31))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_2.hide()

        self.retranslateUi(GroupBox)
        QtCore.QMetaObject.connectSlotsByName(GroupBox)

        self.pushButton.clicked.connect(self.loadCsv)
        self.pushButton_2.clicked.connect(self.run)
        self.fname = None

    def retranslateUi(self, GroupBox):
        _translate = QtCore.QCoreApplication.translate
        GroupBox.setWindowTitle(_translate("GroupBox", "Encrypter"))
        self.label.setText(_translate("GroupBox", " >> Recipient mail"))
        self.pushButton.setText(_translate("GroupBox", "Select a file"))
        self.pushButton_2.setText(_translate("GroupBox", "Run"))



###################################################################################

    def loadCsv(self):
        self.fname = QFileDialog.getOpenFileName(None, "Window name","","xlsx(*.xlsx *.csv *.docx *.doc *.pdf)")

    def run(self):
        print("Running")
        regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' # to validate an Email
        if (re.fullmatch(regex, str(self.lineEdit.text()))):
            try:
                if self.fname[0]:
                    password = None
                    nname = self.fname[0].replace(self.fname[0].split('.')[-2].split('/')[-1] ,self.fname[0].split('.')[-2].split('/')[-1] + '_password_protected') # rename to _password_protected
                    npass = main.generate_random_password() # generate_random_password
                    if self.fname[0].split('/')[-1].split('.')[-1].upper() == 'docx'.upper():
                        password = main.pwd_docx(self.fname[0], nname, npass)
                        if password:
                            main.mail_send(self, os.path.normpath(nname), password)
                    elif self.fname[0].split('/')[-1].split('.')[-1].upper() == 'doc'.upper():
                        password = main.pwd_docx(self.fname[0], nname, npass)
                        if password:
                            main.mail_send(self, os.path.normpath(nname), password)
                    elif self.fname[0].split('/')[-1].split('.')[-1].upper() == 'xlsx'.upper():
                        password =  main.pwd_xlsx(self.fname[0], nname, npass)
                        if password:
                            main.mail_send(self, os.path.normpath(nname), password)
                    elif self.fname[0].split('/')[-1].split('.')[-1].upper() == 'csv'.upper():
                        nname = nname.replace(self.fname[0].split('/')[-1].split('.')[-1], 'xlsx') # CSV files cannot be encrypted and must be converted
                        password =  main.pwd_xlsx(self.fname[0], nname, npass)
                        if password:
                            main.mail_send(self, os.path.normpath(nname), password)
                    elif self.fname[0].split('/')[-1].split('.')[-1].upper() == 'pdf'.upper():
                        password =  main.encrypt_pdf(self.fname[0], nname, npass)
                        if password:
                            main.mail_send(self, os.path.normpath(nname), password)
                    else:
                        self.label.setText('--Error !!')
            except:
             self.label.setText('-- File Error !!')
        else:
            self.label.setText('--Maill Error !!')



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    GroupBox = QtWidgets.QGroupBox()
    ui = Ui_GroupBox()
    ui.setupUi(GroupBox)
    GroupBox.show()
    sys.exit(app.exec_())