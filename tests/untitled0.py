# -*- coding: utf-8 -*-
"""
Created on Sun Jun  1 14:06:58 2025

@author: user
"""

# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'designerCQJxIA.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

#from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import os
import json
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QLineEdit, QTabWidget, QCheckBox, QProgressBar,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView
)
from PyQt5.QtCore import Qt, QTimer

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if not Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(542, 466)
        self.buttonBox = QDialogButtonBox(Dialog)
        self.buttonBox.setObjectName(u"buttonBox")
        self.buttonBox.setGeometry(QRect(221, 360, 341, 32))
        self.buttonBox.setOrientation(Qt.Horizontal)
        self.buttonBox.setStandardButtons(QDialogButtonBox.Cancel|QDialogButtonBox.Ok)
        self.pushButton = QPushButton(Dialog)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(91, 360, 93, 28))
        self.columnView = QColumnView(Dialog)
        self.columnView.setObjectName(u"columnView")
        self.columnView.setGeometry(QRect(20, 70, 256, 192))
        self.horizontalSlider = QSlider(Dialog)
        self.horizontalSlider.setObjectName(u"horizontalSlider")
        self.horizontalSlider.setGeometry(QRect(130, 290, 211, 22))
        self.horizontalSlider.setOrientation(Qt.Horizontal)
        self.lcdNumber = QLCDNumber(Dialog)
        self.lcdNumber.setObjectName(u"lcdNumber")
        self.lcdNumber.setGeometry(QRect(411, 150, 64, 23))

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)

        QMetaObject.connectSlotsByName(Dialog)
    # setupUi

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"Dialog", None))
        self.pushButton.setText(QCoreApplication.translate("Dialog", u"PushButton", None))
    # retranslateUi

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Ui_Dialog()
    window.show()
    sys.exit(app.exec())

