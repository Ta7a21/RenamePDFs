from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtCore import QSettings
import functions as tools
import connects as ct
import sys
import os


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi("app.ui", self)
        self.setWindowIcon(QtGui.QIcon("logo.png"))

        ct.connectFile(self, self.excelButton, self.fileLabel)
        ct.connectFolder(self, self.pdfButton, self.folderLabel)
        ct.connectRename(self, self.renameButton)


app = QtWidgets.QApplication(sys.argv)
window = Ui()
window.show()
app.exec_()
