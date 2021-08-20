from PyQt5 import QtWidgets, uic, QtGui
import connects as ct
import sys


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi("app.ui", self)
        self.setWindowIcon(QtGui.QIcon("images/logo.png"))
        self.excelButton.setIcon(QtGui.QIcon("images/excel.png"))
        self.pdfButton.setIcon(QtGui.QIcon("images/pdf.png"))
        self.renameButton.setIcon(QtGui.QIcon("images/rename.png"))

        ct.connectFile(self, self.excelButton, self.fileLabel)
        ct.connectFolder(self, self.pdfButton, self.folderLabel)
        ct.connectRename(self, self.renameButton)

    def closeEvent(self, event):
        # Confirmation message when the user closes the app
        reply = QtWidgets.QMessageBox()
        reply.setText("Do you really want to close the app?")
        reply.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        reply.setWindowTitle("Close")
        reply.setWindowIcon(QtGui.QIcon("images/logo.png"))
        reply = reply.exec()

        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


app = QtWidgets.QApplication(sys.argv)
window = Ui()
window.show()
app.exec_()
