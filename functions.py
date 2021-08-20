from PyQt5.QtWidgets import QFileDialog, QMessageBox
import os
import xlrd


def browseFiles(self, label):
    fname = QFileDialog.getOpenFileName(
        self, "Open file", "../", "Excel Files (*.xlsx *.xls)"
    )
    filePath = fname[0]
    extensionsToCheck = (".xlsx", ".xls")
    if filePath.endswith(extensionsToCheck):
        setText(label, filePath)


def browseFolders(self, label):
    folderPath = QFileDialog.getExistingDirectory(self, "Open folder", "../")
    if folderPath != "":
        setText(label, folderPath)
    return


def setText(label, text):
    label.setText(text)


def rename(self):
    excelPath = getText(self.fileLabel)
    folderPath = getText(self.folderLabel)
    filesInExcel = extractExcel(self, excelPath)
    if filesInExcel is None:
        return

    try:
        folderIterator = os.scandir(folderPath)
    except:
        errorMssg(self, "No folder selected")
        return

    filesChanged = False
    filesCount = 0
    for file in folderIterator:
        if file.path.endswith(".pdf") and file.is_file():
            pdfNumber = extractNumber(file.name)
            filePath = file.path
            if pdfNumber in filesInExcel and pdfNumber != -1:
                newFilePath = createPath(folderPath, filesInExcel[pdfNumber], ".pdf")
                if filePath != newFilePath:
                    filesCount += 1
                    filesChanged = True
                    os.rename(filePath, newFilePath)
    if filesChanged:
        setText(self.loadingLabel, str(filesCount) + " File(s) Renamed Successfully!!")
    else:
        setText(self.loadingLabel, "No Files Changed")


def getText(label):
    return label.text()


def extractExcel(self, path):
    files = {}
    try:
        workbook = xlrd.open_workbook(path)
    except:
        errorMssg(self, "No file selected")
        return
    sheet = workbook.sheet_by_index(0)
    sheet.cell_value(0, 0)
    for i in range(sheet.nrows):
        files[extractNumber(str(sheet.cell_value(i, 0)))] = sheet.cell_value(i, 0)

    return files


def createPath(directory, basename, extension):
    return os.path.join(directory, basename + extension)


def extractNumber(text):
    number = ""
    for i in range(len(text)):
        if text[i].isnumeric():
            number += text[i]
        else:
            break
    if number == "":
        return -1
    return int(number)


def errorMssg(self, txt):
    QMessageBox.critical(self, "Error", txt)
