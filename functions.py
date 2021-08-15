from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem
import os
import xlrd

def browseFiles(self, label):
    fname = QFileDialog.getOpenFileName(self, "Open file", "../", "Excel Files (*.xlsx *.xlsm *.xlsb)")
    filePath = fname[0]
    extensionsToCheck = (".xlsx", ".xlsm","xlsb")
    if filePath.endswith(extensionsToCheck):
        label.setText(filePath)
    elif filePath != "":
        errorMssg(
            self, "Invalid format. Please select an excel file."
        )
        return

def browseFolders(self, label):
    folderPath = QFileDialog.getExistingDirectory(self, "Open folder", "../")
    if folderPath!="":
        label.setText(folderPath)
    return

def extractExcel(path):
    files = {}
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    sheet.cell_value(0,0)
    for i in range(sheet.nrows):
        files[extractNumber(sheet.cell_value(i,0))] = sheet.cell_value(i,0)

    return files

def getText(label):
    return label.text()

def rename(self):
    filePath = getText(self.fileLabel)
    folderPath = getText(self.folderLabel)
    self.loadingLabel.setText("Loading..")
    filesInExcel = extractExcel(filePath)
    filesChanged = False
    
    for entry in os.scandir(folderPath):
        if entry.path.endswith(".pdf") and entry.is_file():
            pdfNumber = extractNumber(entry.name)
            if pdfNumber in filesInExcel and pdfNumber!=-1:
                filesChanged = True
                os.rename(entry.path,os.path.join(folderPath, filesInExcel[pdfNumber] + ".pdf"))
                
    if filesChanged:
        self.loadingLabel.setText("Files Renamed Successfully!!")
    else:
        self.loadingLabel.setText("No Files Changed")

def extractNumber(text):
    number = ""
    for i in range(len(text)):
        if(text[i].isnumeric()):
            number+=text[i]
        else:
            break
    if number=="":
        return -1
    return int(number)

def errorMssg(self, txt):
    QMessageBox.critical(self, "Error", txt)