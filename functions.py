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

def extractExcel(self,path):
    files = {}
    try:
        workbook = xlrd.open_workbook(path)
    except:
        errorMssg(
            self, "No file selected"
        )
        return
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
    filesInExcel = extractExcel(self,filePath)
    if filesInExcel is None:
        return
        
    try:
        scanItr = os.scandir(folderPath)
    except:
        errorMssg(
            self, "No folder selected"
        )
        return

    self.loadingLabel.setText("Loading..")
    filesChanged = False
    for file in scanItr:
        if file.path.endswith(".pdf") and file.is_file():
            pdfNumber = extractNumber(file.name)
            if pdfNumber in filesInExcel and pdfNumber!=-1:
                filesChanged = True
                os.rename(file.path,os.path.join(folderPath, filesInExcel[pdfNumber] + ".pdf"))
                
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