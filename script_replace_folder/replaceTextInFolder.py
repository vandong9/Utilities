import os
import xlrd
import xlwt

# NOTE NOTE NOTE Note:
# BACKUP <BAKUP, BACKUP>  FILE IN SCAN FOLDER  BEFORE RUN THIS COMMAND
# should use file xls despite xlsx, i dont know why but xlsx i check it fail to load :D :D :D
# need install package xlrd, xlwt to read/write excel file by command

# ##   pip3 install xlrd
# ##   pip3 install xlwt

# Configure Input
# Path to folder need to scan file to replace string
rootPath = "/Volumes/DATA/PROJECTS/VIB/MyVIB2/source/myvib-ios-backup/MyVIB_2.0/MVVM/Account/LoanAccount"
inputExcelFile = "keyvalue.xls"  # file excel that contain key-value
outputExcelFile = "report.xls"  # report file name

# file excel to read key-value
workbook = xlrd.open_workbook(inputExcelFile)
sheet = workbook.sheet_by_index(0)  # default load first sheet
totalRow = sheet.nrows
keyColumn = 1  # column index of key
valueColumn = 0  # column index of value
fileExtensionNeedProcess = ".swift" or ".m"

# file report excel
reportWorkbook = xlwt.Workbook()
reportSheet = reportWorkbook.add_sheet(outputExcelFile)

# Variable State
rowIndex = 0
currentKey = ""
currentValue = ""
effectFile = []


# load files in folder and find,  and recurceive with folder
def processFolder(folderPath):
    for fileName in os.listdir(folderPath):
        newPath = os.path.join(folderPath, fileName)

        if os.path.isdir(newPath):
            processFolder(newPath)
        elif fileName.endswith(fileExtensionNeedProcess):
            replaceTextInFile(currentKey, currentValue, newPath)


# Replace key-text in file
def replaceTextInFile(key, text, filePath):
    with open(filePath, "r+") as filePubspec:
        allLines = filePubspec.readlines()
        filePubspec.seek(0)
        haveUpdate = False
        for line in allLines:
            newLine = line.replace(key, text)
            if newLine.count != line.count:
                haveUpdate = True
            filePubspec.write(newLine)
        if haveUpdate == True:
            effectFile.append(filePath)
        filePubspec.truncate()


# START
while rowIndex < totalRow:
    currentKey = sheet.cell_value(rowIndex, keyColumn).strip()
    currentValue = sheet.cell_value(rowIndex, valueColumn).strip()
    processFolder(rootPath)

    reportSheet.write(rowIndex, 0, currentKey)
    reportSheet.write(rowIndex, 1, currentValue)
    reportSheet.write(rowIndex, 2, str(len(effectFile)))
    reportSheet.write(rowIndex, 3, "\r\n".join(effectFile))

    # break
    rowIndex += 1
    effectFile = []

# Save report
reportWorkbook.save(outputExcelFile)
