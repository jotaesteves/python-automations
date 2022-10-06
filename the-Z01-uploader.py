""" py -m pip install openpyxl lxml pywin32"""

""" imports """

import os
import datetime
import numpy as np
from extendedopenpyxl import load_workbook, save_workbook

""" env variables """

path = "C:/Auto"
sheetName = "2&3 - Signature Sheet"
basisOnly = "BASIS_ONLY"
bezeichnung = "Bezeichnung"
bestellpositionen = "Dieses Abnahmeprotokoll umfasst folgende Bestellpositionen"
summaryOfAddOnPOs = "Summary of add on POs"
list_basis_only = []
list_not_basis_only = []
list_po_not_found = []
list_of_nabm = []
list_of_errors = []

""" set timestamp """

def setTimer():
    global now
    global timestamp
    now = datetime.datetime.now()

    timestamp = now.strftime("%Y-%m-%d__%H-%M-%S")
    print(timestamp)


""" setup """

def setup():
    import warnings
    import glob

    warnings.filterwarnings('ignore')
    os.chdir(path)
    global list_of_files_xlsm
    global list_of_files_xlsx

    list_of_files_xlsm = glob.glob("*.xlsm")
    list_of_files_xlsx = glob.glob("*.xlsx")

    """ ignore temp files """
    list_of_files_xlsm = [
        x for x in list_of_files_xlsm if not x.startswith('~$')]


""" loop trough files in folder """

def runner():
    global list_basis_only
    global list_not_basis_only
    global filePath
    global wb_obj
    global sheet_obj
    global nabm
    global file
    global filePath

    print("----------------------------------------")
    print("Working on:")
    for file in set(list_of_files_xlsm):
        print(file)
        checkFileOpen()

        """ open file """
        filePath = path + "/" + file
        wb_obj = load_workbook(file, read_only=False,
                               keep_vba=True, keep_links=False)
        sheet_obj = wb_obj[sheetName]

        columnA = np.array([x.value for x in sheet_obj['A']])
        """ find index of "bestellpositionen" """
        index = np.where(columnA == bestellpositionen)
        startingRow = index[0][0] + 3 # +1 to get row of bestellpositionen, +2 to get row with first value
        nabm = sheet_obj.cell(row=startingRow, column=10).value
        """ check if nabm is empty or null or not a string """
        if (nabm == None or nabm == "" or not isinstance(nabm, str)):
            list_of_errors.append(file)
            moveFile('/NOT FOUND')
            continue

        """ check if basisOnlyValue is empty or null or not a string """
        basisOnlyValue = sheet_obj.cell(row=startingRow, column=4).value
        if (basisOnlyValue == None or basisOnlyValue == "" or not isinstance(basisOnlyValue, str)):
            list_of_errors.append(file)
            moveFile('/NOT FOUND')
            continue

        isBasisOnly = basisOnlyValue == basisOnly

        if (isBasisOnly):
            list_basis_only.append(file)
            list_of_nabm.append(nabm)

        else:
            list_not_basis_only.append(file)
            editFilesSOAOP()


""" edit files """

def editFilesSOAOP():
    global nabm
    global soaop_obj

    """ open "Summary of add on POs.xlsm" from list_of_files_xlsm """
    soaop = [x for x in list_of_files_xlsx if x.startswith(summaryOfAddOnPOs)]
    soaop_file = path + "/" + soaop[0]
    soaop_obj = load_workbook(
        soaop_file, read_only=False, keep_vba=True, keep_links=False)
    soaop_sheet_obj = soaop_obj.active

    """ read column B and N """
    columnB = np.array([x.value for x in soaop_sheet_obj['B']])
    columnN = soaop_sheet_obj['N']

    i = 0
    j = 10

    indexes = np.where(columnB == nabm)
    """ add +1 to each value of indexes """
    indexes = [x for x in indexes[0]]

    """ check if set is empty """
    isEmpty = (len(indexes) == 0)

    if (isEmpty):
        list_po_not_found.append(file)
        moveFile('/PO WAITING')
        return

    """ only append nabm to files that are edited or basis only """
    list_of_nabm.append(nabm)

    for rowNr in indexes:
        """ get column N from cell """
        tempValue = columnN[rowNr].value
        """ check if cell D23 of sheet_obj is not empty """
        if (sheet_obj.cell(row=23 + i, column=4).value != None):
            """ add value to cell B23 of sheet_obj """
            sheet_obj.cell(row=23 + i, column=2).value = tempValue
            sheet_obj.cell(row=23 + i, column=3).value = j
            i = i + 1
            j = j + 10

    """ print to see data changed """
    """  for i in range(23, 28):
            print(str(sheet_obj.cell(row = i, column = 2).value) + "|" + str(sheet_obj.cell(row = i, column = 3).value) + "|" + str(sheet_obj.cell(row = i, column = 4).value))
    """

    """ save file """
    wb_obj.close()
    """ soaop_obj.close() """
    save_workbook(wb_obj, filePath)


""" move file if PO not found """

def moveFile(folderName="/ERRORS"):
    import shutil
    wb_obj.close()
    """ soaop_obj.close() """
    """ check if folder exists """
    if not os.path.exists(path + folderName):
        os.makedirs(path + folderName)
    shutil.move(filePath, path + folderName + "/" + file)


""" close program if file is open """

def checkFileOpen():
    if (os.path.isfile(path + "/~$" + file)):
        from win32com.client import Dispatch

        xl = Dispatch('Excel.Application')
        xl.Workbooks(file).Close(SaveChanges=True)

def makeFile():
    """ prints all list_of_nabm into NABM_TO_UPLOAD_DD_MM_YY_HH:MM:SS.txt file """
    global list_of_nabm
    global path

    list_of_nabm = list(set(list_of_nabm))
    list_of_nabm.sort()

    """ create file """
    file = open(path + "/NABM_TO_UPLOAD_" + timestamp + ".txt", "w+")

    """ write to file """
    for nabm in list_of_nabm:
        file.write(nabm + "\n")

    """ close file """
    file.close()

""" prints some relevant info """

def analytics():
    print('----------------------------------------')
    print(basisOnly + ": [" + str(len(list_basis_only)) +
          "]  " + str(list_basis_only))
    print("----------------------------------------")
    print("Edited Files: [" + str(len(list_not_basis_only)
                                  ) + "]  " + str(list_not_basis_only))
    if (len(list_po_not_found) > 0):
        print("----------------------------------------")
        print("PO NOT FOUND: [" + str(len(list_po_not_found)
                                      ) + "]  " + str(list_po_not_found))
    if (len(list_of_errors) > 0):
        print("----------------------------------------")
        print("NABM / Bezeichnung Missing: [" + str(len(list_of_errors)
                                      ) + "]  " + str(list_of_errors))


""" stop timer """

def stopTimer():
    end = datetime.datetime.now()
    difference = end - now
    difference = difference - \
        datetime.timedelta(microseconds=difference.microseconds)

    print("***************************************")
    print("Time Elapsed: " + str(difference))


""" end """

setTimer()
setup()
runner()
makeFile()
stopTimer()
analytics()
