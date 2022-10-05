""" Install the package. Run Command: """
""" py -m pip install openpyxl """
""" py -m pip install lxml """
""" py -m pip install pywin32  """

""" env variables """
from extendedopenpyxl import load_workbook, save_workbook
import datetime
import os
path = "C:/Auto"
sheetName = "2&3 - Signature Sheet"
basisOnly = "BASIS_ONLY"
summaryOfAddOnPOs = "Summary of add on POs"
list_basis_only = []
list_not_basis_only = []
list_po_not_found = []

""" imports """

""" set timestamp """


def setTimer():
    global now
    now = datetime.datetime.now()

    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
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
        """ read cell D23 """
        cell_bezeichnung = sheet_obj.cell(row=23, column=4)
        """ read cell J23 """
        cell_nabm = sheet_obj.cell(row=23, column=10)
        nabm = cell_nabm.value

        isBasisOnly = cell_bezeichnung.value == basisOnly

        if (isBasisOnly):
            list_basis_only.append(file)

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

    """ read column B """
    columnB = soaop_sheet_obj['B']
    columnN = soaop_sheet_obj['N']
    i = 0
    j = 10
    rowsThatMatchSet = set(i for i, x in enumerate(columnB) if x.value == nabm)

    """ check if set is empty """
    isEmpty = (len(rowsThatMatchSet) == 0)

    if (isEmpty):
        list_po_not_found.append(nabm)
        moveFile()
        return

    for cellNr in rowsThatMatchSet:
        """ get column N from cell """
        tempValue = columnN[cellNr].value

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
    soaop_obj.close()
    save_workbook(wb_obj, filePath)


""" move file if PO not found """


def moveFile():
    import shutil
    wb_obj.close()
    soaop_obj.close()
    """ check if folder exists """
    if not os.path.exists(path + "/PO WAITING"):
        os.makedirs(path + "/PO WAITING")
    shutil.move(filePath, path + "/PO WAITING/" + file)


""" close program if file is open """


def checkFileOpen():
    if (os.path.isfile(path + "/~$" + file)):
        from win32com.client import Dispatch

        xl = Dispatch('Excel.Application')
        xl.Workbooks(file).Close(SaveChanges=True)


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
stopTimer()
analytics()
