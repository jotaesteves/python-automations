""" py -m pip install pyautogui pywinauto Pillow"""

from importlib.resources import path
import os
import glob
import numpy as np
from tracemalloc import stop
import pyautogui as pa

import pywinauto as pw
from pywinauto import Application

""" env """
path = "C:/Auto"
mdfWildcard = ".*MDF.*"
list_of_coordinates = [[224,24], [285,70], [20,740], [395,222], [455,705], [465,375], [503,234], [598,411]]

def clickCoord(x, y):
    try:
        pa.moveTo(x, y)
        pa.click()
    except:
        print("Error clicking coordinates")

    pa.sleep(2)

def getMDFWindow():
    global app
    global popup

    try:
        """ get MDF window """
        app = Application(backend="win32").connect(found_index=0, title_re=mdfWildcard, timeout=10)
        popup = app.window(found_index=0, title_re=mdfWildcard)

        popup.type_keys('{RIGHT}{ENTER}') # it calls .set_focus() inside

    except:
        print("MDF not found")
        return None

def insertDataToMDF():
    try:
        for filename in glob.glob(path + "/NABM_TO_UPLOAD_*.txt"):
            with open(filename, 'r') as f:
                for line in f:
                    """ get modal from app """
                    print(line)
                    """ press Enter """
    except:
        print("Error inserting data to MDF")
        print("Missing txt file")

    # pa.press("enter")

def main():
    pw.timings.Timings.after_click_wait = 2
    pw.timings.Timings.after_clickinput_wait = 2
    getMDFWindow()

    for i, coord in enumerate(list_of_coordinates, start=1):
        clickCoord(coord[0], coord[1])
        if i == 4: insertDataToMDF()
        # if i == 8: pa.press("enter")

main()