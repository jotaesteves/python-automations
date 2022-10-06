""" py -m pip install pyautogui pywinauto """

import os
import glob
import numpy as np
from tracemalloc import stop
from pyautogui import moveTo, center, locateOnScreen, click, press

import pywinauto
from pywinauto import Application
from pywinauto import findwindows
import win32gui

""" env """
path = "C:/Auto/mdf-images"
mdfWildcard = ".*MDF.*"

def getImages():
    global list_of_images

    os.chdir(path)
    list_of_images = glob.glob("*.png")
    list_of_images = [x for x in list_of_images if not x.startswith('~$')]
    list_of_images = np.sort(list_of_images)

def clickImage(image):
    moveTo(center(locateOnScreen(image)))
    click()

def getMDFWindow():
    global app
    try:
        app = Application(backend="uia").connect(title_re=mdfWildcard).top_window()
        app.set_focus()

        win = app.window(title_re=mdfWildcard)
        print(win)

        """ hwnd = win32gui.FindWindow(None, mdfWildcard)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.ShowWindow(hwnd, 9) """

    except:
        print("MDF not found")
        return None

def insertDataToMDF():
    """ insert data to MDF """
    stop()
    with open("NABM_TO_UPLOAD_.*.txt", "r") as f:
        for line in f:
            """ get modal from app """
            print(line)
            """ press Enter """
    press("enter")

def main():
    pywinauto.timings.Timings.after_clickinput_wait = 2
    getMDFWindow()
    getImages()

    for (file, i) in set(list_of_images):
        clickImage(file)
        print(i)
        print(app)

        """ on step 4 write all nabm from txt file """
        if i == 4: insertDataToMDF()

        """ img 5 click each nabm row, fill in the data, press Enter """
        """ click arrow up upload btn """
        """ click green upload button """
        """ select favorites > auto folder """

main()
