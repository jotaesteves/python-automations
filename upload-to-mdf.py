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
imgPath = path + "/mdf-images"
mdfWildcard = ".*MDF.*"

def getImages():
    global list_of_images

    os.chdir(imgPath)
    list_of_images = glob.glob("*.jpeg")
    list_of_images = [x for x in list_of_images if not x.startswith('~$')]
    list_of_images = np.sort(list_of_images)

def clickImage(image):
    print(image)
    try:
        box = pa.locateOnScreen(image)
        point = pa.center(box)
        print("Image found")
        pa.click(point)
    except:
        print("Image not found")

    """ wait for modal to appear """
    pa.sleep(2)


def getMDFWindow():
    global app
    global popup

    try:
        """ get MDF window """
        """ app = Application(backend="win32").connect(found_index=0, title_re=mdfWildcard, timeout=10)
        popup = app.window(found_index=0, title_re=mdfWildcard) """

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
    getImages()

    print(list_of_images)

    for i, img in enumerate(list_of_images, start=1):   # Python indexes start at zero
        clickImage(img)

        """ on step 4 write all nabm from txt file """
        if i == 4: insertDataToMDF()

        """ img 5 click each nabm row, fill in the data, press Enter """
        """ click arrow up upload btn """
        """ click green upload button """
        """ select favorites > auto folder """

main()
