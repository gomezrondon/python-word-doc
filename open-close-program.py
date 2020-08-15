import os
import pyautogui
import time
import psutil

from get_list_process import isWoducmentOpen

def openFile():
    try:
        os.startfile('D:\gomez\Documents\python-doc.docx')
        time.sleep(4)
        pyautogui.click(1027, 699)
        time.sleep(1)
        pyautogui.click(1027, 699)
        # time.sleep(2)
        # print(pyautogui.position()) #Point(x=1027, y=699)

    except Exception as e:
        print(str(e))


def closeFile():
    try:
        os.system('TASKKILL /F /IM soffice.bin') #is libreOffice.exe

    except Exception as e:
        print(str(e))

if __name__ == '__main__':
    print(isWoducmentOpen())
    if isWoducmentOpen() == True:
        closeFile()
    else:
        openFile()
