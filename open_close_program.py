import os
import pyautogui
import time
from get_list_process import isWoducmentOpen

def findWindowProgram(docx):
    programs = pyautogui.getAllWindows()
    for prog in programs:
        if docx in prog.title:
            print(prog.title)
            return (prog.left + 200 ) , (prog.top + 20 )
    return -1, -1


def openFile(filePath):
    try:
        os.startfile(filePath)
        # time.sleep(4)
        # pyautogui.click(1027, 699)
        # time.sleep(1)
        # pyautogui.click(1027, 699)
        # time.sleep(2)
        # print(pyautogui.position()) #Point(x=1027, y=699)

    except Exception as e:
        print(str(e))


def isOpen(proc_name):
    return isWoducmentOpen(proc_name)

def closeFile(proc_name):
    try:
        os.system('TASKKILL /F /IM '+proc_name) #is libreOffice.exe

    except Exception as e:
        print(str(e))

def ifOpenThenClose(filePath, proc_name):
    print(isWoducmentOpen(proc_name))
    if isWoducmentOpen(proc_name) == True:
        closeFile(proc_name)
    else:
        openFile(filePath)


def clickAndSave(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click(x, y)
    pyautogui.hotkey('ctrl', 's')


if __name__ == '__main__':
    ifOpenThenClose('filepath')
