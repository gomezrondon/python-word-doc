import os
import pyautogui


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
    except Exception as e:
        print(str(e))


def isOpen(docName):
    x, y = findWindowProgram(docName)
    if x > -1 and y > -1:
        return True
    else:
        return False

def closeFile(docName):
    x, y = findWindowProgram(docName)
    pyautogui.click(x, y)
    clickAndSave(x, y)
    pyautogui.hotkey('ctrl', 'F4')


def ifOpenThenClose(filePath, docName):
    print(isOpen(docName))
    if isOpen(docName) == True:
        closeFile(docName)
    else:
        openFile(filePath)


def clickAndSave(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click(x, y)
    pyautogui.hotkey('ctrl', 's')


if __name__ == '__main__':
    filePath = r'C:\Users\jrgm\Documents\GAA\F603.docx'
    docName = 'F603.docx'
    ifOpenThenClose(filePath, docName)
