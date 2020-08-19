import os
import pyautogui


def findWindowProgram(title):
    programs = pyautogui.getWindowsWithTitle(title)
    for prog in programs:
        print(prog.title)
        prog.restore()
        prog.activate()
        return prog

# def findWindowProgram(docx):
#     programs = pyautogui.getAllWindows()
#     for prog in programs:
#         if docx in prog.title:
#             print(prog.title)
#             return (prog.left + 200 ) , (prog.top + 20 )
#     return -1, -1


def openFile(filePath):
    try:
        os.startfile(filePath)
    except Exception as e:
        print(str(e))


def isOpen(title):
    program = findWindowProgram(title)
    if program == None:
        return False
    else:
        return True

def closeFile(title):
    program = clickAndSave(title)
    program.close()

def ifOpenThenClose(filePath, title):
    # print(isOpen(title))
    if isOpen(title) == True:
        closeFile(title)
    else:
        openFile(filePath)


def clickAndSave(title):
    program = findWindowProgram(title)
    pyautogui.hotkey('ctrl', 's')
    return program



if __name__ == '__main__':
    title = 'F603.docx'+' - Word'
    closeFile(title)
    # clickAndSave(title)
    # activateWindow('F603.docx'+' - Word')
    # filePath = r'C:\Users\jrgm\Documents\GAA\F603.docx'
    # docName = 'F603.docx'
    # ifOpenThenClose(filePath, docName)
